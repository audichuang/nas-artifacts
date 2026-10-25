import assert from 'node:assert/strict'
import { existsSync, mkdtempSync, rmSync } from 'node:fs'
import { homedir, tmpdir } from 'node:os'
import { join } from 'node:path'
import { after, before, describe, test } from 'node:test'
import { chromium } from 'playwright-core'
import { ArtifactStore } from '../src/artifact-store.ts'
import { createContentServer } from '../src/content-server.ts'
import { formatGeneral, formatNumber, formatText, renderXlsxPage, xlsxText, XLSX_LIMITS } from '../src/xlsx.ts'
import { zip } from './helpers/zip.ts'

const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
const DRAWING_NS = 'xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"'
const rels = (list: [string, string, string, string?][]) =>
  `<Relationships>${list.map(([id, type, target, mode]) => `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}" Target="${target}"${mode ? ` TargetMode="${mode}"` : ''}/>`).join('')}</Relationships>`

type SheetSpec = { name: string; body: string; state?: string; rels?: [string, string, string, string?][]; kind?: 'worksheet' | 'chartsheet' }

/** A workbook from hand-written parts, the way no single tool writes them. */
function book(sheets: SheetSpec[], options: { styles?: string; strings?: string[] | string; theme?: string; workbookPr?: string; extra?: Record<string, string>; activeTab?: number } = {}): Buffer {
  const files: Record<string, string> = {
    '_rels/.rels': rels([['r1', 'officeDocument', 'xl/workbook.xml']]),
    'xl/workbook.xml': `<workbook ${NS}>${options.workbookPr ?? ''}<bookViews><workbookView activeTab="${options.activeTab ?? 0}"/></bookViews><sheets>${sheets
      .map((sheet, at) => `<sheet name="${sheet.name}" sheetId="${at + 1}" r:id="s${at + 1}"${sheet.state ? ` state="${sheet.state}"` : ''}/>`)
      .join('')}</sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': rels([
      ...sheets.map((sheet, at): [string, string, string] => [`s${at + 1}`, sheet.kind ?? 'worksheet', `${sheet.kind === 'chartsheet' ? 'chartsheets' : 'worksheets'}/sheet${at + 1}.xml`]),
      ['st', 'styles', 'styles.xml'],
      ['ss', 'sharedStrings', 'sharedStrings.xml'],
      ['th', 'theme', 'theme/theme1.xml']
    ])
  }
  sheets.forEach((sheet, at) => {
    const dir = sheet.kind === 'chartsheet' ? 'chartsheets' : 'worksheets'
    files[`xl/${dir}/sheet${at + 1}.xml`] = sheet.body
    if (sheet.rels) files[`xl/${dir}/_rels/sheet${at + 1}.xml.rels`] = rels(sheet.rels)
  })
  if (options.styles !== undefined) files['xl/styles.xml'] = options.styles
  if (options.strings !== undefined) {
    files['xl/sharedStrings.xml'] = typeof options.strings === 'string' ? options.strings : `<sst ${NS}>${options.strings.map((s) => `<si>${s.startsWith('<') ? s : `<t>${s}</t>`}</si>`).join('')}</sst>`
  }
  if (options.theme !== undefined) files['xl/theme/theme1.xml'] = options.theme
  return zip({ ...files, ...options.extra })
}

const sheet = (inner: string, before = '', afterData = '') => `<worksheet ${NS}>${before}<sheetData>${inner}</sheetData>${afterData}</worksheet>`

const THEME = `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme name="c"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2><a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4><a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme><a:fontScheme name="f"><a:majorFont><a:latin typeface="Cambria"/><a:ea typeface="新細明體"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface="新細明體"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>`

/**
 * Styles the way Excel writes them: a Normal cell style whose font cells inherit when their xf names
 * none, theme colours with tints, an indexed colour, a pattern, borders and alignments.
 */
const STYLES = `<styleSheet ${NS}>
<numFmts count="2"><numFmt numFmtId="164" formatCode="&quot;NT$&quot;#,##0.00;[Red]\\-&quot;NT$&quot;#,##0.00"/><numFmt numFmtId="165" formatCode="yyyy&quot;年&quot;m&quot;月&quot;d&quot;日&quot;"/></numFmts>
<fonts count="5">
<font><sz val="12"/><color theme="1"/><name val="新細明體"/><family val="2"/><scheme val="minor"/></font>
<font><b/><i/><u/><sz val="14"/><color rgb="FFC00000"/><name val="Georgia"/></font>
<font><strike/><sz val="12"/><color theme="4" tint="-0.249977111117893"/><name val="Evil';}&lt;/style&gt;&lt;script&gt;x"/></font>
<font><sz val="12"/><color indexed="10"/><name val="Arial"/></font>
<font><u/><sz val="12"/><color theme="10"/><name val="新細明體"/></font>
</fonts>
<fills count="5">
<fill><patternFill patternType="none"/></fill>
<fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor theme="5" tint="0.79998168889431442"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor indexed="13"/></patternFill></fill>
<fill><patternFill patternType="lightGray"><fgColor rgb="FF000000"/><bgColor rgb="FFFFFFFF"/></patternFill></fill>
</fills>
<borders count="3">
<border><left/><right/><top/><bottom/><diagonal/></border>
<border><left style="thin"><color indexed="64"/></left><right style="medium"><color rgb="FF00B050"/></right><top style="dashed"><color theme="4"/></top><bottom style="double"><color auto="1"/></bottom></border>
<border><left/><right/><top/><bottom style="thick"><color rgb="FF7030A0"/></bottom></border>
<border><left/><right/><top style="medium"><color rgb="FFFF0000"/></top><bottom/></border>
</borders>
<cellStyleXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="4" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="11">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="165" fontId="2" fillId="3" borderId="0" xfId="0"/>
<xf numFmtId="10" fontId="3" fillId="4" borderId="2" xfId="0"><alignment horizontal="right" indent="2"/></xf>
<xf fillId="0" borderId="0" xfId="1"/>
<xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0"/>
<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0"><alignment textRotation="90"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"><alignment textRotation="255" shrinkToFit="1"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="3" xfId="0"/>
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"><alignment shrinkToFit="1"/></xf>
</cellXfs>
<cellStyles count="2"><cellStyle name="Normal" xfId="0" builtinId="0"/><cellStyle name="Hyperlink" xfId="1" builtinId="8"/></cellStyles>
</styleSheet>`

const PNG = 'not really a png, the route checks only the part name'

/** One workbook that uses most of the format: what most tests below read from. */
function richBook(): Buffer {
  const first = sheet(
    `<row r="1" ht="30" customHeight="1"><c r="A1" s="1" t="s"><v>0</v></c><c r="B1" s="1"/><c r="D1" t="s"><v>1</v></c></row>` +
      `<row r="2"><c r="A2" s="2"><v>-1234.5</v></c><c r="B2" s="3"><v>45000</v></c><c r="C2" s="4"><v>0.256</v></c><c r="D2" t="b"><v>1</v></c><c r="E2" t="e"><v>#DIV/0!</v></c><c r="G2" t="inlineStr"><is><t>隱藏欄的字</t></is></c></row>` +
      `<row r="3"><c r="A3" s="11" t="inlineStr"><is><t>行內 &amp; &lt;i&gt;</t></is></c><c r="B3" t="str"><f>A3&amp;"!"</f><v>算好的</v></c><c r="C3"><f>SUM(A1:A2)</f></c><c r="D3" s="5" t="s"><v>2</v></c><c r="E3" t="s"><v>3</v></c></row>` +
      `<row r="4" s="6" customFormat="1"><c r="A4" t="s"><v>4</v></c></row>` +
      `<row r="5" hidden="1"><c r="A5" t="s"><v>5</v></c></row>` +
      `<row r="6"><c r="A6" s="8"><v>45658</v></c><c r="B6" s="9"><v>12</v></c><c r="C6" s="10" t="s"><v>6</v></c><c r="D6" t="d"><v>2024-02-29T12:00:00</v></c><c r="E6" s="12" t="inlineStr"><is><t>縮小以適合欄寬</t></is></c></row>` +
      `<row r="7"><c r="A7" t="s"><v>7</v></c><c r="H7" t="s"><v>8</v></c></row>`,
    `<sheetPr><tabColor rgb="FF00B050"/></sheetPr><dimension ref="A1:H7"/><sheetViews><sheetView workbookViewId="0"><pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="16.5"/><cols><col min="1" max="1" width="20" customWidth="1"/><col min="7" max="7" width="0" hidden="1"/><col min="6" max="6" width="9" style="7"/></cols>`,
    `<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells><hyperlinks><hyperlink ref="D3" r:id="h1"/><hyperlink ref="E3" location="'第 二 頁'!A1"/></hyperlinks><drawing r:id="d1"/>`
  )
  const second = sheet(`<row r="1"><c r="A1" t="s"><v>9</v></c><c r="B1"><v>42</v></c></row>`, `<sheetViews><sheetView rightToLeft="1" showGridLines="0" workbookViewId="0"/></sheetViews>`)
  const strings = [
    '標題 &lt;b&gt;粗&lt;/b&gt;',
    '<r><t>一般</t></r><r><rPr><b/><color rgb="FFFF0000"/><sz val="14"/><rFont val="Arial"/></rPr><t>紅粗</t></r><rPh sb="0" eb="1"><t>ふりがな</t></rPh>',
    '連結文字',
    '去第二頁',
    '整列有底色',
    '隱藏列的字',
    '直排',
    '溢出到右邊的很長很長的文字',
    '遠處',
    'שלום'
  ]
  const drawing = `<xdr:wsDr ${DRAWING_NS}>
<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>4</xdr:col><xdr:colOff>9525</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>6</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="Picture 1" descr="公司標誌"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="p1"/><a:stretch/></xdr:blipFill><xdr:spPr><a:prstGeom prst="rect"/></xdr:spPr></xdr:pic><xdr:clientData/></xdr:twoCellAnchor>
<xdr:oneCellAnchor><xdr:from><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="3048000" cy="1905000"/>
<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="3" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="c1"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:oneCellAnchor>
<xdr:absoluteAnchor><xdr:pos x="0" y="3810000"/><xdr:ext cx="1905000" cy="952500"/>
<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="4" name="TextBox 1"/><xdr:cNvSpPr txBox="1"/></xdr:nvSpPr><xdr:spPr><a:prstGeom prst="rect"/><a:solidFill><a:schemeClr val="accent2"/></a:solidFill></xdr:spPr><xdr:txBody><a:bodyPr anchor="ctr"/><a:p><a:pPr algn="ctr"/><a:r><a:rPr sz="1400" b="1"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:rPr><a:t>文字方塊 &lt;x&gt;</a:t></a:r></a:p></xdr:txBody></xdr:sp><xdr:clientData/></xdr:absoluteAnchor>
<xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>8</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="5" name="Diagram"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"/></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>
</xdr:wsDr>`
  const chart = `<c:chartSpace ${DRAWING_NS}><c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>月營收</a:t></a:r></a:p></c:rich></c:tx></c:title><c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>
<c:ser><c:idx val="0"/><c:tx><c:strRef><c:f>Sheet!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>營收</c:v></c:pt></c:strCache></c:strRef></c:tx><c:spPr><a:solidFill><a:schemeClr val="accent6"/></a:solidFill></c:spPr>
<c:cat><c:strRef><c:f>x</c:f><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>一月</c:v></c:pt><c:pt idx="1"><c:v>二月</c:v></c:pt></c:strCache></c:strRef></c:cat>
<c:val><c:numRef><c:f>x</c:f><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>3</c:v></c:pt><c:pt idx="1"><c:v>7</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>`
  return book(
    [
      { name: 'Sheet', body: first, rels: [['h1', 'hyperlink', 'javascript:alert(&quot;x&quot;)', 'External'], ['d1', 'drawing', '../drawings/drawing1.xml'], ['cm', 'comments', '../comments1.xml']] },
      { name: '藏起來', body: sheet('<row r="1"><c r="A1" t="inlineStr"><is><t>不該被搜尋到</t></is></c></row>'), state: 'hidden' },
      { name: '第 二 頁', body: second },
      { name: '很隱藏', body: sheet(''), state: 'veryHidden' }
    ],
    {
      styles: STYLES,
      strings,
      theme: THEME,
      extra: {
        'xl/drawings/drawing1.xml': drawing,
        'xl/drawings/_rels/drawing1.xml.rels': rels([['p1', 'image', '../media/image1.png'], ['c1', 'chart', '../charts/chart1.xml']]),
        'xl/media/image1.png': PNG,
        'xl/charts/chart1.xml': chart,
        'xl/comments1.xml': `<comments ${NS}><authors><author>小華</author></authors><commentList><comment ref="A2" authorId="0"><text><r><t>註解 &lt;script&gt;</t></r></text></comment></commentList></comments>`
      }
    }
  )
}

/** The part of a page that draws one sheet panel. */
const panel = (html: string, n: number) => {
  const start = html.indexOf(`<section class="xl-sheet" id="xs-${n}"`)
  assert.ok(start >= 0, `panel ${n}`)
  const end = html.indexOf('</section>', start)
  return html.slice(start, end)
}

/** The opening tag and content of the cell that shows `text`. */
const cellOf = (html: string, text: string) => {
  const at = html.indexOf(`>${text}<`)
  assert.ok(at >= 0, `cell ${text}`)
  const open = html.lastIndexOf('<td', at)
  return html.slice(open, html.indexOf('</td>', at) + 5)
}

describe('number formats', () => {
  test('sections, signs, colours, scaling, percent, scientific, fractions and literal text', () => {
    const cases: [string, number, string, string | null][] = [
      ['General', 1234.5, '1234.5', null],
      ['General', 0.1 + 0.2, '0.3', null],
      ['General', 123456789012, '1.23457E+11', null],
      ['0.00', -1.005, '-1.01', null],
      ['#,##0', 1234567, '1,234,567', null],
      ['#,##0.00;[Red](#,##0.00)', -1234.567, '(1,234.57)', 'FF0000'],
      ['0.0%', 0.4567, '45.7%', null],
      ['0.00E+00', 12345, '1.23E+04', null],
      ['# ?/?', 1.25, '1 1/4', null],
      ['0.0,,"M"', 12345678, '12.3M', null],
      ['000-0000', 1234567, '123-4567', null],
      ['"NT$"#,##0;"負"#,##0;"零"', 0, '零', null],
      ['"NT$"#,##0;"負"#,##0;"零"', -8, '負8', null],
      ['[Blue][>=100]0;[Red][<0]0;0.0', 150, '150', '0000FF'],
      ['[$€-407] #,##0.00', 99.5, '€ 99.50', null],
      ['_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)', -66000, ' (66,000)', null],
      ['0;-0;;@', 0, '', null]
    ]
    for (const [code, value, text, color] of cases) assert.deepEqual(formatNumber(code, value), { text, color }, `${code} ${value}`)
    assert.equal(formatGeneral(-1.234e-9), '-1.234E-09')
    assert.equal(formatText('"名稱:"@', '甲').text, '名稱:甲')
    assert.equal(formatText('0.00', 'abc').text, 'abc')
  })

  test('dates and times from serials: 1900 (with its leap-year quirk) and 1904 systems, elapsed time, zh-TW built-ins', () => {
    assert.equal(formatNumber('yyyy-mm-dd', 45000).text, '2023-03-15')
    assert.equal(formatNumber('yyyy-mm-dd', 60).text, '1900-02-29')
    assert.equal(formatNumber('yyyy-mm-dd', 61).text, '1900-03-01')
    assert.equal(formatNumber('yyyy-mm-dd', 0, true).text, '1904-01-01')
    assert.equal(formatNumber('d-mmm-yy dddd', 45000).text, '15-Mar-23 Wednesday')
    assert.equal(formatNumber('h:mm AM/PM', 0.75).text, '6:00 PM')
    assert.equal(formatNumber('h:mm:ss', 0.5001).text, '12:00:09')
    assert.equal(formatNumber('[h]:mm:ss', 1.5).text, '36:00:00')
    assert.equal(formatNumber('mm:ss.0', 0.000123).text, '00:10.6')
    assert.equal(formatNumber('[$-404]e/m/d', 45000).text, '112/3/15')
    assert.equal(formatNumber('上午/下午hh"時"mm"分"', 0.3).text, '上午07時12分')
    assert.equal(formatNumber('[$-zh-TW]aaaa', 45000).text, '星期三')
    assert.equal(formatNumber('yyyy/m/d', -1).text, '########')
  })
})

describe('reading Excel workbooks', () => {
  const html = renderXlsxPage(richBook(), '帳本 (1).xlsx', '帳本')
  const first = panel(html, 1)

  test('visible sheets become tabs in order; hidden and veryHidden ones are skipped; the hash keeps the tab', () => {
    const tabs = [...html.matchAll(/<a class="xl-tab"[^>]*href="(#sheet-\d+)"[^>]*>([^<]*)<\/a>/g)].map((m) => [m[1], m[2]])
    assert.deepEqual(tabs, [['#sheet-1', 'Sheet'], ['#sheet-2', '第 二 頁']])
    assert.doesNotMatch(html, /不該被搜尋到|很隱藏|藏起來/)
    assert.match(html, /style="--tab:#00B050"/)
    // The viewer keeps only hashes of letters, digits, '-' and '_', so the tab hash must be one.
    assert.match(html, /\/\^#sheet-\(\\d\+\)\$\//)
    assert.doesNotMatch(html, /id="sheet-\d/)
    assert.match(html, /<meta name="viewport"/)
    assert.match(html, /^<!doctype html><html lang="zh-Hant"><head><script src="\/_\/bridge\.js"><\/script>/)
  })

  test('styles inherit: the Normal font, cell style xfs, row styles over empty cells, column styles', () => {
    // The grid takes the Normal style's font and size.
    assert.match(html, /\.xl-grid\{font-family:'新細明體',[^;]*;font-size:12pt;color:#000000\}/)
    // xf 5 names no font: it inherits the Hyperlink cell style's (underlined, theme hlink colour).
    assert.match(html, /\.xl-grid td\.x5\{text-decoration:underline;color:#0000FF\}/)
    // A row with customFormat paints its empty cells too; a styled column paints its cells.
    const row4 = first.slice(first.indexOf('>4</th>'), first.indexOf('</tr>', first.indexOf('>4</th>')))
    assert.ok((row4.match(/class="x6/g) ?? []).length >= 5, row4)
    assert.match(first, /<td class="x7"/)
  })

  test("fonts and fills: bold/italic/underline/strike, ARGB, theme + tint, indexed colours, patterns, sanitized names", () => {
    assert.match(html, /\.xl-grid td\.x1\{font-weight:700;font-style:italic;text-decoration:underline;color:#C00000;font-size:14pt;font-family:'Georgia',[^}]*background:#F2DCDB;border-inline-end-color:transparent;border-bottom-color:transparent;text-align:center;vertical-align:middle;white-space:pre-wrap/)
    // theme 4 (accent1 4F81BD) darkened by 25%; the font name loses its quotes and markup.
    assert.match(html, /\.xl-grid td\.x3\{text-decoration:line-through;color:#376092;font-family:'Evilstylescriptx',/)
    assert.match(html, /\.xl-grid td\.x3\{[^}]*background:#FFFF00/)
    assert.match(html, /\.xl-grid td\.x4\{color:#FF0000;font-family:'Arial',[^}]*background:#BFBFBF;border-inline-end-color:transparent;border-bottom-color:transparent;text-align:right;padding-inline-end:20px\}/)
    assert.doesNotMatch(html, /Evil'|<\/style><script>x/)
  })

  test('borders per side, settled against the neighbour; gridlines hide under fills', () => {
    // A1:B1 is merged: its end edge is B1's right border; automatic and indexed-64 colours are black.
    const a1 = cellOf(first, '標題 &#60;b&#62;粗&#60;/b&#62;')
    assert.match(a1, /border-inline-end:2px solid #00B050/)
    assert.match(a1, /border-bottom:3px double #000000/)
    assert.match(a1, /border-inline-start:1px solid #000000/)
    assert.match(a1, /border-top:1px dashed #4F81BD/)
    // A2 has no border of its own; A3's top border is drawn as A2's bottom edge.
    assert.match(cellOf(first, '-NT$1,234.50'), /border-bottom:2px solid #FF0000/)
    assert.match(cellOf(first, '25.60%'), /border-bottom:3px solid #7030A0/)
    // Filled cells drop the gridlines inside them.
    assert.match(html, /\.xl-grid td\.x7\{background:#FFFF00;border-inline-end-color:transparent;border-bottom-color:transparent\}/)
    assert.match(first, /<td class="x7"><\/td>/)
  })

  test('values: shared strings with rich runs, inline strings, booleans, errors, cached formula results, formulas without a value', () => {
    assert.match(first, />一般<span style="font-weight:700;color:#FF0000;font-size:14pt;font-family:'Arial',[^"]*">紅粗<\/span><\/span><\/td>/)
    assert.doesNotMatch(first, /ふりがな/)
    // B3 is not empty, so A3's text does not run on into it.
    assert.doesNotMatch(cellOf(first, '行內 &#38; &#60;i&#62;'), /\bov\b/)
    assert.match(cellOf(first, 'TRUE'), /class="c"/)
    assert.match(cellOf(first, '#DIV/0!'), /class="c"/)
    assert.match(first, />算好的</)
    assert.doesNotMatch(first, /=A3&#38;/)
    assert.match(cellOf(first, '=SUM(A1:A2)'), /class="f"/)
  })

  test('number formats from styles: custom codes with colours, built-in ids, dates (ISO and serial), text and alignment', () => {
    assert.match(cellOf(first, '-NT$1,234.50'), /class="x2 n[^"]*"[^>]*color:#FF0000/)
    assert.match(first, />2023年3月15日</)
    assert.match(first, />2025\/1\/1</)
    assert.match(first, />2024\/2\/29 12:00</)
    // Numbers align right by default, text left: an explicitly aligned cell keeps its own.
    assert.match(cellOf(first, '2025/1/1'), /class="x8 n /)
    assert.match(first, /<td class="x9 n rc"><span class="rot" style="transform:translateX\(1\.2em\) rotate\(-90deg\)">12<\/span>/)
    assert.match(first, /class="x10 rc"[^>]*><span class="vt">直排<\/span>/)
    // Shrink to fit: the page script scales the span down to the column once the sheet is shown.
    assert.match(first, /<td class="x12 sh"><span>縮小以適合欄寬<\/span><\/td>/)
    assert.match(html, /querySelectorAll\('td\.sh'\)/)
  })

  test('merged cells, frozen panes, hidden rows and columns, widths and heights', () => {
    assert.match(cellOf(first, '標題 &#60;b&#62;粗&#60;/b&#62;'), /colspan="2"[^>]*class="x1 fr fc"/)
    assert.match(first, /<th class="fc" style="inset-inline-start:32px">A<\/th>/)
    assert.doesNotMatch(first, /<th[^>]*>G<\/th>/)
    assert.doesNotMatch(first, /隱藏列的字|隱藏欄的字|>5<\/th>/)
    // Column A is 20 characters of a 12pt font (8px digits), the rest the default; row 1 is 30pt.
    assert.match(first, /<colgroup><col style="width:32px"><col style="width:160px"><col style="width:72px">/)
    assert.match(first, /<tr style="height:40px"><th class="fr" style="top:20px">1<\/th>/)
    assert.match(first, /<tr style="height:22px"><th>2<\/th>/)
  })

  test('text runs into empty neighbours; links show as text with their target as a title; comments mark the cell', () => {
    assert.match(cellOf(first, '溢出到右邊的很長很長的文字'), /class="ov /)
    assert.match(first, /<span class="xl-link" title="javascript:alert\(&#34;x&#34;\)">連結文字<\/span>/)
    assert.doesNotMatch(html, /href="javascript/)
    assert.match(first, /<a class="xl-link" href="#sheet-2" title="&#39;第 二 頁&#39;!A1">去第二頁<\/a>/)
    assert.match(first, /class="x2 n fc cm"[^>]*title="小華:\n註解 &#60;script&#62;"/)
  })

  test('pictures, charts and text boxes are placed over the grid; what is not drawn is a labelled box', () => {
    assert.match(first, /<img class="xl-draw xl-pic" src="~office\/%E5%B8%B3%E6%9C%AC%20%281%29\.xlsx\/xl\/media\/image1\.png" alt="公司標誌"[^>]*style="inset-inline-start:\d+(\.\d)?px;top:\d+(\.\d)?px;width:/)
    assert.match(first, /<div class="xl-draw xl-chart"[^>]*><svg[^>]*aria-label="月營收"[\s\S]*>一月<\/text>[\s\S]*<rect[^>]*fill="#F79646"/)
    assert.match(first, /class="xl-draw xl-shape"[^>]*background:#C0504D[\s\S]*>文字方塊 &#60;x&#62;<\/span>/)
    assert.match(first, /class="xl-draw xl-missing"[^>]*>物件\(網頁版未支援,請下載原始檔查看\)</)
    // The grid continues under the chart in column I (0-based 8).
    assert.match(first, /<th>[IJK]<\/th>/)
  })

  test('right-to-left sheets and sheets without gridlines', () => {
    const second = panel(html, 2)
    assert.match(second, /<div class="xl-scroll" dir="rtl">/)
    assert.match(second, /<table class="xl-grid nogrid"/)
    assert.match(second, />שלום</)
  })

  test('an empty sheet and a sheet beyond the limits say so', () => {
    const rows = Array.from({ length: XLSX_LIMITS.rows + 100 }, (_, at) => `<row r="${at + 1}"><c r="A${at + 1}"><v>${at}</v></c>${at === 0 ? '<c r="ZZ1"><v>1</v></c>' : ''}</row>`).join('')
    const page = renderXlsxPage(book([{ name: 'empty', body: sheet('') }, { name: 'big', body: sheet(rows + `<row r="${XLSX_LIMITS.rows + 500}"><c r="A${XLSX_LIMITS.rows + 500}"><v>1</v></c></row>`) }]), 'b.xlsx', 'b')
    assert.match(panel(page, 1), /這個工作表是空的/)
    const big = panel(page, 2)
    assert.match(big, /此工作表共 2,500 列 × 702 欄,網頁版只顯示前 2,000 列、前 10 欄\(到 J 欄\)/)
    assert.equal((big.match(/<tr style=[^>]*><th>\d+<\/th>/g) ?? []).length, XLSX_LIMITS.rows)
    assert.match(big, /<th>2000<\/th>/)
    assert.doesNotMatch(big, /<th>2001<\/th>/)
  })

  test('a broken part stays alone: a sheet, the drawing, the styles each fail by themselves', () => {
    const page = renderXlsxPage(
      book(
        [
          { name: 'ok', body: sheet('<row r="1"><c r="A1" t="inlineStr"><is><t>還在</t></is></c></row>', '', '<drawing r:id="d"/>'), rels: [['d', 'drawing', '../drawings/drawing1.xml']] },
          { name: 'broken', body: `<worksheet ${NS}><sheetData><row r="1"><c r="A1"><v>1</v></row></sheetData></worksheet>` },
          { name: 'missing', body: '' }
        ],
        { styles: '<styleSheet><fonts>', extra: { 'xl/drawings/drawing1.xml': '<xdr:wsDr><oops' } }
      ),
      'b.xlsx',
      'b'
    )
    assert.match(page, /樣式無法讀取,以預設樣式顯示/)
    assert.match(panel(page, 1), />還在</)
    assert.match(panel(page, 1), /圖片與圖表無法顯示/)
    assert.match(panel(page, 2), /class="xl-note xl-failed">工作表「broken」無法顯示/)
    assert.match(panel(page, 2), /<script>throw new Error\("Excel: 工作表「broken」無法顯示/)
    assert.match(panel(page, 3), /工作表「missing」無法顯示/)
  })

  test('a broken picture or chart becomes a labelled box; the rest of the drawing still draws', () => {
    const drawing = `<xdr:wsDr ${DRAWING_NS}>
<xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="952500" cy="952500"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="1" name="p"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="emf"/></xdr:blipFill><xdr:spPr/></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>
<xdr:oneCellAnchor><xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="952500" cy="952500"/><xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="c"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="bad"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:oneCellAnchor>
<xdr:oneCellAnchor><xdr:from><xdr:col>6</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="952500" cy="952500"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="3" name="ok"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="png"/></xdr:blipFill><xdr:spPr/></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>
</xdr:wsDr>`
    const page = renderXlsxPage(
      book([{ name: 's', body: sheet('', '', '<drawing r:id="d"/>'), rels: [['d', 'drawing', '../drawings/drawing1.xml']] }], {
        extra: {
          'xl/drawings/drawing1.xml': drawing,
          'xl/drawings/_rels/drawing1.xml.rels': rels([['emf', 'image', '../media/image1.emf'], ['bad', 'chart', '../charts/chart1.xml'], ['png', 'image', '../media/ok.png']]),
          'xl/media/image1.emf': 'x',
          'xl/media/ok.png': PNG,
          'xl/charts/chart1.xml': '<c:chartSpace><c:chart>'
        }
      }),
      'b.xlsx',
      'b'
    )
    assert.match(page, /xl-missing"[^>]*>圖片\(EMF 格式,網頁版未支援,請下載原始檔查看\)</)
    assert.match(page, /xl-missing"[^>]*>圖形無法顯示\(/)
    assert.match(page, /<img class="xl-draw xl-pic" src="~office\/b\.xlsx\/xl\/media\/ok\.png"/)
  })

  test('a chart without cached values (as openpyxl writes them) plots the cells its formulas name', () => {
    const drawing = `<xdr:wsDr ${DRAWING_NS}><xdr:twoCellAnchor><xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>12</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="c"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="c"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>`
    const chart = `<c:chartSpace ${DRAWING_NS}><c:chart><c:plotArea><c:pieChart><c:ser><c:tx><c:strRef><c:f>'資 料'!B1</c:f></c:strRef></c:tx><c:cat><c:numRef><c:f>'資 料'!$A$2:$A$3</c:f></c:numRef></c:cat><c:val><c:numRef><c:f>'資 料'!$B$2:$B$3</c:f></c:numRef></c:val></c:ser></c:pieChart></c:plotArea><c:legend><c:legendPos val="b"/></c:legend></c:chart></c:chartSpace>`
    const page = renderXlsxPage(
      book(
        [{ name: '資 料', body: sheet('<row r="1"><c r="B1" t="s"><v>0</v></c></row><row r="2"><c r="A2" t="s"><v>1</v></c><c r="B2"><v>1</v></c></row><row r="3"><c r="A3" t="s"><v>2</v></c><c r="B3"><v>3</v></c></row>', '', '<drawing r:id="d"/>'), rels: [['d', 'drawing', '../drawings/drawing1.xml']] }],
        { strings: ['比例', '甲', '乙'], extra: { 'xl/drawings/drawing1.xml': drawing, 'xl/drawings/_rels/drawing1.xml.rels': rels([['c', 'chart', '../charts/chart1.xml']]), 'xl/charts/chart1.xml': chart } }
      ),
      'b.xlsx',
      'b'
    )
    assert.equal((page.match(/<path d="M[^"]+ A/g) ?? []).length, 2)
    assert.match(page, />甲<\/text>[\s\S]*>乙<\/text>/)
  })

  test('tables take their style: the file’s own table style, else the built-in look from the theme', () => {
    const tableXml = (name: string) => `<table ${NS} id="1" name="T" displayName="T" ref="A1:B3"><autoFilter ref="A1:B3"/><tableColumns count="2"><tableColumn id="1" name="名"/><tableColumn id="2" name="值"/></tableColumns><tableStyleInfo name="${name}" showRowStripes="1"/></table>`
    const body = sheet('<row r="1"><c r="A1" t="inlineStr"><is><t>名</t></is></c><c r="B1" t="inlineStr"><is><t>值</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>甲</t></is></c><c r="B2"><v>1</v></c></row><row r="3"><c r="A3" t="inlineStr"><is><t>乙</t></is></c><c r="B3"><v>2</v></c></row>', '', '<tableParts count="1"><tablePart r:id="t"/></tableParts>')
    const styles = `<styleSheet ${NS}><fonts count="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellXfs count="1"><xf/></cellXfs><dxfs count="2"><dxf><font><b/><color rgb="FFFFFF00"/></font><fill><patternFill patternType="solid"><bgColor rgb="FF203864"/></patternFill></fill></dxf><dxf><fill><patternFill patternType="solid"><bgColor rgb="FFDDEBF7"/></patternFill></fill></dxf></dxfs><tableStyles count="1"><tableStyle name="我的表格" count="2"><tableStyleElement type="headerRow" dxfId="0"/><tableStyleElement type="firstRowStripe" dxfId="1"/></tableStyle></tableStyles></styleSheet>`
    const page = (name: string) =>
      renderXlsxPage(book([{ name: 's', body, rels: [['t', 'table', '../tables/table1.xml']] }], { styles, theme: THEME, extra: { 'xl/tables/table1.xml': tableXml(name) } }), 'b.xlsx', 'b')
    const custom = page('我的表格')
    assert.match(cellOf(custom, '名'), /background:#203864;color:#FFFF00;font-weight:700/)
    assert.match(cellOf(custom, '甲'), /background:#DDEBF7/)
    const builtIn = page('TableStyleMedium2')
    assert.match(cellOf(builtIn, '名'), /background:#4F81BD;color:#FFFFFF;font-weight:700/)
    assert.match(cellOf(builtIn, '甲'), /background:#DCE6F2/)
  })

  test('conditional formats decided by the value: comparisons with their dxf, colour scales, data bars; formula rules are left alone', () => {
    const rows = [5, 50, 95].map((v, at) => `<row r="${at + 1}"><c r="A${at + 1}"><v>${v}</v></c><c r="B${at + 1}"><v>${v}</v></c><c r="C${at + 1}"><v>${v}</v></c></row>`).join('')
    const rules =
      '<conditionalFormatting sqref="A1:A3"><cfRule type="cellIs" dxfId="0" priority="2" operator="greaterThan"><formula>$D$1</formula></cfRule><cfRule type="expression" dxfId="0" priority="1"><formula>MOD(ROW(),2)=0</formula></cfRule></conditionalFormatting>' +
      '<conditionalFormatting sqref="B1:B3"><cfRule type="colorScale" priority="3"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFFFFFFF"/><color rgb="FF000000"/></colorScale></cfRule></conditionalFormatting>' +
      '<conditionalFormatting sqref="C1:C3"><cfRule type="dataBar" priority="4"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar></cfRule></conditionalFormatting>'
    const styles = `<styleSheet ${NS}><dxfs count="1"><dxf><font><b/><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf></dxfs></styleSheet>`
    const page = renderXlsxPage(book([{ name: 's', body: sheet(rows.replace('</row>', '<c r="D1"><v>40</v></c></row>'), '', rules) }], { styles }), 'b.xlsx', 'b')
    const cell = (ref: string) => {
      const [col, row] = [ref.charCodeAt(0) - 64, Number(ref.slice(1))]
      const tr = page.slice(page.indexOf(`<th>${row}</th>`), page.indexOf('</tr>', page.indexOf(`<th>${row}</th>`)))
      return [...tr.matchAll(/<td[^>]*>[^<]*<\/td>/g)][col - 1][0]
    }
    assert.doesNotMatch(cell('A1'), /FFC7CE/)
    assert.match(cell('A2'), /background:#FFC7CE;color:#9C0006;font-weight:700/)
    assert.match(cell('B1'), /background:#FFFFFF/)
    assert.match(cell('B2'), /background:#808080/)
    assert.match(cell('B3'), /background:#000000/)
    assert.match(cell('C3'), /background-image:linear-gradient\(90deg,#638EC6 0,#[0-9A-F]{6} 100%,transparent 100%\)/)
    assert.match(cell('C2'), /transparent 52\.6%/)
  })

  test('a chart sheet draws its chart; date1904 workbooks count from 1904', () => {
    const chart = `<c:chartSpace ${DRAWING_NS}><c:chart><c:plotArea><c:lineChart><c:ser><c:val><c:numRef><c:numCache><c:ptCount val="3"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt><c:pt idx="2"><c:v>2</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser></c:lineChart></c:plotArea></c:chart></c:chartSpace>`
    const page = renderXlsxPage(
      book(
        [
          { name: '日期', body: sheet('<row r="1"><c r="A1" s="1"><v>0</v></c></row>') },
          { name: '圖', kind: 'chartsheet', body: `<chartsheet ${NS}><drawing r:id="d"/></chartsheet>`, rels: [['d', 'drawing', '../drawings/drawing1.xml']] }
        ],
        {
          workbookPr: '<workbookPr date1904="1"/>',
          styles: `<styleSheet ${NS}><cellXfs count="2"><xf/><xf numFmtId="14"/></cellXfs></styleSheet>`,
          extra: {
            'xl/drawings/drawing1.xml': `<xdr:wsDr ${DRAWING_NS}><xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="8000000" cy="5000000"/><xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="1" name="c"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="c"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:absoluteAnchor></xdr:wsDr>`,
            'xl/drawings/_rels/drawing1.xml.rels': rels([['c', 'chart', '../charts/chart1.xml']]),
            'xl/charts/chart1.xml': chart
          }
        }
      ),
      'b.xlsx',
      'b'
    )
    assert.match(panel(page, 1), />1904\/1\/1</)
    assert.match(panel(page, 2), /<div class="xl-chartsheet"[^>]*><svg[\s\S]*<polyline/)
  })

  test('a file that is not a workbook gets a page that says why, and a throw for the error log', () => {
    const notZip = renderXlsxPage(Buffer.from('PK definitely not'), 'x.xlsx', 'x')
    assert.match(notZip, /class="xl-note xl-failed">Excel: 這不是 \.xlsx 檔,或檔案已損壞/)
    assert.match(notZip, /<script>throw new Error\("Excel: 這不是 \.xlsx 檔/)
    const legacy = Buffer.alloc(64)
    legacy.writeUInt32BE(0xd0cf11e0, 0)
    assert.match(renderXlsxPage(legacy, 'x.xlsx', 'x'), /舊版 \.xls 或加密的活頁簿/)
    assert.match(renderXlsxPage(zip({ 'word/document.xml': '<w:document/>' }), 'x.xlsx', 'x'), /這不是 Excel 活頁簿/)
    const dtd = book([{ name: 's', body: '<!DOCTYPE x [<!ENTITY a "aaaa">]><worksheet><sheetData/></worksheet>' }])
    assert.match(renderXlsxPage(dtd, 'x.xlsx', 'x'), /工作表「s」無法顯示:DTDs are not accepted/)
  })

  test('hostile text from the file is escaped everywhere it lands', () => {
    const evil = '<img src=x onerror=alert(1)>'
    const page = renderXlsxPage(
      book([{ name: '&lt;b onclick=x&gt;', body: sheet(`<row r="1"><c r="A1" t="inlineStr"><is><t>${evil.replace(/</g, '&lt;')}</t></is></c></row>`, `<sheetPr><tabColor rgb="FF0000;background:url(x)"/></sheetPr>`) }]),
      '"><script>.xlsx',
      '</title><script>alert(1)</script>'
    )
    assert.doesNotMatch(page, /<img src=x|<b onclick|<\/title><script>alert|url\(x\)/)
    assert.match(page, />&#60;img src=x onerror=alert\(1\)&#62;</)
    assert.match(page, /<title>&#60;\/title&#62;/)
  })
})

describe('serving Excel workbooks', () => {
  let dir: string
  let store: ArtifactStore
  let content: ReturnType<typeof createContentServer>
  before(() => {
    dir = mkdtempSync(join(tmpdir(), 'xlsx-test-'))
    store = new ArtifactStore(dir)
    content = createContentServer(store, { dataDir: dir, appPort: 0, contentPort: 0, appOrigin: 'http://app.test', contentOrigin: 'http://content.test', tokens: ['t'] })
  })
  after(() => {
    store.close()
    rmSync(dir, { recursive: true, force: true })
  })

  test('a published .xlsx renders as a page, its pictures load from inside the file, and its text is searchable', async () => {
    const { id, entry } = store.publish({ title: '帳本', files: { '帳本 (1).xlsx': { content: richBook() } } })
    assert.equal(entry, '帳本 (1).xlsx')
    const response = await content.request(`http://content.test/a/${id}/v/1/`)
    assert.equal(response.status, 200)
    assert.equal(response.headers.get('cache-control'), 'no-cache')
    const html = await response.text()
    assert.match(html, /<nav class="xl-tabs"/)
    const src = /<img class="xl-draw xl-pic" src="([^"]+)"/.exec(html)![1]
    const picture = await content.request(`http://content.test/a/${id}/v/1/${src}`)
    assert.equal(picture.status, 200)
    assert.equal(picture.headers.get('content-type'), 'image/png')
    assert.match(picture.headers.get('content-security-policy') ?? '', /sandbox allow-scripts/)
    assert.equal((await content.request(`http://content.test/a/${id}/v/1/~office/${encodeURIComponent('帳本 (1).xlsx')}/xl/workbook.xml`)).status, 404)
    for (const query of ['溢出到右邊', '第 二 頁', '月營收', '2023年3月15日', '文字方塊']) assert.deepEqual(store.list({ query }).artifacts.map((a) => a.id), [id], query)
    assert.deepEqual(store.list({ query: '不該被搜尋到' }).artifacts, [])
  })

  test('search text is sheet names and shown cell text, bounded', () => {
    const text = xlsxText(richBook())
    assert.match(text, /^Sheet\n/)
    assert.match(text, /-NT\$1,234\.50/)
    assert.doesNotMatch(text, /=SUM|不該被搜尋到|<td|<span/)
    const rows = Array.from({ length: 2000 }, (_, at) => `<row r="${at + 1}">${Array.from({ length: 20 }, (_, c) => `<c r="${String.fromCharCode(65 + c)}${at + 1}" t="inlineStr"><is><t>${'字'.repeat(20)}</t></is></c>`).join('')}</row>`).join('')
    assert.ok(xlsxText(book([{ name: 'big', body: sheet(rows) }])).length <= XLSX_LIMITS.textChars)
  })
})

// ---------------------------------------------------------------------------------------------
// Files from the stress round: what a crafted or tool-written workbook must still show truthfully.

const chromiumPath = (() => {
  if (process.env.CHROMIUM_PATH) return process.env.CHROMIUM_PATH
  return [
    join(homedir(), 'Library/Caches/ms-playwright/chromium_headless_shell-1248/chrome-headless-shell-mac-arm64/chrome-headless-shell'),
    '/usr/bin/chromium',
    '/usr/bin/chromium-browser'
  ].find((candidate) => existsSync(candidate))
})()

/** A sheet with one chart over it; returns the sheet panel. */
function chartPanel(plot: string): string {
  const drawing = `<xdr:wsDr ${DRAWING_NS}><xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="4762500" cy="2857500"/><xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="1" name="c"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="c"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:absoluteAnchor></xdr:wsDr>`
  const page = renderXlsxPage(
    book([{ name: 's', body: sheet('', '', '<drawing r:id="d"/>'), rels: [['d', 'drawing', '../drawings/drawing1.xml']] }], {
      extra: {
        'xl/drawings/drawing1.xml': drawing,
        'xl/drawings/_rels/drawing1.xml.rels': rels([['c', 'chart', '../charts/chart1.xml']]),
        'xl/charts/chart1.xml': `<c:chartSpace ${DRAWING_NS}><c:chart><c:plotArea>${plot}</c:plotArea></c:chart></c:chartSpace>`
      }
    }),
    'b.xlsx',
    'b'
  )
  return panel(page, 1)
}

const numbers = (values: number[]) => `<c:numRef><c:f>x</c:f><c:numCache><c:ptCount val="${values.length}"/>${values.map((v, at) => `<c:pt idx="${at}"><c:v>${v}</c:v></c:pt>`).join('')}</c:numCache></c:numRef>`
const series = (color: string, values: number[], tag = 'val') => `<c:ser><c:spPr><a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></a:ln></c:spPr><c:${tag}>${numbers(values)}</c:${tag}></c:ser>`

/** Value-axis labels and the y of their gridline; `at(v)` is where value v sits. */
function valueAxis(svg: string) {
  const ticks = [...svg.matchAll(/<text x="[\d.]+" y="(-?[\d.]+)" text-anchor="end" font-size="10">([^<]*)<\/text>/g)].map((m) => ({ label: m[2], y: Number(m[1]) - 3 }))
  const [a, b] = ticks
  const value = (label: string) => Number(label.replace(/[,%]/g, ''))
  const at = (v: number) => a.y + ((v - value(a.label)) / (value(b.label) - value(a.label))) * (b.y - a.y)
  return { labels: ticks.map((t) => t.label), at }
}
const rectsOf = (svg: string, color: string) =>
  [...svg.matchAll(new RegExp(`<rect x="(-?[\\d.]+)" y="(-?[\\d.]+)" width="(-?[\\d.]+)" height="(-?[\\d.]+)" fill="#${color}"`, 'g'))].map((m) => ({ top: Number(m[2]), bottom: Number(m[2]) + Number(m[4]) }))
const close = (actual: number, expected: number, what: string) => assert.ok(Math.abs(actual - expected) < 0.6, `${what}: ${actual} vs ${expected}`)

describe('charts tell the truth about their data', () => {
  test('a 100% stacked chart fills each category to 100%, not to the sum of its values', () => {
    const svg = chartPanel(`<c:barChart><c:barDir val="col"/><c:grouping val="percentStacked"/><c:overlap val="100"/>${series('111111', [1, 30])}${series('222222', [3, 10])}</c:barChart>`)
    const axis = valueAxis(svg)
    assert.equal(axis.labels.at(-1), '100%', axis.labels.join(' '))
    const [a1, a2] = rectsOf(svg, '111111')
    const [b1, b2] = rectsOf(svg, '222222')
    close(a1.bottom, axis.at(0), 'first bar starts at 0')
    close(a1.top, axis.at(25), 'first part is 25%')
    close(b1.top, axis.at(100), 'second part ends at 100%')
    close(a2.top, axis.at(75), 'second category first part is 75%')
    close(b2.top, axis.at(100), 'second category ends at 100%')
  })

  test('a stacked bar keeps negatives below zero, apart from the positives', () => {
    const svg = chartPanel(`<c:barChart><c:barDir val="col"/><c:grouping val="stacked"/>${series('111111', [5])}${series('222222', [-2])}</c:barChart>`)
    const axis = valueAxis(svg)
    const [p] = rectsOf(svg, '111111')
    const [q] = rectsOf(svg, '222222')
    close(p.bottom, axis.at(0), 'P from 0')
    close(p.top, axis.at(5), 'P to 5')
    close(q.top, axis.at(0), 'Q from 0')
    close(q.bottom, axis.at(-2), 'Q down to -2')
  })

  test('stacked lines and areas stack; percent-stacked ones reach 100%', () => {
    const line = chartPanel(`<c:lineChart><c:grouping val="stacked"/>${series('111111', [1, 2])}${series('222222', [1, 2])}</c:lineChart>`)
    const axis = valueAxis(line)
    const ys = (svg: string, color: string) => /<polyline points="([^"]+)"[^>]*stroke="#COLOR"/.exec(svg.replace('COLOR', color)) ?? new RegExp(`<polyline points="([^"]+)"[^>]*stroke="#${color}"`).exec(svg)
    const second = ys(line, '222222')![1].split(' ').map((p) => Number(p.split(',')[1]))
    close(second[0], axis.at(2), 'stacked line first point')
    close(second[1], axis.at(4), 'stacked line second point')
    const area = chartPanel(`<c:areaChart><c:grouping val="percentStacked"/>${series('111111', [1, 3])}${series('222222', [3, 1])}</c:areaChart>`)
    const areaAxis = valueAxis(area)
    assert.equal(areaAxis.labels.at(-1), '100%', areaAxis.labels.join(' '))
    const path = new RegExp(`<path d="([^"]+)" fill="#111111"`).exec(area)![1]
    const ysOf = (d: string) => [...d.matchAll(/(-?[\d.]+),(-?[\d.]+)/g)].map((m) => Number(m[2]))
    assert.ok(ysOf(path).some((y) => Math.abs(y - areaAxis.at(25)) < 0.6), `first area reaches 25%: ${path}`)
    const top = new RegExp(`<path d="([^"]+)" fill="#222222"`).exec(area)![1]
    assert.ok(ysOf(top).filter((y) => Math.abs(y - areaAxis.at(100)) < 0.6).length >= 2, `second area tops out at 100%: ${top}`)
    assert.ok(ysOf(top).some((y) => Math.abs(y - areaAxis.at(75)) < 0.6), `second area sits on the first: ${top}`)
  })

  test('scatter and radar charts are a labelled box, not a line through x values in file order', () => {
    for (const kind of ['scatterChart', 'radarChart']) {
      const tag = kind === 'scatterChart' ? 'yVal' : 'val'
      const svg = chartPanel(`<c:${kind}><c:ser><c:xVal>${numbers([3, 1, 2])}</c:xVal><c:${tag}>${numbers([1, 2, 3])}</c:${tag}></c:ser></c:${kind}>`)
      assert.doesNotMatch(svg, /<svg/, kind)
      assert.match(svg, /class="xl-draw xl-missing"[^>]*>圖表\(此圖表類型網頁版未支援/, kind)
    }
  })
})

describe('cells show what Excel shows', () => {
  test('serif and monospace font names keep their kind of fallback', () => {
    const names = ['SimSun', 'NSimSun', 'Batang', 'Garamond', 'Palatino Linotype', 'Book Antiqua', 'MS Mincho', 'Microsoft Sans Serif', 'Century Gothic', 'Lucida Console', 'Fira Code', 'Source Code Pro', 'Inconsolata', 'Cascadia Code', 'Noto Sans Mono', 'Courier New']
    const kinds = ['serif', 'serif', 'serif', 'serif', 'serif', 'serif', 'serif', 'sans-serif', 'sans-serif', 'monospace', 'monospace', 'monospace', 'monospace', 'monospace', 'monospace', 'monospace']
    const styles = `<styleSheet ${NS}><fonts count="${names.length + 1}"><font><sz val="11"/><name val="Calibri"/></font>${names.map((n) => `<font><sz val="11"/><name val="${n}"/></font>`).join('')}</fonts><cellXfs count="${names.length + 1}"><xf/>${names.map((_, at) => `<xf fontId="${at + 1}"/>`).join('')}</cellXfs></styleSheet>`
    const rows = names.map((n, at) => `<row r="${at + 1}"><c r="A${at + 1}" s="${at + 1}" t="inlineStr"><is><t>${n}</t></is></c></row>`).join('')
    const page = renderXlsxPage(book([{ name: 's', body: sheet(rows) }], { styles }), 'b.xlsx', 'b')
    names.forEach((name, at) => {
      const rule = new RegExp(`td\\.x${at + 1}\\{[^}]*font-family:'${name}',([^;}]*)`).exec(page)
      assert.ok(rule, name)
      assert.equal(rule[1].split(',').at(-1), kinds[at], `${name}: ${rule[1]}`)
    })
  })

  test('a workbook that asks to be recalculated on open shows its formulas, not the 0 placeholders saved for them', () => {
    const body = sheet('<row r="2"><c r="B2"><v>3</v></c><c r="C2"><v>20</v></c><c r="D2"><f>B2*C2</f><v>0</v></c><c r="E2"><f>B2+1</f><v>4</v></c></row>')
    const recalc = renderXlsxPage(book([{ name: 's', body }], { workbookPr: '<calcPr calcId="124519" fullCalcOnLoad="1"/>' }), 'b.xlsx', 'b')
    assert.match(cellOf(recalc, '=B2*C2'), /class="f"/)
    assert.doesNotMatch(panel(recalc, 1), /<td class="n">0<\/td>/)
    assert.match(panel(recalc, 1), /<td class="n">4<\/td>/)
    const saved = renderXlsxPage(book([{ name: 's', body }]), 'b.xlsx', 'b')
    assert.match(panel(saved, 1), /<td class="n">0<\/td>/)
  })

  test('overflowing text stops at the next occupied cell', () => {
    const body = sheet(
      '<row r="1"><c r="A1" t="inlineStr"><is><t>This long title spills across B but must stop at C1</t></is></c><c r="C1" t="inlineStr"><is><t>C1 VALUE</t></is></c></row>',
      '<cols><col min="1" max="3" width="10" customWidth="1"/></cols>'
    )
    const page = renderXlsxPage(book([{ name: 's', body }]), 'b.xlsx', 'b')
    const cell = cellOf(page, 'This long title spills across B but must stop at C1')
    assert.match(cell, /class="ov"/)
    const width = Number(/<span class="ovs" style="width:([\d.]+)px">/.exec(cell)?.[1])
    const colWidth = Number([...page.matchAll(/<col style="width:([\d.]+)px">/g)][1][1])
    close(width, colWidth * 2 - 4, 'runs over A and B only')
  })

  test('rotated and stacked text: markup sits inside its own cell', () => {
    const page = rotationPage()
    assert.match(cellOf(page, 'Rotated 45 in narrow column'), /<td class="x1 rc"/)
  })

  test('rotated text stays visible in its cell and stacked text keeps the row height', { skip: !chromiumPath && 'no Chromium found' }, async () => {
    const browser = await chromium.launch({ executablePath: chromiumPath })
    try {
      const page = await browser.newPage()
      await page.setContent(rotationPage())
      const result = await page.evaluate(() =>
        [...document.querySelectorAll('td')].filter((td) => td.textContent).map((td) => {
          const text = (td.querySelector('.rot,.vt') ?? td).firstChild!
          const deep = text.nodeType === 3 ? text : text.firstChild!
          const range = document.createRange()
          range.setStart(deep, 0)
          range.setEnd(deep, 1)
          const glyph = range.getBoundingClientRect()
          const box = td.getBoundingClientRect()
          const x = glyph.left + glyph.width / 2
          const y = glyph.top + glyph.height / 2
          return { text: td.textContent, inside: x >= box.left && x <= box.right && y >= box.top && y <= box.bottom, row: td.parentElement!.getBoundingClientRect().height }
        })
      )
      for (const cell of result) {
        assert.ok(cell.inside, `first letter of "${cell.text}" is inside its cell`)
        close(cell.row, 120, `row of "${cell.text}" keeps its height`)
      }
    } finally {
      await browser.close()
    }
  })
})

/** Row 1 at 90pt: 45°, 90° up, 90° down (180) and stacked (255) text in narrow columns. */
function rotationPage(): string {
  const styles = `<styleSheet ${NS}><cellXfs count="5"><xf/><xf><alignment textRotation="45"/></xf><xf><alignment textRotation="90"/></xf><xf><alignment textRotation="180"/></xf><xf><alignment textRotation="255"/></xf></cellXfs></styleSheet>`
  const texts = ['Rotated 45 in narrow column', 'Rotated 90 upward', 'Rotated 90 downward', 'Stacked vertical text that is long']
  const cells = texts.map((t, at) => `<c r="${'ABCD'[at]}1" s="${at + 1}" t="inlineStr"><is><t>${t}</t></is></c>`).join('')
  const body = sheet(`<row r="1" ht="90" customHeight="1">${cells}</row>`, '<cols><col min="1" max="4" width="6" customWidth="1"/></cols>')
  return renderXlsxPage(book([{ name: 's', body }], { styles }), 'b.xlsx', 'b')
}
