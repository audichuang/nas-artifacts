import assert from 'node:assert/strict'
import { existsSync, mkdtempSync, rmSync } from 'node:fs'
import { homedir, tmpdir } from 'node:os'
import { join } from 'node:path'
import { after, before, describe, test } from 'node:test'
import { chromium } from 'playwright-core'
import { ArtifactStore } from '../src/artifact-store.ts'
import { createContentServer } from '../src/content-server.ts'
import { colName, formatGeneral, formatNumber, formatText, renderXlsxPage, xlsxText, XLSX_LIMITS } from '../src/xlsx.ts'
import { ZIP_LIMITS } from '../src/zip.ts'
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
<c:val><c:numRef><c:f>x</c:f><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>3</c:v></c:pt><c:pt idx="1"><c:v>7</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:barChart><c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx></c:plotArea></c:chart></c:chartSpace>`
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

  test('X-5: [DBNum1], [DBNum2] and [DBNum3] write the number in Chinese, as Excel does for cheques and ledgers', () => {
    const cases: [string, number, string][] = [
      ['[DBNum1][$-404]General', 1234, '一千二百三十四'],
      ['[DBNum1][$-404]General', 10, '一十'],
      ['[DBNum1][$-404]General', 1001, '一千〇一'],
      ['[DBNum1][$-404]General', 120000, '一十二萬'],
      ['[DBNum1][$-404]General', 12.5, '一十二.五'],
      ['[DBNum1][$-404]General', -3, '-三'],
      ['[DBNum2][$-404]General', 5678, '伍仟陸佰柒拾捌'],
      ['[DBNum2][$-404]#,##0', 1234567, '壹佰貳拾參萬肆仟伍佰陸拾柒'],
      ['[DBNum2][$-404]0.00', 0.5, '零.伍零'],
      ['[DBNum3][$-404]General', 1234, '１２３４'],
      ['[DBNum3][$-404]0.0', 12.34, '１２.３']
    ]
    for (const [code, value, text] of cases) assert.equal(formatNumber(code, value).text, text, `${code} ${value}`)
    // In a date the year is read digit by digit, the month and day as numbers.
    assert.equal(formatNumber('[DBNum1][$-404]yyyy"年"m"月"d"日"', 46027).text, '二〇二六年一月五日')
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
    assert.match(html, /\.xl-grid td\.x4\{color:#FF0000;font-family:'Arial',[^}]*background:url\('data:image\/svg\+xml,[^']*'\) 0 0\/8px 8px repeat;border-inline-end-color:transparent;border-bottom-color:transparent;text-align:right;padding-inline-end:20px\}/)
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
    assert.match(first, /<div class="xl-draw xl-chart"[^>]*><svg[^>]*aria-label="月營收"[\s\S]*<rect[^>]*fill="#F79646"[\s\S]*>一月<\/text>/)
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

  test('conditional formats decided by the value: comparisons with their dxf, colour scales, data bars, formula rules', () => {
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
    assert.match(panel(page, 2), /<div class="xl-chartsheet"[^>]*><svg[\s\S]*<path d="M[^"]*" fill="none" stroke="#4472C4"/)
  })

  test('a file that is not a workbook gets a page that says why, and a throw for the error log', () => {
    const notZip = renderXlsxPage(Buffer.from('PK definitely not'), 'x.xlsx', 'x')
    assert.match(notZip, /<div class="xl-fail" role="alert"><h1>這份 Excel 檔無法顯示<\/h1><p>這不是 \.xlsx 檔,或檔案已損壞[^<]*<\/p><p>可以下載原始檔,用 Excel 開啟查看。<\/p><\/div>/)
    assert.match(notZip, /<script>throw new Error\("Excel: 這不是 \.xlsx 檔/)
    const legacy = Buffer.alloc(64)
    legacy.writeUInt32BE(0xd0cf11e0, 0)
    assert.match(renderXlsxPage(legacy, 'x.xlsx', 'x'), /舊版 \.xls 或加密的活頁簿/)
    assert.match(renderXlsxPage(zip({ 'word/document.xml': '<w:document/>' }), 'x.xlsx', 'x'), /這不是 Excel 活頁簿/)
    const dtd = book([{ name: 's', body: '<!DOCTYPE x [<!ENTITY a "aaaa">]><worksheet><sheetData/></worksheet>' }])
    assert.match(renderXlsxPage(dtd, 'x.xlsx', 'x'), /工作表「s」無法顯示:不接受含 DTD 的 XML/)
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

const close = (actual: number, expected: number, what: string) => assert.ok(Math.abs(actual - expected) < 0.6, `${what}: ${actual} vs ${expected}`)
const AXES = '<c:axId val="1"/><c:axId val="2"/>'
const axis = (kind: string, id: number, cross: number, pos: string) => `<c:${kind}><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${pos}"/><c:crossAx val="${cross}"/></c:${kind}>`
const texts = (svg: string) => [...svg.matchAll(/<text[^>]*>([^<]*)<\/text>/g)].map((m) => m[1])

describe('charts draw through the shared chart renderer', () => {
  test('series take their colours from the file; a stacked bar draws every part', () => {
    const svg = chartPanel(`<c:barChart><c:barDir val="col"/><c:grouping val="percentStacked"/><c:overlap val="100"/>${series('111111', [1, 30])}${series('222222', [3, 10])}${AXES}</c:barChart>${axis('catAx', 1, 2, 'b')}${axis('valAx', 2, 1, 'l')}`)
    assert.equal((svg.match(/<rect [^>]*fill="#111111"/g) ?? []).length, 2)
    assert.equal((svg.match(/<rect [^>]*fill="#222222"/g) ?? []).length, 2)
    assert.ok(texts(svg).includes('100%'), texts(svg).join('|'))
  })

  test('a scatter chart plots its numeric x values on a value axis, not in file order', () => {
    const svg = chartPanel(`<c:scatterChart><c:scatterStyle val="lineMarker"/><c:ser><c:idx val="0"/><c:order val="0"/><c:xVal>${numbers([30, 10, 20])}</c:xVal><c:yVal>${numbers([1, 2, 3])}</c:yVal></c:ser>${AXES}</c:scatterChart>${axis('valAx', 1, 2, 'b')}${axis('valAx', 2, 1, 'l')}`)
    assert.doesNotMatch(svg, /xl-missing/)
    // The x axis is numeric: it labels 10, 20, 30 rather than the categories 1, 2, 3.
    for (const label of ['10', '20', '30']) assert.ok(texts(svg).includes(label), texts(svg).join('|'))
    const markers = [...svg.matchAll(/<path d="M(-?[\d.]+),(-?[\d.]+)[^"]*"[^>]*data-marker/g)].map((m) => Number(m[1]))
    assert.equal(markers.length, 3)
    // x = 30 is drawn right of x = 20, which is right of x = 10, whatever their order in the file.
    assert.ok(markers[0] > markers[2] && markers[2] > markers[1], markers.join(' '))
  })

  test('radar, bubble and combo charts draw instead of a labelled box', () => {
    const radar = chartPanel(`<c:radarChart><c:radarStyle val="marker"/>${series('333333', [1, 2, 3])}${AXES}</c:radarChart>${axis('catAx', 1, 2, 'b')}${axis('valAx', 2, 1, 'l')}`)
    assert.match(radar, /<svg[^>]*>[\s\S]*stroke="#333333"/)
    const bubble = chartPanel(`<c:bubbleChart><c:ser><c:idx val="0"/><c:order val="0"/><c:spPr><a:solidFill><a:srgbClr val="444444"/></a:solidFill></c:spPr><c:xVal>${numbers([1, 2])}</c:xVal><c:yVal>${numbers([1, 2])}</c:yVal><c:bubbleSize>${numbers([1, 4])}</c:bubbleSize></c:ser>${AXES}</c:bubbleChart>${axis('valAx', 1, 2, 'b')}${axis('valAx', 2, 1, 'l')}`)
    assert.match(bubble, /fill="#444444"/)
    const combo = chartPanel(
      `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${series('555555', [100, 140, 90])}${AXES}</c:barChart>` +
        `<c:lineChart><c:grouping val="standard"/>${series('666666', [0.2, 0.35, 0.1])}<c:axId val="3"/><c:axId val="4"/></c:lineChart>` +
        `${axis('catAx', 1, 2, 'b')}${axis('valAx', 2, 1, 'l')}${axis('catAx', 3, 4, 'b')}${axis('valAx', 4, 3, 'r')}`
    )
    assert.equal((combo.match(/<rect [^>]*fill="#555555"/g) ?? []).length, 3)
    assert.match(combo, /<path d="M[^"]*" fill="none" stroke="#666666"/)
    for (const svg of [radar, bubble, combo]) assert.doesNotMatch(svg, /xl-missing/)
  })
})

/** A sheet with one drawing anchored at the origin, 72 pt (96 px) square; returns the sheet panel. */
function shapePanel(objects: string): string {
  const drawing = `<xdr:wsDr ${DRAWING_NS}><xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="914400" cy="914400"/>${objects}<xdr:clientData/></xdr:absoluteAnchor></xdr:wsDr>`
  const page = renderXlsxPage(
    book([{ name: 's', body: sheet('', '', '<drawing r:id="d"/>'), rels: [['d', 'drawing', '../drawings/drawing1.xml']] }], { extra: { 'xl/drawings/drawing1.xml': drawing } }),
    'b.xlsx',
    'b'
  )
  return panel(page, 1)
}
const sp = (geometry: string, text = '', tag = 'sp') =>
  `<xdr:${tag}><xdr:nvSpPr><xdr:cNvPr id="2" name="s"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>${geometry}<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill><a:tailEnd type="triangle"/></a:ln></xdr:spPr>${text ? `<xdr:txBody><a:bodyPr/><a:p><a:r><a:t>${text}</a:t></a:r></a:p></xdr:txBody>` : ''}</xdr:${tag}>`

describe('drawing shapes are their ECMA-376 geometry', () => {
  test('a heart, a callout and a cloud draw their own paths, filled and then outlined', () => {
    const heart = shapePanel(sp('<a:prstGeom prst="heart"><a:avLst/></a:prstGeom>'))
    assert.match(heart, /<div class="xl-draw xl-geo"[^>]*><svg class="xl-geom" viewBox="0 0 72 72"[^>]*><path d="M[^"]*C[^"]*" fill="#FF0000"\/><path d="M[^"]*C[^"]*" fill="none" stroke="#0000FF"/)
    assert.doesNotMatch(heart, /border-radius|background:#FF0000/)
    // The callout's tail leaves the box below it.
    const callout = /<path d="([^"]*)" fill="#FF0000"/.exec(shapePanel(sp('<a:prstGeom prst="wedgeRectCallout"><a:avLst/></a:prstGeom>')))![1]
    assert.ok(Math.max(...(callout.match(/-?[\d.]+/g) ?? []).map(Number)) > 72, callout)
    const cloud = /<path d="([^"]*)" fill="#FF0000"/.exec(shapePanel(sp('<a:prstGeom prst="cloud"><a:avLst/></a:prstGeom>')))![1]
    assert.ok((cloud.match(/A/g) ?? []).length >= 10)
  })

  test("text sits in the geometry's text rectangle; a connector is an open line with its arrowhead", () => {
    const ellipse = shapePanel(sp('<a:prstGeom prst="ellipse"/>', '橢圓的字'))
    assert.match(ellipse, /<div class="xl-shape xl-shape-tx" style="left:14\.6\d*%;top:14\.6\d*%;width:70\.7\d*%;height:70\.7\d*%[^"]*">[\s\S]*橢圓的字/)
    const line = shapePanel(sp('<a:prstGeom prst="straightConnector1"/>', '', 'cxnSp'))
    assert.match(line, /<path d="M0,0 L72,72" fill="none" stroke="#0000FF"[^>]*marker-end="url\(#xl-m\d+\)"/)
    assert.doesNotMatch(line, /fill="#FF0000"/)
    // An unknown preset keeps its box and says so.
    assert.match(shapePanel(sp('<a:prstGeom prst="notAShape"/>')), /class="xl-draw xl-shape"[^>]*background:#FF0000[\s\S]*未支援的圖形:notAShape/)
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

  test('a workbook that asks to be recalculated on open keeps a saved 0 but marks it as possibly not computed', () => {
    // XlsxWriter saves 0 for formulas it did not compute, and also for one that really is 0 (=A1-5 with A1=5):
    // the file cannot tell them apart, so the 0 stays, marked, and charts and rules still see it.
    const body = sheet('<row r="2"><c r="B2"><v>3</v></c><c r="C2"><v>20</v></c><c r="D2"><f>B2*C2</f><v>0</v></c><c r="E2"><f>B2+1</f><v>4</v></c></row>')
    const recalc = renderXlsxPage(book([{ name: 's', body }], { workbookPr: '<calcPr calcId="124519" fullCalcOnLoad="1"/>' }), 'b.xlsx', 'b')
    assert.match(panel(recalc, 1), /<td class="n stale" title="[^"]*重新計算[^"]*">0<\/td>/)
    assert.doesNotMatch(panel(recalc, 1), /=B2\*C2/)
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

/** The text of the row-header cell on the row that holds `text`. */
const rowOf = (html: string, text: string) => {
  const at = html.indexOf(`>${text}<`)
  assert.ok(at >= 0, `cell ${text}`)
  const tr = html.lastIndexOf('<tr', at)
  return Number(/<th[^>]*>(\d+)<\/th>/.exec(html.slice(tr, at))?.[1])
}

describe('worksheet XML the way other tools write it', () => {
  test('row numbers in single quotes or with spaces around = keep their row', () => {
    const body = sheet(`<row r='3'><c r='A3' t='inlineStr'><is><t>單引號</t></is></c></row><row r = "7"><c r="A7" t="inlineStr"><is><t>等號旁有空白</t></is></c></row><row\nr="9"><c r="A9" t="inlineStr"><is><t>換行</t></is></c></row>`)
    const page = panel(renderXlsxPage(book([{ name: 's', body }]), 'b.xlsx', 'b'), 1)
    assert.equal(rowOf(page, '單引號'), 3)
    assert.equal(rowOf(page, '等號旁有空白'), 7)
    assert.equal(rowOf(page, '換行'), 9)
  })

  test('the last row number of a cut sheet is read whatever its quotes', () => {
    const rows = Array.from({ length: XLSX_LIMITS.rows + 1 }, (_, at) => `<row r="${at + 1}"><c r="A${at + 1}"><v>${at}</v></c></row>`).join('')
    const page = panel(renderXlsxPage(book([{ name: 's', body: sheet(`${rows}<row r = '4321'><c r="A4321"><v>1</v></c></row>`) }]), 'b.xlsx', 'b'), 1)
    assert.match(page, /此工作表共 4,321 列/)
  })

  test('<sheetData /> with a space before the slash is an empty sheet, and what follows it still draws', () => {
    const drawing = `<xdr:wsDr ${DRAWING_NS}><xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="952500" cy="952500"/><xdr:sp><xdr:nvSpPr><xdr:cNvPr id="1" name="t"/><xdr:cNvSpPr txBox="1"/></xdr:nvSpPr><xdr:spPr/><xdr:txBody><a:bodyPr/><a:p><a:r><a:t>方塊還在</a:t></a:r></a:p></xdr:txBody></xdr:sp><xdr:clientData/></xdr:absoluteAnchor></xdr:wsDr>`
    for (const empty of ['<sheetData />', '<sheetData/>', '<x:sheetData  />']) {
      const body = `<worksheet ${NS} xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${empty}<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells><drawing r:id="d"/></worksheet>`
      const page = panel(renderXlsxPage(book([{ name: 's', body, rels: [['d', 'drawing', '../drawings/drawing1.xml']] }], { extra: { 'xl/drawings/drawing1.xml': drawing } }), 'b.xlsx', 'b'), 1)
      assert.doesNotMatch(page, /無法顯示/, empty)
      assert.match(page, />方塊還在</, empty)
      assert.match(page, /colspan="2"/, empty)
    }
  })
})

describe('big workbooks', () => {
  test('in a browser, a sheet full of numbers wider than their columns loads quickly, shrunk or shown as ####', { skip: !chromiumPath && 'no Chromium found' }, async () => {
    // 500 rows × 20 columns: every number spills past its column, the way a sales sheet with narrow columns does.
    const rows = Array.from({ length: 500 }, (_, r) => `<row r="${r + 1}">${Array.from({ length: 20 }, (_, c) => `<c r="${colName(c + 1)}${r + 1}"><v>${c % 2 ? 1234567.891 * (c + 1) + r : 12345678 + r}</v></c>`).join('')}</row>`).join('')
    const styles = `<styleSheet ${NS}><cellXfs count="2"><xf/><xf><alignment shrinkToFit="1"/></xf></cellXfs></styleSheet>`
    const shrinkRow = `<row r="501"><c r="A501" s="1" t="inlineStr"><is><t>shrink this text to fit its cell</t></is></c></row>`
    const html = renderXlsxPage(book([{ name: 's', body: sheet(rows + shrinkRow, '<cols><col min="1" max="20" width="9" customWidth="1"/></cols>') }], { styles }), 'b.xlsx', 'b')
    const browser = await chromium.launch({ executablePath: chromiumPath })
    try {
      const page = await browser.newPage()
      const started = Date.now()
      await page.setContent(html)
      const took = Date.now() - started
      const result = await page.evaluate(() => {
        const cells = [...document.querySelectorAll<HTMLElement>('td.n')]
        const sh = document.querySelector<HTMLElement>('td.sh')!
        return {
          total: cells.length,
          hashes: cells.filter((td) => /^#+$/.test(td.textContent!) && /^\d/.test(td.title)).length,
          shrunk: cells.filter((td) => td.style.fontSize && !/#/.test(td.textContent!)).length,
          spill: cells.filter((td) => td.scrollWidth > td.clientWidth + 1).length,
          shFits: sh.firstElementChild!.getBoundingClientRect().width <= sh.clientWidth - 3
        }
      })
      assert.equal(result.total, 10_000)
      assert.ok(result.hashes > 0 && result.shrunk > 0, `some cells shrink, some show ####: ${JSON.stringify(result)}`)
      assert.equal(result.spill, 0, 'no number spills past its column')
      assert.ok(result.shFits, 'shrink-to-fit text fits its cell')
      assert.ok(took < 3000, `the page loads in ${took} ms`)
    } finally {
      await browser.close()
    }
  })

  test('a sheet whose XML is past the per-part limit shows its first rows and says what was cut', () => {
    const saved = ZIP_LIMITS.entryBytes
    ZIP_LIMITS.entryBytes = 60_000
    try {
      const rows = Array.from({ length: 3000 }, (_, at) => `<row r="${at + 1}"><c r="A${at + 1}" t="inlineStr"><is><t>第${at + 1}列</t></is></c><c r="B${at + 1}"><v>${at}</v></c></row>`).join('')
      const tail = '<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>'
      for (const dimension of ['<dimension ref="A1:B3000"/>', '']) {
        const page = renderXlsxPage(book([{ name: '大', body: sheet(rows, dimension, tail) }, { name: '小', body: sheet('<row r="1"><c r="A1" t="inlineStr"><is><t>第二張</t></is></c></row>') }]), 'b.xlsx', 'b')
        const big = panel(page, 1)
        assert.doesNotMatch(big, /無法顯示/)
        assert.match(big, />第1列</)
        assert.match(big, />第100列</)
        assert.doesNotMatch(big, />第3000列</)
        assert.match(big, /這個工作表的檔案太大,網頁版只讀取了前段/)
        assert.match(big, /合併儲存格、超連結、條件式格式與圖片沒有讀到/)
        if (dimension) assert.match(big, /此工作表共 3,000 列 × 2 欄,網頁版只顯示前 [\d,]+ 列/)
        assert.match(panel(page, 2), />第二張</)
      }
    } finally {
      ZIP_LIMITS.entryBytes = saved
    }
  })

  const grid = (rowCount: number, colCount: number) =>
    sheet(Array.from({ length: rowCount }, (_, r) => `<row r="${r + 1}">${Array.from({ length: colCount }, (_, c) => `<c r="${colName(c + 1)}${r + 1}"><v>${r + c}</v></c>`).join('')}</row>`).join(''))
  const shownRows = (html: string) => (html.match(/<tr style=[^>]*><th[^>]*>\d+<\/th>/g) ?? []).length

  test('an ordinary monthly workbook shows every row of every month', () => {
    const months = Array.from({ length: 12 }, (_, m) => ({ name: `${m + 1}月`, body: grid(600, 20) }))
    const page = renderXlsxPage(book(months), 'b.xlsx', 'b')
    for (let n = 1; n <= 12; n++) {
      assert.doesNotMatch(panel(page, n), /網頁版只顯示/, `${n}月`)
      assert.ok(shownRows(panel(page, n)) >= 600, `${n}月: ${shownRows(panel(page, n))} rows`)
    }
  })

  test('a sheet whose chart sits far below its data still draws the grid down to the chart', () => {
    const drawing = `<xdr:wsDr ${DRAWING_NS}><xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>390</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>400</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:sp><xdr:nvSpPr><xdr:cNvPr id="1" name="t"/><xdr:cNvSpPr txBox="1"/></xdr:nvSpPr><xdr:spPr/><xdr:txBody><a:bodyPr/><a:p><a:r><a:t>下面的方塊</a:t></a:r></a:p></xdr:txBody></xdr:sp><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>`
    const page = renderXlsxPage(book([{ name: 's', body: grid(5, 3).replace('</worksheet>', '<drawing r:id="d"/></worksheet>'), rels: [['d', 'drawing', '../drawings/drawing1.xml']] }, { name: 't', body: grid(5, 3) }], { extra: { 'xl/drawings/drawing1.xml': drawing } }), 'b.xlsx', 'b')
    assert.doesNotMatch(panel(page, 1), /網頁版只顯示/)
    assert.ok(shownRows(panel(page, 1)) >= 400, `${shownRows(panel(page, 1))} rows`)
  })

  test('large sheets early in a workbook cannot starve the sheets after them', () => {
    const hog = grid(2000, 60)
    const page = renderXlsxPage(book([{ name: 'a', body: hog }, { name: 'b', body: hog }, { name: 'c', body: hog }, { name: 'small', body: grid(600, 20) }]), 'b.xlsx', 'b')
    assert.ok(shownRows(panel(page, 4)) >= 600, `small: ${shownRows(panel(page, 4))}`)
    const hogs = [1, 2, 3].map((n) => shownRows(panel(page, n)))
    assert.equal(new Set(hogs).size, 1, `the large sheets share alike: ${hogs}`)
    assert.ok(hogs[0] >= XLSX_LIMITS.floorRows, `${hogs[0]}`)
    // However many sheets share the budget, each still draws its first floorRows rows.
    const many = renderXlsxPage(book(Array.from({ length: 30 }, (_, at) => ({ name: `s${at}`, body: grid(400, 60) }))), 'b.xlsx', 'b')
    assert.ok(shownRows(panel(many, 30)) >= XLSX_LIMITS.floorRows, `last of many: ${shownRows(panel(many, 30))}`)
  })
})

describe('what is not drawn says why, in Traditional Chinese', () => {
  test('a missing picture, a linked picture and an undrawable one each have their own label', () => {
    const anchor = (col: number, blip: string) =>
      `<xdr:oneCellAnchor><xdr:from><xdr:col>${col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="952500" cy="952500"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${col + 1}" name="p"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill>${blip}</xdr:blipFill><xdr:spPr/></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>`
    const drawing = `<xdr:wsDr ${DRAWING_NS}>${anchor(0, '<a:blip r:embed="gone"/>')}${anchor(2, '<a:blip r:embed="ext"/>')}${anchor(4, '<a:blip r:link="lnk"/>')}${anchor(6, '<a:blip r:embed="wmf"/>')}${anchor(8, '<a:blip/>')}</xdr:wsDr>`
    const page = renderXlsxPage(
      book([{ name: 's', body: sheet('', '', '<drawing r:id="d"/>'), rels: [['d', 'drawing', '../drawings/drawing1.xml']] }], {
        extra: {
          'xl/drawings/drawing1.xml': drawing,
          'xl/drawings/_rels/drawing1.xml.rels': rels([['gone', 'image', '../media/gone.png'], ['ext', 'image', 'file:///C:/pics/a.png', 'External'], ['lnk', 'image', 'https://example.com/a.png', 'External'], ['wmf', 'image', '../media/a.wmf']]),
          'xl/media/a.wmf': 'x'
        }
      }),
      'b.xlsx',
      'b'
    )
    const labels = [...page.matchAll(/class="xl-draw xl-missing"[^>]*>([^<]*)</g)].map((m) => m[1])
    assert.deepEqual(labels, [
      '圖片(檔案裡找不到這張圖片,可能已被移除或檔案不完整)',
      '圖片(連結到檔案外的圖片,網頁版不會載入;請在 Excel 中開啟原始檔查看)',
      '圖片(連結到檔案外的圖片,網頁版不會載入;請在 Excel 中開啟原始檔查看)',
      '圖片(WMF 格式,網頁版未支援,請下載原始檔查看)',
      '圖片(檔案裡找不到這張圖片,可能已被移除或檔案不完整)'
    ])
  })

  test('a part too deep to walk fails alone with a message in Chinese', () => {
    const deep = `<xdr:wsDr ${DRAWING_NS}><xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="1" cy="1"/>${'<xdr:grpSp>'.repeat(1200)}${'</xdr:grpSp>'.repeat(1200)}<xdr:clientData/></xdr:absoluteAnchor></xdr:wsDr>`
    const page = renderXlsxPage(
      book([{ name: 's', body: sheet('<row r="1"><c r="A1" t="inlineStr"><is><t>格子還在</t></is></c></row>', '', '<drawing r:id="d"/>'), rels: [['d', 'drawing', '../drawings/drawing1.xml']] }], { extra: { 'xl/drawings/drawing1.xml': deep } }),
      'b.xlsx',
      'b'
    )
    assert.match(panel(page, 1), />格子還在</)
    assert.match(panel(page, 1), /圖片與圖表無法顯示\(XML 巢狀超過 1000 層\)/)
  })
})

describe('pictures placed in cells (Excel 365)', () => {
  const META = `<metadata ${NS} xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"><metadataTypes count="1"><metadataType name="XLRICHVALUE"/></metadataTypes><futureMetadata name="XLRICHVALUE" count="3"><bk><extLst><ext uri="{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}"><xlrd:rvb i="0"/></ext></extLst></bk><bk><extLst><ext uri="{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}"><xlrd:rvb i="1"/></ext></extLst></bk><bk><extLst><ext uri="{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}"><xlrd:rvb i="2"/></ext></extLst></bk></futureMetadata><valueMetadata count="3"><bk><rc t="1" v="0"/></bk><bk><rc t="1" v="1"/></bk><bk><rc t="1" v="2"/></bk></valueMetadata></metadata>`
  const RV = '<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="3"><rv s="0"><v>1</v><v>5</v><v>商標 &lt;b&gt;</v></rv><rv s="1"><v>0</v><v>5</v></rv><rv s="0"><v>9</v><v>5</v><v>x</v></rv></rvData>'
  const RVS = '<rvStructures xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="2"><s t="_localImage"><k n="_rvRel:LocalImageIdentifier" t="i"/><k n="CalcOrigin" t="i"/><k n="Text" t="s"/></s><s t="_webimage"><k n="_rvRel:WebImageIdentifier" t="i"/><k n="CalcOrigin" t="i"/></s></rvStructures>'
  const RVR = `<richValueRels xmlns="http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><rel r:id="w"/><rel r:id="p"/></richValueRels>`
  const page = () => {
    const files = book(
      [{ name: 's', body: sheet('<row r="1" ht="60" customHeight="1"><c r="A1" t="e" vm="1"><v>#VALUE!</v></c><c r="B1" t="e" vm="2"><v>#VALUE!</v></c><c r="C1" t="e" vm="3"><v>#VALUE!</v></c><c r="D1" t="e" vm="7"><v>#VALUE!</v></c></row>', '<cols><col min="1" max="1" width="20" customWidth="1"/></cols>') }],
      {
        extra: {
          'xl/metadata.xml': META,
          'xl/richData/rdrichvalue.xml': RV,
          'xl/richData/rdrichvaluestructure.xml': RVS,
          'xl/richData/richValueRel.xml': RVR,
          'xl/richData/_rels/richValueRel.xml.rels': rels([['w', 'image', 'https://example.com/a.png', 'External'], ['p', 'image', '../media/image1.png']]),
          'xl/media/image1.png': PNG
        }
      }
    )
    return panel(renderXlsxPage(files, 'b.xlsx', 'b'), 1)
  }

  test('the picture is drawn inside its cell, with its alt text; the #VALUE! placeholder is not shown', () => {
    const html = page()
    assert.doesNotMatch(html, /#VALUE!/)
    // Column A is 20 characters wide (148px at Calibri 11) and the row 60pt (80px), less the padding.
    assert.match(html, /<td><img class="xl-cellpic" src="~office\/b\.xlsx\/xl\/media\/image1\.png" alt="商標 &#60;b&#62;" loading="lazy" style="width:\d+px;height:78px"><\/td>/)
  })

  test('a web picture and a chain that leads nowhere become labels in the cell', () => {
    const html = page()
    assert.match(html, /<span class="xl-cellmiss">儲存格內的網路圖片\(網頁版不會載入\)<\/span>/)
    assert.equal((html.match(/<span class="xl-cellmiss">儲存格內的圖片或資料類型\(網頁版未支援,請下載原始檔查看\)<\/span>/g) ?? []).length, 2)
  })
})

describe('text that spreads past its cell', () => {
  const spreadPage = () => {
    const styles = `<styleSheet ${NS}><cellXfs count="4"><xf/><xf><alignment horizontal="right"/></xf><xf><alignment horizontal="center"/></xf><xf><alignment horizontal="centerContinuous"/></xf></cellXfs></styleSheet>`
    const t = (ref: string, s: number, text: string) => `<c r="${ref}" s="${s}" t="inlineStr"><is><t>${text}</t></is></c>`
    const body = sheet(
      `<row r="1">${t('D1', 1, 'right aligned long text spills left')}</row>` +
        `<row r="2">${t('C2', 2, 'centered long text spills both ways')}</row>` +
        `<row r="3">${t('A3', 3, 'centre across')}<c r="B3" s="3"/><c r="C3" s="3"/></row>` +
        `<row r="4"><c r="B4"><v>1</v></c>${t('C4', 1, 'right but blocked by B4')}</row>`,
      '<cols><col min="1" max="6" width="12" customWidth="1"/></cols>'
    )
    return renderXlsxPage(book([{ name: 's', body }], { styles }), 'b.xlsx', 'b')
  }

  test('right-aligned text runs back over empty cells; centred text both ways; centre-across over its selection', () => {
    const page = spreadPage()
    assert.match(cellOf(page, 'right aligned long text spills left'), /class="x1 ov"[^>]*><span class="ovs ovr" style="width:\d+px;margin-inline-start:-\d+px"><span>/)
    assert.match(cellOf(page, 'centered long text spills both ways'), /class="x2 ov"[^>]*><span class="ovs ovc" style="width:\d+px;margin-inline-start:-\d+px;clip-path:inset\(0 0px 0 \d+px\)"><span>/)
    assert.match(cellOf(page, 'centre across'), /class="x3 ov"[^>]*><span class="ovs" style="width:\d+px">/)
    assert.doesNotMatch(cellOf(page, 'right but blocked by B4'), /class="x1 ov"/)
  })

  test('in a browser, spread text shows whole, placed as Excel places it', { skip: !chromiumPath && 'no Chromium found' }, async () => {
    const browser = await chromium.launch({ executablePath: chromiumPath })
    try {
      const page = await browser.newPage()
      await page.setContent(spreadPage())
      const boxes = await page.evaluate(() => {
        const out: Record<string, { text: [number, number]; cell: [number, number]; row: [number, number] }> = {}
        for (const td of document.querySelectorAll('td')) {
          if (!td.textContent) continue
          const range = document.createRange()
          range.selectNodeContents(td)
          const text = range.getBoundingClientRect()
          const cell = td.getBoundingClientRect()
          const cells = [...td.parentElement!.querySelectorAll('td')]
          const first = cells[0].getBoundingClientRect()
          const third = cells[2].getBoundingClientRect()
          out[td.textContent] = { text: [text.left, text.right], cell: [cell.left, cell.right], row: [first.left, third.right] }
        }
        return out
      })
      const right = boxes['right aligned long text spills left']
      close(right.text[1], right.cell[1] - 2, 'right-aligned text ends at its cell’s right edge')
      assert.ok(right.text[0] < right.cell[0] - 50, 'and reaches back over the empty cells')
      const centred = boxes['centered long text spills both ways']
      close((centred.text[0] + centred.text[1]) / 2, (centred.cell[0] + centred.cell[1]) / 2, 'centred text is centred on its cell')
      assert.ok(centred.text[0] < centred.cell[0] && centred.text[1] > centred.cell[1], 'and spreads both ways')
      const across = boxes['centre across']
      close((across.text[0] + across.text[1]) / 2, (across.row[0] + across.row[1]) / 2, 'centre-across text is centred over A3:C3')
      const visible = await page.evaluate(() => [...document.querySelectorAll('td.ov .ovs')].every((span) => span.scrollWidth <= span.clientWidth + 1))
      assert.ok(visible, 'no spread text is cut inside its run')
      // Right-aligned text with no room left keeps its end in view, as Excel does, and loses its start.
      await page.setContent(renderXlsxPage(book([{ name: 's', body: sheet('<row r="1"><c r="B1" s="1" t="inlineStr"><is><t>START of a right aligned text far too long for A and B END</t></is></c></row>', '<cols><col min="1" max="2" width="8" customWidth="1"/></cols>') }], { styles: `<styleSheet ${NS}><cellXfs count="2"><xf/><xf><alignment horizontal="right"/></xf></cellXfs></styleSheet>` }), 'b.xlsx', 'b'))
      const cut = await page.evaluate(() => {
        const td = document.querySelector('td.ov')!
        const range = document.createRange()
        range.selectNodeContents(td)
        return { text: range.getBoundingClientRect().right, cell: td.getBoundingClientRect().right }
      })
      close(cut.text, cut.cell - 2, 'the end of the text stays at the right edge')
    } finally {
      await browser.close()
    }
  })
})

describe('conditional formats beyond comparisons and scales', () => {
  const STYLES_CF = `<styleSheet ${NS}><dxfs count="1"><dxf><font><b/><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf></dxfs></styleSheet>`
  /** A sheet of the given rows (arrays of cell XML bodies by column) with the given rules; returns a cell picker. */
  const cfPage = (rows: string[][], rules: string) => {
    const xml = rows.map((cells, r) => `<row r="${r + 1}">${cells.map((cell, c) => (cell ? `<c r="${colName(c + 1)}${r + 1}"${cell}</c>` : '')).join('')}</row>`).join('')
    const page = panel(renderXlsxPage(book([{ name: 's', body: sheet(xml, '', rules) }], { styles: STYLES_CF }), 'b.xlsx', 'b'), 1)
    const cell = (ref: string) => {
      const col = ref.charCodeAt(0) - 64
      const row = Number(ref.slice(1))
      const start = page.indexOf(`<th>${row}</th>`)
      const tr = page.slice(start, page.indexOf('</tr>', start))
      return [...tr.matchAll(/<td[^>]*>(?:(?!<\/td>).)*<\/td>/g)][col - 1][0]
    }
    return { page, cell, styled: (ref: string) => /background:#FFC7CE/.test(cell(ref)) }
  }
  const n = (v: number) => `><v>${v}</v>`
  const t = (v: string) => ` t="inlineStr"><is><t>${v}</t></is>`
  const rule = (sqref: string, body: string) => `<conditionalFormatting sqref="${sqref}">${body}</conditionalFormatting>`

  test('top and bottom N, by count and percent; above and below the average', () => {
    const rows = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10].map((v) => [n(v), n(v), n(v), n(v)])
    const { styled } = cfPage(rows,
      rule('A1:A10', '<cfRule type="top10" dxfId="0" priority="1" rank="3"/>') +
      rule('B1:B10', '<cfRule type="top10" dxfId="0" priority="2" rank="20" percent="1" bottom="1"/>') +
      rule('C1:C10', '<cfRule type="aboveAverage" dxfId="0" priority="3"/>') +
      rule('D1:D10', '<cfRule type="aboveAverage" dxfId="0" priority="4" aboveAverage="0" equalAverage="1"/>'))
    assert.deepEqual([7, 8, 9, 10].map((r) => styled(`A${r}`)), [false, true, true, true])
    assert.deepEqual([1, 2, 3].map((r) => styled(`B${r}`)), [true, true, false])
    assert.deepEqual([5, 6].map((r) => styled(`C${r}`)), [false, true])
    assert.deepEqual([5, 6].map((r) => styled(`D${r}`)), [true, false])
  })

  test('text, blank, error and duplicate rules', () => {
    const rows = [[t('Apple pie'), t('x'), '', t('甲')], [t('banana'), t(' '), ' t="e"><v>#N/A</v>', t('乙')], [t('PINEAPPLE'), n(3), n(1), t('甲')]]
    const { styled } = cfPage(rows,
      rule('A1:A3', '<cfRule type="containsText" dxfId="0" priority="1" operator="containsText" text="apple"><formula>NOT(ISERROR(SEARCH("apple",A1)))</formula></cfRule>') +
      rule('B1:B3', '<cfRule type="containsBlanks" dxfId="0" priority="2"><formula>LEN(TRIM(B1))=0</formula></cfRule>') +
      rule('C1:C3', '<cfRule type="containsErrors" dxfId="0" priority="3"><formula>ISERROR(C1)</formula></cfRule>') +
      rule('D1:D3', '<cfRule type="duplicateValues" dxfId="0" priority="4"/>'))
    assert.deepEqual(['A1', 'A2', 'A3'].map(styled), [true, false, true])
    assert.deepEqual(['B1', 'B2', 'B3'].map(styled), [false, true, false])
    assert.deepEqual(['C2', 'C3'].map(styled), [true, false])
    assert.deepEqual(['D1', 'D2', 'D3'].map(styled), [true, false, true])
  })

  test('formula rules run relative to the first cell of their range, and say so when they cannot', () => {
    const rows = [[t('完成'), n(1)], [t('進行中'), n(2)], [t('完成'), n(3)], [t('取消'), n(4)]]
    const { styled, page } = cfPage(rows,
      rule('A1:B4', '<cfRule type="expression" dxfId="0" priority="1"><formula>$A1="完成"</formula></cfRule>') +
      rule('B1:B4', '<cfRule type="expression" dxfId="0" priority="2"><formula>AND(MOD(ROW(),2)=0,B1&gt;=4)</formula></cfRule>'))
    assert.deepEqual(['A1', 'B1', 'A2', 'B2', 'A3', 'B3', 'B4'].map(styled), [true, true, false, false, true, true, true])
    assert.doesNotMatch(page, /條件式格式/)
    const other = cfPage(rows, rule('A1:A4', '<cfRule type="expression" dxfId="0" priority="1"><formula>VLOOKUP(A1,Sheet2!A:B,2,0)&gt;1</formula></cfRule><cfRule type="timePeriod" dxfId="0" priority="2" timePeriod="today"><formula>FLOOR(A1,1)=TODAY()</formula></cfRule>'))
    assert.match(other.page, /此工作表有部分條件式格式\(公式規則、日期期間規則\)網頁版沒有套用/)
    assert.ok(!other.styled('A1'))
  })

  test('comparisons take text and formulas as operands, and treat a blank cell as Excel does', () => {
    const rows = [[t('Done'), n(5), n(1)], [t('done'), n(12), n(9)], [t('進行中'), n(7), '']]
    const { styled } = cfPage(rows,
      rule('A1:A3', '<cfRule type="cellIs" dxfId="0" priority="1" operator="equal"><formula>"DONE"</formula></cfRule>') +
      rule('B1:B3', '<cfRule type="cellIs" dxfId="0" priority="2" operator="between"><formula>$C$1*6</formula><formula>C1</formula></cfRule>') +
      rule('C1:C3', '<cfRule type="cellIs" dxfId="0" priority="3" operator="lessThan"><formula>5</formula></cfRule>'))
    assert.deepEqual(['A1', 'A2', 'A3'].map(styled), [true, true, false])
    // Between $C$1*6 (6) and C of the same row, in either order: 5 in [1,6], 12 not in [6,9], 7 not in [0,6].
    assert.deepEqual(['B1', 'B2', 'B3'].map(styled), [true, false, false])
    // An empty cell counts as 0, so "less than 5" marks it, as in Excel.
    assert.deepEqual(['C1', 'C2', 'C3'].map(styled), [true, false, true])
  })

  test('icon sets put their icon at the start of the cell; reversed sets and hidden values follow the file', () => {
    const rows = [1, 5, 9].map((v) => [n(v), n(v), n(v)])
    const { cell } = cfPage(rows,
      rule('A1:A3', '<cfRule type="iconSet" priority="1"><iconSet iconSet="3Arrows"><cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/></iconSet></cfRule>') +
      rule('B1:B3', '<cfRule type="iconSet" priority="2"><iconSet iconSet="3TrafficLights1" reverse="1" showValue="0"><cfvo type="percent" val="0"/><cfvo type="num" val="5"/><cfvo type="num" val="9" gte="0"/></iconSet></cfRule>'))
    assert.match(cell('A1'), /<span class="xl-ico" style="color:#D9363E">↓<\/span>1</)
    assert.match(cell('A2'), /<span class="xl-ico" style="color:#F2B705">→<\/span>5</)
    assert.match(cell('A3'), /<span class="xl-ico" style="color:#1E9E4A">↑<\/span>9</)
    // Reversed: the lowest value gets the green light; 9 is not above 9, so it stays in the middle band.
    assert.match(cell('B1'), /<span class="xl-ico" style="color:#1E9E4A">●<\/span><\/td>/)
    assert.match(cell('B3'), /<span class="xl-ico" style="color:#F2B705">●<\/span><\/td>/)
    assert.doesNotMatch(cell('C1'), /xl-ico/)
  })

  test('Excel 2010 icon sets in the extension list are drawn, with custom icons', () => {
    const rows = [1, 5, 9].map((v) => [n(v)])
    const ext = `<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:conditionalFormattings><x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:cfRule type="iconSet" priority="1" id="{1}"><x14:iconSet iconSet="3Triangles" custom="1"><x14:cfvo type="percent"><xm:f>0</xm:f></x14:cfvo><x14:cfvo type="percent"><xm:f>33</xm:f></x14:cfvo><x14:cfvo type="percent"><xm:f>67</xm:f></x14:cfvo><x14:cfIcon iconSet="3Flags" iconId="0"/><x14:cfIcon iconSet="NoIcons" iconId="0"/><x14:cfIcon iconSet="3Stars" iconId="2"/></x14:iconSet></x14:cfRule><xm:sqref>A1:A3</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>`
    const { cell } = cfPage(rows, ext)
    assert.match(cell('A1'), /<span class="xl-ico" style="color:#D9363E">⚑<\/span>1</)
    assert.doesNotMatch(cell('A2'), /xl-ico/)
    assert.match(cell('A3'), /<span class="xl-ico" style="color:#F2B705">★<\/span>9</)
  })
})

describe('numbers too wide for their column (matrix round)', () => {
  test('X-2: in a browser, dates of one format in one column all shrink or all show #, whatever their digits', { skip: !chromiumPath && 'no Chromium found' }, async () => {
    const styles = `<styleSheet ${NS}><numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/></numFmts><cellXfs count="2"><xf/><xf numFmtId="164" applyNumberFormat="1"/></cellXfs></styleSheet>`
    // 2026-01-01, 2026-01-11, 2026-01-15, 2026-08-28, 2026-12-31 …, in a column of the default width (as openpyxl
    // leaves it): the 1s are narrow in a font with proportional digits.
    const serials = [46023, 46033, 46037, 46262, 46387, 46100, 46200, 46300]
    // Column B holds numbers that fit however narrow their 1s are: they keep their size.
    const fits = [1, 31000, 11111]
    const body = sheet(serials.map((v, at) => `<row r="${at + 1}"><c r="A${at + 1}" s="1"><v>${v}</v></c>${at < fits.length ? `<c r="B${at + 1}"><v>${fits[at]}</v></c>` : ''}</row>`).join(''))
    const browser = await chromium.launch({ executablePath: chromiumPath })
    try {
      const page = await browser.newPage()
      await page.setContent(renderXlsxPage(book([{ name: 's', body }], { styles }), 'b.xlsx', 'b'))
      const look = (column: number) =>
        page.evaluate((column) => [...document.querySelectorAll('tbody tr')].map((tr) => tr.children[column] as HTMLElement).filter((td) => td?.classList.contains('n')).map((td) => (/^#+$/.test(td.textContent ?? '') ? '#' : td.style.fontSize || 'as set')), column)
      const shown = await look(1)
      assert.equal(shown.length, serials.length)
      assert.deepEqual(await look(2), fits.map(() => 'as set'))
      assert.equal(new Set(shown).size, 1, shown.join(' | '))
    } finally {
      await browser.close()
    }
  })
})

describe('cell fills and borders (matrix round)', () => {
  const PATTERN_TYPES = ['gray0625', 'gray125', 'lightGray', 'mediumGray', 'darkGray', 'darkHorizontal', 'darkVertical', 'darkDown', 'darkUp', 'darkGrid', 'darkTrellis', 'lightHorizontal', 'lightVertical', 'lightDown', 'lightUp', 'lightGrid', 'lightTrellis']
  const styled = (fills: string, borders: string, xfs: string) =>
    `<styleSheet ${NS}><fills count="99"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>${fills}</fills><borders count="9"><border/>${borders}</borders><cellXfs count="99"><xf/>${xfs}</cellXfs></styleSheet>`
  const cssOf = (html: string, xf: number) => new RegExp(`\\.xl-grid td\\.x${xf}\\{([^}]*)\\}`).exec(html)?.[1] ?? ''

  test('X-3: each of the 17 pattern fills is its hatch in its two colours, not a blend', () => {
    const fills = PATTERN_TYPES.map((type) => `<fill><patternFill patternType="${type}"><fgColor rgb="FF1F4E79"/><bgColor rgb="FFFFE699"/></patternFill></fill>`).join('')
    const xfs = PATTERN_TYPES.map((_, at) => `<xf fillId="${at + 2}" applyFill="1"/>`).join('')
    const body = sheet(`<row r="1">${PATTERN_TYPES.map((_, at) => `<c r="${colName(at + 1)}1" s="${at + 1}"><v>${at}</v></c>`).join('')}</row>`)
    const html = renderXlsxPage(book([{ name: 's', body }], { styles: styled(fills, '', xfs) }), 'b.xlsx', 'b')
    const tiles = new Set<string>()
    PATTERN_TYPES.forEach((type, at) => {
      const css = cssOf(html, at + 1)
      const m = /background:url\('data:image\/svg\+xml,([^']*)'\) 0 0\/8px 8px repeat/.exec(css)
      assert.ok(m, `${type}: ${css}`)
      const svg = decodeURIComponent(m[1])
      assert.match(svg, /<rect width="8" height="8" fill="#FFE699"\/><path d="M[^"]+" fill="#1F4E79"\/>/, type)
      tiles.add(svg)
    })
    assert.equal(tiles.size, PATTERN_TYPES.length, 'every pattern its own tile')
    // A solid fill stays its foreground colour.
    const solid = renderXlsxPage(book([{ name: 's', body: sheet('<row r="1"><c r="A1" s="1"><v>1</v></c></row>') }], { styles: styled('<fill><patternFill patternType="solid"><fgColor rgb="FF00B050"/></patternFill></fill>', '', '<xf fillId="2" applyFill="1"/>') }), 'b.xlsx', 'b')
    assert.match(cssOf(solid, 1), /background:#00B050/)
  })

  test('X-3: diagonalUp and diagonalDown borders draw their lines over the fill', () => {
    const borders = `<border diagonalUp="1"><diagonal style="thin"><color rgb="FFFF0000"/></diagonal></border><border diagonalUp="1" diagonalDown="1"><diagonal style="dashed"><color rgb="FF0000FF"/></diagonal></border><border diagonalDown="1"><diagonal style="none"/></border>`
    const fills = '<fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill>'
    const xfs = '<xf borderId="1" applyBorder="1"/><xf borderId="2" fillId="2" applyBorder="1" applyFill="1"/><xf borderId="3" applyBorder="1"/>'
    const body = sheet('<row r="1"><c r="A1" s="1"/><c r="B1" s="2"/><c r="C1" s="3"/></row>')
    const html = renderXlsxPage(book([{ name: 's', body }], { styles: styled(fills, borders, xfs) }), 'b.xlsx', 'b')
    const lines = (css: string) => [...css.matchAll(/url\('data:image\/svg\+xml,([^']*)'\) 0 0\/100% 100% no-repeat/g)].map((m) => decodeURIComponent(m[1]))
    const up = lines(cssOf(html, 1))
    assert.equal(up.length, 1)
    assert.match(up[0], /<line x1="0" y1="100" x2="100" y2="0" stroke="#FF0000"[^>]*vector-effect="non-scaling-stroke"/)
    const both = cssOf(html, 2)
    assert.equal(lines(both).length, 1)
    assert.match(lines(both)[0], /<line x1="0" y1="100" x2="100" y2="0"[^>]*stroke-dasharray/)
    assert.match(lines(both)[0], /<line x1="0" y1="0" x2="100" y2="100"/)
    // The lines lie over the cell's own fill, the last background layer.
    assert.match(both, /background:url\([^)]*\) 0 0\/100% 100% no-repeat,#FFFF00/)
    assert.equal(lines(cssOf(html, 3)).length, 0)
  })
})

describe('workbooks past the zip limits (matrix round)', () => {
  test('P-1 (Excel): a workbook past the zip limits is called too big, not broken', () => {
    const saved = { ...ZIP_LIMITS }
    ZIP_LIMITS.entries = 2
    try {
      const html = renderXlsxPage(book([{ name: 's', body: sheet('<row r="1"><c r="A1"><v>1</v></c></row>') }]), 'big.xlsx', 'big')
      assert.match(html, /檔案太大/)
      assert.doesNotMatch(html, /損[毀壞]|不是/)
    } finally {
      Object.assign(ZIP_LIMITS, saved)
    }
  })
})

describe('sparklines (matrix round)', () => {
  const X14 = 'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"'
  const group = (attrs: string, f: string, sqref: string) =>
    `<x14:sparklineGroup ${attrs}><x14:colorSeries rgb="FF376092"/><x14:colorNegative rgb="FFD00000"/><x14:colorAxis rgb="FF000000"/><x14:colorMarkers rgb="FF00B050"/><x14:colorHigh rgb="FFFFC000"/><x14:colorLow rgb="FF7030A0"/><x14:sparklines><x14:sparkline><xm:f>${f}</xm:f><xm:sqref>${sqref}</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup>`
  const row = (r: number, values: number[]) => `<row r="${r}">${values.map((v, at) => `<c r="${colName(at + 2)}${r}"><v>${v}</v></c>`).join('')}</row>`
  const page = () => {
    const ext = `<extLst><ext ${X14} uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}"><x14:sparklineGroups ${X14}>` +
      group('markers="1" high="1"', "'銷售 表'!B2:E2", 'F2') + group('type="column" negative="1" low="1"', "'銷售 表'!B3:E3", 'F3') + group('type="stacked" negative="1"', "'銷售 表'!B4:E4", 'F4') +
      group('', 'Nowhere!B2:E2', 'F5') + `</x14:sparklineGroups></ext></extLst>`
    const body = sheet(row(2, [3, 7, 2, 5]) + row(3, [4, -2, 6, -1]) + row(4, [1, -1, 0, 1]) + '<row r="5"><c r="A5" t="inlineStr"><is><t>壞</t></is></c></row>', '', ext)
    return renderXlsxPage(book([{ name: '銷售 表', body }]), 'b.xlsx', 'b')
  }

  test('X-1: line, column and win/loss sparklines draw in their cells, in their colours', () => {
    const html = page()
    const svgIn = (ref: string) => {
      const [, col, r] = /^([A-Z]+)(\d+)$/.exec(ref)!
      const tr = html.split('<tr').find((t) => new RegExp(`^[^>]*><th[^>]*>${r}</th>`).test(t)) ?? ''
      const cells = tr.split('<td').slice(1)
      return cells[col.charCodeAt(0) - 65] ?? ''
    }
    const line = svgIn('F2')
    assert.match(line, /<svg class="xl-spark"[^>]*>/)
    assert.match(line, /<polyline points="[\d.,\s]+" fill="none" stroke="#376092"/)
    // Markers on every point, the high point in its own colour.
    assert.equal((line.match(/<circle [^>]*fill="#00B050"/g) ?? []).length, 3)
    assert.equal((line.match(/<circle [^>]*fill="#FFC000"/g) ?? []).length, 1)
    const column = svgIn('F3')
    assert.equal((column.match(/<rect /g) ?? []).length, 4)
    // The low point's colour goes over the negative colour, as Excel layers them.
    assert.equal((column.match(/<rect [^>]*fill="#D00000"/g) ?? []).length, 1, 'the other negative bar')
    assert.equal((column.match(/<rect [^>]*fill="#7030A0"/g) ?? []).length, 1, 'the low bar')
    const winLoss = svgIn('F4')
    // Win/loss: one bar per non-zero value, up for a win and down for a loss, all one height.
    const bars = [...winLoss.matchAll(/<rect x="[\d.]+" y="([\d.]+)" width="[\d.]+" height="([\d.]+)" fill="(#[0-9A-F]+)"/g)].map((m) => ({ y: +m[1], h: +m[2], fill: m[3] }))
    assert.equal(bars.length, 3)
    assert.equal(new Set(bars.map((b) => b.h)).size, 1)
    assert.deepEqual(bars.map((b) => b.fill), ['#376092', '#D00000', '#376092'])
    assert.ok(bars[1].y > bars[0].y)
  })

  test('X-1: a sparkline whose data cannot be found says so in its cell', () => {
    assert.match(page(), /<span class="xl-cellmiss">走勢圖\(找不到資料\)<\/span>/)
  })
})
