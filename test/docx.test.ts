import assert from 'node:assert/strict'
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { after, before, describe, test } from 'node:test'
import { ArtifactStore } from '../src/artifact-store.ts'
import { createContentServer } from '../src/content-server.ts'
import { docxText, readDocx, renderDocxPage } from '../src/docx.ts'
import { ZIP_LIMITS } from '../src/zip.ts'
import { zip } from './helpers/zip.ts'

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const rels = (list: [string, string, string, boolean?][]) =>
  `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${list
    .map(([id, type, target, external]) => `<Relationship Id="${id}" Type="${REL}/${type}" Target="${target}"${external ? ' TargetMode="External"' : ''}/>`)
    .join('')}</Relationships>`

const theme = `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="t"><a:themeElements>
<a:clrScheme name="c"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
<a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6>
<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme>
<a:fontScheme name="f"><a:majorFont><a:latin typeface="Georgia"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Hant" typeface="標楷體"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Hant" typeface="新細明體"/></a:minorFont></a:fontScheme>
</a:themeElements></a:theme>`

const styles = `<w:styles ${W}>
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:hAnsiTheme="minorHAnsi"/><w:sz w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Body"><w:name w:val="Body"/><w:rPr><w:color w:val="333333"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="H1"><w:name w:val="heading 1"/><w:basedOn w:val="Body"/><w:pPr><w:outlineLvl w:val="0"/><w:spacing w:before="480"/></w:pPr><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia"/><w:b/><w:sz w:val="40"/><w:color w:val="2F5496" w:themeColor="accent1" w:themeShade="BF"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="H2"><w:name w:val="heading 2"/><w:basedOn w:val="H1"/><w:rPr><w:i/><w:sz w:val="28"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Loop1"><w:basedOn w:val="Loop2"/></w:style><w:style w:type="paragraph" w:styleId="Loop2"><w:basedOn w:val="Loop1"/></w:style>
<w:style w:type="paragraph" w:styleId="ListPara"><w:basedOn w:val="Body"/><w:pPr><w:contextualSpacing/><w:ind w:left="720"/></w:pPr></w:style>
<w:style w:type="character" w:default="1" w:styleId="DefaultFont"><w:name w:val="Default Paragraph Font"/></w:style>
<w:style w:type="character" w:styleId="Strong"><w:rPr><w:b/><w:color w:val="C00000"/></w:rPr></w:style>
<w:style w:type="character" w:styleId="Hyperlink"><w:rPr><w:color w:val="0563C1" w:themeColor="hyperlink"/><w:u w:val="single"/></w:rPr></w:style>
<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:tblPr><w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
<w:style w:type="table" w:styleId="Grid"><w:basedOn w:val="TableNormal"/><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:color="auto"/><w:left w:val="single" w:sz="4" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:color="auto"/><w:right w:val="single" w:sz="4" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:color="auto"/></w:tblBorders></w:tblPr>
<w:tblStylePr w:type="firstRow"><w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="4472C4" w:themeFill="accent1"/></w:tcPr></w:tblStylePr></w:style>
</w:styles>`

const numbering = `<w:numbering ${W}>
<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%1.%2)"/><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>
<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="&#xF0B7;"/><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>
<w:abstractNum w:abstractNumId="2"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="taiwaneseCountingThousand"/><w:lvlText w:val="%1、"/></w:lvl><w:lvl w:ilvl="1"><w:start w:val="4"/><w:numFmt w:val="upperRoman"/><w:lvlText w:val="(%2)"/></w:lvl></w:abstractNum>
<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num><w:num w:numId="3"><w:abstractNumId w:val="2"/></w:num>
<w:num w:numId="4"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num>
</w:numbering>`

const p = (body: string, pPr = '') => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${body}</w:p>`
const r = (text: string, rPr = '') => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`
const numbered = (numId: number, ilvl: number, text: string) => p(r(text), `<w:pStyle w:val="ListPara"/><w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${numId}"/></w:numPr>`)
const picture = (rid: string, cx: number, cy: number, anchor = false) => {
  const graphic = `<wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="1" name="p" descr="圖 &lt;一&gt;"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="p"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${rid}"/></pic:blipFill><pic:spPr/></pic:pic></a:graphicData></a:graphic>`
  return anchor
    ? `<w:r><w:drawing><wp:anchor behindDoc="0" distT="0" distB="0" distL="114300" distR="114300"><wp:positionH relativeFrom="column"><wp:align>right</wp:align></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>${graphic.replace('<wp:docPr', '<wp:wrapSquare wrapText="bothSides"/><wp:docPr')}</wp:anchor></w:drawing></w:r>`
    : `<w:r><w:drawing><wp:inline>${graphic}</wp:inline></w:drawing></w:r>`
}
const textBox = (inner: string) =>
  `<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:inline><wp:extent cx="1828800" cy="914400"/><wp:docPr id="2" name="tb"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:spPr><a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="FFF2CC"/></a:solidFill><a:ln w="12700"><a:solidFill><a:schemeClr val="accent2"/></a:solidFill></a:ln></wps:spPr><wps:txbx><w:txbxContent>${inner}</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice><mc:Fallback><w:pict><v:shape style="width:144pt;height:72pt"><v:textbox><w:txbxContent>${p(r('備援文字方塊'))}</w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r>`

const tc = (text: string, tcPr = '') => `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/>${tcPr}</w:tcPr>${p(r(text))}</w:tc>`
const table = `<w:tbl><w:tblPr><w:tblStyle w:val="Grid"/><w:tblW w:w="6000" w:type="dxa"/><w:tblLook w:firstRow="1" w:noHBand="1" w:noVBand="1"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>
<w:tr>${tc('表頭甲')}${tc('表頭乙')}${tc('表頭丙')}</w:tr>
<w:tr>${tc('跨兩列', '<w:vMerge w:val="restart"/><w:vAlign w:val="center"/>')}<w:tc><w:tcPr><w:gridSpan w:val="2"/><w:shd w:val="clear" w:fill="E2EFDA"/></w:tcPr>${p(r('跨兩欄'))}</w:tc></w:tr>
<w:tr>${tc('', '<w:vMerge/>')}${tc('乙3', '<w:tcBorders><w:bottom w:val="double" w:sz="6" w:color="FF0000"/></w:tcBorders>')}<w:tc><w:tcPr/><w:tbl><w:tblGrid><w:gridCol w:w="900"/></w:tblGrid><w:tr>${tc('巢狀表格')}</w:tr></w:tbl>${p('')}</w:tc></w:tr>
</w:tbl>`

const field = (instr: string, result: string) =>
  `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve">${instr}</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${result}<w:r><w:fldChar w:fldCharType="end"/></w:r>`

function sampleDocx(overrides: Record<string, string> = {}): Buffer {
  const body = [
    p(r('文件標題'), '<w:pStyle w:val="H1"/>'),
    p(`<w:bookmarkStart w:id="0" w:name="_Toc1"/>${r('第二層標題')}<w:bookmarkEnd w:id="0"/>`, '<w:pStyle w:val="H2"/>'),
    p(r('迴圈樣式也能畫'), '<w:pStyle w:val="Loop1"/>'),
    p(
      r('一般 ') + `<w:r><w:rPr><w:rStyle w:val="Strong"/></w:rPr><w:t>強調</w:t></w:r>` + r(' 主題色', '<w:color w:val="000000" w:themeColor="accent2"/>') + r('螢光', '<w:highlight w:val="yellow"/>') +
        r('底紋', '<w:color w:val="auto"/><w:shd w:val="clear" w:color="auto" w:fill="1F3864"/>') + r('雙刪', '<w:dstrike/>') + r('上標', '<w:vertAlign w:val="superscript"/>') + r('small', '<w:smallCaps/>') + r('寬', '<w:spacing w:val="40"/>') +
        r('東亞字型', '<w:rFonts w:ascii="Arial" w:eastAsia="微軟正黑體"/><w:sz w:val="32"/>') + `<w:r><w:sym w:font="Wingdings" w:char="F0FC"/></w:r>`,
      '<w:jc w:val="both"/><w:ind w:firstLine="480"/><w:spacing w:before="120" w:after="240" w:line="360" w:lineRule="auto"/>'
    ),
    p(r('分散對齊加框線'), '<w:jc w:val="distribute"/><w:pBdr><w:top w:val="single" w:sz="8" w:space="1" w:color="00B050"/><w:bottom w:val="single" w:sz="8" w:space="1" w:color="00B050"/></w:pBdr><w:shd w:val="clear" w:fill="F2F2F2"/>'),
    numbered(1, 0, '第一項'),
    numbered(1, 1, '子項甲'),
    numbered(1, 1, '子項乙'),
    numbered(1, 0, '第二項'),
    numbered(1, 1, '重新起算'),
    numbered(2, 0, '圓點項目'),
    numbered(4, 0, '另一清單從一開始'),
    numbered(3, 0, '中文編號'),
    numbered(3, 0, '中文編號二'),
    numbered(3, 1, '羅馬'),
    table,
    p(r('連到外面', '<w:rStyle w:val="Hyperlink"/>').replace('<w:r>', '<w:hyperlink r:id="rLink"><w:r>') + '</w:hyperlink>' + `<w:hyperlink w:anchor="_Toc1">${r('跳到標題')}</w:hyperlink>`),
    p(field(' TOC \\o "1-3" \\h ', r('目錄項') + '<w:r><w:tab/></w:r>' + r('3')), '<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9350"/></w:tabs>'),
    p(field(' HYPERLINK "https://field.example/x" ', r('欄位連結'))),
    p(r('插入的字') .replace('<w:r>', '<w:ins w:id="1" w:author="a"><w:r>') + '</w:ins>' + `<w:del w:id="2" w:author="a"><w:r><w:delText>刪掉的字</w:delText></w:r></w:del>` + r('隱藏', '<w:vanish/>')),
    p(r('註腳在此') + `<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="1"/></w:r>` + r('附註') + `<w:r><w:endnoteReference w:id="1"/></w:r>`),
    p(picture('rImg', 1905000, 952500) + picture('rImg', 952500, 952500, true) + picture('rEmf', 952500, 952500)),
    p(textBox(p(r('方塊裡的字')))),
    p(`<m:oMath><m:f><m:num><m:r><m:t>a</m:t></m:r></m:num><m:den><m:r><m:t>b</m:t></m:r></m:den></m:f><m:sSup><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup></m:oMath>`),
    p(r('右到左', '<w:rtl/>'), '<w:bidi/>'),
    p(r('前') + '<w:r><w:br w:type="page"/></w:r>' + r('後')),
    p(r('&lt;script&gt;alert(1)&lt;/script&gt; &amp; "q"', `<w:rFonts w:ascii="Arial';}body{x:url(evil)" w:eastAsia="x&lt;/style&gt;"/><w:color w:val="red;background:url(x)"/>`)),
    p(r('第一節結束'), '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr>'),
    p(r('橫向一節')),
    `<w:sectPr><w:headerReference w:type="default" r:id="rHdr"/><w:footerReference w:type="default" r:id="rFtr"/><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/><w:pgMar w:top="1000" w:right="1000" w:bottom="1000" w:left="1000" w:header="500" w:footer="500" w:gutter="0"/></w:sectPr>`
  ].join('')
  return zip({
    '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>',
    '_rels/.rels': rels([['r1', 'officeDocument', 'word/document.xml']]),
    'docProps/core.xml': '<cp:coreProperties xmlns:cp="x" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>範例文件</dc:title></cp:coreProperties>',
    'word/document.xml': `<w:document ${W}><w:body>${body}</w:body></w:document>`,
    'word/_rels/document.xml.rels': rels([
      ['rSt', 'styles', 'styles.xml'], ['rNum', 'numbering', 'numbering.xml'], ['rTh', 'theme', 'theme/theme1.xml'], ['rFn', 'footnotes', 'footnotes.xml'], ['rEn', 'endnotes', 'endnotes.xml'],
      ['rHdr', 'header', 'header1.xml'], ['rFtr', 'footer', 'footer1.xml'], ['rImg', 'image', 'media/image1.png'], ['rEmf', 'image', 'media/image2.emf'],
      ['rLink', 'hyperlink', 'javascript:alert(1)&quot; onmouseover=&quot;x', true]
    ]),
    'word/styles.xml': styles,
    'word/numbering.xml': numbering,
    'word/theme/theme1.xml': theme,
    'word/footnotes.xml': `<w:footnotes ${W}><w:footnote w:type="separator" w:id="-1">${p('<w:r><w:separator/></w:r>')}</w:footnote><w:footnote w:id="1">${p('<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r>' + r(' 註腳內容'))}</w:footnote></w:footnotes>`,
    'word/endnotes.xml': `<w:endnotes ${W}><w:endnote w:id="1">${p('<w:r><w:endnoteRef/></w:r>' + r(' 章節附註內容'))}</w:endnote></w:endnotes>`,
    'word/header1.xml': `<w:hdr ${W}>${p(r('頁首文字') + picture('rLogo', 952500, 476250))}</w:hdr>`,
    'word/_rels/header1.xml.rels': rels([['rLogo', 'image', 'media/logo.png']]),
    'word/footer1.xml': `<w:ftr ${W}>${p(r('頁尾 ') + field(' PAGE ', r('1')))}</w:ftr>`,
    'word/media/image1.png': 'png-bytes',
    'word/media/logo.png': 'png-bytes',
    'word/media/image2.emf': 'emf',
    ...overrides
  })
}

/** The CSS a page gives the element whose class list precedes `text`. */
function styleBefore(html: string, text: string): string {
  const at = html.indexOf(text)
  assert.ok(at >= 0, `missing ${text}`)
  const open = html.lastIndexOf('class="', at)
  const classes = html.slice(open + 7, html.indexOf('"', open + 7)).split(' ')
  return classes
    .map((name) => new RegExp(`\\.${name}\\{([^}]*)\\}`).exec(html)?.[1] ?? '')
    .join(';')
}

describe('reading Word files', () => {
  const html = renderDocxPage(sampleDocx(), '報告 (v2).docx', '範例')

  test('styles inherit through basedOn chains, theme fonts and colours, and docDefaults', () => {
    const h1 = styleBefore(html, '>文件標題<')
    assert.match(h1, /font-family:'Georgia','標楷體'/)
    assert.match(h1, /font-weight:700/)
    assert.match(h1, /font-size:20pt/)
    // accent1 4472C4 shaded by BF.
    assert.match(h1, /color:rgb\(4\d,8\d,1[45]\d\)/)
    const h2 = styleBefore(html, '>第二層標題<')
    assert.match(h2, /font-style:italic/)
    assert.match(h2, /font-weight:700/)
    assert.match(h2, /font-size:14pt/)
    assert.match(html, /role="heading" aria-level="1"[^>]*>.*文件標題/)
    // docDefaults: theme minor fonts (Calibri + the Hant script font), 11pt, the default paragraph style's colour.
    const body = styleBefore(html, '>一般 <')
    assert.match(body, /font-family:'Calibri','新細明體'/)
    assert.match(body, /font-size:11pt/)
    assert.match(body, /color:rgb\(51,51,51\)/)
    assert.match(html, />迴圈樣式也能畫</)
  })

  test('run formatting: character styles, theme colours, highlight, shading, strikes, sizes and East Asian fonts', () => {
    assert.match(styleBefore(html, '>強調<'), /font-weight:700.*color:rgb\(192,0,0\)/)
    assert.match(styleBefore(html, '> 主題色<'), /color:rgb\(237,125,49\)/)
    assert.match(styleBefore(html, '>螢光<'), /background-color:rgb\(255,255,0\)/)
    // Dark shading with automatic text colour turns the text white, as Word does.
    assert.match(styleBefore(html, '>底紋<'), /background-color:rgb\(31,56,100\);color:rgb\(255,255,255\)/)
    assert.match(styleBefore(html, '>雙刪<'), /text-decoration-line:line-through;text-decoration-style:double/)
    assert.match(styleBefore(html, '>上標<'), /vertical-align:super/)
    assert.match(styleBefore(html, '>small<'), /font-variant:small-caps/)
    assert.match(styleBefore(html, '>寬<'), /letter-spacing:2pt/)
    assert.match(styleBefore(html, '>東亞字型<'), /font-family:'Arial','微軟正黑體'.*font-size:16pt/)
    assert.match(html, />✓</)
  })

  test('paragraph formatting: justification, indents, spacing, borders and shading', () => {
    const block = /<div class="dx-p (s\d+)"><span class="s\d+">一般 /.exec(html)
    assert.ok(block)
    const css = new RegExp(`\\.${block[1]}\\{([^}]*)\\}`).exec(html)![1]
    assert.match(css, /text-align:justify/)
    assert.match(css, /text-indent:24pt/)
    assert.match(css, /line-height:1\.8/)
    const boxed = /<div class="dx-p (s\d+)"><span class="s\d+">分散對齊加框線/.exec(html)!
    const boxCss = new RegExp(`\\.${boxed[1]}\\{([^}]*)\\}`).exec(html)![1]
    assert.match(boxCss, /text-align-last:justify/)
    assert.match(boxCss, /border-top:1pt solid rgb\(0,176,80\)/)
    assert.match(boxCss, /background-color:rgb\(242,242,242\)/)
  })

  test('numbering: multi-level counters restart, bullets map from Symbol, Chinese counting, overrides restart', () => {
    const labels = [...html.matchAll(/class="dx-num[^"]*">([^<]*)<\/span>([^<]*<span[^>]*>)?([^<]*)/g)].map((m) => `${m[1].trim()}${m[3]}`)
    assert.deepEqual(labels, ['1.第一項', '1.a)子項甲', '1.b)子項乙', '2.第二項', '2.a)重新起算', '•圓點項目', '1.另一清單從一開始', '一、中文編號', '二、中文編號二', '(IV)羅馬'])
  })

  test('tables: grid widths, gridSpan, vMerge rowspan, table style first row, borders, shading, nested tables', () => {
    assert.match(html, /<td rowspan="2" class="(s\d+)"><div class="dx-cell"><div[^>]*><span[^>]*>跨兩列/)
    assert.match(html, /<td colspan="2" class="s\d+"><div class="dx-cell"><div[^>]*><span[^>]*>跨兩欄/)
    assert.equal(html.match(/<td/g)!.length, 3 + 2 + 2 + 1)
    assert.match(styleBefore(html, '>表頭甲<'), /font-weight:700.*color:rgb\(255,255,255\)/)
    const header = /<td class="(s\d+)"><div class="dx-cell"><div[^>]*><span[^>]*>表頭甲/.exec(html)!
    assert.match(new RegExp(`\\.${header[1]}\\{([^}]*)\\}`).exec(html)![1], /background-color:rgb\(68,114,196\);?/)
    const merged = /<td rowspan="2" class="(s\d+)"/.exec(html)!
    assert.match(new RegExp(`\\.${merged[1]}\\{([^}]*)\\}`).exec(html)![1], /border-top:0\.5pt solid rgb\(0,0,0\).*vertical-align:middle/)
    const red = /<td class="(s\d+)"><div class="dx-cell"><div[^>]*><span[^>]*>乙3/.exec(html)!
    assert.match(new RegExp(`\\.${red[1]}\\{([^}]*)\\}`).exec(html)![1], /border-bottom:2\.25pt double rgb\(255,0,0\)/)
    assert.match(html, /<colgroup><col style="width:33\.333%">/)
    assert.match(html, /<table class="dx-tbl"[\s\S]*<table class="dx-tbl"[\s\S]*巢狀表格/)
  })

  test('links, bookmarks, fields, tracked changes, hidden text', () => {
    // External targets are only ever a title, never an href.
    assert.match(html, /<span class="dx-lnk" title="javascript:alert\(1\)&#34; onmouseover=&#34;x"><span[^>]*>連到外面<\/span><\/span>/)
    assert.doesNotMatch(html, /href="javascript/)
    assert.match(html, /<span id="bm-_Toc1"><\/span>/)
    assert.match(html, /<a class="dx-lnk" href="#bm-_Toc1">/)
    // Field codes never show; their cached results do. A HYPERLINK field result is a link title.
    assert.doesNotMatch(html, /TOC \\o|HYPERLINK|PAGE /)
    assert.match(html, /目錄項[\s\S]*class="dx-fill dx-ld-dot"[\s\S]*>3</)
    assert.match(html, /title="https:\/\/field\.example\/x"><span[^>]*>欄位連結/)
    assert.match(html, />插入的字</)
    assert.doesNotMatch(html, /刪掉的字|隱藏/)
  })

  test('pictures, text boxes (Choice over Fallback), equations, notes, page breaks, header and footer, sections', () => {
    assert.match(html, /<img class="dx-img" src="~office\/%E5%A0%B1%E5%91%8A%20%28v2%29\.docx\/word\/media\/image1\.png" alt="圖 &#60;一&#62;" style="width:150pt;max-width:100%;aspect-ratio:150\/75">/)
    assert.match(html, /<span class="dx-float" style="float:right[^"]*"><img/)
    assert.match(html, /class="dx-unsup"[^>]*>EMF 圖片\(網頁版未支援,請下載原始檔查看\)/)
    assert.match(html, /class="dx-txbx" style="[^"]*background:rgb\(255,242,204\);border:1pt solid rgb\(237,125,49\)">[\s\S]*方塊裡的字/)
    assert.doesNotMatch(html, /備援文字方塊/)
    assert.match(html, /<span class="m-frac"><span>a<\/span><span>b<\/span><\/span>x<sup>2<\/sup>/)
    assert.match(html, /<a class="dx-fnref" href="#fn-1" id="fnref-1">1<\/a>/)
    assert.match(html, /<a class="dx-fnref" href="#en-1" id="enref-1">i<\/a>/)
    assert.match(html, /<div class="dx-note" id="fn-1">[\s\S]*href="#fnref-1">1<\/a>[\s\S]*註腳內容/)
    assert.match(html, /<div class="dx-note" id="en-1">[\s\S]*章節附註內容/)
    assert.match(html, />前<\/span><span class="dx-pgbr"/)
    assert.match(html, /direction:rtl/)
    // Header (with its own picture relationship) on top, footer with the cached PAGE result at the bottom.
    assert.match(html, /<div class="dx-header"[^>]*>[\s\S]*頁首文字[\s\S]*word\/media\/logo\.png/)
    assert.match(html, /<div class="dx-footer"[^>]*>[\s\S]*頁尾 [\s\S]*>1</)
    // Two sheets: A4 portrait, then landscape.
    const sheets = [...html.matchAll(/<div class="dx-sheet[^"]*" style="--pw:([\d.]+)pt/g)].map((m) => Number(m[1]))
    assert.deepEqual(sheets, [595.3, 841.9])
    assert.equal(html.indexOf('class="dx-header"') < html.indexOf('文件標題'), true)
    assert.equal(html.indexOf('class="dx-footer"') > html.indexOf('橫向一節'), true)
  })

  test('hostile text, fonts and colours never escape into markup or CSS', () => {
    assert.match(html, /&#60;script&#62;alert\(1\)&#60;\/script&#62; &#38; &#34;q&#34;/)
    assert.doesNotMatch(html, /<script>alert/)
    assert.doesNotMatch(html, /url\(evil\)|url\(x\)|x<\/style>/)
    assert.match(html, /'Arialbodyxurlevil'/)
    assert.equal(html.match(/<\/style>/g)!.length, 1)
  })

  test('control characters from the file (one marks tabs internally) cannot split markup', () => {
    const doc = sampleDocx({ 'word/document.xml': `<w:document ${W}><w:body>${p(picture('rImg', 10, 10).replace('圖 &lt;一&gt;', 'a&#1;b&#1;c') + r('x&#1;y'))}</w:body></w:document>` })
    const page = renderDocxPage(doc, 'x.docx', 'x')
    assert.match(page, /alt="abc"/)
    assert.match(page, />xy</)
    assert.doesNotMatch(page, /\u0001/)
  })

  test('the page is an HTML document with the bridge, a viewport and explicit light background', () => {
    assert.match(html, /^<!doctype html><html lang="zh-Hant"><head><script src="[^"]+"><\/script><meta charset="utf-8"><meta name="viewport"/)
    assert.match(html, /color-scheme:light/)
    assert.match(html, /<title>範例<\/title>/)
  })

  test('search text covers body, header, footer, notes, text boxes and table cells, not deleted or field code text', () => {
    const text = docxText(sampleDocx())
    for (const word of ['範例文件', '文件標題', '頁首文字', '頁尾', '註腳內容', '章節附註內容', '方塊裡的字', '巢狀表格', '欄位連結', '插入的字']) assert.ok(text.includes(word), word)
    assert.ok(!text.includes('刪掉的字'))
    assert.ok(!text.includes('TOC'))
  })
})

describe('Word files that are partly or wholly broken', () => {
  test('a broken styles part, numbering part or header leaves a notice; the body still draws', () => {
    const doc = readDocx(sampleDocx({ 'word/styles.xml': '<w:styles><w:style>', 'word/header1.xml': '<<<', 'word/numbering.xml': '<!DOCTYPE x><x/>' }), '')
    assert.match(doc.html, /class="dx-notice">樣式 \(styles\.xml\)無法讀取/)
    assert.match(doc.html, /class="dx-notice">編號 \(numbering\.xml\)無法讀取/)
    assert.match(doc.html, /class="dx-notice">頁首無法讀取/)
    assert.match(doc.html, />文件標題</)
    assert.match(doc.html, />巢狀表格</)
  })

  test('a malformed table (an orphan vertical merge, a picture whose relationship is missing) still draws, the picture labelled', () => {
    const broken = sampleDocx({
      'word/document.xml': `<w:document ${W}><w:body>${p(r('之前'))}<w:tbl><w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr>${p(picture('nope', 10, 10))}</w:tc></w:tr></w:tbl>${p(r('之後'))}</w:body></w:document>`
    })
    const html = renderDocxPage(broken, 'b.docx', 'b')
    assert.match(html, />之前<[\s\S]*class="dx-unsup"[^>]*>圖片[\s\S]*>之後</)
  })

  test('a file that is not a zip gets a page that says why, and a throw for the error log', () => {
    const html = renderDocxPage(Buffer.from('this is not a word file'), 'x.docx', 'x')
    assert.match(html, /這份 Word 檔無法顯示/)
    assert.match(html, /檔案不是有效的 \.docx/)
    assert.match(html, /<script>throw new Error\("Word: [^"]+"\)<\/script>/)
    const ole = Buffer.alloc(64)
    ole.writeUInt32BE(0xd0cf11e0, 0)
    assert.match(renderDocxPage(ole, 'x.docx', 'x'), /舊版 Word \(\.doc\) 或受密碼保護/)
    assert.match(renderDocxPage(zip({ 'a.txt': 'x' }), 'x.docx', 'x'), /找不到 word\/document\.xml/)
  })

  test('huge documents are cut and say so', () => {
    const many = Array.from({ length: 20_050 }, (_, at) => p(r(`段${at}`))).join('')
    const html = renderDocxPage(zip({ 'word/document.xml': `<w:document ${W}><w:body>${many}</w:body></w:document>` }), 'x.docx', 'x')
    assert.match(html, /文件太長:只顯示前 20000 個段落與表格/)
    assert.match(html, />段19999</)
    assert.doesNotMatch(html, />段20000</)
  })
})

describe('serving Word files', () => {
  let dir: string
  let store: ArtifactStore
  let content: ReturnType<typeof createContentServer>
  before(() => {
    dir = mkdtempSync(join(tmpdir(), 'docx-test-'))
    store = new ArtifactStore(dir)
    const config = { dataDir: dir, appPort: 0, contentPort: 0, appOrigin: 'http://app.test', contentOrigin: 'http://content.test', tokens: ['t'] }
    content = createContentServer(store, config)
  })
  after(() => {
    store.close()
    rmSync(dir, { recursive: true, force: true })
  })

  test('a published .docx renders as a page; its pictures load through the media route; it is searchable', async () => {
    const { id, entry } = store.publish({ title: '文件', files: { '季報 (final).docx': { content: sampleDocx() } } })
    assert.equal(entry, '季報 (final).docx')
    const response = await content.request(`http://content.test/a/${id}/v/1/`)
    assert.equal(response.status, 200)
    assert.equal(response.headers.get('cache-control'), 'no-cache')
    const html = await response.text()
    assert.match(html, /<body class="docx">/)
    const src = /<img class="dx-img" src="([^"]+image1\.png)"/.exec(html)![1]
    assert.match(src, /^~office\/%E5%AD%A3%E5%A0%B1%20%28final%29\.docx\/word\/media\/image1\.png$/)
    const picture = await content.request(`http://content.test/a/${id}/v/1/${src}`)
    assert.equal(picture.status, 200)
    assert.equal(picture.headers.get('content-type'), 'image/png')
    for (const path of ['~office/季報 (final).docx/word/document.xml', '~office/季報 (final).docx/word/media/image2.emf']) {
      assert.equal((await content.request(`http://content.test/a/${id}/v/1/${encodeURI(path)}`)).status, 404, path)
    }
    assert.deepEqual(store.list({ query: '方塊裡的字' }).artifacts.map((a) => a.id), [id])
  })

  test('a .docx that is not readable still publishes and its page says why', async () => {
    const { id } = store.publish({ title: 'Broken', files: { 'broken.docx': { content: Buffer.from('PK nope') } } })
    const html = await (await content.request(`http://content.test/a/${id}/v/1/`)).text()
    assert.match(html, /這份 Word 檔無法顯示/)
    assert.match(html, /<script>throw new Error\("Word: [^"]+"\)<\/script>/)
  })
})

describe('crafted or very large Word files (stress round 1)', () => {
  const docWith = (body: string, parts: Record<string, string> = {}, relList: [string, string, string][] = []) =>
    zip({
      'word/document.xml': `<w:document ${W}><w:body>${body}</w:body></w:document>`,
      'word/_rels/document.xml.rels': rels(relList),
      ...parts
    })

  test('a page break inside a list item whose numbering comes from its style continues the item, unnumbered', () => {
    const xml = `<w:numbering ${W}><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="5"><w:abstractNumId w:val="0"/></w:num></w:numbering>`
    const styles = `<w:styles ${W}><w:style w:type="paragraph" w:styleId="ListNumber"><w:name w:val="List Number"/><w:pPr><w:numPr><w:numId w:val="5"/></w:numPr></w:pPr></w:style></w:styles>`
    const item = (body: string) => `<w:p><w:pPr><w:pStyle w:val="ListNumber"/></w:pPr>${body}</w:p>`
    const body = item(r('第一項') + '<w:r><w:br w:type="page"/></w:r>' + r('第一項續')) + item(r('第二項'))
    const doc = readDocx(docWith(body, { 'word/numbering.xml': xml, 'word/styles.xml': styles }, [['rNum', 'numbering', 'numbering.xml'], ['rSt', 'styles', 'styles.xml']]), '')
    const labels = [...doc.html.matchAll(/>(\d+)\.</g)].map((m) => m[1])
    assert.deepEqual(labels, ['1', '2'])
    assert.match(doc.html, /第一項續[\s\S]*>2\.<[\s\S]*第二項/)
  })

  test('a tracked deleted page break does not cut the page; only the live one does', () => {
    const body = p(r('甲甲') + '<w:del w:id="1" w:author="a"><w:r><w:br w:type="page"/></w:r></w:del>' + r('乙乙') + '<w:r><w:br w:type="page"/></w:r>' + r('丙丙'))
    const html = readDocx(docWith(body), '').html
    // 甲 and 乙 stay in one paragraph on the first page; the live break starts the page with 丙.
    assert.match(html, /甲甲<\/span><span[^>]*>乙乙/)
    assert.match(html, /乙乙[\s\S]*class="dx-pgbr"[\s\S]*丙丙/)

    // Only deleted text after a live break: nothing on the next page, so no empty continuation.
    const trailing = readDocx(docWith(p(r('丁丁') + '<w:r><w:br w:type="page"/></w:r><w:del w:id="2" w:author="a"><w:r><w:delText>刪掉的</w:delText></w:r></w:del>') + p(r('戊戊'))), '').html
    assert.doesNotMatch(trailing, /<div class="dx-p[^"]*"><br><\/div>/)
  })

  test('F3: a huge list start stays a short label, not megabytes of letters', () => {
    const lvl = (fmt: string) => `<w:lvl w:ilvl="0"><w:start w:val="300000000"/><w:numFmt w:val="${fmt}"/><w:lvlText w:val="%1."/></w:lvl>`
    const xml = `<w:numbering ${W}><w:abstractNum w:abstractNumId="0">${lvl('upperLetter')}</w:abstractNum><w:abstractNum w:abstractNumId="1">${lvl('lowerLetter')}</w:abstractNum>
<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="10000000000"/></w:lvlOverride></w:num></w:numbering>`
    const body = Array.from({ length: 20 }, (_, at) => numbered(at % 2 ? 2 : 1, 0, `項${at}`)).join('')
    const started = Date.now()
    const doc = readDocx(docWith(body, { 'word/numbering.xml': xml }, [['rNum', 'numbering', 'numbering.xml']]), '')
    assert.ok(doc.html.length < 200_000, `html is ${doc.html.length} bytes`)
    assert.ok(Date.now() - started < 2000)
    assert.match(doc.html, />項19</)
    assert.match(doc.html, />32767\.</, 'the start is clamped to Word\'s largest, drawn as a number')
  })

  test('F19: a table style row region keeps its own bottom rule when insideH is nil', () => {
    const style = `<w:styles ${W}><w:style w:type="table" w:styleId="ML2"><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="8" w:color="000000"/><w:bottom w:val="single" w:sz="8" w:color="000000"/></w:tblBorders></w:tblPr>
<w:tblStylePr w:type="firstRow"><w:tcPr><w:tcBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="single" w:sz="24" w:color="4F81BD"/><w:right w:val="nil"/><w:insideH w:val="nil"/><w:insideV w:val="nil"/></w:tcBorders></w:tcPr></w:tblStylePr>
<w:tblStylePr w:type="firstCol"><w:tcPr><w:tcBorders><w:left w:val="single" w:sz="12" w:color="00FF00"/><w:right w:val="single" w:sz="12" w:color="00FF00"/><w:insideV w:val="nil"/></w:tcBorders></w:tcPr></w:tblStylePr></w:style></w:styles>`
    const table = `<w:tbl><w:tblPr><w:tblStyle w:val="ML2"/><w:tblLook w:firstRow="1" w:firstColumn="1" w:noHBand="1" w:noVBand="1"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>
<w:tr>${tc('表頭一')}${tc('表頭二')}</w:tr><w:tr>${tc('首欄')}${tc('內容')}</w:tr><w:tr>${tc('末首欄')}${tc('末列')}</w:tr></w:tbl>`
    const doc = readDocx(docWith(table, { 'word/styles.xml': style }, [['rSt', 'styles', 'styles.xml']]), '')
    const css = doc.html + `<style>${doc.css}</style>`
    const cell = (text: string) => {
      const at = css.lastIndexOf('<td', css.indexOf(`>${text}<`))
      const name = /class="([^"]*)"/.exec(css.slice(at))![1]
      return new RegExp(`\\.${name}\\{([^}]*)\\}`).exec(css)![1]
    }
    assert.match(cell('表頭二'), /border-bottom:(?!none)[^;]*3pt/)
    assert.match(cell('首欄'), /border-right:(?!none)[^;]*1\.5pt/)
  })

  test('a row band two rows tall keeps its inside line between its own rows, not at its edges', () => {
    const style = `<w:styles ${W}><w:style w:type="table" w:styleId="B2"><w:tblPr><w:tblStyleRowBandSize w:val="2"/><w:tblStyleColBandSize w:val="2"/></w:tblPr>
<w:tblStylePr w:type="band1Horz"><w:tcPr><w:tcBorders><w:insideH w:val="single" w:sz="8" w:color="FF0000"/></w:tcBorders></w:tcPr></w:tblStylePr>
<w:tblStylePr w:type="band1Vert"><w:tcPr><w:tcBorders><w:insideV w:val="single" w:sz="8" w:color="00FF00"/></w:tcBorders></w:tcPr></w:tblStylePr></w:style></w:styles>`
    const row = (n: number) => `<w:tr>${['a', 'b', 'c', 'd'].map((c) => tc(`r${n}${c}`)).join('')}</w:tr>`
    const table = `<w:tbl><w:tblPr><w:tblStyle w:val="B2"/><w:tblLook w:firstRow="0" w:firstColumn="0" w:noHBand="0" w:noVBand="0"/></w:tblPr><w:tblGrid>${'<w:gridCol w:w="1500"/>'.repeat(4)}</w:tblGrid>${[1, 2, 3, 4].map(row).join('')}</w:tbl>`
    const doc = readDocx(docWith(table, { 'word/styles.xml': style }, [['rSt', 'styles', 'styles.xml']]), '')
    const css = doc.html + `<style>${doc.css}</style>`
    const cell = (text: string) => {
      const at = css.lastIndexOf('<td', css.indexOf(`>${text}<`))
      const tag = css.slice(at, css.indexOf('>', at))
      const name = /class="([^"]*)"/.exec(tag)?.[1]
      return `${/style="([^"]*)"/.exec(tag)?.[1] ?? ''};${name ? new RegExp(`\\.${name.split(' ').join('|\\.')}\\{([^}]*)\\}`).exec(css)?.[1] ?? '' : ''}`
    }
    // Rows 1–2 are one band: the red line runs between them, not under row 2 (the band's edge).
    assert.match(cell('r1a'), /border-bottom:[^;]*rgb\(255,0,0\)|border-bottom:[^;]*#FF0000/i)
    assert.doesNotMatch(cell('r2a'), /border-bottom:[^;]*(rgb\(255,0,0\)|#FF0000)/i)
    // Columns a–b are one band: the green line runs between them.
    assert.match(cell('r1a'), /border-right:[^;]*(rgb\(0,255,0\)|#00FF00)/i)
    assert.doesNotMatch(cell('r1b'), /border-right:[^;]*(rgb\(0,255,0\)|#00FF00)/i)
  })

  test('F20: a table past the paragraph limit draws whole rows, says how many rows it left out, and the text after it still shows', () => {
    const row = (n: number) => `<w:tr>${Array.from({ length: 10 }, (_, c) => tc(`r${n}c${c}`)).join('')}</w:tr>`
    const grid = `<w:tblGrid>${'<w:gridCol w:w="900"/>'.repeat(10)}</w:tblGrid>`
    const big = (rows: number) => `<w:tbl>${grid}${Array.from({ length: rows }, (_, n) => row(n)).join('')}</w:tbl>${p(r('表格之後'))}`
    const html2000 = readDocx(docWith(big(2000)), '').html
    assert.match(html2000, />r1999c9</)
    assert.match(html2000, />表格之後</)
    assert.doesNotMatch(html2000, /文件太長/)

    const html8000 = readDocx(docWith(big(8000)), '').html
    const rows = html8000.match(/<tr[^>]*>[\s\S]*?<\/tr>/g) ?? []
    const drawn = rows.filter((tr) => !tr.includes('dx-cut'))
    for (const tr of drawn) assert.match(tr, /r\d+c9</, 'a drawn row lost its cells')
    const left = Number(/其餘 (\d+) 列未顯示/.exec(html8000)?.[1])
    assert.equal(drawn.length + left, 8000)
  })

  test('F22: Ming, Mincho and Song fonts draw serif even when the font table calls them modern', () => {
    const fontTable = `<w:fonts ${W}><w:font w:name="細明體"><w:family w:val="modern"/></w:font><w:font w:name="MS Mincho"><w:family w:val="modern"/></w:font><w:font w:name="SimSun"><w:family w:val="auto"/></w:font><w:font w:name="MS Gothic"><w:family w:val="modern"/></w:font></w:fonts>`
    const run = (font: string, text: string) => p(r(text, `<w:rFonts w:ascii="${font}" w:eastAsia="${font}" w:hAnsi="${font}"/>`))
    const body = run('細明體', '明體字') + run('MS Mincho', '明朝字') + run('SimSun', '宋體字') + run('NSimSun', '新宋體字') + run('MS Gothic', '黑體字')
    const doc = readDocx(docWith(body, { 'word/fontTable.xml': fontTable }, [['rF', 'fontTable', 'fontTable.xml']]), '')
    const page = doc.html + `<style>${doc.css}</style>`
    for (const text of ['明體字', '明朝字', '宋體字', '新宋體字']) assert.match(styleBefore(page, text), /font-family:[^;]*,serif(;|$)/, text)
    assert.doesNotMatch(styleBefore(page, '黑體字'), /monospace/)
  })

  test('F23: a list number wider than its hanging indent keeps a gap before the text', () => {
    const xml = `<w:numbering ${W}><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="Chapter %1."/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`
    const body = p(r('正文'), '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:ind w:left="720" w:hanging="360"/>')
    const doc = readDocx(docWith(body, { 'word/numbering.xml': xml }, [['rNum', 'numbering', 'numbering.xml']]), '')
    const css = styleBefore(doc.html + `<style>${doc.css}</style>`, 'Chapter 1.')
    assert.match(css, /min-width:18pt/)
    assert.match(css, /box-sizing:border-box/)
    assert.match(css, /padding-inline-end:/)
  })
})

test('F3: a list override on a far-out level number does not spin through billions of levels', { timeout: 5000 }, () => {
  const xml = `<w:numbering ${W}><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>
<w:num w:numId="1"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="-3000000000"><w:startOverride w:val="5"/></w:lvlOverride></w:num></w:numbering>`
  const doc = readDocx(
    zip({
      'word/document.xml': `<w:document ${W}><w:body>${numbered(1, 0, '項目')}</w:body></w:document>`,
      'word/_rels/document.xml.rels': rels([['rNum', 'numbering', 'numbering.xml']]),
      'word/numbering.xml': xml
    }),
    ''
  )
  assert.match(doc.html, />1\.<[\s\S]*>項目</)
})

describe('any Word file draws faithfully (stress round 2)', () => {
  const docWith = (body: string, parts: Record<string, string> = {}, relList: [string, string, string, boolean?][] = []) =>
    zip({
      'word/document.xml': `<w:document ${W}><w:body>${body}</w:body></w:document>`,
      'word/_rels/document.xml.rels': rels(relList),
      ...parts
    })
  const EMU = 12700
  type AnchorOptions = { behind?: boolean; relH: string; relV: string; offH?: number; offV?: number; alignH?: string; w: number; h: number; content: string; wrap?: string }
  const anchor = (o: AnchorOptions) =>
    `<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="2" behindDoc="${o.behind ? 1 : 0}" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>` +
    `<wp:positionH relativeFrom="${o.relH}">${o.alignH ? `<wp:align>${o.alignH}</wp:align>` : `<wp:posOffset>${Math.round((o.offH ?? 0) * EMU)}</wp:posOffset>`}</wp:positionH>` +
    `<wp:positionV relativeFrom="${o.relV}"><wp:posOffset>${Math.round((o.offV ?? 0) * EMU)}</wp:posOffset></wp:positionV>` +
    `<wp:extent cx="${o.w * EMU}" cy="${o.h * EMU}"/>${o.wrap ?? '<wp:wrapNone/>'}<wp:docPr id="9" name="x"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">${o.content}</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`
  const pic = (rid: string) => `<pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="p"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${rid}"/></pic:blipFill><pic:spPr/></pic:pic>`
  const box = (text: string) => `<wps:wsp><wps:spPr><a:prstGeom prst="rect"/></wps:spPr><wps:txbx><w:txbxContent>${p(r(text))}</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp>`
  const media: [string, string, string][] = [['rBg', 'image', 'media/bg.png']]
  const mediaParts = { 'word/media/bg.png': 'png-bytes' }
  // A US Letter page (the default): 612 × 792 pt, one-inch margins. Positions are in cqw of the sheet's width.
  const fullPage = (extra: Partial<AnchorOptions> = {}) => anchor({ behind: true, relH: 'page', relV: 'page', w: 612, h: 792, content: pic('rBg'), ...extra })

  test('F30: a behind-text picture anchored to the page sits under the cover text, placed on the sheet', () => {
    const body = p(fullPage() + r('CERTIFICATE OF COMPLETION', '<w:sz w:val="72"/>')) + p(r('Awarded to Jane Doe'), '<w:jc w:val="center"/>')
    const html = readDocx(docWith(body, mediaParts, media), 'm/').html
    assert.match(html, /<span class="dx-abs" style="left:0cqw;top:0cqw;width:100cqw;height:129\.412cqw;z-index:-1"><img class="dx-img" src="m\/word\/media\/bg\.png"[^>]*><\/span><span[^>]*>CERTIFICATE OF COMPLETION</)
    assert.doesNotMatch(html, /class="dx-pg"/, 'the first page needs no segment of its own')
  })

  test('F30: a text box in front of text anchored to the margin and column lands where its offsets say', () => {
    const body = p(anchor({ relH: 'column', relV: 'margin', offH: 72, offV: 144, w: 180, h: 72, content: box('浮動方塊') }) + r('正文'))
    const html = readDocx(docWith(body), '').html
    assert.match(html, /<span class="dx-abs" style="left:calc\(var\(--il\) \+ 11\.765cqw\);top:35\.294cqw;width:29\.412cqw;height:11\.765cqw;z-index:1"><span class="dx-txbx"[^>]*>[\s\S]*浮動方塊/)
  })

  test('F30: after a page break, page anchors are placed on that page, which keeps its margin and full height', () => {
    const afterBreak = p(r('第一頁') + '<w:r><w:br w:type="page"/></w:r>') + p(fullPage() + r('第二頁封面'))
    const html = readDocx(docWith(afterBreak, mediaParts, media), '').html
    assert.match(html, /class="dx-pgbr"[\s\S]*?<\/div><div class="dx-pg" style="padding-top:11\.765cqw;min-height:129\.412cqw"><div class="dx-p[^"]*"><span class="dx-abs" style="left:calc\(0cqw - var\(--il\)\);top:0cqw;/)
    assert.match(html, /第二頁封面<\/span><\/div><\/div>/)

    const before = p(r('前一頁')) + p(fullPage() + r('新的一頁'), '<w:pageBreakBefore/>')
    const html2 = readDocx(docWith(before, mediaParts, media), '').html
    assert.match(html2, /<span class="dx-pgbr" role="separator" aria-label="分頁"><\/span><div class="dx-pg" style="padding-top:11\.765cqw;min-height:129\.412cqw"><div class="dx-p/)

    const plain = p(r('甲')) + p(r('乙'), '<w:pageBreakBefore/>')
    assert.doesNotMatch(readDocx(docWith(plain), '').html, /padding-top:11\.765cqw/, 'a page without page anchors keeps the plain layout')
  })

  test('a page break inside a paragraph moves what follows it, page anchors included, to the next page', () => {
    const inline = p(r('第一頁的字') + '<w:r><w:br w:type="page"/></w:r>' + r('第二頁的字') + fullPage())
    const html = readDocx(docWith(inline, mediaParts, media), '').html
    // The drawing after the break lands in the second page's segment, not over the first page's text.
    assert.match(html, /第一頁的字[\s\S]*class="dx-pgbr"[\s\S]*<div class="dx-pg"[^>]*>[\s\S]*第二頁的字/)
    assert.match(html, /<div class="dx-pg"[^>]*>[\s\S]*class="dx-abs"/)
    assert.equal((html.match(/第二頁的字/g) ?? []).length, 1)
  })

  test('F30: a paragraph-anchored picture is pinned to its paragraph; a paragraph holding only an anchor keeps its line', () => {
    const body = p(anchor({ relH: 'column', relV: 'paragraph', offH: 36, offV: 18, w: 72, h: 36, content: box('旁註') }) + r('段落文字'), '<w:ind w:left="720"/>') + p(fullPage())
    const html = readDocx(docWith(body, mediaParts, media), '').html
    assert.match(html, /<div class="dx-p [^"]*"><span class="dx-pin" style="margin-inline-start:calc\(-1 \* 36pt\)"><span class="dx-abs" style="left:5\.882cqw;top:2\.941cqw;width:11\.765cqw;height:5\.882cqw;z-index:1">[\s\S]*旁註[\s\S]*<\/span><\/span><span[^>]*>段落文字</)
    assert.match(html, /z-index:-1"><img[^>]*><\/span><br><\/div>/)
  })

  test('F30: a VML watermark in the header is centred on the margin box, rotated and behind the text', () => {
    const shape = `<w:r><w:pict><v:shape id="PowerPlusWaterMarkObject" type="#_x0000_t136" style="position:absolute;margin-left:0;margin-top:0;width:412.4pt;height:247.45pt;rotation:315;z-index:-251657216;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin" fillcolor="silver" stroked="f"><v:fill opacity=".5"/><v:textpath style="font-family:&quot;Calibri&quot;;font-size:1pt" string="CONFIDENTIAL"/></v:shape></w:pict></w:r>`
    const doc = docWith(p(r('內文')) + `<w:sectPr><w:headerReference w:type="default" r:id="rH"/></w:sectPr>`, { 'word/header1.xml': `<w:hdr ${W}>${p(shape + r('信頭'))}</w:hdr>` }, [['rH', 'header', 'header1.xml']])
    const html = readDocx(doc, '').html
    assert.match(html, /<span class="dx-abs" style="left:calc\(var\(--il\) \+ 4\.542cqw\);top:44\.489cqw;width:67\.386cqw;height:40\.433cqw;z-index:-1;transform:rotate\(315deg\)"><span class="dx-wordart[^"]*" style="[^"]*">CONFIDENTIAL<\/span><\/span>/)
  })

  const altChunkDoc = (target: string, data: string, contentType?: string) =>
    zip({
      '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">${contentType ? `<Override PartName="${target}" ContentType="${contentType}"/>` : ''}</Types>`,
      'word/document.xml': `<w:document ${W}><w:body>${p(r('塊之前'))}<w:altChunk r:id="rA"/>${p(r('塊之後'))}</w:body></w:document>`,
      'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rA" Type="${REL}/aFChunk" Target="${target}"/></Relationships>`,
      [target.replace(/^\//, '')]: data
    })

  test('F31: an HTML altChunk keeps its text and structure through a whitelist; everything active is dropped', () => {
    const chunk = `<!DOCTYPE html><html><head><title>標題不該出現</title><style>p{color:red}</style><script>alert(1)</script></head><body onload="x()">
<h1 style="color:red" onclick="evil()">報告 &amp; 摘要</h1><p class="x">第一段 <b>粗</b><i>斜</i><u>底</u><s>刪</s>H<sub>2</sub>O x<sup>2</sup> &lt;tag&gt; &#x4E2D;&#25991;</p>
<div>區塊一</div><div>區塊二</div>
<ul><li>甲<li>乙<ol><li>巢狀</ol></ul>
<table border=1><thead><tr><th colspan="2">表頭</th></tr></thead><tbody><tr><td rowspan=2 onmouseover="x">左</td><td>右</td></tr></tbody></table>
<blockquote>引言</blockquote><pre><code>code  here</code></pre><table><tr><td>格<b>粗</p>續</td></tr></table>
<a href="javascript:alert(1)" onclick="y">腳本連結</a> <a href="https://example.com/x?a=1&amp;b=2">外部</a>
<img src="javascript:alert(1)"><img src="https://evil.example/t.png"><img src="data:image/png;base64,iVBORw0KGgo=" alt="小圖" onerror="z()" width="120"><img src="data:text/html;base64,PHNjcmlwdD4=">
<iframe src="https://x"></iframe><object data="x"></object><svg><script>alert(2)</script><text>svg文字</text></svg><img src=x onerror=alert(3)//
<form><button>按鈕</button></form><noscript>請開啟</noscript><template><p>模板</p></template><!-- <script>alert(4)</script> -->
<p>未閉合 <b>粗體
</body></html>`
    const data = altChunkDoc('/word/afchunk.htm', chunk)
    const html = readDocx(data, '').html
    assert.doesNotMatch(html, /<style|<script|<iframe|<object|<svg|<form|<input|<template|<button|\son\w+=|javascript:|evil\.example|data:text|alert/)
    assert.doesNotMatch(html, /標題不該出現|svg文字|模板|請開啟|dx-unsup/)
    assert.match(html, />塊之前<[\s\S]*<div class="dx-html[^"]*">[\s\S]*<\/div>[\s\S]*>塊之後</)
    for (const piece of ['<h1>報告 &#38; 摘要</h1>', '<p>第一段 <b>粗</b><i>斜</i><u>底</u><s>刪</s>H<sub>2</sub>O x<sup>2</sup> &#60;tag&#62; 中文</p>', '<th colspan="2">表頭</th>', '<td rowspan="2">左</td>', '<blockquote>引言</blockquote>', '<pre><code>code  here</code></pre>', '<span class="dx-lnk">腳本連結</span>', '<span class="dx-lnk" title="https://example.com/x?a=1&#38;b=2">外部</span>', '<img class="dx-img" src="data:image/png;base64,iVBORw0KGgo=" alt="小圖" width="120">', '<li>甲</li><li>乙<ol><li>巢狀</li></ol></li>', '<td>格<b>粗續</b></td>', '按鈕']) {
      assert.ok(html.includes(piece), `missing ${piece}`)
    }
    assert.doesNotMatch(html, /區塊一區塊二/)
    assert.match(html, /粗體\s*<\/b><\/p><\/div>/)
    assert.match(docxText(data), /報告 & 摘要[\s\S]*右/)
  })

  test('F31: an html-docx-js MHT chunk decodes quoted-printable UTF-8 and inlines its pictures as data URIs', () => {
    const mht = [
      'MIME-Version: 1.0',
      'Content-Type: multipart/related;',
      '    type="text/html";',
      '    boundary="----=mhtDocumentPart"',
      '',
      '',
      '------=mhtDocumentPart',
      'Content-Type: text/html;',
      '    charset="utf-8"',
      'Content-Transfer-Encoding: quoted-printable',
      'Content-Location: file:///C:/fake/document.html',
      '',
      '<html><body><h1 style=3D"color:red">=E6=A8=99=E9=A1=8C</h1><p>long line that is soft=',
      'ly wrapped</p><p><img src=3D"file:///C:/fake/image0.png" width=3D"200"><img src=3D"file:///C:/fake/none.png"></p></body></html>',
      '',
      '------=mhtDocumentPart',
      'Content-Type: image/png',
      'Content-Transfer-Encoding: base64',
      'Content-Location: file:///C:/fake/image0.png',
      '',
      'iVBORw0KGgoAAAANSUhEUg==',
      '',
      '------=mhtDocumentPart--',
      ''
    ].join('\r\n')
    const html = readDocx(altChunkDoc('/word/afchunk.mht', mht, 'message/rfc822'), '').html
    assert.ok(html.includes('<h1>標題</h1>'), html)
    assert.ok(html.includes('long line that is softly wrapped'))
    assert.ok(html.includes('<img class="dx-img" src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUg==" alt="" width="200">'))
    assert.doesNotMatch(html, /file:\/\/|none\.png/)
  })

  test('F31: a chunk that is neither HTML nor MHT stays a labelled box; plain text shows as text', () => {
    assert.match(readDocx(altChunkDoc('/word/a.rtf', '{\\rtf1 hello}', 'application/rtf'), '').html, /dx-unsup[^>]*>嵌入的外部內容/)
    assert.match(readDocx(altChunkDoc('/word/a.txt', '第一行\n<b>第二行</b>', 'text/plain'), '').html, /<div class="dx-html[^"]*"><p>第一行<br>&#60;b&#62;第二行&#60;\/b&#62;<\/p><\/div>/)
  })

  test('F44: a missing picture and an externally linked one are named apart; the box says why', () => {
    const vml = (rid: string) => `<w:r><w:pict><v:shape style="width:72pt;height:72pt"><v:imagedata r:id="${rid}"/></v:shape></w:pict></w:r>`
    const body = p(picture('nope', 10, 10)) + p(picture('rGone', 10, 10)) + p(picture('rExt', 10, 10)) + p(picture('rEmf', 10, 10)) + p(vml('nope')) + p(vml('rExt'))
    const html = readDocx(docWith(body, { 'word/media/b.emf': 'emf' }, [['rGone', 'image', 'media/gone.png'], ['rExt', 'image', 'https://example.com/a.png', true], ['rEmf', 'image', 'media/b.emf']]), '').html
    const boxes = [...html.matchAll(/class="dx-unsup"[^>]*>([^<]*)</g)].map((match) => match[1])
    assert.equal(boxes.length, 6)
    assert.match(boxes[0], /^圖片遺失/)
    assert.match(boxes[1], /^圖片遺失/)
    assert.match(boxes[2], /^外部連結的圖片/)
    assert.match(boxes[3], /^EMF 圖片\(網頁版未支援/)
    assert.match(boxes[4], /^圖片遺失/)
    assert.match(boxes[5], /^外部連結的圖片/)
    for (const label of [boxes[0], boxes[2]]) assert.doesNotMatch(label, /網頁版未支援/)
  })

  test('F44: notices and error boxes speak Traditional Chinese, never a parser\'s English', () => {
    const doc = readDocx(sampleDocx({ 'word/styles.xml': '<w:styles><w:style>', 'word/numbering.xml': '<!DOCTYPE x><x/>' }), '')
    const notices = [...doc.html.matchAll(/class="dx-notice">([^<]*)</g)].map((match) => match[1])
    assert.ok(notices.length >= 2)
    for (const notice of notices) assert.doesNotMatch(notice.replace(/\([^)]*\.xml\)/, '').replace(/&#60;[^&]*&#62;/g, ''), /[A-Za-z]{4,}/, notice)
    const page = renderDocxPage(zip({ 'word/document.xml': `<w:document ${W}><w:body><w:p><w:r><w:t>x</w:r></w:p></w:body></w:document>` }), 'x.docx', 'x')
    const visible = page.replace(/<script>[\s\S]*?<\/script>/g, '').replace(/<style>[\s\S]*?<\/style>/g, '').replace(/<[^>]+>/g, ' ')
    assert.doesNotMatch(visible.replace(/Word|\.docx|x\.docx/g, ''), /[A-Za-z]{4,}/, visible)
  })
  test('N4: a square- or tight-wrapped anchor sits at its paragraph\'s top wherever its run is; a paragraph offset moves it down', () => {
    const square = '<wp:wrapSquare wrapText="bothSides"/>'
    const tight = '<wp:wrapTight wrapText="bothSides"><wp:wrapPolygon edited="0"><wp:start x="0" y="0"/><wp:lineTo x="0" y="21600"/></wp:wrapPolygon></wp:wrapTight>'
    const at = (o: Partial<AnchorOptions>) => anchor({ relH: 'column', relV: 'paragraph', w: 144, h: 100, content: pic('rBg'), ...o })
    const body =
      p(r('左圖的字。') + at({ alignH: 'left', wrap: square })) +
      p(r('右圖的字。') + at({ alignH: 'right', wrap: tight })) +
      p(r('上下的字。') + at({ alignH: 'center', wrap: '<wp:wrapTopAndBottom/>' })) +
      p(r('下移的字。') + at({ alignH: 'left', offV: 18, wrap: square }))
    const html = readDocx(docWith(body, mediaParts, media), '').html
    assert.match(html, /<div class="dx-p [^"]*"><span class="dx-float" style="float:left;margin:0pt 0pt 0pt 0pt"><img[^>]*><\/span><span[^>]*>左圖的字。</)
    assert.match(html, /<div class="dx-p [^"]*"><span class="dx-float" style="float:right;[^"]*"><img[^>]*><\/span><span[^>]*>右圖的字。</)
    assert.match(html, /<div class="dx-p [^"]*"><span class="dx-figure" style="text-align:center"><img[^>]*><\/span><span[^>]*>上下的字。</)
    assert.match(html, /<div class="dx-p [^"]*"><span class="dx-float" style="float:left;margin:18pt 0pt 0pt 0pt"><img[^>]*><\/span><span[^>]*>下移的字。</)
  })

  test('N4: a square-wrapped VML picture of an older file floats to its side at its paragraph\'s top too', () => {
    const vml = (style: string) => `<w:r><w:pict><v:shape style="position:absolute;${style};width:100pt;height:50pt"><v:imagedata r:id="rBg"/><w10:wrap xmlns:w10="urn:schemas-microsoft-com:office:word" type="square"/></v:shape></w:pict></w:r>`
    const body = p(r('舊檔右圖。') + vml('margin-left:340pt;margin-top:12pt')) + p(r('舊檔左圖。') + vml('margin-left:0;margin-top:0;mso-position-horizontal:left'))
    const html = readDocx(docWith(body, mediaParts, media), '').html
    assert.match(html, /<div class="dx-p [^"]*"><span class="dx-float" style="float:right;margin:12pt 9pt 0pt 9pt"><img[^>]*><\/span><span[^>]*>舊檔右圖。</)
    assert.match(html, /<div class="dx-p [^"]*"><span class="dx-float" style="float:left;margin:0pt 9pt 0pt 9pt"><img[^>]*><\/span><span[^>]*>舊檔左圖。</)
  })

  test('a VML picture wrapped above and below keeps its left, centre or right alignment', () => {
    const vml = (align: string) => `<w:r><w:pict><v:shape style="position:absolute;margin-left:0;margin-top:0;mso-position-horizontal:${align};width:100pt;height:50pt"><v:imagedata r:id="rBg"/><w10:wrap xmlns:w10="urn:schemas-microsoft-com:office:word" type="topAndBottom"/></v:shape></w:pict></w:r>`
    const html = readDocx(docWith(p(r('左') + vml('left')) + p(r('中') + vml('center')) + p(r('右') + vml('right')), mediaParts, media), '').html
    assert.deepEqual([...html.matchAll(/class="dx-figure" style="text-align:(\w+)/g)].map((m) => m[1]), ['left', 'center', 'right'])
  })
})

describe('Word drawings: charts and shape geometry', () => {
  const docWith = (body: string, parts: Record<string, string> = {}, relList: [string, string, string][] = []) =>
    zip({
      'word/document.xml': `<w:document ${W}><w:body>${body}</w:body></w:document>`,
      'word/_rels/document.xml.rels': rels(relList),
      'word/theme/theme1.xml': theme,
      ...parts
    })
  const inline = (cx: number, cy: number, graphic: string, uri: string) =>
    p(`<w:r><w:drawing><wp:inline><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="3" name="d"/><a:graphic><a:graphicData uri="${uri}">${graphic}</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`)
  const C = 'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
  const ser = (values: number[]) =>
    `<c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:strRef><c:f>S!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>銷售</c:v></c:pt></c:strCache></c:strRef></c:tx><c:cat><c:strRef><c:f>S!$A$2:$A$4</c:f><c:strCache><c:ptCount val="${values.length}"/>${values.map((_, i) => `<c:pt idx="${i}"><c:v>季${i + 1}</c:v></c:pt>`).join('')}</c:strCache></c:strRef></c:cat><c:val><c:numRef><c:f>S!$B$2:$B$4</c:f><c:numCache><c:ptCount val="${values.length}"/>${values.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('')}</c:numCache></c:numRef></c:val></c:ser>`
  const axes = '<c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>'

  test('a chart in a drawing is drawn as an SVG at its extent, in the theme colours; a missing part is a labelled box', () => {
    const chart = `<c:chartSpace ${C}><c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>季報</a:t></a:r></a:p></c:rich></c:tx></c:title><c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${ser([3, 5, 4])}<c:axId val="1"/><c:axId val="2"/></c:barChart>${axes}</c:plotArea></c:chart></c:chartSpace>`
    const graphic = '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rC"/>'
    const uri = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
    const doc = readDocx(
      docWith(inline(5486400, 3200400, graphic, uri) + inline(914400, 914400, graphic.replace('rC', 'gone'), uri), { 'word/charts/chart1.xml': chart }, [['rC', 'chart', 'charts/chart1.xml'], ['rT', 'theme', 'theme/theme1.xml']]),
      ''
    )
    // 5486400 × 3200400 EMU is 432 × 252 pt.
    assert.match(doc.html, /<span class="dx-chart" style="width:432pt[^"]*aspect-ratio:432\/252"><svg [^>]*viewBox="0 0 432 252"[^>]*aria-label="季報"/)
    assert.equal((doc.html.match(/<rect [^>]*fill="#4472C4"/g) ?? []).length, 3)
    assert.match(doc.html, />季1<\/text>/)
    assert.match(doc.html, /class="dx-unsup"[^>]*>圖表\(檔案裡找不到這張圖表\)</)
    // A file with no theme still colours its series apart, with Office's default accents; the line group of a combo is drawn too.
    const combo = `<c:chartSpace ${C}><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${ser([3, 5, 4])}<c:axId val="1"/><c:axId val="2"/></c:barChart><c:lineChart><c:grouping val="standard"/>${ser([4, 2, 6]).replace(/<c:idx val="0"\/><c:order val="0"\/>/, '<c:idx val="1"/><c:order val="1"/>')}<c:axId val="1"/><c:axId val="2"/></c:lineChart>${axes}</c:plotArea></c:chart></c:chartSpace>`
    const bare = zip({ 'word/document.xml': `<w:document ${W}><w:body>${inline(5486400, 3200400, graphic, uri)}</w:body></w:document>`, 'word/_rels/document.xml.rels': rels([['rC', 'chart', 'charts/chart1.xml']]), 'word/charts/chart1.xml': combo })
    const html = readDocx(bare, '').html
    assert.equal((html.match(/<rect [^>]*fill="#4472C4"/g) ?? []).length, 3)
    assert.match(html, /<path d="M[^"]*" fill="none" stroke="#ED7D31"/)
  })

  test("a non-rectangular shape draws its geometry behind its text, which sits in the shape's text rectangle", () => {
    const shape = (prst: string) =>
      `<wps:wsp><wps:spPr><a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:ln></wps:spPr><wps:txbx><w:txbxContent>${p(r('心裡的字'))}</w:txbxContent></wps:txbx><wps:bodyPr lIns="0" tIns="0" rIns="0" bIns="0"/></wps:wsp>`
    const uri = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape'
    const html = readDocx(docWith(inline(914400, 914400, shape('heart'), uri) + inline(914400, 914400, shape('ellipse'), uri) + inline(914400, 914400, shape('notAShape'), uri)), '').html
    const [heart, ellipse, unknown] = [...html.matchAll(/<span class="dx-txbx" style="([^"]*)">([\s\S]*?)<\/span>/g)].map((m) => ({ style: m[1], body: m[2] }))
    assert.match(heart.body, /^<svg class="dx-geom" viewBox="0 0 72 72"[^>]*><path d="M[^"]*C[^"]*" fill="rgb\(255,0,0\)"\/><path d="M[^"]*C[^"]*" fill="none" stroke="rgb\(0,0,255\)"/)
    assert.doesNotMatch(heart.style, /background:|border:|border-radius/)
    // An ellipse's text rectangle is inset 14.6% (10.54 pt of 72) on every side.
    assert.match(ellipse.style, /padding:10\.5\d*pt 10\.5\d*pt 10\.5\d*pt 10\.5\d*pt/)
    assert.match(unknown.style, /background:rgb\(255,0,0\)/)
    assert.match(html, /未支援的圖形:notAShape/)
  })

  test('N3: a pattern fill draws its hatch and a picture fill its picture, in any geometry; a patterned outline takes its foreground', () => {
    const uri = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape'
    const patt = '<a:pattFill prst="smCheck"><a:fgClr><a:srgbClr val="1F4E79"/></a:fgClr><a:bgClr><a:srgbClr val="FFE699"/></a:bgClr></a:pattFill>'
    const blip = '<a:blipFill><a:blip r:embed="rFill"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>'
    const shape = (prst: string, fill: string, ln = '<a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>') =>
      `<wps:wsp><wps:spPr><a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom>${fill}${ln}</wps:spPr><wps:txbx><w:txbxContent>${p(r('白字', '<w:color w:val="FFFFFF"/>'))}</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp>`
    const body = [shape('ellipse', patt), shape('rect', patt), shape('hexagon', blip), shape('rect', blip), shape('chevron', '<a:noFill/>', `<a:ln w="25400">${patt}</a:ln>`)].map((s) => inline(914400, 914400, s, uri)).join('')
    const html = readDocx(docWith(body, { 'word/media/fill.png': 'png' }, [['rFill', 'image', 'media/fill.png']]), 'm/').html
    const boxes = [...html.matchAll(/<span class="dx-txbx" style="([^"]*)">([\s\S]*?)<\/span>/g)].map((m) => ({ style: m[1], body: m[2] }))
    assert.equal(boxes.length, 5)
    // The ellipse: an 8 × 8 hatch tile of 6 pt, background then foreground pixels, filling the geometry.
    const [, id] = /<pattern id="([^"]+)" patternUnits="userSpaceOnUse" width="6" height="6">/.exec(boxes[0].body) ?? []
    assert.ok(id, boxes[0].body)
    assert.match(boxes[0].body, /<rect width="8" height="8" fill="rgb\(255,230,153\)"\/><path d="M[^"]*" fill="rgb\(31,78,121\)"\/>/)
    assert.match(boxes[0].body, new RegExp(`<path d="M[^"]*" fill="url\\(#${id}\\)"/>`))
    // The rectangle: the same tile as a repeating CSS background.
    assert.match(boxes[1].style, /background:url\('data:image\/svg\+xml,[^']*rgb\(255%2C230%2C153\)[^']*rgb\(31%2C78%2C121\)[^']*'\) 0 0\/6pt 6pt repeat/)
    // The hexagon: the picture stretched over the shape's box, clipped to its geometry.
    assert.match(boxes[2].body, /<pattern id="([^"]+)" patternUnits="userSpaceOnUse" width="72" height="72"><image href="m\/word\/media\/fill\.png" width="72" height="72" preserveAspectRatio="none"\/><\/pattern>/)
    assert.match(boxes[2].body, /<path d="M[^"]*" fill="url\(#[^"]+\)"\/>/)
    assert.match(boxes[3].style, /background:url\('m\/word\/media\/fill\.png'\) 0 0\/100% 100% no-repeat/)
    // A patterned outline is drawn in its foreground colour.
    assert.match(boxes[4].body, /stroke="rgb\(31,78,121\)" stroke-width="2"/)
    // A shape in a group that takes the group's fill draws the group's hatch.
    const group = `<wpg:wgp xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/><a:chOff x="0" y="0"/><a:chExt cx="914400" cy="914400"/></a:xfrm>${patt}</wpg:grpSpPr>${shape('ellipse', '<a:grpFill/>').replace('<wps:spPr>', '<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>')}</wpg:wgp>`
    const grouped = readDocx(docWith(inline(914400, 914400, group, 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup')), '').html
    assert.match(grouped, /<pattern id="[^"]+" patternUnits="userSpaceOnUse" width="6" height="6"><g[^>]*><rect width="8" height="8" fill="rgb\(255,230,153\)"/)
    // Two shapes never share a pattern id.
    const ids = [...html.matchAll(/<pattern id="([^"]+)"/g)].map((m) => m[1])
    assert.equal(new Set(ids).size, ids.length)
  })
})

describe('Word shapes and layout (matrix round)', () => {
  const docOf = (body: string, parts: Record<string, string> = {}, relList: [string, string, string][] = []) =>
    zip({ 'word/document.xml': `<w:document ${W}><w:body>${body}</w:body></w:document>`, 'word/_rels/document.xml.rels': rels(relList), 'word/theme/theme1.xml': theme, ...parts })
  const shapeIn = (xfrm: string, prst: string, bodyPr = '') =>
    p(`<w:r><w:drawing><wp:inline><wp:extent cx="1828800" cy="914400"/><wp:docPr id="3" name="d"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:spPr><a:xfrm${xfrm}><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></wps:spPr><wps:txbx><w:txbxContent>${p(r('箭頭字'))}</w:txbxContent></wps:txbx><wps:bodyPr${bodyPr}/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`)
  const boxes = (html: string) => [...html.matchAll(/<span class="dx-txbx" style="([^"]*)">([\s\S]*?)<\/span>/g)].map((m) => ({ style: m[1], body: m[2] }))

  test('P-1 (Word): a document past the zip limits is called too big, not broken', () => {
    const saved = { ...ZIP_LIMITS }
    ZIP_LIMITS.entries = 1
    try {
      const html = renderDocxPage(docOf(p(r('x'))), 'big.docx', 'big')
      assert.match(html, /檔案太大/)
      assert.doesNotMatch(html, /損毀|不是有效/)
    } finally {
      Object.assign(ZIP_LIMITS, saved)
    }
  })

  test('W-1: a flipped shape mirrors its geometry, not its text; a rotation and vert270 share one transform', () => {
    const html = readDocx(docOf(shapeIn(' flipH="1"', 'rightArrow') + shapeIn(' flipV="1"', 'triangle') + shapeIn(' rot="5400000"', 'rect', ' vert="vert270"') + shapeIn(' rot="2700000" flipH="1"', 'rightArrow', ' vert="vert270"')), '').html
    const [h, v, rv, all] = boxes(html)
    assert.match(h.body, /^<svg class="dx-geom"[^>]*style="transform:scaleX\(-1\)"/)
    assert.doesNotMatch(h.style, /transform/)
    assert.match(v.body, /^<svg class="dx-geom"[^>]*style="transform:scaleY\(-1\)"/)
    for (const box of [rv, all]) assert.equal(box.style.match(/transform:/g)?.length, 1, box.style)
    assert.match(rv.style, /transform:rotate\(90deg\) rotate\(180deg\)/)
    assert.match(all.style, /transform:rotate\(45deg\) rotate\(180deg\)/)
    // The box turns over for vert270; the geometry turns back, so the arrow still points the way the file says (flipped: left).
    assert.match(all.body, /^<svg class="dx-geom"[^>]*style="transform:rotate\(180deg\) scaleX\(-1\)"/)
  })

  test('W-1: a flipped group mirrors where its members sit and how each is drawn, its text still readable', () => {
    const member = (x: number, prst: string) =>
      `<wps:wsp><wps:spPr><a:xfrm><a:off x="${x}" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></wps:spPr><wps:txbx><w:txbxContent>${p(r(prst))}</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp>`
    const group = `<wpg:wgp xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"><wpg:grpSpPr><a:xfrm flipH="1"><a:off x="0" y="0"/><a:ext cx="3657600" cy="914400"/><a:chOff x="0" y="0"/><a:chExt cx="3657600" cy="914400"/></a:xfrm></wpg:grpSpPr>${member(0, 'rightArrow')}${member(2743200, 'triangle')}</wpg:wgp>`
    const html = readDocx(docOf(p(`<w:r><w:drawing><wp:inline><wp:extent cx="3657600" cy="914400"/><wp:docPr id="3" name="g"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">${group}</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`)), '').html
    const items = [...html.matchAll(/<span class="dx-gi" style="left:([\d.]+)%[^"]*">([\s\S]*?)<\/span><\/span>/g)].map((m) => ({ left: Number(m[1]), body: m[2] }))
    // The arrow, first at the left, sits at the right and points left; the triangle moves to the left.
    assert.equal(items.length, 2)
    assert.equal(items[0].left, 75)
    assert.match(items[0].body, /<svg class="dx-geom"[^>]*style="transform:scaleX\(-1\)"/)
    assert.equal(items[1].left, 0)
    assert.match(items[0].body, />rightArrow</)
  })

  test('W-2: a tbRl section starts at the right margin, as Word lays vertical text out', () => {
    const html = renderDocxPage(docOf(p(r('直書內容')) + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:textDirection w:val="tbRl"/></w:sectPr>'), 'v.docx', 'v')
    assert.match(html, /<div class="dx-vert"/)
    const rule = /\.dx-vert\{([^}]*)\}/.exec(html)?.[1] ?? ''
    // The box spans the text column, so vertical-rl lines begin at its right edge rather than shrinking to the left.
    assert.match(rule, /(^|;)width:100%/)
  })

  test('W-3: with a different first page (titlePg), the header at the top is the first-page header', () => {
    const hdr = (text: string) => `<w:hdr ${W}>${p(r(text))}</w:hdr>`
    const parts = { 'word/header1.xml': hdr('首頁頁首 FIRST'), 'word/header2.xml': hdr('奇數頁頁首 ODD') }
    const relList: [string, string, string][] = [['h1', 'header', 'header1.xml'], ['h2', 'header', 'header2.xml']]
    const sect = (titlePg: string) => `<w:sectPr><w:headerReference w:type="default" r:id="h2"/><w:headerReference w:type="first" r:id="h1"/>${titlePg}</w:sectPr>`
    const top = (html: string) => /<div class="dx-header"[^>]*>([\s\S]*?)<\/div><div class="dx-body">/.exec(html)?.[1] ?? ''
    assert.match(top(readDocx(docOf(p(r('本文')) + sect('<w:titlePg/>'), parts, relList), '').html), /首頁頁首 FIRST/)
    // Without titlePg (or with it turned off) the first page shows the default header.
    assert.match(top(readDocx(docOf(p(r('本文')) + sect(''), parts, relList), '').html), /奇數頁頁首 ODD/)
    assert.match(top(readDocx(docOf(p(r('本文')) + sect('<w:titlePg w:val="0"/>'), parts, relList), '').html), /奇數頁頁首 ODD/)
  })

  test('W-4: kana, English-word and other list formats number as Word does, not as 1, 2, 3', () => {
    const cases: [string, number, string][] = [
      ['aiueo', 1, 'ｱ'], ['aiueo', 2, 'ｲ'], ['aiueo', 6, 'ｶ'], ['aiueoFullWidth', 1, 'ア'], ['aiueoFullWidth', 46, 'ン'],
      ['iroha', 1, 'ｲ'], ['iroha', 2, 'ﾛ'], ['irohaFullWidth', 1, 'イ'], ['irohaFullWidth', 3, 'ハ'],
      ['cardinalText', 1, 'One'], ['cardinalText', 21, 'Twenty-one'], ['cardinalText', 115, 'One hundred fifteen'],
      ['ordinalText', 1, 'First'], ['ordinalText', 12, 'Twelfth'], ['ordinalText', 23, 'Twenty-third'], ['ordinalText', 100, 'One hundredth'],
      ['ordinal', 2, '2nd'], ['ordinal', 11, '11th'], ['hex', 26, '1A'], ['chicago', 1, '*'], ['chicago', 5, '**'],
      ['ideographEnclosedCircle', 3, '㊂'], ['ganada', 2, '나'], ['chosung', 3, 'ㄷ'], ['russianLower', 2, 'б'], ['russianUpper', 3, 'В']
    ]
    const abstract = cases.map(([fmt, start], i) => `<w:abstractNum w:abstractNumId="${i}"><w:lvl w:ilvl="0"><w:start w:val="${start}"/><w:numFmt w:val="${fmt}"/><w:lvlText w:val="%1|"/></w:lvl></w:abstractNum>`).join('')
    const nums = cases.map((_, i) => `<w:num w:numId="${i + 1}"><w:abstractNumId w:val="${i}"/></w:num>`).join('')
    const body = cases.map(([fmt], i) => p(r(fmt), `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="${i + 1}"/></w:numPr>`)).join('')
    const html = readDocx(docOf(body, { 'word/numbering.xml': `<w:numbering ${W}>${abstract}${nums}</w:numbering>` }, [['n', 'numbering', 'numbering.xml']]), '').html
    const labels = [...html.matchAll(/class="dx-num[^"]*">([^<|]*)\|/g)].map((m) => m[1])
    // The standard spells cardinalText/ordinalText in lower case; Word capitalises the first word in a list. The casing
    // inside a compound (Twenty-one or Twenty-One) is not pinned.
    const fold = (label: string, at: number) => (cases[at][0].endsWith('Text') ? label.toLowerCase() : label)
    assert.deepEqual(labels.map(fold), cases.map(([, , label], at) => fold(label, at)))
    assert.ok(labels.filter((_, at) => cases[at][0].endsWith('Text')).every((label) => /^[A-Z]/.test(label)), labels.join())
  })
})
