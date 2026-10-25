import assert from 'node:assert/strict'
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { after, before, describe, test } from 'node:test'
import { ArtifactStore } from '../src/artifact-store.ts'
import { createContentServer } from '../src/content-server.ts'
import { docxText, readDocx, renderDocxPage } from '../src/docx.ts'
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
    assert.match(html, /<script>throw new Error\("Word: not a zip file"\)<\/script>/)
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
    assert.match(html, /<script>throw new Error\("Word: not a zip file"\)<\/script>/)
  })
})
