#!/usr/bin/env node
// Writes three small Office files to publish: verify.pptx (2 slides: text, a picture, speaker notes; made with
// pptxgenjs), verify.docx (heading + paragraph) and verify.xlsx (two sheets), the last two zipped by
// test/helpers/zip.ts. They prove the pipeline, not that arbitrary uploads render: feed a real file with
// VERIFY_OFFICE_FILE to the office-files journey for that.
// Usage: node .claude/skills/verify/scripts/make-office.mjs <dir>
import { mkdirSync, writeFileSync } from 'node:fs'
import { createRequire } from 'node:module'
import { join } from 'node:path'
import { crc32, deflateSync } from 'node:zlib'
import { fileURLToPath } from 'node:url'
import { ROOT } from './browser.mjs'

/** A solid-colour RGB PNG, built here so the bytes are known to be valid. */
function png(width, height, [r, g, b]) {
  const chunk = (type, data) => {
    const head = Buffer.alloc(8)
    head.writeUInt32BE(data.length, 0)
    head.write(type, 4, 'latin1')
    const crc = Buffer.alloc(4)
    crc.writeUInt32BE(crc32(Buffer.concat([Buffer.from(type, 'latin1'), data])), 0)
    return Buffer.concat([head, data, crc])
  }
  const ihdr = Buffer.alloc(13)
  ihdr.writeUInt32BE(width, 0)
  ihdr.writeUInt32BE(height, 4)
  ihdr.set([8, 2, 0, 0, 0], 8)
  const row = Buffer.concat([Buffer.from([0]), Buffer.from(Array.from({ length: width }, () => [r, g, b]).flat())])
  const raw = Buffer.concat(Array.from({ length: height }, () => row))
  return Buffer.concat([Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), chunk('IHDR', ihdr), chunk('IDAT', deflateSync(raw)), chunk('IEND', Buffer.alloc(0))])
}
const PNG = png(64, 64, [240, 120, 20]).toString('base64')

export async function makeOffice(dir) {
  mkdirSync(dir, { recursive: true })
  const pptxgen = createRequire(join(ROOT, 'package.json'))('pptxgenjs')
  const PptxGenJS = pptxgen.default ?? pptxgen
  const deck = new PptxGenJS()
  const one = deck.addSlide()
  one.addText('Verify deck slide one', { x: 0.5, y: 0.4, w: 9, h: 1, fontSize: 32 })
  one.addNotes('speaker notes for slide one')
  const two = deck.addSlide()
  two.addText('Slide two with a picture', { x: 0.5, y: 0.4, w: 9, h: 1, fontSize: 28 })
  two.addImage({ data: `image/png;base64,${PNG}`, x: 1, y: 1.6, w: 2, h: 2 })
  two.addNotes('speaker notes for slide two')
  writeFileSync(join(dir, 'verify.pptx'), await deck.write({ outputType: 'nodebuffer' }))

  const { zip } = await import(join(ROOT, 'test', 'helpers', 'zip.ts'))
  const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
  const rels = (list) => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${list.map(([id, type, target]) => `<Relationship Id="${id}" Type="${REL}/${type}" Target="${target}"/>`).join('')}</Relationships>`
  const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
  const para = (text, style = '') => `<w:p>${style ? `<w:pPr><w:pStyle w:val="${style}"/></w:pPr>` : ''}<w:r><w:t>${text}</w:t></w:r></w:p>`
  writeFileSync(join(dir, 'verify.docx'), zip({
    '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
    '_rels/.rels': rels([['r1', 'officeDocument', 'word/document.xml']]),
    'word/document.xml': `<w:document ${W}><w:body>${para('Verify document heading', 'Heading1')}${para('A paragraph of selectable verify text.')}</w:body></w:document>`
  }))

  const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
  const cellRow = (r, cells) => `<row r="${r}">${cells.join('')}</row>`
  const str = (ref, index) => `<c r="${ref}" t="s"><v>${index}</v></c>`
  const num = (ref, value) => `<c r="${ref}"><v>${value}</v></c>`
  writeFileSync(join(dir, 'verify.xlsx'), zip({
    '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>',
    '_rels/.rels': rels([['r1', 'officeDocument', 'xl/workbook.xml']]),
    'xl/workbook.xml': `<workbook ${NS}><sheets><sheet name="Totals" sheetId="1" r:id="s1"/><sheet name="Details" sheetId="2" r:id="s2"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': rels([['s1', 'worksheet', 'worksheets/sheet1.xml'], ['s2', 'worksheet', 'worksheets/sheet2.xml'], ['ss', 'sharedStrings', 'sharedStrings.xml']]),
    'xl/sharedStrings.xml': `<sst ${NS}><si><t>Item</t></si><si><t>Amount</t></si><si><t>verify-first-sheet</t></si><si><t>verify-second-sheet</t></si></sst>`,
    'xl/worksheets/sheet1.xml': `<worksheet ${NS}><sheetData>${cellRow(1, [str('A1', 0), str('B1', 1)])}${cellRow(2, [str('A2', 2), num('B2', 42)])}</sheetData></worksheet>`,
    'xl/worksheets/sheet2.xml': `<worksheet ${NS}><sheetData>${cellRow(1, [str('A1', 3), num('B1', 7)])}</sheetData></worksheet>`
  }))
  return { pptx: join(dir, 'verify.pptx'), docx: join(dir, 'verify.docx'), xlsx: join(dir, 'verify.xlsx') }
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  const dir = process.argv[2]
  if (!dir) {
    console.error('usage: make-office.mjs <dir>')
    process.exit(64)
  }
  const files = await makeOffice(dir)
  for (const path of Object.values(files)) console.log(path)
}
