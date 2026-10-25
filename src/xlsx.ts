/**
 * A Excel file published as is (`*.xlsx`), drawn on the server as one HTML page. Not written yet:
 * the page says so, and the throw lands in the viewer's error log.
 */
import { escapeHtml, injectBridge } from './content-render.ts'
import { failureScript } from './office.ts'

/** The whole page for a Excel file. Pictures go through officeMediaBase(fileName) (office.ts). */
export function renderXlsxPage(data: Buffer, fileName: string, title: string): string {
  void data
  void fileName
  const message = 'Excel: not supported yet'
  return injectBridge(`<!doctype html><html lang="zh-Hant"><head><meta charset="utf-8"><title>${escapeHtml(title)}</title></head><body><p>${escapeHtml(message)}</p>${failureScript(message)}</body></html>`)
}

/** The searchable text of a Excel file. */
export function xlsxText(data: Buffer): string {
  void data
  return ''
}
