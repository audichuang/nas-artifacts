// Office's 8 × 8 hatch patterns, shared by the PowerPoint, Word and Excel readers so a pattern looks the same in all three.

/** DrawingML pattern presets (a:pattFill prst): each row a byte from the top, high bit leftmost. */
export const PATTERNS: Record<string, string> = {
  pct5: '8000000008000000', pct10: '8000080080000800', pct20: '8800220088002200', pct25: '8822882288228822', pct30: 'aa44aa11aa44aa11', pct40: 'aa55aa51aa55aa15',
  pct50: 'aa55aa55aa55aa55', pct60: 'ee55bb55ee55bb55', pct70: '77dd77dd77dd77dd', pct75: '77ffddff77ffddff', pct80: 'effffeffeffffeff', pct90: 'fffffff7ffffff7f',
  horz: 'ff00000000000000', vert: '8080808080808080', ltHorz: 'ff000000ff000000', ltVert: '8888888888888888', dkHorz: 'ffff0000ffff0000', dkVert: 'cccccccccccccccc',
  narHorz: 'ff00ff00ff00ff00', narVert: '5555555555555555', dashHorz: 'f00000000f000000', dashVert: '8080808008080808', cross: 'ff80808080808080',
  dnDiag: '8040201008040201', upDiag: '0102040810204080', ltDnDiag: '8844221188442211', ltUpDiag: '1122448811224488', dkDnDiag: 'cc663399cc663399',
  dkUpDiag: '3366cc993366cc99', wdDnDiag: 'c1e070381c0e0783', wdUpDiag: '83070e1c3870e0c1', dashDnDiag: '0000884422110000', dashUpDiag: '0000112244880000',
  diagCross: '8142241818244281', smCheck: '9966669999666699', lgCheck: 'f0f0f0f00f0f0f0f', smGrid: 'ff888888ff888888', lgGrid: 'ff80808080808080',
  dotGrid: 'aa00800080008000', smConfetti: '8008400210012004', lgConfetti: 'b130031bd8c00c8d', horzBrick: 'ff808080ff080808', diagBrick: '0102040818244281',
  solidDmnd: '10387cfe7c381000', openDmnd: '8244281028448201', dotDmnd: '8000220008002200', plaid: 'aa55aa55f0f0f0f0', sphere: '77898f8f7798f8f8',
  weave: '8854224588142251', divot: '0010081000800180', shingle: '038448300c020101', wave: '001825c0001825c0', trellis: 'ff66ff99ff66ff99', zigZag: '8142241881422418'
}

/** Excel's cell patterns (patternFill patternType), in the same 8 × 8 form; the names are Excel's own. */
export const CELL_PATTERNS: Record<string, string> = {
  gray0625: '8000080080000800', gray125: '8800220088002200', lightGray: '8822882288228822', mediumGray: 'aa55aa55aa55aa55', darkGray: '77dd77dd77dd77dd',
  darkHorizontal: 'ffff0000ffff0000', darkVertical: 'cccccccccccccccc', darkDown: 'cc663399cc663399', darkUp: '3366cc993366cc99',
  darkGrid: 'cccc3333cccc3333', darkTrellis: 'ff66ff99ff66ff99', lightHorizontal: 'ff000000ff000000', lightVertical: '8888888888888888',
  lightDown: '8844221188442211', lightUp: '1122448811224488', lightGrid: 'ff888888ff888888', lightTrellis: '9966669999666699'
}

/** One hatch tile is 8 screen pixels at 100 % zoom: 6 points. */
export const PATTERN_PT = 6

/** A hatch tile's SVG content in an 8 × 8 box: the background, then the foreground pixels (CSS colours). */
export function patternTile(fg: string, bg: string, bits: string): string {
  let d = ''
  for (let y = 0; y < 8; y++) {
    const row = parseInt(bits.slice(y * 2, y * 2 + 2), 16)
    for (let x = 0; x < 8; x++) if (row & (0x80 >> x)) d += `M${x},${y}h1v1h-1z`
  }
  return `<rect width="8" height="8" fill="${bg}"/>${d ? `<path d="${d}" fill="${fg}"/>` : ''}`
}

/** The tile as a standalone SVG, for a CSS background. */
export const patternSvg = (fg: string, bg: string, bits: string): string =>
  `<svg xmlns="http://www.w3.org/2000/svg" width="8" height="8" shape-rendering="crispEdges">${patternTile(fg, bg, bits)}</svg>`

/** The share of a tile's pixels in the foreground colour (0–1): how dark the pattern reads from a distance. */
export const patternCover = (bits: string): number => [...bits].reduce((sum, hex) => sum + [...parseInt(hex, 16).toString(2)].filter((bit) => bit === '1').length, 0) / 64
