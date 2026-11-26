/**
 * Host half of the right-Sidebar print action.
 *
 * The document preview renders DOCX/PPTX through `officeKit` and only keeps the
 * converted PDF in browser memory, so printing cannot reuse those bytes. This
 * module re-runs the same conversion through the same Host service — which is
 * content-addressed, so an unchanged file yields the very PDF the preview shows.
 *
 * Three modes:
 * - `prepare` rasterizes the chosen pages into the workspace, so the browser can
 *   lay them out on paper and hand the result to the Windows print dialog.
 * - `print` falls back to the system `Print` verb of the file's registered
 *   handler, which prints without any dialog.
 * - `save` copies the same PDF beside the source.
 *
 * @module dsh-preview-printer
 */

import { spawn } from 'node:child_process'
import { copyFile, mkdir, readFile, readdir, rm, stat, writeFile } from 'node:fs/promises'
import { basename, extname, join, resolve, sep } from 'node:path'

/** Host services used by the command. */
export const inject = ['commands', 'workingDirectory']

/** Slash command the Client action dispatches. */
const COMMAND = 'print-preview'

/** Marker the Client uses to recognize a prepared page manifest. */
const MANIFEST_MARKER = 'PREVIEW-PRINT '

/** Extensions the action offers, matching the Client-side gate. */
const PRINTABLE = new Set(['pdf', 'docx', 'doc', 'pptx', 'ppt'])

/** Output formats the engine converts directly to PDF. */
const CONVERTIBLE = new Set(['docx', 'doc', 'pptx', 'ppt'])

/** Legacy inputs a preview can rescue by converting to a modern format. */
const LEGACY_FORMATS = { xls: 'xlsx', doc: 'docx', ppt: 'pptx' }

/** Default raster resolution for the in-app print document. */
const DPI = 300

/** Workspace-local directory holding prepared page images. */
const PRINT_DIR = '.dsh-print'

/** Prepared page sets older than this are removed on the next request. */
const PRINT_TTL_MS = 2 * 60 * 60 * 1000

/**
 * Register the print command.
 * @param ctx - Host context carrying the command registry and working directory.
 */
export function apply(ctx) {
  ctx.effect(() => ctx.commands.register({
    name: COMMAND,
    description: 'Print previewed PDF, DOCX or PPTX files, or save the same PDF beside the source.',
    // The payload is a machine request from the Sidebar action, not user prose.
    recordInput: false,
    handler: invocation => printPreview(ctx, invocation),
  }), 'preview-printer: print command')
}

/**
 * Resolve one request, produce the PDF, then prepare, print, or save it.
 * @param ctx - Host context.
 * @param invocation - command invocation carrying the target Session and signal.
 * @returns a settled command result whose text is shown to the user.
 */
async function printPreview(ctx, invocation) {
  const request = decodeRequest(invocation.rawInput)
  if (request === undefined) return { kind: 'error', text: '打印请求无效，请重新点击按钮。' }
  const filePath = resolve(request.path)
  const extension = extname(filePath).slice(1).toLowerCase()
  if (!PRINTABLE.has(extension) && !(request.mode === 'convert' && LEGACY_FORMATS[extension] !== undefined)) {
    return { kind: 'error', text: `只支持 PDF、DOCX、PPTX 文件，收到 .${extension || '未知'}。` }
  }

  let info
  try {
    info = await stat(filePath)
  } catch {
    return { kind: 'error', text: `找不到文件：${filePath}` }
  }
  if (!info.isFile()) return { kind: 'error', text: `不是普通文件：${filePath}` }

  let cwd
  try {
    cwd = await ctx.workingDirectory.ensure(invocation.agent, invocation.signal)
  } catch {
    return { kind: 'error', text: '无法确定当前工作区。' }
  }
  if (!within(cwd, filePath)) {
    return { kind: 'error', text: '只能打印当前工作区内的文件。' }
  }

  // Repair never needs a conversion: pad the container and stop there. Doing it
  // before `convertedPdf` also keeps a stale-Client request from ever printing.
  if (request.mode === 'convert') {
    try {
      return { kind: 'success', text: `已转换：${await convertForPreview(ctx, invocation, filePath, info)}` }
    } catch (error) {
      return { kind: 'error', text: `转换失败：${describe(error)}` }
    }
  }

  if (request.mode === 'repair') {
    try {
      const padded = await paddedSource(invocation, filePath, info)
      if (padded === undefined) return { kind: 'error', text: '文件已是扇区对齐，无需修复。' }
      const repairedPath = await saveRepaired(filePath, padded.bytes)
      return { kind: 'success', text: `已生成修复版：${repairedPath}` }
    } catch (error) {
      return { kind: 'error', text: `修复失败：${describe(error)}` }
    }
  }

  let prepared
  try {
    prepared = extension === 'pdf'
      ? { path: filePath, pageCount: undefined, repaired: false }
      : await convertedPdf(ctx, invocation, filePath, info)
  } catch (error) {
    return { kind: 'error', text: `转成 PDF 失败：${describe(error)}` }
  }
  const pdfPath = prepared.path
  const pages = prepared.pageCount === undefined ? '' : `（${prepared.pageCount} 页）`
  const repaired = prepared.repaired
    ? (prepared.repairedPath === undefined
        ? '（文件容器不完整，已自动补齐）'
        : `（文件容器不完整，已自动补齐；修复版：${prepared.repairedPath}）`)
    : ''

  if (request.mode === 'save') {
    try {
      return { kind: 'success', text: `已保存 PDF${pages}${repaired}：${await saveCopy(pdfPath, filePath)}` }
    } catch (error) {
      return { kind: 'error', text: `保存 PDF 失败：${describe(error)}` }
    }
  }

  if (request.mode === 'prepare') {
    try {
      const manifestPath = await preparePages(ctx, invocation, cwd, pdfPath, request.range, request.dpi, prepared.repaired === true, prepared.repairedPath)
      return { kind: 'success', text: `${MANIFEST_MARKER}${manifestPath}` }
    } catch (error) {
      return { kind: 'error', text: `准备打印失败：${describe(error)}` }
    }
  }

  try {
    await printPdf(pdfPath)
    return { kind: 'success', text: `已交给系统打印${pages}${repaired}。` }
  } catch (error) {
    try {
      const saved = await saveCopy(pdfPath, filePath)
      return { kind: 'success', text: `系统打印没有启动（${describe(error)}），已存为 PDF${pages}${repaired}：${saved}` }
    } catch (saveError) {
      return { kind: 'error', text: `打印失败：${describe(error)}；存为 PDF 也失败：${describe(saveError)}` }
    }
  }
}

/**
 * Produce the PDF the preview shows for an Office source: the same service and
 * the same content-addressed conversion, saved to the output store.
 *
 * A legacy `.doc`/`.ppt` whose compound file lost its final partial sector is
 * rejected by the converter as `invalid-document`. Such a file is also what the
 * preview refuses to open, while WPS still reads it. One retry pads a copy in
 * memory to the sector boundary; the source file is never modified.
 *
 * @param ctx - Host context.
 * @param invocation - command invocation carrying the Session identity.
 * @param filePath - absolute Office source path.
 * @param info - authorized stat of the source.
 * @returns absolute path of the saved PDF, the converter's page count, and whether padding rescued it.
 */
async function convertedPdf(ctx, invocation, filePath, info) {
  const extension = extname(filePath).slice(1).toLowerCase()
  if (!CONVERTIBLE.has(extension)) throw new Error(`不支持转换 .${extension}`)
  const kit = ctx.get('officeKit')
  if (kit === undefined) throw new Error('当前环境没有可用的 Office 转换服务')

  const stem = basename(filePath, extname(filePath))
  const output = { owner: invocation.agent.session.id, name: `${stem}-${shortHash(filePath)}.pdf` }
  const source = await fileSource(invocation, filePath, info)
  try {
    const result = await kit.convertTo({ extension, format: 'pdf', source, output }, invocation.signal)
    return { path: result.path, pageCount: result.pageCount, repaired: false }
  } catch (error) {
    const legacy = extension === 'doc' || extension === 'ppt'
    if (!legacy || !isCompoundFailure(error)) throw error
    const padded = await paddedSource(invocation, filePath, info)
    if (padded === undefined) throw error
    const result = await kit.convertTo({ extension, format: 'pdf', source: padded.source, output }, invocation.signal)
    // Persist the padded copy so the user also gets a file the preview can open.
    const repairedPath = await saveRepaired(filePath, padded.bytes)
    return { path: result.path, pageCount: result.pageCount, repaired: true, repairedPath }
  }
}

/**
 * Convert one legacy file the preview cannot render into a modern sibling it can.
 * Real compound documents go through office-kit; text exports that quoting tools
 * mislabel as `.xls` become a UTF-8 CSV beside the source.
 * @param ctx - Host context.
 * @param invocation - command invocation carrying the Session identity.
 * @param filePath - absolute legacy source path.
 * @param info - stat of the source.
 * @returns absolute path of the converted file inside the workspace.
 */
async function convertForPreview(ctx, invocation, filePath, info) {
  const extension = extname(filePath).slice(1).toLowerCase()
  const target = LEGACY_FORMATS[extension]
  if (target === undefined) throw new Error(`不支持转换 .${extension}`)
  const directory = resolve(filePath, '..')
  const stem = basename(filePath, extname(filePath))
  const bytes = await readFile(filePath)

  if (isCompoundFile(bytes)) {
    const kit = ctx.get('officeKit')
    if (kit === undefined) throw new Error('当前环境没有可用的 Office 转换服务')
    const output = { owner: invocation.agent.session.id, name: `${stem}-转换.${target}` }
    const source = await fileSource(invocation, filePath, info)
    let result
    try {
      result = await kit.convertTo({ extension, format: target, source, output }, invocation.signal)
    } catch (error) {
      // The same truncated container that blocks printing also blocks this
      // conversion, so pad it once and retry.
      const padded = isCompoundFailure(error) ? await paddedSource(invocation, filePath, info) : undefined
      if (padded === undefined) throw error
      result = await kit.convertTo({ extension, format: target, source: padded.source, output }, invocation.signal)
      await saveRepaired(filePath, padded.bytes)
    }
    const destination = join(directory, `${stem}-转换.${target}`)
    await copyFile(result.path, destination)
    return destination
  }

  const csv = toCsv(decodeLegacyText(bytes))
  if (csv.trim() === '') throw new Error('文件内容为空，无法转换')
  // Prefer a real workbook: it keeps types, survives round trips and is what
  // people open next. CSV stays as the fallback when packing fails.
  try {
    const workbook = buildXlsx(parseDelimited(decodeLegacyText(bytes)), stem)
    const destination = join(directory, `${stem}-转换.xlsx`)
    await writeFile(destination, workbook)
    return destination
  } catch {
    const destination = join(directory, `${stem}-转换.csv`)
    await writeFile(destination, csv, 'utf8')
    return destination
  }
}

/**
 * @param bytes - file head.
 * @returns whether the bytes open like an OLE compound document.
 */
function isCompoundFile(bytes) {
  const magic = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]
  if (bytes.length < magic.length) return false
  return magic.every((value, index) => bytes[index] === value)
}

/**
 * Decode a text-based legacy export: UTF-8 when it is valid, otherwise the GBK
 * the exporting tools on this platform actually write.
 * @param bytes - complete file bytes.
 * @returns decoded text.
 */
function decodeLegacyText(bytes) {
  try { return new TextDecoder('utf-8', { fatal: true }).decode(bytes) } catch { /* fall through */ }
  try { return new TextDecoder('gbk').decode(bytes) } catch { /* fall through */ }
  return new TextDecoder('latin1').decode(bytes)
}

/**
 * Turn a delimited text export into CSV, unwrapping Excel's `="code"` guards so
 * identifiers keep their leading zeros without showing the wrapper.
 * @param text - decoded export.
 * @returns CSV text.
 */
function toCsv(text) {
  const lines = String(text).replace(/\r\n?/g, '\n').split('\n').filter(line => line.trim() !== '')
  if (lines.length === 0) return ''
  const tabs = lines.slice(0, 20).reduce((total, line) => total + line.split('\t').length - 1, 0)
  const delimiter = tabs > 0 ? '\t' : ','
  return lines
    .map(line => line.split(delimiter).map(field => quoteCsv(unwrapCell(field))).join(','))
    .join('\n')
}

/** @param field - raw delimited field. @returns the identifier inside `="…"`, else the field. */
function unwrapCell(field) {
  const trimmed = String(field).trim()
  const match = /^="(.*)"$/.exec(trimmed)
  return match === null ? field : match[1]
}

/** @param value - one CSV field. @returns the field quoted when it needs to be. */
function quoteCsv(value) {
  const text = String(value)
  return /[",\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text
}

/**
 * Split a delimited export into a cell matrix, turning plain integers and
 * decimals into numbers while keeping identifiers such as `002580` as text.
 * @param text - decoded export.
 * @returns rows of cell values.
 */
function parseDelimited(text) {
  const lines = String(text).replace(/\r\n?/g, '\n').split('\n').filter(line => line.trim() !== '')
  const tabs = lines.slice(0, 20).reduce((total, line) => total + line.split('\t').length - 1, 0)
  const delimiter = tabs > 0 ? '\t' : ','
  return lines.map(line => line.split(delimiter).map(field => {
    const value = unwrapCell(field).trim()
    if (value === '') return null
    if (/^-?(?:0|[1-9]\d*)(?:\.\d+)?$/.test(value)) return Number(value)
    return value
  }))
}

/** @param index - one-based column number. @returns its spreadsheet letters. */
function columnName(index) {
  let name = ''
  let value = index
  while (value > 0) {
    const remainder = (value - 1) % 26
    name = String.fromCharCode(65 + remainder) + name
    value = Math.floor((value - 1) / 26)
  }
  return name
}

/** @param value - text for XML. @returns escaped text. */
function escapeXml(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}

const CRC_TABLE = (() => {
  const table = new Uint32Array(256)
  for (let index = 0; index < 256; index++) {
    let value = index
    for (let bit = 0; bit < 8; bit++) value = value & 1 ? (0xedb88320 ^ (value >>> 1)) : (value >>> 1)
    table[index] = value >>> 0
  }
  return table
})()

/** @param buffer - bytes. @returns their CRC-32. */
function crc32(buffer) {
  let crc = 0xffffffff
  for (const byte of buffer) crc = CRC_TABLE[(crc ^ byte) & 0xff] ^ (crc >>> 8)
  return (crc ^ 0xffffffff) >>> 0
}

/**
 * Pack files into a stored (uncompressed) ZIP, which is all a minimal xlsx needs.
 * @param entries - name and text pairs.
 * @returns the archive bytes.
 */
function zipStore(entries) {
  const parts = []
  const central = []
  let offset = 0
  for (const [name, text] of entries) {
    const nameBytes = Buffer.from(name, 'utf8')
    const data = Buffer.from(text, 'utf8')
    const crc = crc32(data)
    const local = Buffer.alloc(30)
    local.writeUInt32LE(0x04034b50, 0)
    local.writeUInt16LE(20, 4)
    local.writeUInt16LE(0x0800, 6)
    local.writeUInt16LE(0, 8)
    local.writeUInt16LE(0, 10)
    local.writeUInt16LE(0x0021, 12)
    local.writeUInt32LE(crc, 14)
    local.writeUInt32LE(data.length, 18)
    local.writeUInt32LE(data.length, 22)
    local.writeUInt16LE(nameBytes.length, 26)
    local.writeUInt16LE(0, 28)
    parts.push(local, nameBytes, data)
    central.push({ nameBytes, crc, size: data.length, offset })
    offset += local.length + nameBytes.length + data.length
  }

  const directory = []
  let directorySize = 0
  for (const entry of central) {
    const header = Buffer.alloc(46)
    header.writeUInt32LE(0x02014b50, 0)
    header.writeUInt16LE(20, 4)
    header.writeUInt16LE(20, 6)
    header.writeUInt16LE(0x0800, 8)
    header.writeUInt16LE(0, 10)
    header.writeUInt16LE(0, 12)
    header.writeUInt16LE(0x0021, 14)
    header.writeUInt32LE(entry.crc, 16)
    header.writeUInt32LE(entry.size, 20)
    header.writeUInt32LE(entry.size, 24)
    header.writeUInt16LE(entry.nameBytes.length, 28)
    header.writeUInt32LE(entry.offset, 42)
    directory.push(header, entry.nameBytes)
    directorySize += header.length + entry.nameBytes.length
  }

  const end = Buffer.alloc(22)
  end.writeUInt32LE(0x06054b50, 0)
  end.writeUInt16LE(central.length, 8)
  end.writeUInt16LE(central.length, 10)
  end.writeUInt32LE(directorySize, 12)
  end.writeUInt32LE(offset, 16)
  return Buffer.concat([...parts, ...directory, end])
}

/**
 * Build a single-sheet xlsx workbook from a cell matrix, with no dependency on a
 * spreadsheet library.
 * @param rows - cell values, `null` for an empty cell.
 * @param sheetName - worksheet name.
 * @returns the workbook bytes.
 */
function buildXlsx(rows, sheetName) {
  const name = String(sheetName).replace(/[\\/?*[\]:]/g, ' ').slice(0, 31).trim() || 'Sheet1'

  // Shared strings are what Excel itself writes; readers that mishandle inline
  // strings still resolve these, so text survives every consumer.
  const strings = []
  const indexOfString = new Map()
  const stringIndex = (value) => {
    let index = indexOfString.get(value)
    if (index === undefined) {
      index = strings.length
      strings.push(value)
      indexOfString.set(value, index)
    }
    return index
  }

  const sheetRows = rows.map((cells, rowIndex) => {
    const entries = cells.map((value, columnIndex) => {
      if (value === null || value === undefined || value === '') return ''
      const reference = `${columnName(columnIndex + 1)}${rowIndex + 1}`
      if (typeof value === 'number') return `<c r="${reference}"><v>${value}</v></c>`
      return `<c r="${reference}" t="s"><v>${stringIndex(String(value))}</v></c>`
    }).join('')
    return `<row r="${rowIndex + 1}">${entries}</row>`
  }).join('')

  const dimension = `A1:${columnName(Math.max(1, rows.reduce((width, cells) => Math.max(width, cells.length), 0)))}${Math.max(1, rows.length)}`
  const columns = Math.max(1, rows.reduce((width, cells) => Math.max(width, cells.length), 0))
  const sheet = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    + `<dimension ref="${dimension}"/>`
    + '<sheetViews><sheetView workbookViewId="0"/></sheetViews>'
    // Readers derive row height from this element; without it exceljs reports
    // `defaultRowHeight` as undefined and the preview draws no rows at all.
    + '<sheetFormatPr defaultRowHeight="15" baseColWidth="10"/>'
    + `<cols><col min="1" max="${columns}" width="14" customWidth="1"/></cols>`
    + `<sheetData>${sheetRows}</sheetData></worksheet>`

  const sharedStrings = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${strings.length}" uniqueCount="${strings.length}">`
    + strings.map(text => `<si><t xml:space="preserve">${escapeXml(text)}</t></si>`).join('')
    + '</sst>'

  const workbook = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
    + 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    + `<sheets><sheet name="${escapeXml(name)}" sheetId="1" state="visible" r:id="rId1"/></sheets></workbook>`

  const workbookRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
    + '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
    + '</Relationships>'

  const packageRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
    + '</Relationships>'

  const contentTypes = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    + '<Default Extension="xml" ContentType="application/xml"/>'
    + '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
    + '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
    + '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
    + '</Types>'

  return zipStore([
    ['[Content_Types].xml', contentTypes],
    ['_rels/.rels', packageRels],
    ['xl/workbook.xml', workbook],
    ['xl/_rels/workbook.xml.rels', workbookRels],
    ['xl/sharedStrings.xml', sharedStrings],
    ['xl/worksheets/sheet1.xml', sheet],
  ])
}

/**
 * Write the zero-padded legacy file beside its source, reusing an identical copy
 * that a previous repair already produced.
 * @param sourcePath - absolute path of the damaged source.
 * @param bytes - padded file bytes.
 * @returns absolute path of the repaired copy.
 */
async function saveRepaired(sourcePath, bytes) {
  const directory = resolve(sourcePath, '..')
  const stem = basename(sourcePath, extname(sourcePath))
  const extension = extname(sourcePath)
  const buffer = Buffer.from(bytes.buffer, bytes.byteOffset, bytes.length)
  for (let attempt = 0; attempt < 100; attempt++) {
    const suffix = attempt === 0 ? '' : `-${attempt + 1}`
    const target = join(directory, `${stem}-修复${suffix}${extension}`)
    try {
      const existing = await readFile(target)
      if (existing.length === buffer.length && existing.equals(buffer)) return target
      continue
    } catch (error) {
      if (error === null || error.code !== 'ENOENT') throw error
    }
    await writeFile(target, buffer, { flag: 'wx' })
    return target
  }
  throw new Error('同名修复版过多，无法确定保存文件名')
}

/**
 * Whether a conversion failure names a damaged or truncated compound file.
 * @param error - thrown conversion error.
 * @returns true when sector padding may rescue the document.
 */
function isCompoundFailure(error) {
  if (error === undefined || error === null) return false
  if (error.code === 'invalid-document') return true
  return /compound[- ]file|truncated/i.test(String(error.message ?? ''))
}

/**
 * Describe the same legacy file with its final sector zero-filled, so a strict
 * compound-file reader accepts it. Returns `undefined` when the file is already
 * sector-aligned or too small to be a compound file.
 * @param invocation - command invocation carrying the Session identity.
 * @param filePath - absolute source path.
 * @param info - stat of the source.
 * @returns the source over the padded bytes plus those bytes, or `undefined`.
 */
async function paddedSource(invocation, filePath, info) {
  const bytes = await readFile(filePath)
  const sector = 512
  if (bytes.length <= sector) return undefined
  const target = sector + Math.ceil((bytes.length - sector) / sector) * sector
  if (target === bytes.length) return undefined
  const padded = new Uint8Array(target)
  padded.set(bytes, 0)
  const version = `padded:${bytes.length}:${Math.trunc(info.mtimeMs)}`
  return {
    bytes: padded,
    source: {
      key: JSON.stringify([invocation.agent.session.id, filePath, 'padded']),
      version,
      bytes: padded.length,
      async read(signal) {
        signal.throwIfAborted()
        return { bytes: padded, version }
      },
    },
  }
}

/**
 * Rasterize the requested pages of one PDF into the Session workspace and write
 * a manifest the browser can read through the workspace-files endpoint.
 * @param ctx - Host context.
 * @param invocation - command invocation carrying the Session identity.
 * @param cwd - Session working directory.
 * @param pdfPath - absolute path of the PDF to print.
 * @param rangeText - requested pages such as `1-3,5`; empty renders every page.
 * @param dpi - raster resolution for the printed pages; 300 keeps text crisp.
 * @param repaired - whether the source needed sector padding before conversion.
 * @param repairedPath - path of the saved padded copy, when one exists.
 * @returns absolute path of the manifest.
 */
async function preparePages(ctx, invocation, cwd, pdfPath, rangeText, dpi, repaired = false, repairedPath) {
  const kit = ctx.get('officeKit')
  if (kit === undefined) throw new Error('当前环境没有可用的 Office 转换服务')
  const info = await stat(pdfPath)
  const source = await fileSource(invocation, pdfPath, info)
  const resolution = resolutionOf(dpi)
  // A large page at 400 DPI can exceed the rasterizer's pixel ceiling; step down
  // until one resolution fits instead of failing the whole request.
  let rendered
  let lastError
  for (const candidate of [resolution, 300, 200, 150].filter((value, index, all) => value <= resolution && all.indexOf(value) === index)) {
    try {
      rendered = await kit.renderImages({
        extension: 'pdf',
        source,
        dpi: candidate,
        maxImages: 40,
        maxPixels: 16_000_000,
        maxDimension: 6000,
        maxBytes: 64 * 1024 * 1024,
      }, invocation.signal)
      break
    } catch (error) {
      lastError = error
      if (!/pixel|dimension|too.?large|invalid-request/i.test(describe(error))) throw error
    }
  }
  if (rendered === undefined) throw lastError ?? new Error('无法渲染打印页面')

  const total = rendered.pageCount ?? rendered.images.length
  const wanted = parsePages(rangeText, total)
  const chosen = rendered.images.filter(image => wanted.includes(image.page ?? 0))
  if (chosen.length === 0) throw new Error(`没有可打印的页面（共 ${total} 页）`)

  const directory = join(cwd, PRINT_DIR, `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`)
  await mkdir(directory, { recursive: true })
  const pages = []
  for (const image of chosen) {
    const file = `page-${image.page ?? pages.length + 1}.png`
    const target = join(directory, file)
    await writeFile(target, image.png)
    pages.push({ path: target, page: image.page ?? pages.length + 1, width: image.width, height: image.height })
  }
  const manifestPath = join(directory, 'manifest.json')
  await writeFile(manifestPath, JSON.stringify({ dpi: rendered.dpi ?? resolution, total, pdf: pdfPath, pages, repaired, repairedPath }, null, 2), 'utf8')
  await prunePrepared(cwd, directory)
  return manifestPath
}

/**
 * Clamp a requested print resolution to the range the rasterizer accepts.
 * @param dpi - requested resolution, possibly absent or nonsensical.
 * @returns a resolution between 150 and 600.
 */
function resolutionOf(dpi) {
  const value = Number(dpi)
  if (!Number.isFinite(value)) return DPI
  return Math.min(600, Math.max(150, Math.round(value)))
}

/**
 * Describe one file as an Office-kit source: authorized size and a deferred read
 * that re-states the version the provider reserves capacity from.
 * @param invocation - command invocation carrying the Session identity.
 * @param filePath - absolute source path.
 * @param info - stat of the source.
 * @returns the source the provider reads.
 */
async function fileSource(invocation, filePath, info) {
  const version = `${info.size}:${Math.trunc(info.mtimeMs)}`
  return {
    key: JSON.stringify([invocation.agent.session.id, filePath]),
    version,
    bytes: info.size,
    async read(signal) {
      signal.throwIfAborted()
      const bytes = await readFile(filePath)
      return { bytes, version }
    },
  }
}

/**
 * Hand one saved PDF to the shell print verb. The default PDF application owns
 * the dialog, so printer selection and "Save as PDF" come from the system.
 * @param pdfPath - absolute path of an existing PDF.
 * @returns completion after the shell accepted the verb.
 */
function printPdf(pdfPath) {
  // Start-Process -Verb Print resolves the file's registered `print` verb and
  // throws, instead of failing silently, when no application handles the file.
  const quoted = String(pdfPath).replace(/'/g, "''")
  const script = `Start-Process -FilePath '${quoted}' -Verb Print -ErrorAction Stop`
  return new Promise((settle, fail) => {
    const child = spawn('powershell.exe', ['-NoProfile', '-NonInteractive', '-Command', script], {
      windowsHide: true,
      stdio: ['ignore', 'pipe', 'pipe'],
    })
    let stderr = ''
    child.stderr.on('data', (chunk) => { stderr += String(chunk) })
    child.on('error', fail)
    child.on('close', (code) => {
      if (code === 0) settle()
      else fail(new Error(stderr.trim() || `PowerShell 以退出码 ${code} 结束`))
    })
  })
}

/**
 * Copy the produced PDF beside its source, with a numeric suffix when the name
 * is taken, so the fallback never overwrites an existing file.
 * @param pdfPath - absolute path of the produced PDF.
 * @param sourcePath - absolute path of the previewed source file.
 * @returns absolute path of the written copy.
 */
async function saveCopy(pdfPath, sourcePath) {
  const stem = basename(sourcePath, extname(sourcePath))
  const directory = resolve(sourcePath, '..')
  for (let attempt = 0; attempt < 100; attempt++) {
    const suffix = attempt === 0 ? '' : `-${attempt + 1}`
    const target = join(directory, `${stem}-print${suffix}.pdf`)
    try {
      // `COPYFILE_EXCL` keeps the loop honest if the file appears meanwhile.
      await copyFile(pdfPath, target, 1)
      return target
    } catch (error) {
      if (error && error.code === 'EEXIST') continue
      throw error
    }
  }
  throw new Error('同名 PDF 过多，无法确定保存文件名')
}

/**
 * Parse a page selection against a document's page count.
 * @param rangeText - `1-3,5` style selection; empty means every page.
 * @param total - pages in the document.
 * @returns sorted, distinct, one-based page numbers.
 */
function parsePages(rangeText, total) {
  const text = normalizeSelection(rangeText)
  if (text === '') return Array.from({ length: total }, (_, index) => index + 1)
  const pages = new Set()
  for (const part of text.split(/[,\s]+/)) {
    if (part === '') continue
    const match = /^(\d+)(?:-(\d+))?$/.exec(part)
    if (match === null) throw new Error(`页码格式无效："${part}"（支持 1-3 这样的范围，或 1、3、5 这样的列举）`)
    const start = Number(match[1])
    const end = match[2] === undefined ? start : Number(match[2])
    if (start < 1 || end < start) throw new Error(`页码范围无效："${part}"`)
    if (end > total) throw new Error(`页码超出范围："${part}"，本文档共 ${total} 页`)
    for (let page = start; page <= end; page++) pages.add(page)
  }
  if (pages.size === 0) throw new Error('页码为空')
  return [...pages].sort((left, right) => left - right)
}

/**
 * Fold the punctuation one IME may produce into the separators the parser reads:
 * full-width digits, every common dash or the word 至, and every list separator.
 * @param value - raw page selection from the panel.
 * @returns the normalized selection.
 */
function normalizeSelection(value) {
  return String(value ?? '')
    .replace(/[\uFF10-\uFF19]/g, digit => String.fromCharCode(digit.charCodeAt(0) - 0xfee0))
    .replace(/[，,、;；/|]+/g, ',')
    .replace(/[~～\-–—－−至]+/g, '-')
    .replace(/\s*-\s*/g, '-')
    .trim()
}

/**
 * Remove prepared page sets that outlived their retention.
 * @param cwd - Session working directory.
 * @param keep - directory of the current request.
 * @returns completion once stale directories are gone.
 */
async function prunePrepared(cwd, keep) {
  const root = join(cwd, PRINT_DIR)
  let entries
  try {
    entries = await readdir(root, { withFileTypes: true })
  } catch {
    return
  }
  const deadline = Date.now() - PRINT_TTL_MS
  for (const entry of entries) {
    if (!entry.isDirectory()) continue
    const directory = join(root, entry.name)
    if (directory === keep) continue
    try {
      const info = await stat(directory)
      if (info.mtimeMs < deadline) await rm(directory, { recursive: true, force: true })
    } catch {
      // A directory that vanished or is in use stays for the next pass.
    }
  }
}

/**
 * Decode the Client action payload: base64url JSON.
 * @param rawInput - text following the command name.
 * @returns the parsed request, or `undefined` when malformed.
 */
function decodeRequest(rawInput) {
  const token = String(rawInput).trim()
  if (token === '') return undefined
  try {
    const json = Buffer.from(token.replace(/-/g, '+').replace(/_/g, '/'), 'base64').toString('utf8')
    const value = JSON.parse(json)
    if (typeof value !== 'object' || value === null) return undefined
    if (typeof value.path !== 'string' || value.path.trim() === '') return undefined
    // An unknown mode is a protocol error, never a silent print: a stale Client
    // must not turn a repair request into paper.
    const modes = new Set(['save', 'prepare', 'repair', 'print', 'convert'])
    if (typeof value.mode !== 'string' || !modes.has(value.mode)) return undefined
    return {
      path: value.path,
      mode: value.mode,
      range: typeof value.range === 'string' ? value.range : '',
      dpi: typeof value.dpi === 'number' ? value.dpi : DPI,
    }
  } catch {
    return undefined
  }
}

/**
 * Whether a path is the directory itself or inside it.
 * @param directory - absolute directory.
 * @param target - absolute candidate path.
 * @returns true when the candidate is contained.
 */
function within(directory, target) {
  const base = resolve(directory).toLowerCase()
  const candidate = resolve(target).toLowerCase()
  if (candidate === base) return true
  return candidate.startsWith(base.endsWith(sep) ? base : base + sep)
}

/**
 * Short stable digest so two files with the same base name keep separate
 * entries in the provider's output store.
 * @param value - path to hash.
 * @returns eight hexadecimal digits.
 */
function shortHash(value) {
  let hash = 5381
  for (let index = 0; index < value.length; index++) {
    hash = ((hash * 33) ^ value.charCodeAt(index)) >>> 0
  }
  return hash.toString(16).padStart(8, '0')
}

/**
 * @param error - unknown thrown value.
 * @returns a one-line description for the user.
 */
function describe(error) {
  if (error instanceof Error) return error.message
  return String(error)
}
