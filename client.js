/**
 * Browser half of the right-Sidebar print action.
 *
 * Registers one toolbar button into the document preview's action slot. The
 * button opens a small panel; printing asks the Host to rasterize exactly the
 * chosen pages, lays them out as paper-sized images, and hands the result to the
 * Windows print dialog — no external PDF application takes part. Images, Markdown
 * and plain text never touch the Host: the browser prints them directly.
 */
window.__ModuleLoader__.load({
  id: 'dsh-preview-printer',
  factory(require) {
    const React = require('react')
    const h = React.createElement

    const NS = 'previewPrinter'
    const COMMAND = 'print-preview'
    /** Shown in the panel header; the harness keeps it equal to package.json. */
    const VERSION = '0.2.2'
    const MANIFEST_MARKER = 'PREVIEW-PRINT '
    const OFFICE = /\.(pdf|docx?|pptx?)$/i
    const IMAGE = /\.(png|jpe?g|webp|bmp)$/i
    const MARKDOWN = /\.(md|markdown)$/i
    const PLAIN_TEXT = /\.(txt|py|json|js|bat)$/i

    /** Guard rails: a machine-generated JSON can be megabytes and hundreds of pages. */
    const MAX_TEXT_BYTES = 2 * 1024 * 1024
    const MAX_TEXT_LINES = 5000

    const zh = {
      print: '打印',
      save: '存为 PDF',
      savedGeneric: 'PDF 转存已完成！',
      rangeLabel: '页码',
      rangePlaceholder: '不填=全部，如 1-3 或 2，5，7',
      preparing: '正在准备页面…',
      dialogOpened: '已打开打印对话框',
      noDialog: '未检测到打印对话框，已改用直接打印。',
      noDialogImage: '未检测到打印对话框。',
      repairedNote: '（文件容器不完整，已自动补齐）',
      repairedSaved: '（已自动补齐，修复版：{path}）',
      legacyHint: '预览失败多为旧版文档容器残缺；已自动补齐并另存修复版（原文件未动），可继续打印。',
      salvageChip: '已修复',
      salvageBusy: '修复中…',
      salvageFallback: '点此打印',
      salvageTitle: '预览失败，旧版文档容器残缺已由本插件自动补齐，点击继续打印',
      salvageFallbackTitle: '自动修复未成功；点击仍可打印，打印时会再次尝试补齐',
      salvageSaved: '已生成修复版：{path}',
      convertChip: '已转化',
      convertBusy: '转化中…',
      convertTitle: '格式已转化，点击按钮继续预览',
      convertFailedChip: '转化失败',
      convertFailedTitle: '自动转化未成功；可手动另存为 xlsx 或 csv 后再打开',
      textTooLarge: '文件 {size} 超过 {limit}，不适合打印；请改用编辑器打开或另存为 PDF。',
      textTooManyLines: '文件有 {lines} 行，超过 {limit} 行上限，不适合打印。',
      unsupportedSheet: '表格不通过本插件打印；请在 Excel 或 WPS 中打印',
      unsupportedGif: '该格式暂不支持预览与打印',
      unsupportedCode: '源码文本可预览，但打印暂不支持（可另存为 PDF 再打印）',
      unsupportedGeneric: '该格式暂不支持打印',
      saved: '已存为 PDF',
      failed: '操作失败',
      resolution: '清晰度',
      resolutionStandard: '标准 200 DPI',
      resolutionHigh: '高清 300 DPI',
      resolutionUltra: '超清 400 DPI',
      direct: '高清全速打印',
      directHint: '矢量输出，不弹对话框，直接打到默认打印机',
      appTitle: '预览打印',
    }
    const en = {
      print: 'Print',
      save: 'Save as PDF',
      savedGeneric: 'PDF saved.',
      rangeLabel: 'Pages',
      rangePlaceholder: 'Empty = all, e.g. 1-3 or 2,5,7',
      preparing: 'Preparing pages…',
      dialogOpened: 'Print dialog opened',
      noDialog: 'No print dialog appeared; printed directly.',
      noDialogImage: 'No print dialog appeared.',
      repairedNote: '(container was damaged; padded automatically)',
      repairedSaved: '(padded automatically; repaired copy: {path})',
      legacyHint: 'The preview failed because a legacy container is damaged; it is padded already and a repaired copy was saved. The original is untouched.',
      salvageChip: 'Repaired',
      salvageBusy: 'Repairing…',
      salvageFallback: 'Print instead',
      salvageTitle: 'The preview failed; this plugin padded the damaged legacy container already. Click to keep printing.',
      salvageFallbackTitle: 'The automatic repair failed; click to print anyway, the print path pads again.',
      salvageSaved: 'Repaired copy: {path}',
      convertChip: 'Converted',
      convertBusy: 'Converting…',
      convertTitle: 'The format was converted; click to preview the result',
      convertFailedChip: 'Convert failed',
      convertFailedTitle: 'Automatic conversion failed; save the file as xlsx or csv and open that instead',
      textTooLarge: 'This file is {size}, over the {limit} limit; open it in an editor or export to PDF instead.',
      textTooManyLines: 'This file has {lines} lines, over the {limit} line limit; printing is skipped.',
      unsupportedSheet: 'Spreadsheets do not print through this plugin; print them from Excel or WPS',
      unsupportedGif: 'This format is not previewable or printable yet',
      unsupportedCode: 'Source text previews fine, but printing is not supported yet (export to PDF first)',
      unsupportedGeneric: 'This format cannot be printed yet',
      saved: 'Saved as PDF',
      failed: 'Failed',
      resolution: 'Resolution',
      resolutionStandard: 'Standard 200 DPI',
      resolutionHigh: 'High 300 DPI',
      resolutionUltra: 'Ultra 400 DPI',
      direct: 'Full-speed HD print',
      directHint: 'Vector output, no dialog, straight to the default printer',
      appTitle: 'Preview printer',
    }

    /** Which print path one file takes. */
    function pathKind(filePath) {
      if (OFFICE.test(filePath)) return 'office'
      if (/\.(xls|xlt)$/i.test(filePath)) return 'sheet'
      if (IMAGE.test(filePath)) return 'image'
      if (MARKDOWN.test(filePath)) return 'markdown'
      if (PLAIN_TEXT.test(filePath)) return 'text'
      return undefined
    }

    /** Address of one workspace file inside its Session, as the Sidebar expects. */
    function fileAddress(sessionId, filePath) {
      const normalized = String(filePath).replace(/\\/g, '/')
      const encoded = normalized
        .split('/')
        .map(segment => encodeURIComponent(segment).replace(/%3A/gi, ':'))
        .join('/')
      return `dsh-resource://file/session/${encodeURIComponent(sessionId)}/${encoded}`
    }

    /** MIME type of a bitmap the action prints. */
    function mimeOf(filePath) {
      const extension = filePath.split('.').pop().toLowerCase()
      if (extension === 'png') return 'image/png'
      if (extension === 'webp') return 'image/webp'
      if (extension === 'bmp') return 'image/bmp'
      return 'image/jpeg'
    }

    /** File name of one path, without the directory. */
    function fileNameOf(filePath) {
      return String(filePath).split(/[\\/]/).pop() || 'image'
    }

    /**
     * Decode a text file the way this platform writes them: UTF-8 when valid,
     * otherwise the GBK/ANSI that older scripts still use.
     * @param bytes - complete file bytes.
     * @returns decoded text.
     */
    function decodeText(bytes) {
      try { return new TextDecoder('utf-8', { fatal: true }).decode(bytes) } catch { /* fall through */ }
      try { return new TextDecoder('gbk').decode(bytes) } catch { /* fall through */ }
      return new TextDecoder('latin1').decode(bytes)
    }

    /**
     * Pretty-print JSON when it parses, so minified files stay readable on paper.
     * @param text - file contents.
     * @returns indented JSON, or the original text when it is not valid JSON.
     */
    function prettyJson(text) {
      try { return JSON.stringify(JSON.parse(text), null, 2) } catch { return text }
    }

    /** @param bytes - byte count. @returns a short human-readable size. */
    function humanBytes(bytes) {
      if (bytes >= 1024 * 1024) return `${(bytes / (1024 * 1024)).toFixed(1)} MB`
      if (bytes >= 1024) return `${Math.round(bytes / 1024)} KB`
      return `${bytes} B`
    }

    /** Encode the Host payload as base64url JSON, so spaces and CJK survive the command line. */
    function encodeRequest(value) {
      const bytes = new TextEncoder().encode(JSON.stringify(value))
      let binary = ''
      for (const byte of bytes) binary += String.fromCharCode(byte)
      return btoa(binary).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '')
    }

    /** Escape text for safe HTML insertion. */
    function escapeHtml(value) {
      return String(value)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
    }

    /** Inline Markdown: code, bold, emphasis, and links reduced to their text. */
    function inlineMarkdown(value) {
      let text = escapeHtml(value)
      text = text.replace(/`([^`]+)`/g, '<code>$1</code>')
      text = text.replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>')
      text = text.replace(/(^|[^*])\*([^*\n]+)\*/g, '$1<em>$2</em>')
      text = text.replace(/\[([^\]]+)\]\([^)]*\)/g, '$1')
      return text
    }

    /**
     * Render the Markdown subset the preview commonly shows: headings, paragraphs,
     * lists, quotes, fenced code, rules, and inline emphasis.
     * @param source - Markdown text.
     * @returns HTML.
     */
    function renderMarkdown(source) {
      const lines = String(source).replace(/\r\n?/g, '\n').split('\n')
      const out = []
      let paragraph = []
      let list = null
      let quote = []
      let code = null
      const flushParagraph = () => {
        if (paragraph.length === 0) return
        out.push(`<p>${inlineMarkdown(paragraph.join(' '))}</p>`)
        paragraph = []
      }
      const flushList = () => {
        if (list === null) return
        out.push(`<${list.type}>${list.items.map(item => `<li>${inlineMarkdown(item)}</li>`).join('')}</${list.type}>`)
        list = null
      }
      const flushQuote = () => {
        if (quote.length === 0) return
        out.push(`<blockquote>${quote.map(line => `<p>${inlineMarkdown(line)}</p>`).join('')}</blockquote>`)
        quote = []
      }
      const flush = () => { flushParagraph(); flushList(); flushQuote() }

      for (const line of lines) {
        if (code !== null) {
          if (/^```/.test(line)) { out.push(`<pre><code>${escapeHtml(code.join('\n'))}</code></pre>`); code = null }
          else code.push(line)
          continue
        }
        if (/^```/.test(line)) { flush(); code = []; continue }
        if (/^\s*$/.test(line)) { flush(); continue }
        const heading = /^(#{1,6})\s+(.*)$/.exec(line)
        if (heading !== null) { flush(); out.push(`<h${heading[1].length}>${inlineMarkdown(heading[2])}</h${heading[1].length}>`); continue }
        if (/^\s*[-*+]\s+/.test(line)) {
          if (list === null || list.type !== 'ul') { flush(); list = { type: 'ul', items: [] } }
          list.items.push(line.replace(/^\s*[-*+]\s+/, ''))
          continue
        }
        if (/^\s*\d+[.)]\s+/.test(line)) {
          if (list === null || list.type !== 'ol') { flush(); list = { type: 'ol', items: [] } }
          list.items.push(line.replace(/^\s*\d+[.)]\s+/, ''))
          continue
        }
        if (/^>\s?/.test(line)) { flushParagraph(); flushList(); quote.push(line.replace(/^>\s?/, '')); continue }
        if (/^\s*([-*_]\s*){3,}$/.test(line)) { flush(); out.push('<hr>'); continue }
        paragraph.push(line.trim())
      }
      if (code !== null) out.push(`<pre><code>${escapeHtml(code.join('\n'))}</code></pre>`)
      flush()
      return out.join('\n')
    }

    /** zlib-wrapped deflate, as PDF's FlateDecode expects; absent in older engines. */
    async function deflateBytes(bytes) {
      if (typeof CompressionStream === 'undefined') return undefined
      const stream = new Blob([bytes]).stream().pipeThrough(new CompressionStream('deflate'))
      return new Uint8Array(await new Response(stream).arrayBuffer())
    }

    /**
     * Assemble a one-page PDF holding one RGB image, fitted inside the page margins.
     * @param rgb - raw RGB samples, row-major.
     * @param width - image width in pixels.
     * @param height - image height in pixels.
     * @param landscape - whether the sheet is landscape.
     * @param marginPt - page margin in points.
     * @returns the PDF bytes.
     */
    async function buildImagePdf(rgb, width, height, landscape, marginPt) {
      const pageWidth = landscape ? 841.89 : 595.28
      const pageHeight = landscape ? 595.28 : 841.89
      const scale = Math.min((pageWidth - marginPt * 2) / width, (pageHeight - marginPt * 2) / height)
      const drawWidth = width * scale
      const drawHeight = height * scale
      const offsetX = (pageWidth - drawWidth) / 2
      const offsetY = (pageHeight - drawHeight) / 2

      const compressed = await deflateBytes(rgb)
      const samples = compressed ?? rgb
      const content = `q\n${drawWidth.toFixed(2)} 0 0 ${drawHeight.toFixed(2)} ${offsetX.toFixed(2)} ${offsetY.toFixed(2)} cm\n/Im0 Do\nQ\n`

      const chunks = []
      let length = 0
      const push = (value) => {
        const bytes = typeof value === 'string' ? new TextEncoder().encode(value) : value
        chunks.push(bytes)
        length += bytes.length
      }
      const offsets = []
      const object = (body) => { offsets.push(length); push(body) }
      const streamObject = (header, data) => {
        offsets.push(length)
        push(`${header}/Length ${data.length}>>stream\n`)
        push(data)
        push('\nendstream\nendobj\n')
      }

      push('%PDF-1.4\n%\u00e2\u00e3\u00cf\u00d3\n')
      object('1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n')
      object('2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 >>\nendobj\n')
      object(`3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${pageWidth.toFixed(2)} ${pageHeight.toFixed(2)}] `
        + '/Resources << /XObject << /Im0 4 0 R >> >> /Contents 5 0 R >>\nendobj\n')
      streamObject(`4 0 obj\n<< /Type /XObject /Subtype /Image /Width ${width} /Height ${height} `
        + `/ColorSpace /DeviceRGB /BitsPerComponent 8${compressed === undefined ? '' : ' /Filter /FlateDecode'} `, samples)
      streamObject('5 0 obj\n<< ', new TextEncoder().encode(content))

      const xrefOffset = length
      let xref = `xref\n0 ${offsets.length + 1}\n0000000000 65535 f \n`
      for (const offset of offsets) xref += `${String(offset).padStart(10, '0')} 00000 n \n`
      push(xref)
      push(`trailer\n<< /Size ${offsets.length + 1} /Root 1 0 R >>\nstartxref\n${xrefOffset}\n%%EOF\n`)

      const output = new Uint8Array(length)
      let position = 0
      for (const chunk of chunks) { output.set(chunk, position); position += chunk.length }
      return output
    }

    /**
     * Save one previewed bitmap as a single-page PDF in the browser's download folder.
     * @param data - image bytes.
     * @param mime - MIME type of the bitmap.
     * @param name - file name for the download.
     * @returns the written file name.
     */
    async function saveImageAsPdf(data, mime, name) {
      const url = URL.createObjectURL(new Blob([data], { type: mime }))
      try {
        const image = new Image()
        image.src = url
        await image.decode()
        const limit = 12_000_000
        const shrink = image.naturalWidth * image.naturalHeight > limit
          ? Math.sqrt(limit / (image.naturalWidth * image.naturalHeight))
          : 1
        const width = Math.max(1, Math.round(image.naturalWidth * shrink))
        const height = Math.max(1, Math.round(image.naturalHeight * shrink))

        const canvas = document.createElement('canvas')
        canvas.width = width
        canvas.height = height
        const context = canvas.getContext('2d')
        if (context === null) throw new Error('canvas 2D context unavailable')
        // Flatten transparency onto white so the PDF has no alpha channel.
        context.fillStyle = '#ffffff'
        context.fillRect(0, 0, width, height)
        context.drawImage(image, 0, 0, width, height)
        const pixels = context.getImageData(0, 0, width, height).data
        const rgb = new Uint8Array(width * height * 3)
        for (let index = 0, target = 0; index < pixels.length; index += 4) {
          rgb[target++] = pixels[index]
          rgb[target++] = pixels[index + 1]
          rgb[target++] = pixels[index + 2]
        }

        const pdf = await buildImagePdf(rgb, width, height, width > height, 34.02)
        const pdfUrl = URL.createObjectURL(new Blob([pdf], { type: 'application/pdf' }))
        const anchor = document.createElement('a')
        anchor.href = pdfUrl
        anchor.download = name
        document.body.appendChild(anchor)
        anchor.click()
        anchor.remove()
        setTimeout(() => URL.revokeObjectURL(pdfUrl), 30_000)
        return name
      } finally {
        URL.revokeObjectURL(url)
      }
    }

    /** Shared print isolation: only our page root survives the print stylesheet. */
    const PRINT_ISOLATION = [
      '@media print {',
      '  html, body { margin: 0 !important; padding: 0 !important; background: #fff !important; }',
      '  body > *:not([data-preview-printer-print-root]) { display: none !important; }',
      '  [data-preview-printer-print-root] { position: static !important; left: auto !important; top: auto !important; }',
      '}',
    ].join('\n')

    /**
     * Attach one prepared printout, open the print dialog, then clean up.
     * @param style - stylesheet carrying the page rules and print isolation.
     * @param root - element holding the pages, parked outside the visible area.
     * @param urls - object URLs created for this printout.
     * @param ready - resolves once the printout's images are decoded.
     * @returns whether the dialog reported opening, plus the cleanup.
     */
    async function openPrintDialog(style, root, urls, ready) {
      document.head.appendChild(style)
      document.body.appendChild(root)
      await ready
      const dispose = () => {
        root.remove()
        style.remove()
        for (const url of urls) URL.revokeObjectURL(url)
      }

      // `beforeprint` is the only signal the dialog actually opened; without it
      // the caller falls back to handing the file to the system print verb.
      const started = new Promise((resolve) => {
        let fired = false
        const onBefore = () => { fired = true }
        window.addEventListener('beforeprint', onBefore, { once: true })
        setTimeout(() => {
          window.removeEventListener('beforeprint', onBefore)
          resolve(fired)
        }, 1500)
      })
      window.print()
      window.addEventListener('afterprint', dispose, { once: true })
      setTimeout(dispose, 10 * 60 * 1000)
      return { printed: await started, dispose }
    }

    /**
     * Lay prepared document pages out as paper-sized images, one sheet each.
     * @param pages - decoded page images with pixel dimensions.
     * @param dpi - resolution the images were rendered at.
     * @returns the dialog outcome.
     */
    async function printPages(pages, dpi) {
      const sizes = pages.map(page => ({
        width: page.width / dpi,
        height: page.height / dpi,
        landscape: page.width > page.height,
      }))
      const portraits = sizes.filter(size => !size.landscape)
      const landscapes = sizes.filter(size => size.landscape)
      const widest = list => list.reduce((max, size) => Math.max(max, size.width), 0)
      const tallest = list => list.reduce((max, size) => Math.max(max, size.height), 0)
      const portraitWidth = widest(portraits)
      const portraitHeight = tallest(portraits)
      const landscapeWidth = widest(landscapes)
      const landscapeHeight = tallest(landscapes)
      const mixed = portraits.length > 0 && landscapes.length > 0
      const first = sizes[0]

      // Each orientation gets its own named page so a landscape section prints
      // sideways instead of being squeezed onto the first page's sheet.
      const style = document.createElement('style')
      style.textContent = [
        `@page { size: ${(first.landscape ? landscapeWidth : portraitWidth)}in ${(first.landscape ? landscapeHeight : portraitHeight)}in; margin: 0; }`,
        ...(mixed ? [
          `@page pp-portrait { size: ${portraitWidth}in ${portraitHeight}in; margin: 0; }`,
          `@page pp-landscape { size: ${landscapeWidth}in ${landscapeHeight}in; margin: 0; }`,
        ] : []),
        PRINT_ISOLATION,
        '@media print {',
        '  [data-preview-printer-print-root] .pp-page { page-break-after: always; break-after: page; }',
        '  [data-preview-printer-print-root] .pp-page:last-child { page-break-after: auto; break-after: auto; }',
        '}',
      ].join('\n')

      const root = document.createElement('div')
      root.setAttribute('data-preview-printer-print-root', '')
      root.style.cssText = 'position: fixed; left: -20000px; top: 0; background: #fff;'
      const urls = []
      const images = []
      for (let index = 0; index < pages.length; index++) {
        const page = pages[index]
        const size = sizes[index]
        const url = URL.createObjectURL(new Blob([page.data], { type: 'image/png' }))
        urls.push(url)
        const frame = document.createElement('div')
        frame.className = 'pp-page'
        const named = mixed ? `page:${size.landscape ? 'pp-landscape' : 'pp-portrait'};` : ''
        frame.style.cssText = `${named}width:${size.width}in;height:${size.height}in;overflow:hidden;background:#fff;`
        const image = document.createElement('img')
        image.src = url
        // `contain` keeps the whole page visible when a mixed document prints a
        // page on the other orientation's sheet.
        image.style.cssText = 'display:block;width:100%;height:100%;object-fit:contain;'
        images.push(image)
        frame.appendChild(image)
        root.appendChild(frame)
      }
      return openPrintDialog(style, root, urls, Promise.all(images.map(image => image.decode().catch(() => undefined))))
    }

    /**
     * Lay one bitmap out on an A4 sheet, rotated to match its own aspect ratio.
     * @param data - image bytes.
     * @param type - MIME type of the bitmap.
     * @returns the dialog outcome.
     */
    async function openImagePrintout(data, type) {
      const url = URL.createObjectURL(new Blob([data], { type }))
      const probe = new Image()
      probe.src = url
      await probe.decode().catch(() => undefined)
      const landscape = probe.naturalWidth > probe.naturalHeight
      const widthIn = landscape ? 11.69 : 8.27
      const heightIn = landscape ? 8.27 : 11.69

      const style = document.createElement('style')
      style.textContent = [
        `@page { size: ${widthIn}in ${heightIn}in; margin: 0; }`,
        PRINT_ISOLATION,
        '@media print {',
        '  [data-preview-printer-print-root] .pp-page { padding: 12mm; box-sizing: border-box; }',
        '  [data-preview-printer-print-root] .pp-page img { display: block; width: 100%; height: 100%; object-fit: contain; }',
        '}',
      ].join('\n')

      const root = document.createElement('div')
      root.setAttribute('data-preview-printer-print-root', '')
      root.style.cssText = 'position: fixed; left: -20000px; top: 0; background: #fff;'
      const frame = document.createElement('div')
      frame.className = 'pp-page'
      frame.style.cssText = `width: ${widthIn}in; height: ${heightIn}in; overflow: hidden; background: #fff;`
      const image = document.createElement('img')
      image.src = url
      frame.appendChild(image)
      root.appendChild(frame)
      return openPrintDialog(style, root, [url], image.decode().catch(() => undefined))
    }

    /**
     * Build one flowing A4 document from already-rendered HTML.
     * @param innerHtml - body markup.
     * @param extraCss - rules appended after the shared typography.
     * @returns the dialog outcome.
     */
    async function openFlowingPrintout(innerHtml, extraCss = '') {
      const style = document.createElement('style')
      style.textContent = [
        '@page { size: A4; margin: 16mm 14mm; }',
        PRINT_ISOLATION,
        '[data-preview-printer-print-root] .pp-flow { color: #000; background: #fff; font: 11pt/1.7 "Microsoft YaHei", "PingFang SC", system-ui, sans-serif; }',
        '[data-preview-printer-print-root] .pp-flow h1 { font-size: 20pt; margin: 0 0 8pt; }',
        '[data-preview-printer-print-root] .pp-flow h2 { font-size: 16pt; margin: 14pt 0 6pt; }',
        '[data-preview-printer-print-root] .pp-flow h3 { font-size: 13pt; margin: 12pt 0 4pt; }',
        '[data-preview-printer-print-root] .pp-flow h4, [data-preview-printer-print-root] .pp-flow h5, [data-preview-printer-print-root] .pp-flow h6 { font-size: 11.5pt; margin: 10pt 0 4pt; }',
        '[data-preview-printer-print-root] .pp-flow p { margin: 0 0 6pt; }',
        '[data-preview-printer-print-root] .pp-flow ul, [data-preview-printer-print-root] .pp-flow ol { margin: 0 0 6pt; padding-left: 18pt; }',
        '[data-preview-printer-print-root] .pp-flow li { margin: 0 0 2pt; }',
        '[data-preview-printer-print-root] .pp-flow code { font: 10pt "Cascadia Mono", Consolas, monospace; background: #f3f4f6; padding: 0 2pt; }',
        '[data-preview-printer-print-root] .pp-flow pre { background: #f6f8fa; border: 1px solid #e5e7eb; border-radius: 4pt; padding: 8pt; break-inside: avoid; }',
        '[data-preview-printer-print-root] .pp-flow pre code { background: none; padding: 0; }',
        '[data-preview-printer-print-root] .pp-flow blockquote { margin: 0 0 6pt; padding-left: 10pt; border-left: 3pt solid #d0d7de; color: #444; }',
        '[data-preview-printer-print-root] .pp-flow h1, [data-preview-printer-print-root] .pp-flow h2, [data-preview-printer-print-root] .pp-flow h3 { break-after: avoid; }',
        ...(extraCss === '' ? [] : [extraCss]),
      ].join('\n')

      const root = document.createElement('div')
      root.setAttribute('data-preview-printer-print-root', '')
      root.style.cssText = 'position: fixed; left: -20000px; top: 0; background: #fff;'
      const flow = document.createElement('div')
      flow.className = 'pp-flow'
      flow.style.cssText = 'width: 182mm;'
      flow.innerHTML = innerHtml
      root.appendChild(flow)
      return openPrintDialog(style, root, [], Promise.resolve())
    }

    /** Render Markdown source into a flowing, page-broken HTML document. */
    async function openMarkdownPrintout(text) {
      return openFlowingPrintout(renderMarkdown(text))
    }

    /** Print plain text verbatim: whitespace preserved, no Markdown interpretation. */
    async function openPlainTextPrintout(text) {
      return openFlowingPrintout(
        `<pre class="pp-plain">${escapeHtml(text)}</pre>`,
        '[data-preview-printer-print-root] .pp-flow .pp-plain { margin: 0; white-space: pre-wrap; word-break: break-word;'
        + ' font: 10.5pt/1.6 "Cascadia Mono", Consolas, "Microsoft YaHei", monospace; }',
      )
    }

    /** Toolbar glyph. */
    function PrinterIcon() {
      return h('svg', {
        width: 16, height: 16, viewBox: '0 0 16 16', fill: 'none',
        stroke: 'currentColor', strokeWidth: 1.3, strokeLinecap: 'round', strokeLinejoin: 'round',
        'aria-hidden': true, style: { display: 'block' },
      }, [
        h('path', { key: 'paper', d: 'M4.5 6V2.75h7V6' }),
        h('rect', { key: 'body', x: 2, y: 6, width: 12, height: 5, rx: 1.2 }),
        h('path', { key: 'out', d: 'M4.5 9.5h7v3.75h-7z' }),
      ])
    }

    /** Speed mark for the no-dialog print path. */
    function BoltIcon() {
      return h('svg', {
        width: 14, height: 14, viewBox: '0 0 16 16', fill: 'none',
        stroke: 'currentColor', strokeWidth: 1.4, strokeLinecap: 'round', strokeLinejoin: 'round',
        'aria-hidden': true, style: { display: 'block', flex: '0 0 auto' },
      }, [h('path', { key: 'bolt', d: 'M9 1.5 4 9h3.2L7 14.5 12 7H8.8z' })])
    }

    /** What one file is, independent of which print path it takes. */
    function iconKindOf(filePath) {
      if (OFFICE.test(filePath)) return 'office'
      if (/\.(xls|xlt)$/i.test(filePath)) return 'sheet'
      if (IMAGE.test(filePath)) return 'image'
      if (MARKDOWN.test(filePath)) return 'markdown'
      if (/\.json$/i.test(filePath)) return 'data'
      if (/\.(py|js|bat)$/i.test(filePath)) return 'code'
      return 'plain'
    }

    /** Locale key explaining why a previewable file has no print action. */
    function unsupportedReasonOf(filePath) {
      if (/\.(xlsx|csv|tsv)$/i.test(filePath)) return 'unsupportedSheet'
      if (/\.gif$/i.test(filePath)) return 'unsupportedGif'
      if (/\.(c|h|cc|cpp|cxx|hpp|hxx|inl|asm|inc|sh|mk)$/i.test(filePath)) return 'unsupportedCode'
      return 'unsupportedGeneric'
    }

    /** File-type mark beside the file name, so the panel states what it targets. */
    function KindIcon(props) {
      const paths = {
        office: [
          h('path', { key: 'sheet', d: 'M3 1.5h4l2 2v7H3z' }),
          h('path', { key: 'fold', d: 'M7 1.5v2h2' }),
          h('path', { key: 'lines', d: 'M4.3 5.4h3.4M4.3 7.2h3.4M4.3 9h2.2' }),
        ],
        sheet: [
          h('rect', { key: 'box', x: 1.5, y: 2, width: 9, height: 8, rx: 1 }),
          h('path', { key: 'grid', d: 'M1.5 4.8h9M1.5 7.4h9M5 2v8' }),
        ],
        image: [
          h('rect', { key: 'box', x: 1.5, y: 2.5, width: 9, height: 7, rx: 1 }),
          h('circle', { key: 'sun', cx: 4.2, cy: 5, r: 0.8 }),
          h('path', { key: 'hill', d: 'M2.2 8.6 4.8 6l1.6 1.6 1.3-1.1L10 8.6' }),
        ],
        // The Markdown mark (M with a down arrow), spread wider than the other
        // glyphs so the letterform survives the small render size.
        markdown: [
          h('path', { key: 'm', d: 'M1.6 9V4.2l2 2.6 2-2.6V9' }),
          h('path', { key: 'stem', d: 'M9.2 4.2v4.6' }),
          h('path', { key: 'head', d: 'M7.9 7.6 9.2 8.9l1.3-1.3' }),
        ],
        // Plain text reads as lines, data as braces, code as angle brackets.
        plain: [
          h('path', { key: 'a', d: 'M2.5 3.5h7M2.5 6h7M2.5 8.5h4.5' }),
        ],
        // Angular braces: straight segments rasterize cleanly at this size,
        // where the curved form turns to mush.
        data: [
          h('path', { key: 'curly', d: 'M4.6 2.6c-1.1 0-1.6.5-1.6 1.6v1.2c0 .7-.4 1.1-1 1.1.6 0 1 .4 1 1.1v1.2c0 1.1.5 1.6 1.6 1.6' }),
          h('path', { key: 'curly2', d: 'M7.4 2.6c1.1 0 1.6.5 1.6 1.6v1.2c0 .7.4 1.1 1 1.1-.6 0-1 .4-1 1.1v1.2c0 1.1-.5 1.6-1.6 1.6' }),
        ],
        code: [
          h('path', { key: 'left', d: 'M4.6 3.4 2 6l2.6 2.6' }),
          h('path', { key: 'right', d: 'M7.4 3.4 10 6l-2.6 2.6' }),
        ],
      }
      const shapes = paths[props.kind] ?? paths.plain
      // A 12px box cannot hold the markdown mark legibly: the glyph needs the
      // room, so every icon renders at 14px while the stroke stays hairline.
      const stroke = props.kind === 'markdown' ? 1 : 0.95
      return h('svg', {
        width: 14, height: 14, viewBox: '0 0 12 12', fill: 'none',
        stroke: 'currentColor', strokeWidth: stroke, strokeLinecap: 'round', strokeLinejoin: 'round',
        'aria-hidden': true, style: { display: 'block', flex: '0 0 auto' },
      }, shapes)
    }

    function PrintAction(props) {
      const { absolutePath, request, readBytes, openFile, t } = props
      const [open, setOpen] = React.useState(false)
      const [busy, setBusy] = React.useState(false)
      const [status, setStatus] = React.useState('')
      const [failed, setFailed] = React.useState(false)
      const [range, setRange] = React.useState('')
      const [dpi, setDpi] = React.useState(300)
      const [salvage, setSalvage] = React.useState(false)
      const [repairedPath, setRepairedPath] = React.useState('')
      const [repairFailed, setRepairFailed] = React.useState(false)
      const [convertedPath, setConvertedPath] = React.useState('')
      const [convertFailed, setConvertFailed] = React.useState(false)
      const [closeAt, setCloseAt] = React.useState(0)
      const buttonRef = React.useRef(null)
      const panelRef = React.useRef(null)
      const legacy = typeof absolutePath === 'string' && /\.(doc|ppt)$/i.test(absolutePath)
      const sheetLegacy = typeof absolutePath === 'string' && /\.(xls|xlt)$/i.test(absolutePath)
      const anchor = React.useMemo(() => {
        if (!open || buttonRef.current === null) return null
        const rect = buttonRef.current.getBoundingClientRect()
        return { top: rect.bottom + 6, right: Math.max(8, window.innerWidth - rect.right) }
      }, [open])

      React.useEffect(() => {
        if (!open) return undefined
        const onKey = (event) => { if (event.key === 'Escape') setOpen(false) }
        const onDown = (event) => {
          const panel = panelRef.current
          const button = buttonRef.current
          if (panel !== null && panel.contains(event.target)) return
          if (button !== null && button.contains(event.target)) return
          setOpen(false)
        }
        window.addEventListener('keydown', onKey)
        window.addEventListener('mousedown', onDown, true)
        return () => {
          window.removeEventListener('keydown', onKey)
          window.removeEventListener('mousedown', onDown, true)
        }
      }, [open])

      // A failed conversion leaves a failure mark inside this tab's frame: the
      // office body adds `data-textpreview-failed`, the spreadsheet body only a
      // `role="alert"` line. Watching them lets the toolbar react where the user
      // is already looking, instead of waiting for them to open the panel.
      React.useEffect(() => {
        const anchor = buttonRef.current
        if (anchor === null) return undefined
        const root = anchor.closest('[data-textpreview-state]')
        if (root === null) return undefined
        const check = () => {
          setSalvage(root.querySelector('[data-textpreview-failed], [role="alert"]') !== null)
        }
        check()
        const observer = new MutationObserver(check)
        observer.observe(root, { childList: true, subtree: true })
        return () => observer.disconnect()
      }, [])

      // Re-read the failure mark when the tab switches to another file, since the
      // observer only reports later mutations.
      React.useEffect(() => {
        const anchor = buttonRef.current
        const root = anchor === null ? null : anchor.closest('[data-textpreview-state]')
        setSalvage(root === null ? false : root.querySelector('[data-textpreview-failed], [role="alert"]') !== null)
        setRepairedPath('')
        setRepairFailed(false)
        setConvertedPath('')
        setConvertFailed(false)
      }, [absolutePath])

      // Repair on detection rather than on click, so the chip can state a fact
      // and the user ends up with a copy the preview can open.
      React.useEffect(() => {
        if (!legacy || !salvage || repairedPath !== '') return undefined
        let cancelled = false
        void (async () => {
          try {
            const result = await request({ mode: 'repair', path: absolutePath })
            if (cancelled) return
            if (!result.ok) { setRepairFailed(true); return }
            const text = String(result.text || '')
            const separator = text.indexOf('：')
            setRepairedPath(separator >= 0 ? text.slice(separator + 1) : text)
          } catch {
            // Printing still pads on demand when this pre-emptive pass fails.
            if (!cancelled) setRepairFailed(true)
          }
        })()
        return () => { cancelled = true }
      }, [legacy, salvage, absolutePath, repairedPath])

      // A spreadsheet the preview cannot parse gets converted once, then waits
      // for one click that opens the converted file in the same pane.
      React.useEffect(() => {
        if (!sheetLegacy || !salvage || convertedPath !== '' || convertFailed) return undefined
        let cancelled = false
        void (async () => {
          try {
            const result = await request({ mode: 'convert', path: absolutePath })
            if (cancelled) return
            if (!result.ok) { setConvertFailed(true); return }
            const text = String(result.text || '')
            const separator = text.indexOf('：')
            setConvertedPath(separator >= 0 ? text.slice(separator + 1) : text)
          } catch {
            if (!cancelled) setConvertFailed(true)
          }
        })()
        return () => { cancelled = true }
      }, [sheetLegacy, salvage, absolutePath, convertedPath, convertFailed])

      // A finished print or save dismisses itself; the status is readable for a
      // moment and the panel stops covering the preview afterwards.
      React.useEffect(() => {
        if (closeAt === 0) return undefined
        const timer = setTimeout(() => {
          setOpen(false)
          setStatus('')
          setFailed(false)
          setCloseAt(0)
        }, 3000)
        return () => clearTimeout(timer)
      }, [closeAt])

      if (typeof absolutePath !== 'string') return null
      const kind = pathKind(absolutePath)

      // Every previewable file keeps the control, so its presence proves the
      // plugin is loaded; an unsupported one renders dimmed and explains why
      // instead of vanishing.
      if (kind === undefined) {
        const reason = unsupportedReasonOf(absolutePath)
        return h('button', {
          type: 'button',
          disabled: true,
          'data-preview-printer': 'unsupported',
          'data-preview-printer-reason': reason,
          title: t(reason),
          'aria-label': t(reason),
          style: {
            display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
            width: 28, height: 28, padding: 0, border: 'none', borderRadius: 6,
            background: 'transparent', color: 'inherit', lineHeight: 0,
            opacity: 0.45, cursor: 'not-allowed',
          },
        }, h(PrinterIcon))
      }

      const office = kind === 'office'

      // Legacy spreadsheets never print through this plugin; their only entry is
      // the converted-copy chip. The wrapper renders even before a failure is
      // known, because the failure watcher needs a node inside the tab frame.
      if (kind === 'sheet') {
        const ready = convertedPath !== ''
        const chip = !salvage ? null : h('button', {
          type: 'button',
          'data-preview-printer-convert': convertFailed ? 'failed' : ready ? 'converted' : 'busy',
          title: convertFailed ? t('convertFailedTitle') : t('convertTitle'),
          disabled: !ready,
          onClick: () => { if (ready) openFile(convertedPath) },
          style: {
            padding: '2px 6px', borderRadius: 999,
            border: '1px solid #d4a72c', background: '#fffdf4', color: '#9a6700',
            font: 'inherit', fontSize: 11, lineHeight: 1.4, whiteSpace: 'nowrap',
            cursor: ready ? 'pointer' : 'progress',
          },
        }, convertFailed ? t('convertFailedChip') : ready ? t('convertChip') : t('convertBusy'))
        return h('span', {
          ref: buttonRef,
          'data-preview-printer-anchor': '',
          // `contents` keeps the wrapper boxless: no extra gap in the toolbar
          // before a chip exists, while the node still anchors the watcher.
          style: { display: 'contents' },
        }, chip)
      }

      const report = (text, isFailed) => {
        setStatus(text)
        setFailed(isFailed === true)
      }

      /** Report a finished action, then let the panel dismiss itself. */
      const done = (text) => {
        setStatus(text)
        setFailed(false)
        setCloseAt(Date.now())
      }

      const save = async () => {
        setBusy(true)
        report(t('preparing'), false)
        try {
          const result = await request({ mode: 'save', path: absolutePath })
          if (result.ok) { done(result.text || t('saved')) } else { report(result.text || t('failed'), true) }
        } catch (error) {
          report(error instanceof Error ? error.message : String(error), true)
        } finally {
          setBusy(false)
        }
      }

      const direct = async () => {
        setBusy(true)
        report(t('preparing'), false)
        try {
          const result = await request({ mode: 'print', path: absolutePath })
          if (result.ok) { done(result.text || t('dialogOpened')) } else { report(result.text || t('failed'), true) }
        } catch (error) {
          report(error instanceof Error ? error.message : String(error), true)
        } finally {
          setBusy(false)
        }
      }

      const printImage = async () => {
        setBusy(true)
        report(t('preparing'), false)
        try {
          const file = await readBytes(absolutePath)
          if (!file.ok) { report(file.text || t('failed'), true); return }
          const outcome = await openImagePrintout(file.data, mimeOf(absolutePath))
          if (!outcome.printed) { outcome.dispose(); report(t('noDialogImage'), true); return }
          done(t('dialogOpened'))
        } catch (error) {
          report(error instanceof Error ? error.message : String(error), true)
        } finally {
          setBusy(false)
        }
      }

      const saveImagePdf = async () => {
        setBusy(true)
        report(t('preparing'), false)
        try {
          const file = await readBytes(absolutePath)
          if (!file.ok) { report(file.text || t('failed'), true); return }
          const base = fileNameOf(absolutePath).replace(/\.[^.]+$/, '')
          await saveImageAsPdf(file.data, mimeOf(absolutePath), `${base}-print.pdf`)
          // The browser may save somewhere other than Downloads, so report the
          // outcome only, never a path this plugin cannot confirm.
          done(t('savedGeneric'))
        } catch (error) {
          report(error instanceof Error ? error.message : String(error), true)
        } finally {
          setBusy(false)
        }
      }

      const printMarkdownFile = async () => {
        setBusy(true)
        report(t('preparing'), false)
        try {
          const file = await readBytes(absolutePath)
          if (!file.ok) { report(file.text || t('failed'), true); return }
          const text = decodeText(file.data)
          const outcome = await openMarkdownPrintout(text)
          if (!outcome.printed) { outcome.dispose(); report(t('noDialogImage'), true); return }
          done(t('dialogOpened'))
        } catch (error) {
          report(error instanceof Error ? error.message : String(error), true)
        } finally {
          setBusy(false)
        }
      }

      const printPlainTextFile = async () => {
        setBusy(true)
        report(t('preparing'), false)
        try {
          const file = await readBytes(absolutePath)
          if (!file.ok) { report(file.text || t('failed'), true); return }
          if (file.data.length > MAX_TEXT_BYTES) {
            report(t('textTooLarge').replace('{size}', humanBytes(file.data.length)).replace('{limit}', humanBytes(MAX_TEXT_BYTES)), true)
            return
          }
          let text = decodeText(file.data)
          const lines = text.split('\n').length
          if (lines > MAX_TEXT_LINES) {
            report(t('textTooManyLines').replace('{lines}', String(lines)).replace('{limit}', String(MAX_TEXT_LINES)), true)
            return
          }
          if (/\.json$/i.test(absolutePath)) text = prettyJson(text)
          const outcome = await openPlainTextPrintout(text)
          if (!outcome.printed) { outcome.dispose(); report(t('noDialogImage'), true); return }
          done(t('dialogOpened'))
        } catch (error) {
          report(error instanceof Error ? error.message : String(error), true)
        } finally {
          setBusy(false)
        }
      }

      const printOffice = async () => {
        setBusy(true)
        report(t('preparing'), false)
        try {
          const prepared = await request({ mode: 'prepare', path: absolutePath, range, dpi })
          if (!prepared.ok) { report(prepared.text || t('failed'), true); return }
          const text = String(prepared.text || '')
          const manifestPath = text.startsWith(MANIFEST_MARKER) ? text.slice(MANIFEST_MARKER.length).trim() : ''
          if (manifestPath === '') { report(text || t('failed'), true); return }

          const manifestFile = await readBytes(manifestPath)
          if (!manifestFile.ok) { report(manifestFile.text || t('failed'), true); return }
          const manifest = JSON.parse(new TextDecoder().decode(manifestFile.data))

          const pages = []
          for (const page of manifest.pages) {
            const file = await readBytes(page.path)
            if (!file.ok) { report(file.text || t('failed'), true); return }
            pages.push({ ...page, data: file.data })
          }
          const outcome = await printPages(pages, manifest.dpi)
          if (!outcome.printed) {
            outcome.dispose()
            const fallback = await request({ mode: 'print', path: absolutePath })
            if (fallback.ok) { done(`${t('noDialog')} ${fallback.text}`) } else { report(fallback.text, true) }
            return
          }
          const repairedText = manifest.repaired === true
            ? (typeof manifest.repairedPath === 'string' && manifest.repairedPath !== ''
                ? t('repairedSaved').replace('{path}', manifest.repairedPath)
                : t('repairedNote'))
            : ''
          done(repairedText === '' ? t('dialogOpened') : `${t('dialogOpened')} ${repairedText}`)
        } catch (error) {
          report(error instanceof Error ? error.message : String(error), true)
        } finally {
          setBusy(false)
        }
      }

      const runPrimary = () => {
        if (kind === 'office') { void printOffice(); return }
        if (kind === 'image') { void printImage(); return }
        if (kind === 'text') { void printPlainTextFile(); return }
        void printMarkdownFile()
      }

      const button = h('button', {
        ref: buttonRef,
        type: 'button',
        title: t('print'),
        'aria-label': t('print'),
        'aria-expanded': open,
        'data-preview-printer': busy ? 'busy' : 'idle',
        onClick: async () => {
          setOpen(value => !value)
          setStatus('')
          setFailed(false)
        },
        style: {
          display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
          width: 28, height: 28, padding: 0, border: 'none', borderRadius: 6,
          background: open ? 'rgba(0,0,0,0.08)' : 'transparent',
          color: 'inherit', lineHeight: 0, cursor: busy ? 'progress' : 'pointer',
        },
      }, h(PrinterIcon))

      const salvageChip = (legacy && salvage)
        ? h('button', {
          key: 'salvage',
          type: 'button',
          'data-preview-printer-salvage': repairedPath === '' ? (repairFailed ? 'failed' : 'repairing') : 'repaired',
          title: repairFailed ? t('salvageFallbackTitle') : t('salvageTitle'),
          onClick: () => {
            setOpen(true)
            setFailed(false)
            if (repairedPath !== '') { setStatus(t('salvageSaved').replace('{path}', repairedPath)); return }
            setStatus(repairFailed ? t('salvageFallbackTitle') : t('salvageBusy'))
          },
          style: {
            marginLeft: 4, padding: '2px 6px', borderRadius: 999,
            border: '1px solid #d4a72c', background: '#fffdf4', color: '#9a6700',
            font: 'inherit', fontSize: 11, lineHeight: 1.4, cursor: 'pointer', whiteSpace: 'nowrap',
          },
        }, repairedPath === '' ? (repairFailed ? t('salvageFallback') : t('salvageBusy')) : t('salvageChip'))
        : null

      if (!open || anchor === null) {
        return h('div', { style: { display: 'inline-flex', alignItems: 'center' } }, [button, salvageChip])
      }

      const children = [
        h('div', {
          key: 'header',
          style: {
            padding: '11px 12px 9px',
            borderBottom: '1px solid #eaeef2',
            background: '#f6f8fa',
            textAlign: 'center',
          },
        }, [
          h('div', { key: 'title', style: { fontWeight: 600, fontSize: 12.5 } }, [
            t('appTitle'),
            h('span', {
              key: 'version',
              style: { marginLeft: 5, fontSize: 10.5, fontWeight: 400, color: '#8b949e', letterSpacing: '0.02em' },
            }, `v${VERSION}`),
          ]),
          h('div', {
            key: 'file',
            title: absolutePath,
            style: {
              marginTop: 3, display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 4,
              color: '#6e7781', fontSize: 11, minWidth: 0,
            },
          }, [
            h(KindIcon, { key: 'icon', kind: iconKindOf(absolutePath) }),
            h('span', {
              key: 'name',
              style: { overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' },
            }, fileNameOf(absolutePath)),
          ]),
        ]),
      ]

      if (office) {
        children.push(h('label', { key: 'range', style: { display: 'block', marginBottom: 6 } }, [
          h('span', { key: 'text', style: { display: 'block', color: '#57606a', marginBottom: 3 } }, t('rangeLabel')),
          h('input', {
            key: 'input',
            type: 'text',
            value: range,
            placeholder: t('rangePlaceholder'),
            disabled: busy,
            onChange: (event) => { setRange(event.target.value) },
            onKeyDown: (event) => { if (event.key === 'Enter' && !busy) void printOffice() },
            style: {
              width: '100%', boxSizing: 'border-box', padding: '6px 8px', borderRadius: 6,
              border: '1px solid #d0d7de', font: 'inherit', background: '#fff', color: 'inherit',
            },
          }),
        ]))
        children.push(h('label', { key: 'dpi', style: { display: 'block', marginBottom: 6 } }, [
          h('span', { key: 'text', style: { display: 'block', color: '#57606a', marginBottom: 3 } }, t('resolution')),
          h('select', {
            key: 'select',
            value: String(dpi),
            disabled: busy,
            onChange: (event) => { setDpi(Number(event.target.value)) },
            style: {
              width: '100%', boxSizing: 'border-box', padding: '6px 8px', borderRadius: 6,
              border: '1px solid #d0d7de', font: 'inherit', background: '#fff', color: 'inherit',
            },
          }, [
            h('option', { key: '200', value: '200' }, t('resolutionStandard')),
            h('option', { key: '300', value: '300' }, t('resolutionHigh')),
            h('option', { key: '400', value: '400' }, t('resolutionUltra')),
          ]),
        ]))
      }

      children.push(h('div', { key: 'actions', style: { display: 'flex', gap: 8 } }, [
        h('button', {
          key: 'print',
          type: 'button',
          disabled: busy,
          onClick: runPrimary,
          style: {
            flex: 1, padding: '6px 10px', borderRadius: 6, border: '1px solid #1f6feb',
            background: busy ? '#9dbcf5' : '#1f6feb', color: '#fff', font: 'inherit', cursor: busy ? 'progress' : 'pointer',
          },
        }, t('print')),
        (office || kind === 'image') ? h('button', {
          key: 'save',
          type: 'button',
          disabled: busy,
          onClick: () => { void (office ? save() : saveImagePdf()) },
          style: {
            flex: 1, padding: '6px 10px', borderRadius: 6, border: '1px solid #d0d7de',
            background: '#f6f8fa', color: '#1f2328', font: 'inherit', cursor: busy ? 'progress' : 'pointer',
          },
        }, t('save')) : null,
      ]))

      if (office) {
        children.push(h('div', {
          key: 'direct-block',
          style: { marginTop: 8, paddingTop: 7, borderTop: '1px solid #eaeef2' },
        }, [
          h('button', {
            key: 'direct',
            type: 'button',
            disabled: busy,
            'data-preview-printer-direct': '',
            onClick: () => { void direct() },
            style: {
              display: 'flex', alignItems: 'center', gap: 6, width: '100%', boxSizing: 'border-box',
              padding: '6px 10px', borderRadius: 6,
              border: '1px dashed #d4a72c', background: '#fffdf4', color: '#9a6700',
              font: 'inherit', fontWeight: 600, cursor: busy ? 'progress' : 'pointer',
            },
          }, [h(BoltIcon, { key: 'icon' }), h('span', { key: 'text' }, t('direct'))]),
          h('div', { key: 'hint', style: { marginTop: 4, color: '#8b949e', fontSize: 11 } }, t('directHint')),
        ]))
      }

      if (status !== '') {
        children.push(h('div', {
          key: 'status',
          'data-preview-printer-status': failed ? 'failed' : 'ok',
          style: { marginTop: 8, color: failed ? '#cf222e' : '#57606a', wordBreak: 'break-word' },
        }, status))
      }

      // Only a failed legacy preview needs the explanation; a healthy `.doc`
      // shows nothing, so the note never reads as a problem that is not there.
      if (legacy && salvage) {
        children.push(h('div', {
          key: 'legacy',
          'data-preview-printer-legacy': '',
          style: { marginTop: 8, color: '#8b949e', fontSize: 11, lineHeight: 1.5 },
        }, t('legacyHint')))
      }

      const panel = h('div', {
        ref: panelRef,
        role: 'dialog',
        'data-preview-printer-panel': '',
        style: {
          position: 'fixed', top: anchor.top, right: anchor.right, width: 268, zIndex: 2147483000,
          background: '#ffffff', color: '#1f2328', border: '1px solid #d8dee4', borderRadius: 10,
          overflow: 'hidden',
          boxShadow: '0 8px 24px rgba(15, 23, 42, 0.16)', padding: 0, font: '12px/1.5 system-ui, sans-serif',
        },
      }, [
        children[0],
        h('div', { key: 'body', style: { padding: '10px 12px 12px' } }, children.slice(1)),
      ])

      return h('div', { style: { display: 'inline-flex', alignItems: 'center' } }, [button, salvageChip, panel])
    }

    function apply(ctx) {
      ctx.effect(() => ctx.locale.register(NS, { zh, en }), 'preview-printer: dictionaries')
      const dispose = ctx.slots.inject('sidebar.right.tab.document.actions', () => ctx.slots.register({
        name: 'sidebar.right.tab.document.actions',
        id: 'preview-printer',
        order: 10,
        locale: NS,
        inject: (sessionId) => ({
          /**
           * Run one Host request for the previewed file.
           * @param payload - `{ mode, path, range, dpi }` describing the operation.
           * @returns whether the Host settled it successfully, with its text.
           */
          request: async (payload) => {
            const line = `/${COMMAND} ${encodeRequest(payload)}`
            const result = await ctx.remote.commands.execute(sessionId, line, [])
            if (!result.ok) return { ok: false, text: `${result.error.message} (${result.error.code})` }
            if (result.value === undefined) return { ok: false, text: `未找到命令 /${COMMAND}` }
            const settled = result.value.result
            return settled.kind === 'success'
              ? { ok: true, text: settled.text }
              : { ok: false, text: settled.text }
          },
          /**
           * Read one file of the Session workspace.
           * @param filePath - absolute path inside the workspace.
           * @returns the bytes, or the Host's failure text.
           */
          readBytes: async (filePath) => {
            const result = await ctx.remote.workspaceFiles.readBytes(sessionId, filePath, {})
            if (!result.ok) return { ok: false, text: `${result.error.message} (${result.error.code})` }
            return { ok: true, data: result.value.data }
          },
          /**
           * Open one workspace file in this pane, replacing the current preview.
           * The Sidebar service is looked up lazily so a missing provider cannot
           * deactivate the whole plugin through `inject`.
           * @param filePath - absolute path inside the Session workspace.
           * @returns whether the Sidebar accepted the request.
           */
          openFile: (filePath) => {
            try {
              const sidebar = typeof ctx.get === 'function' ? ctx.get('sidebarRight') : undefined
              if (sidebar === undefined || typeof sidebar.openResource !== 'function') return false
              sidebar.openResource(fileAddress(sessionId, filePath), { replaceTab: true })
              return true
            } catch {
              return false
            }
          },
        }),
      }, PrintAction))
      ctx.effect(() => dispose, 'preview-printer: action')
    }

    return { inject: ['slots', 'locale', 'remote', 'remote.commands', 'remote.workspaceFiles'], apply }
  },
})
