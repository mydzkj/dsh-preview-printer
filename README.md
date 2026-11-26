# dsh-preview-printer

English | [中文](README.zh-CN.md)

Print documents, images, Markdown and source text from the right Sidebar document preview — and rescue legacy files the preview cannot open.

## What it adds

One control in the document toolbar (`sidebar.right.tab.document.actions`). Which files it appears for, and what it does:

| Previewed file | Control | Behaviour |
|---|---|---|
| `.pdf`, `.docx`, `.doc`, `.pptx`, `.ppt` | panel | Pages + resolution, then the Windows print dialog, or save the PDF beside the source, or Full-speed HD print |
| `.png`, `.jpg`/`.jpeg`, `.webp`, `.bmp` | panel | One image per A4 sheet, oriented by its own aspect ratio, fitted without cropping; **Save as PDF** builds a one-page PDF in the browser and downloads it |
| `.md`, `.markdown` | panel | Rendered (headings, lists, quotes, fenced code, emphasis) into a flowing A4 document |
| `.txt`, `.py`, `.json`, `.js`, `.bat` | panel | Printed as text: UTF-8 first, GBK when the bytes are not valid UTF-8; JSON is re-indented when it parses; files over 2 MB or 5,000 lines are refused with a reason |
| `.c`, `.h`, `.cc`, `.cpp`, `.cxx`, `.hpp`, `.asm`, `.sh` — and extensionless files such as `Makefile` | dimmed button | **Printing is not supported (yet).** The control still renders so the plugin is visibly present, but it is disabled and its tooltip explains that the source can be exported to PDF first |
| `.xls`, `.xlt` | chip only | Never printed. When the preview cannot parse the file, the plugin converts it once (real OLE via office-kit, text exports via a built-in xlsx writer) and offers an 「已转化」 chip that opens the converted workbook |
| `.xlsx`, `.csv`, `.tsv`, `.gif` | dimmed button | Native preview only. The control is still rendered — so the plugin is visibly present — but disabled, with a tooltip pointing at Excel/WPS or explaining that the format is not supported yet |

Panel options for Office files:

- **Pages** — empty prints every page. Separators may be an ASCII or full-width comma, the enumeration comma (顿号), a semicolon, a slash, or whitespace; ranges accept `-`, `~`, en/em dashes, the full-width hyphen, or 至. For example `1-3`, `1、3、5`, `1 至 3` and `１－２` all work.
- **Resolution** — 200 / 300 (default) / 400 DPI rasterization.
- **Print** — the Host rasterizes exactly those pages, the browser lays them out as paper-sized images and calls `window.print()`, so the **Windows print dialog** appears: pick a printer, or choose *Microsoft Print to PDF*. No external PDF application takes part.
- **Save as PDF** — for documents, copies the same PDF beside the source as `<name>-print.pdf` (`-2`, `-3`, … when taken). For images, **Save as PDF** assembles a one-page PDF in the browser (RGB samples, Flate-compressed) and downloads it to the system Downloads folder; it never touches the print pipeline.
- **Full-speed HD print** — hands the source PDF to the registered system `Print` verb: vector output, no dialog, straight to the default printer. The PDF application decides whether to show anything; WPS prints silently.
- The toolbar button only opens the panel; every action starts from a button inside it, so nothing happens on a stray click. A finished print or save reports its outcome and closes the panel after three seconds.

## Where the PDF comes from

- **PDF** — the original file; no conversion.
- **DOCX / PPTX** — the Host converts with the same `officeKit` service the preview renders through. The conversion is content-addressed, so an unchanged file produces the very PDF the preview shows.
- **DOC / PPT** — the same converter, which already renders legacy formats for the preview (`doc` and `ppt` are among the preview's own extensions); page selection and DPI apply as usual.
- **Images, Markdown, text** — no PDF at all: the browser prints the bitmap or the rendered document directly.

## Behaviour notes

- The dialog path prints raster images; 300 DPI is crisp for worksheets on a laser printer, 400 DPI is sharper and slower. A page too large for the rasterizer's pixel ceiling steps down automatically (300 → 200 → 150).
- Mixed page orientation is preserved: portrait and landscape pages get their own named `@page` rule.
- Without a `beforeprint` event within 1.5 s for Office files the browser falls back to the system print verb and says so.
- If the print verb cannot start, the same PDF is copied beside the source.
- A legacy `.doc`/`.ppt` whose compound file lost its final partial sector is rejected by the converter (`invalid-document`). The plugin then retries once with a zero-padded copy and **saves that copy beside the source as `<name>-修复.<ext>`**; the original file is never modified. Print and save report the repaired path.
- When the preview itself failed for such a file, the toolbar grows an 「已修复」 chip (「修复中…」 while the padding runs, 「点此打印」 if the pre-emptive repair failed): it watches the preview's own failure mark, so the remedy appears where the user is already looking instead of hiding inside the panel.
- Prepared page images live in `<workspace>\.dsh-print\` and are pruned after two hours.
- Every Office click issues one Host command with `recordInput: false`; the request payload is not written to the session log. Nothing in this plugin calls a model, so printing consumes no tokens.
- Only files inside the current workspace are printed or converted.

## Known gaps

- `.c`, `.cpp`, `.h`, `.asm`, `.sh` and extensionless files preview through DSH where the extension is recognised but cannot be printed; the toolbar control is dimmed and explains why.
- `.gif` is not supported at all (neither preview nor print).
- Markdown tables print as plain paragraphs; there is no table layout yet.
- `.xlsx`, `.csv` and `.tsv` are preview-only; their toolbar control is dimmed with a "print from Excel or WPS" hint.
- The dialog path rasterises documents; only Full-speed HD print is vector.
- The plugin targets DSH 0.2.2-rc.5 and its client APIs; a future DSH release may change them.

## Manual

`docs/dsh-preview-printer-manual-zh.docx` is a three-page Chinese guide that ships with the package; it doubles as a ready-made print sample. The same document as a PDF is published as a release asset rather than committed to the package.
## Privacy

The plugin makes no network requests and runs entirely on the local machine:

- No telemetry, analytics, or crash reporting.
- No conversation or file content is uploaded anywhere. Printing goes to the local Windows print dialog; document conversion uses the LibreOffice sidecar bundled with DSH.
- Files are read only when you open the preview or press a control, and only from the current workspace.

## License

MIT — see [LICENSE](LICENSE). Copyright (c) 2026 Qingguo Li.
