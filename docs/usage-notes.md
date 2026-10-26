# Usage notes

## Basic reading flow

1. Put a `.docx`, `.pptx`, or `.xlsx` file into your Obsidian vault.
2. Click the file in the file explorer.
3. Read the document in the Obsidian tab opened by the plugin.
4. Use the toolbar and navigation panels to reload, navigate, zoom, search,
   copy text, inspect notes, create a summary note, or open the file externally.

## Zoom

- Type a percentage in the zoom input.
- Use `Ctrl` + mouse wheel over the document preview for continuous zoom.
- Use fit width when you want the document to match the current pane.

## PowerPoint reading

- Use the thumbnail and title sidebar to inspect the presentation and jump
  directly to a slide.
- Search all slide titles, body text, tables, and speaker notes from the
  sidebar. Press `Ctrl`/`Cmd` + `F` to focus presentation search.
- Use the arrow buttons, `Page Up`/`Page Down`, `Left`/`Right`, or
  `Space`/`Shift` + `Space` to change slides. Press `Home` or `End` to jump to
  the first or last slide.
- Enter a slide number to jump directly to it.
- Use the percentage input or `Ctrl` + mouse wheel to zoom.
- Use fit window to keep the complete slide visible while resizing a pane.
- Use fullscreen for presentation-focused reading.
- Copy selected text from the rendered slide, or copy all extracted text from
  the current slide when no selection is active.
- Show or hide the current slide's speaker notes.
- Create a same-name presentation summary note with the current slide and
  numbered references for every slide.
- Copy performance diagnostics from the toolbar or command palette when
  reporting an Office compatibility or performance issue.
- The current slide, zoom, fit mode, scroll position, navigation visibility,
  and speaker-note visibility are restored per file.

## Spreadsheet reading

- Switch among visible worksheets with the tabs below the grid. A separate
  indicator lists hidden and very-hidden worksheets and opens them only when
  requested without changing the source workbook.
- Enter `A1`, a rectangular range, a worksheet-qualified reference, or a
  supported workbook/sheet-scoped named range in the formula-bar name box to
  jump directly to it.
- Scroll large sparse worksheets without mounting their complete row and
  column range. The grid keeps only the viewport, finite overscan, and visible
  frozen rows/columns mounted.
- Use the mouse to drag a rectangular selection, `Shift` + click to extend it,
  or the arrow, page, `Home`, and `End` keys to move the active cell.
- Search values and stored formula text across the complete workbook,
  including hidden worksheets. Press `Ctrl`/`Cmd` + `F` to focus search and
  use `Enter`/`Shift` + `Enter` to move between cross-sheet results.
- Copy displayed values as TSV, export the selected rectangle as a Markdown
  table, or use the separate formula-copy action. Clipboard operations are
  capped at 250,000 cells.
- Inspect the selected cell's complete stored formula text and
  workbook-cached result in the read-only formula bar.
- Read legacy cell-comment authors and complete text from the formula bar.
  Comment markers remain discoverable even when the commented cell is blank.
- Preview package-local PNG, JPEG, GIF, and BMP drawings at their worksheet
  anchors. Common bar, line, area, and pie charts use only workbook-cached
  series data and show an explicit fallback for unsupported chart types.
- Preserve a bounded subset of numeric conditional formatting: literal
  `cellIs` rules, two-/three-color scales, and data bars.
- Create a same-name spreadsheet summary note containing sheet dimensions and
  visibility, named ranges, rich-content counts, the current selection, and up
  to 200 displayed cell previews collected from workbook content.
- Workbook hyperlinks open only after an explicit click and confirmation.
  External workbook references and data connections are never fetched.
- The active worksheet, zoom/fit preference, and scroll position are restored
  per file without persisting cell values or formulas.

## Image preview

Click an image rendered from the Word document to open a larger preview modal.
Press `Esc` or close the modal to return to the document.

Inside the image preview:

- Use the mouse wheel to zoom.
- Drag the image to pan.
- Double-click to fit the image back into the preview.
- Use the toolbar to fit to window, view actual size, copy the image, or save
  the image.
- Press `Ctrl` + `C` on Windows/Linux or `Cmd` + `C` on macOS while the image
  preview is open to copy the image.
- Saved images use a default name based on the source document and image
  dimensions when available.
- The toolbar shows the original image dimensions and current zoom.

## Search

- Type in the search field to highlight matching rendered text.
- Use the up and down buttons to move between results.
- Press `Enter` for the next result or `Shift` + `Enter` for the previous
  result.
- The result counter shows the current result and total matches.

## Outline

- Use the outline button to show or hide the heading outline.
- Click an outline item to scroll the rendered document to that heading.
- Collapse or expand outline sections with the chevron beside parent headings.
- The outline highlights the current section as the document scrolls.
- The outline is built from headings found in the rendered Word document.

## Copy

- Copy selected rendered text as plain text.
- Copy the whole document as plain text when no selection is active.
- Copy selected rendered content as Markdown.
- Copy the whole `.docx` as Markdown when no selection is active.

## Settings

Open Obsidian settings, then go to `Community plugins` -> `Office Reader`.

Available settings:

- Plugin interface language.
- Default zoom percentage for newly opened Word previews.
- Whether newly opened Word previews should fit the pane width by default.
- Whether newly opened worksheets should fit the grid width by default.
- Whether the outline is visible by default.
- Whether rendered images can be clicked for larger preview.
- Large file warning threshold in MB.
- External opening note. The plugin uses the operating system default
  application for source Office files.

## Summary notes

The summary note action creates a same-name Markdown file next to a Word
document, PowerPoint presentation, or Excel workbook.

Example:

```text
Report.docx
Report.md
Quarterly review.pptx
Quarterly review.md
Data register.xlsx
Data register.md
```

The generated note includes frontmatter and starter sections:

```markdown
---
source: "Report.docx"
type: word-note
reader: office-reader
reader_format: docx
created: 2026-05-28
---

# Report

Source: [[Report.docx]]

## Summary

## Key findings

## Follow-ups

## Quoted excerpts
```

Presentation notes use `type: presentation-note` and
`reader_format: pptx`, record the current slide, and include a numbered
reference list for every slide. If the same-name Markdown file already exists,
the plugin opens it without overwriting content.

Spreadsheet notes use `type: spreadsheet-note` and `reader_format: xlsx`,
record the current worksheet and selection, list worksheet dimensions,
visibility, supported named ranges, and comment/image/chart/conditional-rule
counts, and include at most 200 displayed-value preview cells. Existing
same-name Markdown files are opened without being overwritten.

## Compatibility and errors

- Legacy `.doc` files open to an explanation page instead of being rendered
  directly.
- Encrypted or password-protected documents show a dedicated encrypted
  document message.
- Failed `.docx` previews are classified as encrypted, format mismatched,
  damaged ZIP packages, invalid XML structures, unsupported document
  structures, or unknown failures where possible.
- Failed `.pptx` previews distinguish format mismatch, encryption, damaged
  XML or ZIP data, unsupported structures, and safe-preview limit violations.
- Failed `.xlsx` previews distinguish format mismatch, encryption, damaged
  workbook XML/ZIP data, unsupported active content, and safe-preview limits.
- Legacy `.xls` files show conversion guidance and an external-open action.
- PPTX archives are checked before decompression for file count, per-entry
  expanded size, total expanded size, ZIP64 usage, encryption, and abnormal
  compression ratios.
- Error pages provide a collapsed diagnostic section and a copy action for
  issue reports.
- Copied diagnostics use a shared, versioned JSON envelope and include only
  the format, diagnostic kind, file name, size, modification time, a
  privacy-safe summary, and bounded format metrics. Raw renderer errors,
  document content, internal XML, and absolute vault paths are excluded.
- Successful DOCX, PPTX, and XLSX previews expose the same JSON performance
  snapshot with package/parse/first-readable/search/navigation/scroll
  timings, peak heap when Chromium exposes it, actual DOM, bounded-cache,
  retained resource, cancellation, cleanup, maximum work slice,
  yield/preemption count, and memory-pressure observations. Uninterruptible
  third-party DOCX rendering is labeled separately. Document text, speaker
  notes, cell values, internal XML, and absolute vault paths are excluded.
- Large file warnings include the file size and use the configured threshold
  from settings.

## Security notes

- **Local-only operations**: The plugin reads `.docx`, `.pptx`, and `.xlsx`
  from your local Obsidian vault. It never automatically fetches workbook
  links, external references, data connections, or remote resources. A safe
  hyperlink may open only after the user clicks and confirms it.
- **No external resources**: The plugin never loads scripts, styles, or
  assets from the internet. All rendering logic runs locally.
- **Read-only access**: The plugin never modifies, overwrites, or writes back
  to the original Office file. It uses Obsidian's binary vault API for
  rendering and text extraction only.
- **No dynamic script injection**: The plugin creates only structural DOM
  elements (`div`, `span`, `button`, `input`) for document rendering. No
  `<script>` elements are created or injected at any point.
- **Safe PPTX archives**: ZIP metadata is validated before decompression, and
  external package relationships are ignored instead of being loaded.
- **Safe XLSX formulas**: Formula text and cached results are displayed
  without running macros, external formula code, or recalculation.
- **Safe XLSX rich content**: Images are restricted to package-local safe
  raster formats; chart previews use cached series only. External drawing,
  chart, and workbook relationships are never resolved. The reviewed
  [XLSM compatibility matrix](XLSM_COMPATIBILITY.md) keeps `.xlsm`
  unregistered.
- **Sandboxed rendering**: Office content is rendered into structural DOM
  containers with no script execution context.
- **Desktop-only**: The plugin requires desktop Obsidian because it uses
  Electron APIs for image clipboard operations and file dialogs. This is
  declared in `manifest.json` as `isDesktopOnly: true`.

## Known limits

- This is not an Office editor.
- The plugin never saves changes back to `.docx`, `.pptx`, or `.xlsx`.
- Legacy `.xls` files are not rendered directly, and `.xlsm` is not
  registered. Modern threaded comments, floating-object editing, pivot
  tables, slicers, pixel-perfect chart fidelity, and complex conditional
  formatting are outside the 3.2.0 XLSX scope.
- Legacy `.doc` files are not rendered directly, but the plugin shows
  external-open and conversion guidance.
- Complex Word layouts may not render exactly like Microsoft Word.
- PPTX animation, transitions, audio/video playback, macros, editing, charts,
  SmartArt, SVG/GIF/WebP media, and pixel-perfect PowerPoint fidelity are not
  supported.
- Charts and SmartArt are identified with local placeholders and diagnostics;
  their native PowerPoint appearance is not rendered.
- Presentation thumbnails are generated locally on demand. Very large or
  image-heavy slides can still take longer when they enter the visible range.
- Word previews follow the current Obsidian theme, but explicit colors stored
  in the Word document may still influence the rendered result.
- Very large files or files with many images may render slowly.
- Mobile support is not included in this version.

## Recommended workflow

Keep original Office files in the vault, open them for reading, and create
Markdown summary notes for long-term knowledge management.

This keeps the original document format intact while bringing summaries,
decisions, tasks, and quotes into your Obsidian workflow.
