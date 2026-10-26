# Office Reader

[中文文档](README.zh-CN.md) | [Roadmap](docs/ROADMAP.md)

> **Upstream:** This project is maintained at
> [`yulimfish/obsidian-word-reader`](https://github.com/yulimfish/obsidian-word-reader)
> and is based on the upstream project
> [`qianwei4712/obsidian-word-reader`](https://github.com/qianwei4712/obsidian-word-reader).
> The upstream authors retain credit for the original work.

Office Reader is a desktop-only Obsidian plugin for opening `.docx`, `.pptx`,
and `.xlsx` files directly inside Obsidian as safe, read-only documents.

The plugin is not an Office editor. It provides local reading workflows while
keeping the original files unchanged.

Requires Obsidian Desktop 1.12.7 or newer.

## Features

- Open `.docx` files in an Obsidian tab with headings, paragraphs, lists,
  tables, images, page breaks, outline navigation, search, zoom, and
  text/Markdown copy.
- Open `.pptx` files in an Obsidian tab with local rendering for text,
  images, common shapes, tables, themes, layouts, and masters, plus
  thumbnails, slide search, speaker notes, fullscreen reading, and summary
  notes.
- Open `.xlsx` files in a bounded virtual grid with worksheet tabs, frozen
  panes, merged cells, basic styles, cached formula results, workbook-wide
  search, hidden-sheet indicators, TSV/Markdown export, and summary notes.
- Restore zoom, fit mode, navigation state, and reading position for the 100
  most recently used Office documents.
- Switch the plugin interface between Chinese and English.
- Create same-name Markdown summary notes linked back to source Office files.
- Copy privacy-safe diagnostics for compatibility or performance reports
  without including document content, internal XML, or absolute vault paths.

## Supported files

| Extension | Status | Notes |
| --- | --- | --- |
| `.docx` | Supported | Rendered inside Obsidian through `docx-preview`. |
| `.pptx` | Supported | Rendered locally as a read-only presentation. |
| `.xlsx` | Supported | Rendered locally as a read-only virtualized spreadsheet. |
| `.doc` | Guidance page | Shown inside Obsidian with external-open and `.docx` conversion guidance. |
| `.xls` | Guidance page | Shown inside Obsidian with external-open and `.xlsx` conversion guidance. |
| `.xlsm` | Not registered | Macro-enabled workbooks are outside the supported scope. |

## Usage notes

- Put a `.docx`, `.pptx`, or `.xlsx` file into your vault and click it in the
  file explorer to open a read-only tab.
- The plugin is read-only and never saves changes back to the original Office
  file.
- `.xlsm` is not registered; `.doc` and `.xls` show guidance pages instead of
  rendering.
- Encrypted or password-protected documents show dedicated error pages.
- Four release files (`main.js`, `docx-engine.cjs`, `manifest.json`,
  `styles.css`) must stay together when installing or upgrading manually.
- Workbook hyperlinks open only after an explicit click and confirmation.

Detailed usage, compatibility, security, and known-limit notes live in
[docs/usage-notes.md](docs/usage-notes.md).

## Development notes

- Install dependencies with `npm install`.
- Run the full local gate with `npm run check`.
- Run release metadata validation with `npm run release:check`.
- Production builds enforce a 500 KiB `dist/main.js` budget.
- CI and release workflows use Node.js 20.19.0.
- Release packages contain only `main.js`, `docx-engine.cjs`,
  `manifest.json`, and `styles.css`.

Development, install, and release details live in
[docs/development-release.md](docs/development-release.md).

## Upstream and license notes

- Upstream attribution is shown above; this repository continues the project
  under the MIT License.
- Runtime dependency notices are inventoried in
  [docs/THIRD_PARTY_LICENSES.md](docs/THIRD_PARTY_LICENSES.md).
- Release scope, support boundaries, and the manual checklist are maintained
  in [docs/STABILITY.md](docs/STABILITY.md).
