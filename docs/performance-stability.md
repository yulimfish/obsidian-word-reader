# Performance and stability notes

For release-scope details and support boundaries, see [STABILITY.md](STABILITY.md).

## Automated stability gates

- `npm run lifecycle:check` runs 100 open/reload/navigation/switch/close
  cycles for DOCX, PPTX, XLSX, and a mixed sequence under forced GC. It gates
  Blob URLs, caches, tasks, timers, observers, listeners, stale UI commits,
  retained-heap trends, recoverable package errors, and unhandled rejections.
  Full JSON, a compact trend point, and a Markdown summary are written under
  `benchmark-results/` and uploaded by CI for 90 days.
- `npm run performance:check` executes five isolated DOCX, PPTX, and XLSX
  samples, gates p95 regressions by format and stage, and writes full JSON,
  compact CI trend JSON, and a Markdown summary under `benchmark-results/`.
  CI uploads these artifacts for cross-run comparison. The manual Obsidian
  Desktop calibration worksheet is in `benchmarks/`.
- `npm run large-files:extended` runs three isolated samples for 500-page
  DOCX, 2,000-slide PPTX, an Excel-row-limit sparse XLSX, and a
  50,000-row/400,000-cell dense XLSX. It runs weekly and before release, not
  on every commit, and writes matching machine-readable trend artifacts.

## Shared architecture and runtime policies

- DOCX DOM commit, PPTX metadata/thumbnail work, and XLSX stream parsing
  share an adaptive 8 ms work-slice policy. Foreground navigation preempts
  background work; elevated memory pressure lowers batches/concurrency, and
  critical pressure stops preload and releases inactive caches.
- A shared reader shell plus `OfficeReaderAdapter`, `ReaderSession`, and
  `ReaderCapabilities` contracts keep toolbars, status, errors, diagnostics,
  and cleanup behavior consistent. DOCX, PPTX, and XLSX rendering logic lives
  in format-specific adapters and sessions.
- Persisted settings use an explicit schema version with separate `common`,
  `docx`, `pptx`, and `xlsx` sections. Pre-2.4 flat settings migrate
  automatically. Corrupt or partial data falls back field by field; a raw
  load failure is never overwritten by an automatic repair write.
- File modification, deletion, rename, and restored-workspace format mismatch
  invalidate the active generation before stale parse, search, thumbnail,
  worksheet, drawing, or DOM work can commit. Duplicate valid restored views
  remain isolated and do not block startup.
- Rendering work is guarded by cancellation and file-revision tokens so stale
  results are discarded.
- Production builds enforce a compressed `dist/main.js` budget of 500 KiB.
- Reading state uses a bounded 100-file LRU store. Persisted entries contain
  only path, modification time, format, position, zoom, and navigation state.
  When a source file changes, stale position/navigation data is cleared while
  the user's zoom and fit preference is retained.
- Closing or unloading a file releases document buffers, generated Blob URLs,
  search timers, and search result references.

## Format-specific runtime behavior

- PPTX file and slide rendering use separate cancellation generations so
  rapid file switching and page navigation cannot commit stale slides.
- The current PPTX slide renders before full metadata indexing finishes.
  Background metadata parsing is cancellable and limited to four workers.
- Presentation search reuses a normalized metadata index. The navigation list
  is virtualized to at most 60 mounted rows, and thumbnail rendering is
  limited to two lower-priority workers.
- Off-screen thumbnails are unmounted and their generated image resources are
  released; stale slide and thumbnail renders are cancelled cooperatively.
- PPTX shape, text, and table rendering yields between bounded work slices,
  while XML, relationship, slide-context, and binary caches have fixed
  limits.
- Generated presentation resources are also released when a presentation is
  reloaded, switched, or closed.
- XLSX worksheets use a 2,500-cell mount ceiling with finite row/column
  overscan. Frozen panes remain visible without materializing the full sheet.
- Worksheet and shared-string XML is decompressed and validated in chunks.
  Loading reports progress, dense worksheets do not retain the complete
  `sheetData` XML, and parsing can be cancelled between chunks and rows.
- Worksheet parsing, switching, workbook-wide search, summary collection, and
  scheduled grid rendering are cancellable. Package caches, mounted drawings,
  chart sample counts, and image Blob URLs have fixed bounds and are released
  on reload, sheet/file switch, and view close.
- Word content is rendered into a temporary buffer before replacing the
  visible preview.
- Long documents commit rendered pages and build navigation in cancellable
  chunks so the interface can update between batches.
- Long previews defer off-screen page painting until pages approach the
  viewport.
- Embedded images use lazy loading, asynchronous decoding, and recyclable
  Blob URLs instead of persistent base64 data URLs.
- The current rendered file state is tracked to avoid unnecessary repeated
  renders.
- DOCX render post-processing collects images, pages, headings, Blob URLs,
  and searchable text in one pass.
- DOCX search builds one reusable text index per render and uses CSS
  highlights without rewriting the document DOM. Reading-state writes are
  coalesced and unchanged states are skipped.
- Fit-width mode is a CSS state change and does not rerender the document.
- Development builds log Word file reading, rendering, DOM commit, outline,
  and total preview timings, plus PPTX render duration, object counts,
  resources, and fonts, to the developer console.
