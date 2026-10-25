# Performance baselines

`npm run performance:check` runs DOCX, PPTX, and XLSX workloads. Every format
is measured five times in a new Node.js process, then checked against the
format-specific budget in this directory. Timing and heap gates use p95;
machine results also retain the median and every raw sample.

The shared metrics are:

- OOXML package loading and content parsing;
- time until the first readable content window;
- search and navigation latency;
- p95 scroll-window calculation or render time;
- sampled peak JavaScript heap growth;
- maximum parsed, estimated, or actual DOM nodes, with the measurement kind
  recorded alongside the value;
- current cache entries and declared cache limits; and
- cancellation and cleanup outcomes.
- maximum cooperative work slice and yield count, with output consistency and
  the `< 50 ms` response budget asserted for every format.

XLSX additionally keeps the existing 3.2.0 dense-streaming gates: `3,000 ms`
parse, `192 MiB` heap growth, and a `256 KiB` worksheet-data buffer. Production
artifacts remain limited to a `500 KiB` `dist/main.js` and an `8 MiB` release
zip.

Generated files are written to `benchmark-results/`:

- `performance.latest.json`: full machine-readable samples and aggregates;
- `performance-trend.json`: compact CI trend point with commit/run metadata;
- `performance-summary.md`: human-readable CI job summary.

GitHub Actions uploads all three files for 90 days on every CI run, including
failed budget runs. A failure identifies the format, aggregate, stage/metric,
observed value, and threshold.

`npm run large-files:extended` is the slower second tier. It runs three fresh
processes per format with 500-page DOCX, 2,000-slide PPTX, an exact
1,048,576-row sparse XLSX, and a 50,000-row/400,000-cell dense XLSX. The weekly
and release workflows retain:

- `large-files.latest.json`: full extended samples and aggregates;
- `large-files-trend.json`: compact extended trend point; and
- `large-files-summary.md`: extended CI/release summary.

Extended thresholds live in `*-extended-budget.json`. This tier is deliberately
excluded from `npm run check`, but is included in `npm run release`.

`npm run lifecycle:check` is the companion long-cycle gate. It runs four
100-cycle sequences (DOCX, PPTX, XLSX, and mixed) with forced garbage
collection, plus 300 repeated damaged/encrypted/over-limit/unsupported package
attempts. It verifies that Blob URLs, format caches, background tasks, timers,
observers, listeners, and stale UI commits return to zero; retained heap has no
20-sample monotonic run; and no Promise rejection escapes cancellation.

Lifecycle files are also written to `benchmark-results/` and retained by CI:

- `lifecycle-stress.latest.json`: complete operations, resources, heap samples,
  stale-task outcomes, and error matrix;
- `lifecycle-stress-trend.json`: compact commit/run trend point; and
- `lifecycle-stress-summary.md`: CI job summary.

The thresholds are versioned in `lifecycle-stress-budget.json`. Fixture labels
and artifacts contain no document content, cell values, speaker notes, internal
XML, or vault paths.

Node.js measurements do not model Electron layout, paint, or third-party
renderer long tasks. Use
the companion [Obsidian Desktop baseline](./OBSIDIAN_DESKTOP_BASELINE.md) to
calibrate these gates against actual application behavior.
