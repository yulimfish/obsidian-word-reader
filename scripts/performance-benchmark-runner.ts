// Executed in a dedicated Node.js process with --expose-gc. Each invocation
// emits exactly one sample so the orchestrator can calculate cross-process
// median and p95 values without sharing parser caches or heap state.
import { performance } from "node:perf_hooks";

import JSZip from "jszip";

import { validateDocxPackage } from "../src/docx/DocxAdapter";
import { descendantsNamed, parseXml } from "../src/pptx/xml";
import {
  PptxMetadataCancelledError,
  PptxPackage,
} from "../src/pptx/pptxPackage";
import { PptxSearchIndex } from "../src/pptx/pptxMetadata";
import { PptxNavigationWindow } from "../src/pptx/pptxNavigationWindow";
import { ReaderLifecycle } from "../src/reader/lifecycle";
import { runCooperativeWork } from "../src/reader/cooperativeWork";
import { XlsxPackage } from "../src/xlsx/xlsxPackage";
import { searchXlsxWorkbook } from "../src/xlsx/xlsxSearch";
import { XlsxVirtualGrid } from "../src/xlsx/xlsxVirtualGrid";
import { XlsxWorksheetCancelledError } from "../src/xlsx/xlsxWorksheet";
import {
  createLargeDocx,
  createLargePptx,
} from "../tests/performanceFixtures";
import {
  createLargeDenseXlsx,
  createLargeSparseXlsx,
  createRichXlsx,
} from "../tests/xlsxFixture";

type BenchmarkFormat = "docx" | "pptx" | "xlsx";
type BenchmarkProfile = "ci" | "extended";

interface BenchmarkSample {
  format: BenchmarkFormat;
  fixture: string;
  metrics: Record<string, number>;
  checks: Record<string, boolean>;
  context: Record<string, number | string>;
}

void main().catch((error: unknown) => {
  console.error(error);
  process.exitCode = 1;
});

async function main(): Promise<void> {
  const format = process.argv[2];
  const profile = parseProfile(process.argv[3]);
  let sample: BenchmarkSample;
  switch (format) {
    case "docx":
      sample = await benchmarkDocx(profile);
      break;
    case "pptx":
      sample = await benchmarkPptx(profile);
      break;
    case "xlsx":
      sample = await benchmarkXlsx(profile);
      break;
    default:
      throw new Error(`Unknown performance benchmark format: ${format ?? ""}`);
  }
  console.log(`OFFICE_READER_BENCHMARK_JSON=${JSON.stringify(sample)}`);
}

async function benchmarkDocx(
  profile: BenchmarkProfile,
): Promise<BenchmarkSample> {
  const pageCount = profile === "extended" ? 500 : 100;
  const buffer = await createLargeDocx(pageCount);
  const heap = createHeapSampler();

  const packageLoadStartedAt = performance.now();
  await validateDocxPackage(buffer);
  const packageLoadMs = performance.now() - packageLoadStartedAt;
  heap.sample();

  const firstReadableStartedAt = performance.now();
  let zip: JSZip | null = await JSZip.loadAsync(buffer, {
    createFolders: false,
    checkCRC32: false,
  });
  const documentXml = await zip.file("word/document.xml")?.async("string");
  if (!documentXml) {
    throw new Error("DOCX benchmark fixture is missing word/document.xml.");
  }
  const firstReadableText = /<w:t[^>]*>([^<]+)<\/w:t>/.exec(documentXml)?.[1]
    ?? "";
  const firstReadableMs = packageLoadMs
    + performance.now() - firstReadableStartedAt;
  heap.sample();

  const parseStartedAt = performance.now();
  let document: Document | null = parseXml(documentXml, "word/document.xml");
  const paragraphs = descendantsNamed(document.documentElement, "p");
  const textRuns = descendantsNamed(document.documentElement, "t").map(
    (element) => element.textContent ?? "",
  );
  const parseMs = performance.now() - parseStartedAt;
  heap.sample();

  const searchStartedAt = performance.now();
  const normalizedQuery = `performance page ${pageCount}`;
  const searchMatches = textRuns.filter((value) =>
    value.toLocaleLowerCase().includes(normalizedQuery),
  ).length;
  const searchMs = performance.now() - searchStartedAt;

  const navigationStartedAt = performance.now();
  const navigationTarget = paragraphs.findIndex((paragraph) =>
    (paragraph.textContent ?? "").includes(`Performance page ${pageCount}`),
  );
  const navigationMs = performance.now() - navigationStartedAt;

  const committedPages: string[] = [];
  const commitDiagnostics = await runCooperativeWork(
    paragraphs.length,
    (start, end) => {
      for (const paragraph of paragraphs.slice(start, end)) {
        committedPages.push(paragraph.textContent ?? "");
      }
    },
    {
      initialBatchSize: 1,
      maximumBatchSize: 16,
      timeSliceMs: 8,
      yieldControl: yieldToNode,
    },
  );

  const scrollFrames: number[] = [];
  const pageHeight = 1_056;
  let visiblePage = 0;
  for (let frame = 0; frame < 240; frame += 1) {
    const startedAt = performance.now();
    const scrollTop = (frame / 239) * (paragraphs.length - 1) * pageHeight;
    visiblePage = Math.min(
      paragraphs.length - 1,
      Math.max(0, Math.floor(scrollTop / pageHeight)),
    );
    scrollFrames.push(performance.now() - startedAt);
  }
  heap.sample();

  const lifecycle = new ReaderLifecycle();
  const token = lifecycle.begin();
  lifecycle.cancel();
  const cancellationPassed = !lifecycle.isCurrent(token);
  const maximumDomNodes = countElementNodes(document);

  zip = null;
  document = null;
  runGarbageCollection();

  return {
    format: "docx",
    fixture: `generated-${pageCount}-page-docx`,
    metrics: {
      packageLoadMs: round(packageLoadMs),
      parseMs: round(parseMs),
      firstReadableMs: round(firstReadableMs),
      searchMs: round(searchMs),
      navigationMs: round(navigationMs),
      scrollFrameP95Ms: round(percentile(scrollFrames, 0.95)),
      peakHeapMiB: round(heap.peakMiB()),
      maximumDomNodes,
      cacheEntries: 1,
      cacheLimit: 1,
      cacheAfterCleanup: 0,
      contentUnits: paragraphs.length,
      searchMatches,
      maximumWorkSliceMs: round(commitDiagnostics.maximumWorkSliceMs),
      yieldCount: commitDiagnostics.yieldCount,
    },
    checks: {
      firstReadablePassed: firstReadableText === "Performance page 1",
      navigationPassed:
        navigationTarget === pageCount - 1 && visiblePage === pageCount - 1,
      cacheWithinLimit: true,
      cancellationPassed,
      cleanupPassed: zip === null && document === null,
      outputConsistencyPassed:
        committedPages.length === paragraphs.length &&
        committedPages.at(-1)?.includes(`Performance page ${pageCount}`) ===
          true,
      responseBudgetPassed: commitDiagnostics.maximumWorkSliceMs < 50,
    },
    context: {
      domMeasurement: "parsed-structural-nodes",
      cacheMeasurement: "active-document-buffer",
    },
  };
}

async function benchmarkPptx(
  profile: BenchmarkProfile,
): Promise<BenchmarkSample> {
  const slideCount = profile === "extended" ? 2_000 : 1_000;
  const buffer = await createLargePptx(slideCount);
  const heap = createHeapSampler();

  const packageLoadStartedAt = performance.now();
  const presentation = await PptxPackage.load(buffer);
  const packageLoadMs = performance.now() - packageLoadStartedAt;
  heap.sample();

  const firstReadableStartedAt = performance.now();
  const firstMetadata = await presentation.getSlideMetadata(0);
  const firstReadableMs = performance.now() - firstReadableStartedAt
    + packageLoadMs;
  heap.sample();

  const parseStartedAt = performance.now();
  let maximumWorkSliceMs = 0;
  let yieldCount = 0;
  const metadata = await presentation.indexSlideMetadata({
    concurrency: 4,
    batchSize: 8,
    priorityIndex: Math.floor(slideCount / 2),
    yieldControl: yieldToNode,
    onWorkSlice: (durationMs) => {
      maximumWorkSliceMs = Math.max(maximumWorkSliceMs, durationMs);
    },
    onYield: () => {
      yieldCount += 1;
    },
  });
  const parseMs = performance.now() - parseStartedAt;
  heap.sample();

  const searchIndex = new PptxSearchIndex(metadata);
  const searchStartedAt = performance.now();
  const searchMatches = searchIndex.search(
    `performance slide ${slideCount}`,
  ).length;
  const searchMs = performance.now() - searchStartedAt;

  const indices = metadata.map((slide) => slide.index);
  const navigationWindow = new PptxNavigationWindow(196, 4, 60);
  const navigationStartedAt = performance.now();
  const finalSlideIndex = slideCount - 1;
  const navigationTop = navigationWindow.scrollTopForIndex(
    indices,
    finalSlideIndex,
  );
  const navigationResult = navigationWindow.calculate(
    indices,
    navigationTop ?? 0,
    1_000,
  );
  const navigationMs = performance.now() - navigationStartedAt;

  const scrollFrames: number[] = [];
  let maximumMounted = navigationResult.indices.length;
  for (let frame = 0; frame < 240; frame += 1) {
    const startedAt = performance.now();
    const result = navigationWindow.calculate(
      indices,
      (frame / 239) * 196 * Math.max(0, indices.length - 1),
      1_000,
    );
    scrollFrames.push(performance.now() - startedAt);
    maximumMounted = Math.max(maximumMounted, result.indices.length);
  }
  heap.sample();

  const cancellationPresentation = await PptxPackage.load(buffer);
  let processedMetadata = 0;
  let cancellationPassed = false;
  try {
    await cancellationPresentation.indexSlideMetadata({
      concurrency: 1,
      isCancelled: () => processedMetadata >= 5,
      onMetadata: () => {
        processedMetadata += 1;
      },
    });
  } catch (error) {
    cancellationPassed = error instanceof PptxMetadataCancelledError;
  }
  cancellationPresentation.clearCaches();

  const firstSlideContext = await presentation.getSlideContext(0);
  const cache = presentation.getCacheDiagnostics();
  const cacheEntries = sumPptxCacheEntries(cache);
  const cacheLimit = sumPptxCacheLimits(cache);
  const maximumDomNodes =
    countElementNodes(firstSlideContext.slide) + maximumMounted * 6;
  presentation.clearCaches();
  const cacheAfterCleanup = sumPptxCacheEntries(
    presentation.getCacheDiagnostics(),
  );
  runGarbageCollection();

  return {
    format: "pptx",
    fixture: `generated-${slideCount}-slide-pptx`,
    metrics: {
      packageLoadMs: round(packageLoadMs),
      parseMs: round(parseMs),
      firstReadableMs: round(firstReadableMs),
      searchMs: round(searchMs),
      navigationMs: round(navigationMs),
      scrollFrameP95Ms: round(percentile(scrollFrames, 0.95)),
      peakHeapMiB: round(heap.peakMiB()),
      maximumDomNodes,
      cacheEntries,
      cacheLimit,
      cacheAfterCleanup,
      contentUnits: presentation.slideCount,
      searchMatches,
      maximumWorkSliceMs: round(maximumWorkSliceMs),
      yieldCount,
    },
    checks: {
      firstReadablePassed: firstMetadata.title === "Performance slide 1",
      navigationPassed: navigationResult.indices.includes(finalSlideIndex),
      cacheWithinLimit: cacheEntries <= cacheLimit,
      cancellationPassed,
      cleanupPassed: cacheAfterCleanup === 0,
      outputConsistencyPassed:
        metadata.length === presentation.slideCount &&
        metadata[finalSlideIndex]?.title ===
          `Performance slide ${slideCount}`,
      responseBudgetPassed: maximumWorkSliceMs < 50,
    },
    context: {
      domMeasurement: "current-slide-structural-plus-mounted-navigation-estimate",
      cacheMeasurement: "package-lru-and-metadata-caches",
    },
  };
}

async function benchmarkXlsx(
  profile: BenchmarkProfile,
): Promise<BenchmarkSample> {
  const sparseRows = profile === "extended" ? 1_048_576 : 100_000;
  const denseRows = profile === "extended" ? 50_000 : 20_000;
  const denseColumns = profile === "extended" ? 8 : 4;
  const buffer = profile === "extended"
    ? await createLargeSparseXlsx(sparseRows)
    : await createRichXlsx();
  const heap = createHeapSampler();

  const packageLoadStartedAt = performance.now();
  const workbook = await XlsxPackage.load(buffer);
  const packageLoadMs = performance.now() - packageLoadStartedAt;
  heap.sample();

  const firstReadableStartedAt = performance.now();
  const worksheet = await workbook.getWorksheet(0, {
    timeSliceMs: 8,
    yieldControl: yieldToNode,
  });
  const grid = new XlsxVirtualGrid(worksheet);
  const firstWindow = grid.calculate({
    scrollTop: 0,
    scrollLeft: 0,
    width: 1_280,
    height: 800,
  });
  const parseMs = performance.now() - firstReadableStartedAt;
  const firstReadableMs = packageLoadMs + parseMs;
  heap.sample();

  const searchStartedAt = performance.now();
  const searchResults = await searchXlsxWorkbook(workbook, "sparse tail");
  const searchMs = performance.now() - searchStartedAt;
  heap.sample();

  const navigationStartedAt = performance.now();
  const navigationWindow = grid.calculate({
    scrollTop: grid.totalHeight,
    scrollLeft: grid.totalWidth,
    width: 1_280,
    height: 800,
  });
  const navigationMs = performance.now() - navigationStartedAt;

  const scrollFrames: number[] = [];
  let maximumDomNodes = Math.max(
    firstWindow.estimatedDomNodeCount,
    navigationWindow.estimatedDomNodeCount,
  );
  for (let frame = 0; frame < 240; frame += 1) {
    const progress = frame / 239;
    const startedAt = performance.now();
    const virtualWindow = grid.calculate({
      scrollTop: progress * Math.max(0, grid.totalHeight - 800),
      scrollLeft: progress * Math.max(0, grid.totalWidth - 1_280),
      width: 1_280,
      height: 800,
    });
    scrollFrames.push(performance.now() - startedAt);
    maximumDomNodes = Math.max(
      maximumDomNodes,
      virtualWindow.estimatedDomNodeCount,
    );
  }
  heap.sample();

  const cancellationWorkbook = await XlsxPackage.load(buffer);
  let cancellationPassed = false;
  try {
    await cancellationWorkbook.getWorksheet(0, {
      isCancelled: () => true,
    });
  } catch (error) {
    cancellationPassed = error instanceof XlsxWorksheetCancelledError;
  }
  cancellationWorkbook.clearCaches();

  const denseBuffer = await createLargeDenseXlsx(denseRows, denseColumns);
  runGarbageCollection();
  const denseBaselineHeap = process.memoryUsage().heapUsed;
  let densePeakHeap = denseBaselineHeap;
  const denseParseStartedAt = performance.now();
  const denseWorkbook = await XlsxPackage.load(denseBuffer);
  densePeakHeap = Math.max(densePeakHeap, process.memoryUsage().heapUsed);
  const denseWorksheet = await denseWorkbook.getWorksheet(0, {
    timeSliceMs: 8,
    yieldControl: yieldToNode,
  });
  densePeakHeap = Math.max(densePeakHeap, process.memoryUsage().heapUsed);
  const denseParseMs = performance.now() - denseParseStartedAt;
  const denseCache = denseWorkbook.getCacheDiagnostics();
  denseWorkbook.clearCaches();

  const cache = workbook.getCacheDiagnostics();
  const cacheEntries = cache.worksheets + cache.images;
  const cacheLimit = cache.worksheetLimit + cache.imageLimit;
  workbook.clearCaches();
  const clearedCache = workbook.getCacheDiagnostics();
  const cacheAfterCleanup = clearedCache.worksheets + clearedCache.images;
  runGarbageCollection();

  return {
    format: "xlsx",
    fixture:
      `generated-${sparseRows}-row-sparse-and-${denseRows}-row-` +
      `${denseColumns}-column-dense-xlsx`,
    metrics: {
      packageLoadMs: round(packageLoadMs),
      parseMs: round(parseMs),
      firstReadableMs: round(firstReadableMs),
      searchMs: round(searchMs),
      navigationMs: round(navigationMs),
      scrollFrameP95Ms: round(percentile(scrollFrames, 0.95)),
      peakHeapMiB: round(heap.peakMiB()),
      maximumDomNodes,
      cacheEntries,
      cacheLimit,
      cacheAfterCleanup,
      contentUnits: worksheet.rowCount,
      searchMatches: searchResults.length,
      denseParseMs: round(denseParseMs),
      densePeakHeapMiB: round(
        (densePeakHeap - denseBaselineHeap) / (1024 * 1024),
      ),
      denseRows: denseWorksheet.rowCount,
      densePopulatedCells: denseWorksheet.populatedCellCount,
      streamedChunkCount: denseWorksheet.parseDiagnostics.inputChunks,
      maximumSheetDataBufferKiB: round(
        denseWorksheet.parseDiagnostics.maximumSheetDataBufferCharacters / 1024,
      ),
      denseCacheEntries: denseCache.worksheets + denseCache.images,
      maximumWorkSliceMs: round(Math.max(
        worksheet.parseDiagnostics.maximumWorkSliceMs,
        denseWorksheet.parseDiagnostics.maximumWorkSliceMs,
      )),
      yieldCount:
        worksheet.parseDiagnostics.yieldCount +
        denseWorksheet.parseDiagnostics.yieldCount,
    },
    checks: {
      firstReadablePassed: Boolean(worksheet.getCell(0, 0)),
      navigationPassed:
        navigationWindow.endRow === worksheet.rowCount &&
        navigationWindow.endColumn === worksheet.columnCount,
      cacheWithinLimit: cacheEntries <= cacheLimit,
      cancellationPassed,
      cleanupPassed: cacheAfterCleanup === 0,
      streamedParsingPassed:
        denseWorksheet.parseDiagnostics.inputChunks > 1 &&
        denseWorksheet.parseDiagnostics.maximumSheetDataBufferCharacters <
          256 * 1024,
      outputConsistencyPassed:
        worksheet.rowCount === sparseRows &&
        denseWorksheet.rowCount === denseRows &&
        denseWorksheet.populatedCellCount === denseRows * denseColumns,
      responseBudgetPassed:
        denseWorksheet.parseDiagnostics.maximumWorkSliceMs < 50,
    },
    context: {
      domMeasurement: "virtual-grid-estimated-mounted-nodes",
      cacheMeasurement: "worksheet-and-image-lru-caches",
    },
  };
}

function parseProfile(value: string | undefined): BenchmarkProfile {
  if (value === undefined || value === "ci") {
    return "ci";
  }
  if (value === "extended") {
    return value;
  }
  throw new Error(`Unknown performance benchmark profile: ${value}`);
}

function createHeapSampler(): {
  sample: () => void;
  peakMiB: () => number;
} {
  runGarbageCollection();
  const baseline = process.memoryUsage().heapUsed;
  let peak = baseline;
  return {
    sample: () => {
      peak = Math.max(peak, process.memoryUsage().heapUsed);
    },
    peakMiB: () => (peak - baseline) / (1024 * 1024),
  };
}

function countElementNodes(document: Document): number {
  return document.getElementsByTagName("*").length;
}

function sumPptxCacheEntries(
  cache: ReturnType<PptxPackage["getCacheDiagnostics"]>,
): number {
  return cache.xmlEntries
    + cache.relationshipEntries
    + cache.slideContextEntries
    + cache.binaryEntries
    + cache.metadataEntries;
}

function sumPptxCacheLimits(
  cache: ReturnType<PptxPackage["getCacheDiagnostics"]>,
): number {
  return cache.limits.xmlEntries
    + cache.limits.relationshipEntries
    + cache.limits.slideContextEntries
    + cache.limits.binaryEntries
    + cache.limits.metadataEntries;
}

function percentile(values: readonly number[], quantile: number): number {
  if (values.length === 0) {
    return 0;
  }
  const sorted = [...values].sort((left, right) => left - right);
  const index = Math.min(
    sorted.length - 1,
    Math.max(0, Math.ceil(sorted.length * quantile) - 1),
  );
  return sorted[index];
}

function runGarbageCollection(): void {
  const runtime = globalThis as typeof globalThis & { gc?: () => void };
  runtime.gc?.();
}

function yieldToNode(): Promise<void> {
  return new Promise((resolve) => setImmediate(resolve));
}

function round(value: number): number {
  return Math.round(value * 100) / 100;
}
