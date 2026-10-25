import { performance } from "node:perf_hooks";

import JSZip from "jszip";

import { validateDocxPackage } from "../src/docx/DocxAdapter";
import { PptxSearchIndex } from "../src/pptx/pptxMetadata";
import { PptxPackage } from "../src/pptx/pptxPackage";
import { descendantsNamed, parseXml } from "../src/pptx/xml";
import { ReaderFileRevisionTracker } from "../src/reader/fileRevision";
import { ReaderLifecycle } from "../src/reader/lifecycle";
import {
  ReaderResourceScope,
  type ReaderResourceSnapshot,
} from "../src/reader/resourceScope";
import { XlsxPackage } from "../src/xlsx/xlsxPackage";
import { searchXlsxWorkbook } from "../src/xlsx/xlsxSearch";
import { XlsxVirtualGrid } from "../src/xlsx/xlsxVirtualGrid";
import {
  createLargeDocx,
  createLargePptx,
} from "../tests/performanceFixtures";
import { addOoxmlEntry, createRichXlsx } from "../tests/xlsxFixture";

type StressFormat = "docx" | "pptx" | "xlsx";
type StaleTaskKind =
  | "parse"
  | "search"
  | "thumbnail"
  | "worksheet"
  | "drawing";

interface Fixtures {
  docx: ArrayBuffer;
  pptx: ArrayBuffer;
  xlsx: ArrayBuffer;
}

interface FormatHandle {
  navigate: () => Promise<boolean>;
  clear: () => void;
  cacheEntries: () => number;
}

interface OperationCounts {
  open: number;
  reload: number;
  quickNavigation: number;
  fileSwitch: number;
  close: number;
}

interface SequenceResult {
  name: string;
  formatOrder: StressFormat[];
  cycles: number;
  durationMs: number;
  operations: OperationCounts;
  formatCycles: Record<StressFormat, number>;
  resourcePeak: ReaderResourceSnapshot;
  resourceResidue: ReaderResourceSnapshot;
  formatCachePeak: number;
  formatCacheResidue: number;
  staleUiCommits: Record<StaleTaskKind, number>;
  handledTaskErrors: number;
  revisionInvalidations: number;
  navigationFailures: number;
  heap: {
    gcAvailable: boolean;
    samples: number;
    startMedianMiB: number;
    finalMedianMiB: number;
    retainedGrowthMiB: number;
    maxConsecutiveIncreases: number;
  };
}

interface ErrorMatrixResult {
  iterationsPerCase: number;
  attempts: number;
  recoverableErrors: number;
  unexpectedSuccesses: number;
  resourceResidue: ReaderResourceSnapshot;
  observedKinds: Record<string, number>;
}

interface LifecycleStressResult {
  product: "Office Reader";
  schemaVersion: 1;
  generatedAt: string;
  runtime: {
    node: string;
    platform: NodeJS.Platform;
    architecture: string;
  };
  configuration: {
    cyclesPerSequence: number;
    heapWarmupCycles: number;
    sustainedGrowthWindow: number;
    errorIterationsPerCase: number;
  };
  sequences: SequenceResult[];
  errorMatrix: ErrorMatrixResult;
  unhandledRejections: number;
}

const CYCLES_PER_SEQUENCE = 100;
const HEAP_WARMUP_CYCLES = 10;
const SUSTAINED_GROWTH_WINDOW = 20;
const ERROR_ITERATIONS_PER_CASE = 25;
const STALE_TASK_KINDS: readonly StaleTaskKind[] = [
  "parse",
  "search",
  "thumbnail",
  "worksheet",
  "drawing",
];
const EMPTY_RESOURCES: ReaderResourceSnapshot = {
  blobUrls: 0,
  caches: 0,
  tasks: 0,
  timers: 0,
  observers: 0,
  listeners: 0,
};

void main().catch((error: unknown) => {
  console.error(error);
  process.exitCode = 1;
});

async function main(): Promise<void> {
  if (!getGarbageCollector()) {
    throw new Error("Lifecycle stress requires Node.js --expose-gc.");
  }
  const unhandledRejections: unknown[] = [];
  const onUnhandledRejection = (reason: unknown): void => {
    unhandledRejections.push(reason);
  };
  process.on("unhandledRejection", onUnhandledRejection);
  try {
    const fixtures = await createFixtures();
    forceGarbageCollection();
    const sequences = [];
    sequences.push(await runSequence(
      "docx",
      ["docx"],
      fixtures,
    ));
    sequences.push(await runSequence(
      "pptx",
      ["pptx"],
      fixtures,
    ));
    sequences.push(await runSequence(
      "xlsx",
      ["xlsx"],
      fixtures,
    ));
    sequences.push(await runSequence(
      "mixed",
      ["docx", "pptx", "xlsx"],
      fixtures,
    ));
    const errorMatrix = await runErrorMatrix(fixtures);
    await new Promise<void>((resolve) => setImmediate(resolve));

    const result: LifecycleStressResult = {
      product: "Office Reader",
      schemaVersion: 1,
      generatedAt: new Date().toISOString(),
      runtime: {
        node: process.version,
        platform: process.platform,
        architecture: process.arch,
      },
      configuration: {
        cyclesPerSequence: CYCLES_PER_SEQUENCE,
        heapWarmupCycles: HEAP_WARMUP_CYCLES,
        sustainedGrowthWindow: SUSTAINED_GROWTH_WINDOW,
        errorIterationsPerCase: ERROR_ITERATIONS_PER_CASE,
      },
      sequences,
      errorMatrix,
      unhandledRejections: unhandledRejections.length,
    };
    console.log(`OFFICE_READER_LIFECYCLE_JSON=${JSON.stringify(result)}`);
  } finally {
    process.off("unhandledRejection", onUnhandledRejection);
  }
}

async function createFixtures(): Promise<Fixtures> {
  const [docx, pptx, xlsx] = await Promise.all([
    createLargeDocx(4),
    createLargePptx(6),
    createRichXlsx(),
  ]);
  return { docx, pptx, xlsx };
}

async function runSequence(
  name: string,
  formatOrder: StressFormat[],
  fixtures: Fixtures,
): Promise<SequenceResult> {
  const startedAt = performance.now();
  const operations = emptyOperationCounts();
  const formatCycles = emptyFormatCounts();
  const staleUiCommits = emptyStaleTaskCounts();
  const resourcePeak = { ...EMPTY_RESOURCES };
  const resourceResidue = { ...EMPTY_RESOURCES };
  const heapSamples: number[] = [];
  let formatCacheResidue = 0;
  let formatCachePeak = 0;
  let handledTaskErrors = 0;
  let revisionInvalidations = 0;
  let navigationFailures = 0;

  for (let cycle = 0; cycle < CYCLES_PER_SEQUENCE; cycle += 1) {
    const format = formatOrder[cycle % formatOrder.length];
    formatCycles[format] += 1;
    const scope = new ReaderResourceScope(() => {
      handledTaskErrors += 1;
    });
    registerTransientResources(scope, `${name}-${cycle}`);
    mergeResourceMaximum(resourcePeak, scope.snapshot());

    const revisions = new ReaderFileRevisionTracker();
    const primaryFile = fileDescriptor(name, format, "primary", cycle, fixtures);
    let revision = revisions.capture(primaryFile, format);
    let handle = await openFormat(format, fixtures[format]);
    operations.open += 1;
    let releaseCache = scope.track("caches", handle, () => handle.clear());
    formatCachePeak = Math.max(formatCachePeak, handle.cacheEntries());
    mergeResourceMaximum(resourcePeak, scope.snapshot());

    revisions.markChanged(primaryFile.path);
    if (!revisions.isCurrent(revision, primaryFile, format)) {
      revisionInvalidations += 1;
    }
    releaseCache();
    formatCacheResidue = Math.max(formatCacheResidue, handle.cacheEntries());
    handle = await openFormat(format, fixtures[format]);
    operations.reload += 1;
    revision = revisions.capture(primaryFile, format);
    releaseCache = scope.track("caches", handle, () => handle.clear());
    formatCachePeak = Math.max(formatCachePeak, handle.cacheEntries());

    if (!await handle.navigate()) {
      navigationFailures += 1;
    }
    operations.quickNavigation += 1;

    const staleWork = scheduleStaleTasks(scope, staleUiCommits);
    mergeResourceMaximum(resourcePeak, scope.snapshot());
    revisions.markChanged(primaryFile.path);
    if (!revisions.isCurrent(revision, primaryFile, format)) {
      revisionInvalidations += 1;
    }
    staleWork.cancel();
    releaseCache();
    formatCacheResidue = Math.max(formatCacheResidue, handle.cacheEntries());

    const switchedFile = fileDescriptor(
      name,
      format,
      "secondary",
      cycle,
      fixtures,
    );
    revision = revisions.capture(switchedFile, format);
    handle = await openFormat(format, fixtures[format]);
    releaseCache = scope.track("caches", handle, () => handle.clear());
    formatCachePeak = Math.max(formatCachePeak, handle.cacheEntries());
    operations.fileSwitch += 1;
    if (!revisions.isCurrent(revision, switchedFile, format)) {
      throw new Error(`${name} cycle ${cycle} lost the switched revision.`);
    }

    staleWork.close();
    releaseCache();
    handle.clear();
    formatCacheResidue = Math.max(formatCacheResidue, handle.cacheEntries());
    await scope.drainTasks();
    scope.clear();
    mergeResourceMaximum(resourceResidue, scope.snapshot());
    operations.close += 1;

    forceGarbageCollection();
    heapSamples.push(process.memoryUsage().heapUsed / (1024 * 1024));
  }

  return {
    name,
    formatOrder,
    cycles: CYCLES_PER_SEQUENCE,
    durationMs: round(performance.now() - startedAt),
    operations,
    formatCycles,
    resourcePeak,
    resourceResidue,
    formatCachePeak,
    formatCacheResidue,
    staleUiCommits,
    handledTaskErrors,
    revisionInvalidations,
    navigationFailures,
    heap: summarizeHeap(heapSamples),
  };
}

function registerTransientResources(
  scope: ReaderResourceScope,
  label: string,
): void {
  const blobUrl = URL.createObjectURL(new Blob([label]));
  scope.track("blobUrls", blobUrl, () => URL.revokeObjectURL(blobUrl));

  const timer = setTimeout(() => undefined, 60_000);
  scope.track("timers", timer, () => clearTimeout(timer));

  let observing = true;
  const observer = { label };
  scope.track("observers", observer, () => {
    observing = false;
  });
  if (!observing) {
    throw new Error("The synthetic observer was released too early.");
  }

  const target = new EventTarget();
  scope.listen(target, "office-reader-stress", () => undefined);
}

function scheduleStaleTasks(
  scope: ReaderResourceScope,
  staleUiCommits: Record<StaleTaskKind, number>,
): { cancel: () => void; close: () => void } {
  const lifecycles = STALE_TASK_KINDS.map((kind) => {
    const lifecycle = new ReaderLifecycle();
    const token = lifecycle.begin();
    scope.runTask(async () => {
      await Promise.resolve();
      if (lifecycle.isCurrent(token)) {
        staleUiCommits[kind] += 1;
      }
    });
    return lifecycle;
  });
  scope.runTask(async () => {
    await Promise.resolve();
    throw new Error("Synthetic cancellation");
  });
  return {
    cancel: () => {
      for (const lifecycle of lifecycles) {
        lifecycle.cancel();
      }
    },
    close: () => {
      for (const lifecycle of lifecycles) {
        lifecycle.cancel();
      }
    },
  };
}

async function openFormat(
  format: StressFormat,
  buffer: ArrayBuffer,
): Promise<FormatHandle> {
  switch (format) {
    case "docx":
      return openDocx(buffer);
    case "pptx":
      return openPptx(buffer);
    case "xlsx":
      return openXlsx(buffer);
  }
}

async function openDocx(buffer: ArrayBuffer): Promise<FormatHandle> {
  await validateDocxPackage(buffer);
  let zip: JSZip | null = await JSZip.loadAsync(buffer, {
    createFolders: false,
    checkCRC32: false,
  });
  const xml = await zip.file("word/document.xml")?.async("string");
  if (!xml) {
    throw new Error("DOCX stress fixture has no document part.");
  }
  let document: Document | null = parseXml(xml, "word/document.xml");
  const paragraphs = descendantsNamed(document.documentElement, "p");
  return {
    navigate: async () => {
      const last = paragraphs.at(-1)?.textContent ?? "";
      const matches = descendantsNamed(
        document?.documentElement ?? null,
        "t",
      ).filter((entry) => entry.textContent?.includes("page 4"));
      return last.includes("page 4") && matches.length === 1;
    },
    clear: () => {
      zip = null;
      document = null;
      paragraphs.length = 0;
    },
    cacheEntries: () => Number(zip !== null) + Number(document !== null),
  };
}

async function openPptx(buffer: ArrayBuffer): Promise<FormatHandle> {
  let presentation: PptxPackage | null = await PptxPackage.load(buffer);
  const metadata = await presentation.indexSlideMetadata({ concurrency: 2 });
  const search = new PptxSearchIndex(metadata);
  return {
    navigate: async () => {
      const active = presentation;
      if (!active) {
        return false;
      }
      const context = await active.getSlideContext(active.slideCount - 1);
      return context.slidePath.endsWith("slide6.xml") &&
        search.search("performance slide 6").length === 1;
    },
    clear: () => {
      presentation?.clearCaches();
      presentation = null;
      search.clear();
      metadata.length = 0;
    },
    cacheEntries: () => presentation
      ? sumPptxCacheEntries(presentation.getCacheDiagnostics())
      : 0,
  };
}

async function openXlsx(buffer: ArrayBuffer): Promise<FormatHandle> {
  let workbook: XlsxPackage | null = await XlsxPackage.load(buffer);
  let worksheet = await workbook.getWorksheet(0);
  let grid: XlsxVirtualGrid | null = new XlsxVirtualGrid(worksheet);
  return {
    navigate: async () => {
      const active = workbook;
      if (!active || !grid) {
        return false;
      }
      const matches = await searchXlsxWorkbook(active, "sparse tail");
      const tail = grid.calculate({
        scrollTop: grid.totalHeight,
        scrollLeft: grid.totalWidth,
        width: 1_280,
        height: 800,
      });
      return matches.length === 1 && tail.endRow === worksheet.rowCount;
    },
    clear: () => {
      workbook?.clearCaches();
      workbook = null;
      grid = null;
      worksheet = null as never;
    },
    cacheEntries: () => {
      if (!workbook) {
        return 0;
      }
      const cache = workbook.getCacheDiagnostics();
      return cache.worksheets + cache.images;
    },
  };
}

async function runErrorMatrix(fixtures: Fixtures): Promise<ErrorMatrixResult> {
  const encrypted = new Uint8Array(32);
  encrypted.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
  const damaged = Uint8Array.from([1, 2, 3, 4]).buffer;
  const overLimit = await createCompressionBomb();
  const unsupported = {
    docx: await addOoxmlEntry(fixtures.docx, "word/vbaProject.bin", "macro"),
    pptx: await addOoxmlEntry(fixtures.pptx, "ppt/vbaProject.bin", "macro"),
    xlsx: await addOoxmlEntry(fixtures.xlsx, "xl/vbaProject.bin", "macro"),
  } satisfies Fixtures;
  const resourceResidue = { ...EMPTY_RESOURCES };
  const observedKinds: Record<string, number> = {};
  let attempts = 0;
  let recoverableErrors = 0;
  let unexpectedSuccesses = 0;

  for (const format of ["docx", "pptx", "xlsx"] as const) {
    const cases = [
      ["damaged", damaged],
      ["encrypted", encrypted.buffer],
      ["limit-exceeded", overLimit],
      ["unsupported", unsupported[format]],
    ] as const;
    for (const [expectedKind, buffer] of cases) {
      for (
        let iteration = 0;
        iteration < ERROR_ITERATIONS_PER_CASE;
        iteration += 1
      ) {
        attempts += 1;
        const scope = new ReaderResourceScope();
        registerTransientResources(
          scope,
          `error-${format}-${expectedKind}-${iteration}`,
        );
        let handle: FormatHandle | null = null;
        try {
          handle = await openFormat(format, buffer);
          unexpectedSuccesses += 1;
        } catch (error) {
          recoverableErrors += 1;
          const observedKind = getErrorKind(error, expectedKind);
          observedKinds[`${format}:${observedKind}`] =
            (observedKinds[`${format}:${observedKind}`] ?? 0) + 1;
        } finally {
          handle?.clear();
          scope.clear();
          await scope.drainTasks();
          scope.clear();
          mergeResourceMaximum(resourceResidue, scope.snapshot());
        }
      }
    }
  }
  forceGarbageCollection();
  return {
    iterationsPerCase: ERROR_ITERATIONS_PER_CASE,
    attempts,
    recoverableErrors,
    unexpectedSuccesses,
    resourceResidue,
    observedKinds,
  };
}

async function createCompressionBomb(): Promise<ArrayBuffer> {
  const zip = new JSZip();
  zip.file("compressed.bin", "0".repeat(1024 * 1024));
  return zip.generateAsync({ type: "arraybuffer", compression: "DEFLATE" });
}

function getErrorKind(error: unknown, fallback: string): string {
  if (
    typeof error === "object" &&
    error !== null &&
    "kind" in error &&
    typeof error.kind === "string"
  ) {
    return error.kind;
  }
  return fallback;
}

function fileDescriptor(
  sequence: string,
  format: StressFormat,
  name: string,
  cycle: number,
  fixtures: Fixtures,
): {
  path: string;
  extension: string;
  stat: { mtime: number; size: number };
} {
  return {
    path: `stress/${sequence}/${name}.${format}`,
    extension: format,
    stat: {
      mtime: cycle * 2 + (name === "primary" ? 1 : 2),
      size: fixtures[format].byteLength,
    },
  };
}

function summarizeHeap(samples: number[]): SequenceResult["heap"] {
  const steady = samples.slice(HEAP_WARMUP_CYCLES);
  const segmentSize = Math.min(10, Math.floor(steady.length / 2));
  const startMedian = median(steady.slice(0, segmentSize));
  const finalMedian = median(steady.slice(-segmentSize));
  return {
    gcAvailable: getGarbageCollector() !== null,
    samples: samples.length,
    startMedianMiB: round(startMedian),
    finalMedianMiB: round(finalMedian),
    retainedGrowthMiB: round(finalMedian - startMedian),
    maxConsecutiveIncreases: maximumConsecutiveIncreases(steady),
  };
}

function maximumConsecutiveIncreases(values: number[]): number {
  let current = 0;
  let maximum = 0;
  for (let index = 1; index < values.length; index += 1) {
    if (values[index] > values[index - 1] + 0.01) {
      current += 1;
      maximum = Math.max(maximum, current);
    } else {
      current = 0;
    }
  }
  return maximum;
}

function mergeResourceMaximum(
  target: ReaderResourceSnapshot,
  observed: ReaderResourceSnapshot,
): void {
  for (const key of Object.keys(target) as Array<keyof ReaderResourceSnapshot>) {
    target[key] = Math.max(target[key], observed[key]);
  }
}

function sumPptxCacheEntries(
  cache: ReturnType<PptxPackage["getCacheDiagnostics"]>,
): number {
  return cache.xmlEntries +
    cache.relationshipEntries +
    cache.slideContextEntries +
    cache.binaryEntries +
    cache.metadataEntries;
}

function emptyOperationCounts(): OperationCounts {
  return { open: 0, reload: 0, quickNavigation: 0, fileSwitch: 0, close: 0 };
}

function emptyFormatCounts(): Record<StressFormat, number> {
  return { docx: 0, pptx: 0, xlsx: 0 };
}

function emptyStaleTaskCounts(): Record<StaleTaskKind, number> {
  return { parse: 0, search: 0, thumbnail: 0, worksheet: 0, drawing: 0 };
}

function median(values: number[]): number {
  if (values.length === 0) {
    return 0;
  }
  const sorted = [...values].sort((left, right) => left - right);
  const middle = Math.floor(sorted.length / 2);
  return sorted.length % 2 === 0
    ? (sorted[middle - 1] + sorted[middle]) / 2
    : sorted[middle];
}

function forceGarbageCollection(): void {
  const garbageCollector = getGarbageCollector();
  if (!garbageCollector) {
    throw new Error("Garbage collection is unavailable.");
  }
  garbageCollector();
  garbageCollector();
}

function getGarbageCollector(): (() => void) | null {
  const runtime = globalThis as typeof globalThis & { gc?: () => void };
  return runtime.gc ?? null;
}

function round(value: number): number {
  return Math.round(value * 100) / 100;
}
