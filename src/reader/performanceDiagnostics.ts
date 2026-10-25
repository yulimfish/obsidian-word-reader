import {
  createReaderDiagnosticReport,
  type ReaderDiagnosticFile,
  type ReaderDiagnosticReport,
} from "./diagnostics";
import type { ReaderFormat } from "./readingState";
import type { ReaderMemoryPressureLevel } from "./memoryPressure";

export type PerformanceTimingName =
  | "packageLoadMs"
  | "parseMs"
  | "firstReadableMs"
  | "searchMs"
  | "navigationMs";

export interface ReaderPerformanceSnapshot {
  timings: Record<PerformanceTimingName, number | null> & {
    scrollFrameP95Ms: number | null;
  };
  memory: {
    peakHeapMiB: number | null;
    measurement: "chromium-used-js-heap" | "unavailable";
  };
  dom: {
    maximumNodes: number;
  };
  cache: {
    entries: number;
    limit: number;
  };
  resources: {
    retained: number;
  };
  lifecycle: {
    cancellationPassed: boolean | null;
    cleanupPassed: boolean | null;
  };
  responsiveness: {
    maximumWorkSliceMs: number;
    longTaskCount: number;
    yieldCount: number;
    backgroundPreemptions: number;
    memoryPressure: ReaderMemoryPressureLevel;
    uninterruptiblePhase: {
      name: "docx-preview-render";
      durationMs: number;
    } | null;
  };
}

export interface ReaderPerformanceDiagnosticDetails
extends ReaderPerformanceSnapshot {
  measurement: "obsidian-desktop-runtime";
  privacy: {
    containsDocumentContent: false;
    containsSpeakerNotes: false;
    containsCellValues: false;
    containsInternalXml: false;
    containsVaultPath: false;
  };
}

export type ReaderPerformanceDiagnosticReport =
  ReaderDiagnosticReport<ReaderPerformanceDiagnosticDetails>;

const MAX_SCROLL_SAMPLES = 240;

export class ReaderPerformanceTracker {
  private readonly timings: Record<PerformanceTimingName, number | null> = {
    packageLoadMs: null,
    parseMs: null,
    firstReadableMs: null,
    searchMs: null,
    navigationMs: null,
  };
  private scrollSamples: number[] = [];
  private heapBaselineBytes: number | null = null;
  private peakHeapBytes: number | null = null;
  private cancellationPassed: boolean | null = null;
  private cleanupPassed: boolean | null = null;
  private maximumWorkSliceMs = 0;
  private longTaskCount = 0;
  private yieldCount = 0;
  private backgroundPreemptions = 0;
  private memoryPressure: ReaderMemoryPressureLevel = "normal";
  private uninterruptiblePhase: {
    name: "docx-preview-render";
    durationMs: number;
  } | null = null;

  constructor() {
    this.reset();
  }

  reset(): void {
    for (const timing of Object.keys(this.timings) as PerformanceTimingName[]) {
      this.timings[timing] = null;
    }
    this.scrollSamples = [];
    this.cancellationPassed = null;
    this.cleanupPassed = null;
    this.maximumWorkSliceMs = 0;
    this.longTaskCount = 0;
    this.yieldCount = 0;
    this.backgroundPreemptions = 0;
    this.memoryPressure = "normal";
    this.uninterruptiblePhase = null;
    const heapBytes = readRuntimeHeapBytes();
    this.heapBaselineBytes = heapBytes;
    this.peakHeapBytes = heapBytes;
  }

  recordTiming(name: PerformanceTimingName, durationMs: number): void {
    this.timings[name] = normalizeDuration(durationMs);
    this.sampleHeap();
  }

  recordScrollFrame(durationMs: number): void {
    this.scrollSamples.push(normalizeDuration(durationMs));
    if (this.scrollSamples.length > MAX_SCROLL_SAMPLES) {
      this.scrollSamples.shift();
    }
    this.sampleHeap();
  }

  recordCancellation(passed: boolean): void {
    this.cancellationPassed = passed;
  }

  recordCleanup(passed: boolean): void {
    this.cleanupPassed = passed;
  }

  recordWorkSlice(durationMs: number): void {
    const duration = normalizeDuration(durationMs);
    this.maximumWorkSliceMs = Math.max(this.maximumWorkSliceMs, duration);
    if (duration >= 50) {
      this.longTaskCount += 1;
    }
    this.sampleHeap();
  }

  recordYield(count = 1): void {
    this.yieldCount += normalizeCount(count);
  }

  recordBackgroundPreemption(count = 1): void {
    this.backgroundPreemptions += normalizeCount(count);
  }

  recordMemoryPressure(level: ReaderMemoryPressureLevel): void {
    if (
      level === "critical" ||
      (level === "elevated" && this.memoryPressure === "normal")
    ) {
      this.memoryPressure = level;
    }
  }

  recordUninterruptibleDocxRender(durationMs: number): void {
    this.uninterruptiblePhase = {
      name: "docx-preview-render",
      durationMs: normalizeDuration(durationMs),
    };
  }

  sampleHeap(): void {
    const heapBytes = readRuntimeHeapBytes();
    if (heapBytes === null) {
      return;
    }
    this.heapBaselineBytes ??= heapBytes;
    this.peakHeapBytes = Math.max(this.peakHeapBytes ?? heapBytes, heapBytes);
  }

  snapshot(resources: {
    domNodes: number;
    cacheEntries: number;
    cacheLimit: number;
    retainedResources: number;
  }): ReaderPerformanceSnapshot {
    this.sampleHeap();
    const heapGrowthBytes =
      this.peakHeapBytes === null || this.heapBaselineBytes === null
        ? null
        : Math.max(0, this.peakHeapBytes - this.heapBaselineBytes);
    return {
      timings: {
        ...this.timings,
        scrollFrameP95Ms:
          this.scrollSamples.length > 0
            ? round(percentile(this.scrollSamples, 0.95))
            : null,
      },
      memory: {
        peakHeapMiB:
          heapGrowthBytes === null
            ? null
            : round(heapGrowthBytes / (1024 * 1024)),
        measurement:
          heapGrowthBytes === null
            ? "unavailable"
            : "chromium-used-js-heap",
      },
      dom: {
        maximumNodes: normalizeCount(resources.domNodes),
      },
      cache: {
        entries: normalizeCount(resources.cacheEntries),
        limit: normalizeCount(resources.cacheLimit),
      },
      resources: {
        retained: normalizeCount(resources.retainedResources),
      },
      lifecycle: {
        cancellationPassed: this.cancellationPassed,
        cleanupPassed: this.cleanupPassed,
      },
      responsiveness: {
        maximumWorkSliceMs: round(this.maximumWorkSliceMs),
        longTaskCount: this.longTaskCount,
        yieldCount: this.yieldCount,
        backgroundPreemptions: this.backgroundPreemptions,
        memoryPressure: this.memoryPressure,
        uninterruptiblePhase: this.uninterruptiblePhase,
      },
    };
  }
}

export function createReaderPerformanceDiagnosticReport(
  format: ReaderFormat,
  file: ReaderDiagnosticFile,
  snapshot: ReaderPerformanceSnapshot,
): ReaderPerformanceDiagnosticReport {
  return createReaderDiagnosticReport(
    format,
    "performance",
    {
      ...file,
      name: basename(file.name),
    },
    `${format.toUpperCase()} performance snapshot`,
    {
      measurement: "obsidian-desktop-runtime",
      ...snapshot,
      privacy: {
        containsDocumentContent: false,
        containsSpeakerNotes: false,
        containsCellValues: false,
        containsInternalXml: false,
        containsVaultPath: false,
      },
    },
  );
}

function readRuntimeHeapBytes(): number | null {
  if (typeof window === "undefined") {
    return null;
  }
  const runtimePerformance = window.performance as Performance & {
    memory?: { usedJSHeapSize?: number };
  };
  const value = runtimePerformance.memory?.usedJSHeapSize;
  return Number.isFinite(value) && value !== undefined && value >= 0
    ? value
    : null;
}

function normalizeDuration(value: number): number {
  return round(Number.isFinite(value) ? Math.max(0, value) : 0);
}

function normalizeCount(value: number): number {
  return Number.isFinite(value) ? Math.max(0, Math.floor(value)) : 0;
}

function percentile(values: readonly number[], quantile: number): number {
  const sorted = [...values].sort((left, right) => left - right);
  const index = Math.min(
    sorted.length - 1,
    Math.max(0, Math.ceil(sorted.length * quantile) - 1),
  );
  return sorted[index] ?? 0;
}

function basename(value: string): string {
  return value.split(/[\\/]/).at(-1) ?? "";
}

function round(value: number): number {
  return Math.round(value * 100) / 100;
}
