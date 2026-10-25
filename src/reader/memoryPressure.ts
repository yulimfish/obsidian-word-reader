export type ReaderMemoryPressureLevel = "normal" | "elevated" | "critical";

export type ReaderMemoryPressureSource =
  | "heap"
  | "package"
  | "images"
  | "dom"
  | "cache";

export interface ReaderMemoryPressureInput {
  heapUsedBytes?: number;
  heapLimitBytes?: number;
  packageBytes?: number;
  packageLimitBytes?: number;
  imageBytes?: number;
  imageLimitBytes?: number;
  domNodes?: number;
  domLimit?: number;
  cacheEntries?: number;
  cacheLimit?: number;
}

export interface ReaderMemoryPolicy {
  level: ReaderMemoryPressureLevel;
  ratio: number;
  sources: ReaderMemoryPressureSource[];
  allowPreload: boolean;
  backgroundConcurrency: number;
  batchScale: number;
}

const ELEVATED_RATIO = 0.7;
const CRITICAL_RATIO = 0.9;

/** Converts local resource budgets into one deterministic degradation policy. */
export function evaluateReaderMemoryPressure(
  input: ReaderMemoryPressureInput,
  maximumBackgroundConcurrency = 2,
): ReaderMemoryPolicy {
  const ratios = new Map<ReaderMemoryPressureSource, number>();
  addRatio(ratios, "heap", input.heapUsedBytes, input.heapLimitBytes);
  addRatio(ratios, "package", input.packageBytes, input.packageLimitBytes);
  addRatio(ratios, "images", input.imageBytes, input.imageLimitBytes);
  addRatio(ratios, "dom", input.domNodes, input.domLimit);
  addRatio(ratios, "cache", input.cacheEntries, input.cacheLimit);
  const ratio = Math.max(0, ...ratios.values());
  const level: ReaderMemoryPressureLevel = ratio >= CRITICAL_RATIO
    ? "critical"
    : ratio >= ELEVATED_RATIO
      ? "elevated"
      : "normal";
  const sources = [...ratios.entries()]
    .filter(([, value]) => value === ratio && value > 0)
    .map(([source]) => source);
  const maximum = Math.max(1, Math.floor(maximumBackgroundConcurrency));
  return {
    level,
    ratio: round(ratio),
    sources,
    allowPreload: level !== "critical",
    backgroundConcurrency: level === "critical"
      ? 0
      : level === "elevated"
        ? 1
        : maximum,
    batchScale: level === "critical" ? 0.25 : level === "elevated" ? 0.5 : 1,
  };
}

export function readRuntimeHeapBudget(): {
  heapUsedBytes?: number;
  heapLimitBytes?: number;
} {
  if (typeof window === "undefined") {
    return {};
  }
  const runtimePerformance = window.performance as Performance & {
    memory?: {
      usedJSHeapSize?: number;
      jsHeapSizeLimit?: number;
    };
  };
  const heapUsedBytes = readFinite(runtimePerformance.memory?.usedJSHeapSize);
  const heapLimitBytes = readFinite(runtimePerformance.memory?.jsHeapSizeLimit);
  return { heapUsedBytes, heapLimitBytes };
}

function addRatio(
  target: Map<ReaderMemoryPressureSource, number>,
  source: ReaderMemoryPressureSource,
  value: number | undefined,
  limit: number | undefined,
): void {
  if (
    !Number.isFinite(value) ||
    !Number.isFinite(limit) ||
    value === undefined ||
    limit === undefined ||
    value < 0 ||
    limit <= 0
  ) {
    return;
  }
  target.set(source, value / limit);
}

function readFinite(value: number | undefined): number | undefined {
  return Number.isFinite(value) && value !== undefined && value >= 0
    ? value
    : undefined;
}

function round(value: number): number {
  return Math.round(value * 1_000) / 1_000;
}
