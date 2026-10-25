export interface CooperativeWorkOptions {
  initialBatchSize?: number;
  maximumBatchSize?: number;
  timeSliceMs?: number;
  isCancelled?: () => boolean;
  shouldPreempt?: () => boolean;
  yieldControl?: () => Promise<void>;
  now?: () => number;
}

export interface CooperativeWorkDiagnostics {
  processedUnits: number;
  yieldCount: number;
  preemptionCount: number;
  maximumWorkSliceMs: number;
  cancelled: boolean;
}

export type CooperativeBatch = (
  start: number,
  end: number,
) => void | Promise<void>;

const DEFAULT_TIME_SLICE_MS = 8;
const DEFAULT_INITIAL_BATCH_SIZE = 4;
const DEFAULT_MAXIMUM_BATCH_SIZE = 64;

/** Runs bounded, adaptive batches and yields a macrotask between every batch. */
export async function runCooperativeWork(
  totalUnits: number,
  processBatch: CooperativeBatch,
  options: CooperativeWorkOptions = {},
): Promise<CooperativeWorkDiagnostics> {
  const total = normalizeCount(totalUnits);
  const maximumBatchSize = normalizeBatchSize(
    options.maximumBatchSize,
    DEFAULT_MAXIMUM_BATCH_SIZE,
  );
  let batchSize = Math.min(
    maximumBatchSize,
    normalizeBatchSize(
      options.initialBatchSize,
      DEFAULT_INITIAL_BATCH_SIZE,
    ),
  );
  const timeSliceMs = normalizeDuration(
    options.timeSliceMs,
    DEFAULT_TIME_SLICE_MS,
  );
  const now = options.now ?? (() => performance.now());
  const yieldControl = options.yieldControl ?? defaultYieldControl;
  let processedUnits = 0;
  let yieldCount = 0;
  let preemptionCount = 0;
  let maximumWorkSliceMs = 0;

  while (processedUnits < total) {
    if (options.isCancelled?.()) {
      return createDiagnostics(true);
    }
    const preempted = options.shouldPreempt?.() ?? false;
    if (preempted) {
      preemptionCount += 1;
      await yieldControl();
      yieldCount += 1;
      continue;
    }

    const end = Math.min(total, processedUnits + batchSize);
    const startedAt = now();
    await processBatch(processedUnits, end);
    const elapsed = Math.max(0, now() - startedAt);
    maximumWorkSliceMs = Math.max(maximumWorkSliceMs, elapsed);
    processedUnits = end;

    if (elapsed > timeSliceMs && batchSize > 1) {
      batchSize = Math.max(1, Math.floor(batchSize / 2));
    } else if (
      elapsed < timeSliceMs * 0.4 &&
      batchSize < maximumBatchSize
    ) {
      batchSize = Math.min(maximumBatchSize, batchSize * 2);
    }

    if (processedUnits < total) {
      await yieldControl();
      yieldCount += 1;
    }
  }

  return createDiagnostics(false);

  function createDiagnostics(cancelled: boolean): CooperativeWorkDiagnostics {
    return {
      processedUnits,
      yieldCount,
      preemptionCount,
      maximumWorkSliceMs,
      cancelled,
    };
  }
}

export interface ReaderWorkCoordinatorSnapshot {
  foregroundTasks: number;
  backgroundPreemptions: number;
}

/** Coordinates foreground interaction with cooperatively yielding background work. */
export class ReaderWorkCoordinator {
  private foregroundTasks = 0;
  private backgroundPreemptions = 0;

  get foregroundActive(): boolean {
    return this.foregroundTasks > 0;
  }

  beginForeground(): () => void {
    this.foregroundTasks += 1;
    let active = true;
    return () => {
      if (!active) {
        return;
      }
      active = false;
      this.foregroundTasks = Math.max(0, this.foregroundTasks - 1);
    };
  }

  async runForeground<Value>(task: () => Promise<Value>): Promise<Value> {
    const finish = this.beginForeground();
    try {
      return await task();
    } finally {
      finish();
    }
  }

  async yieldBackground(
    yieldControl: () => Promise<void> = defaultYieldControl,
    isCancelled: () => boolean = () => false,
  ): Promise<boolean> {
    let preempted = false;
    do {
      if (isCancelled()) {
        return false;
      }
      preempted ||= this.foregroundActive;
      await yieldControl();
    } while (this.foregroundActive);
    if (preempted) {
      this.backgroundPreemptions += 1;
    }
    return !isCancelled();
  }

  resetDiagnostics(): void {
    this.backgroundPreemptions = 0;
  }

  snapshot(): ReaderWorkCoordinatorSnapshot {
    return {
      foregroundTasks: this.foregroundTasks,
      backgroundPreemptions: this.backgroundPreemptions,
    };
  }
}

function defaultYieldControl(): Promise<void> {
  return Promise.resolve();
}

function normalizeCount(value: number): number {
  return Number.isFinite(value) ? Math.max(0, Math.floor(value)) : 0;
}

function normalizeBatchSize(
  value: number | undefined,
  fallback: number,
): number {
  return Number.isFinite(value) && value !== undefined && value >= 1
    ? Math.floor(value)
    : fallback;
}

function normalizeDuration(
  value: number | undefined,
  fallback: number,
): number {
  return Number.isFinite(value) && value !== undefined && value > 0
    ? value
    : fallback;
}
