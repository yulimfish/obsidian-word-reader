export type ReaderResourceKind =
  | "blobUrls"
  | "caches"
  | "timers"
  | "observers"
  | "listeners";

export interface ReaderResourceSnapshot {
  blobUrls: number;
  caches: number;
  tasks: number;
  timers: number;
  observers: number;
  listeners: number;
}

type ResourceRelease = () => void;

interface TrackedResource {
  release: ResourceRelease;
}

const RESOURCE_KINDS: readonly ReaderResourceKind[] = [
  "blobUrls",
  "caches",
  "timers",
  "observers",
  "listeners",
];

/**
 * Owns resources that outlive a single synchronous reader operation. Cleanup
 * is idempotent and best-effort so one broken third-party disposer cannot
 * prevent the remaining resources from being released.
 */
export class ReaderResourceScope {
  private readonly resources = new Map<
    ReaderResourceKind,
    Map<unknown, TrackedResource>
  >();
  private readonly tasks = new Set<Promise<void>>();

  constructor(
    private readonly onError: (error: unknown) => void = () => undefined,
  ) {
    for (const kind of RESOURCE_KINDS) {
      this.resources.set(kind, new Map());
    }
  }

  track(
    kind: ReaderResourceKind,
    resource: unknown,
    release: ResourceRelease,
  ): ResourceRelease {
    const entries = this.resources.get(kind);
    const previous = entries?.get(resource);
    if (previous) {
      entries?.delete(resource);
      this.releaseSafely(previous.release);
    }
    const tracked = { release };
    entries?.set(resource, tracked);
    let active = true;
    return () => {
      if (!active) {
        return;
      }
      active = false;
      if (entries?.get(resource) === tracked) {
        entries.delete(resource);
        this.releaseSafely(release);
      }
    };
  }

  forget(kind: ReaderResourceKind, resource: unknown): void {
    this.resources.get(kind)?.delete(resource);
  }

  runTask(
    task: () => void | Promise<unknown>,
    onError: (error: unknown) => void = this.onError,
  ): void {
    let pending: Promise<void>;
    pending = Promise.resolve()
      .then(task)
      .then(() => undefined)
      .catch((error: unknown) => {
        onError(error);
      })
      .finally(() => {
        this.tasks.delete(pending);
      });
    this.tasks.add(pending);
  }

  async drainTasks(): Promise<void> {
    while (this.tasks.size > 0) {
      await Promise.all([...this.tasks]);
    }
  }

  listen<K extends keyof DocumentEventMap>(
    target: Document,
    type: K,
    listener: (event: DocumentEventMap[K]) => void,
    options?: boolean | AddEventListenerOptions,
  ): ResourceRelease;
  listen<K extends keyof HTMLElementEventMap>(
    target: HTMLElement,
    type: K,
    listener: (event: HTMLElementEventMap[K]) => void,
    options?: boolean | AddEventListenerOptions,
  ): ResourceRelease;
  listen(
    target: EventTarget,
    type: string,
    listener: EventListener,
    options?: boolean | AddEventListenerOptions,
  ): ResourceRelease;
  listen(
    target: EventTarget,
    type: string,
    listener: EventListener,
    options?: boolean | AddEventListenerOptions,
  ): ResourceRelease {
    target.addEventListener(type, listener, options);
    const key = { target, type, listener };
    return this.track("listeners", key, () => {
      target.removeEventListener(type, listener, options);
    });
  }

  observe(
    observer: Pick<ResizeObserver, "observe" | "disconnect">,
    target: Element,
  ): ResourceRelease {
    observer.observe(target);
    return this.track("observers", observer, () => {
      observer.disconnect();
    });
  }

  setTimer(
    target: Window,
    callback: () => void,
    delay: number,
  ): number {
    const timer = target.setTimeout(() => {
      this.forget("timers", `timeout:${timer}`);
      callback();
    }, delay);
    this.track("timers", `timeout:${timer}`, () => {
      target.clearTimeout(timer);
    });
    return timer;
  }

  clearTimer(target: Window, timer: number): void {
    this.forget("timers", `timeout:${timer}`);
    target.clearTimeout(timer);
  }

  requestFrame(target: Window, callback: () => void): number {
    const frame = target.requestAnimationFrame(() => {
      this.forget("timers", `frame:${frame}`);
      callback();
    });
    this.track("timers", `frame:${frame}`, () => {
      target.cancelAnimationFrame(frame);
    });
    return frame;
  }

  cancelFrame(target: Window, frame: number): void {
    this.forget("timers", `frame:${frame}`);
    target.cancelAnimationFrame(frame);
  }

  clear(kind?: ReaderResourceKind): void {
    const kinds = kind ? [kind] : RESOURCE_KINDS;
    for (const currentKind of kinds) {
      const entries = this.resources.get(currentKind);
      if (!entries) {
        continue;
      }
      const releases = [...entries.values()].reverse();
      entries.clear();
      for (const tracked of releases) {
        this.releaseSafely(tracked.release);
      }
    }
  }

  snapshot(): ReaderResourceSnapshot {
    return {
      blobUrls: this.count("blobUrls"),
      caches: this.count("caches"),
      tasks: this.tasks.size,
      timers: this.count("timers"),
      observers: this.count("observers"),
      listeners: this.count("listeners"),
    };
  }

  private count(kind: ReaderResourceKind): number {
    return this.resources.get(kind)?.size ?? 0;
  }

  private releaseSafely(release: ResourceRelease): void {
    try {
      release();
    } catch (error) {
      this.onError(error);
    }
  }
}
