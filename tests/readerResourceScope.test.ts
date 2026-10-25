import assert from "node:assert/strict";
import test from "node:test";

import { ReaderResourceScope } from "../src/reader/resourceScope";

void test("ReaderResourceScope returns every resource category to baseline", async () => {
  const errors: unknown[] = [];
  const released: string[] = [];
  const target = new EventTarget();
  const scope = new ReaderResourceScope((error) => {
    errors.push(error);
  });
  scope.track("blobUrls", "blob:test", () => released.push("blob"));
  scope.track("caches", "cache", () => released.push("cache"));
  scope.track("timers", 1, () => released.push("timer"));
  scope.track("observers", "observer", () => released.push("observer"));
  scope.listen(target, "reader-test", () => released.push("event"));
  scope.runTask(async () => {
    await Promise.resolve();
  });

  assert.deepEqual(scope.snapshot(), {
    blobUrls: 1,
    caches: 1,
    tasks: 1,
    timers: 1,
    observers: 1,
    listeners: 1,
  });
  await scope.drainTasks();
  scope.clear();
  assert.deepEqual(scope.snapshot(), {
    blobUrls: 0,
    caches: 0,
    tasks: 0,
    timers: 0,
    observers: 0,
    listeners: 0,
  });
  assert.deepEqual(released.sort(), ["blob", "cache", "observer", "timer"]);
  assert.deepEqual(errors, []);
});

void test("ReaderResourceScope catches task and cleanup failures", async () => {
  const errors: unknown[] = [];
  const scope = new ReaderResourceScope((error) => {
    errors.push(error);
  });
  scope.runTask(async () => {
    throw new Error("cancelled task");
  });
  scope.track("caches", "broken", () => {
    throw new Error("broken cleanup");
  });

  await scope.drainTasks();
  scope.clear();
  assert.equal(errors.length, 2);
  assert.equal(scope.snapshot().tasks, 0);
  assert.equal(scope.snapshot().caches, 0);
});

void test("replacing a tracked resource cannot release its successor", () => {
  const released: string[] = [];
  const scope = new ReaderResourceScope();
  const releaseFirst = scope.track("caches", "document", () => {
    released.push("first");
  });
  scope.track("caches", "document", () => {
    released.push("second");
  });

  assert.deepEqual(released, ["first"]);
  releaseFirst();
  assert.equal(scope.snapshot().caches, 1);
  scope.clear();
  assert.deepEqual(released, ["first", "second"]);
});

void test("tracked timers and animation frames settle or cancel at baseline", () => {
  const callbacks = new Map<number, () => void>();
  const cancelled: number[] = [];
  let nextId = 0;
  const scheduler = {
    setTimeout: (callback: () => void): number => {
      const id = ++nextId;
      callbacks.set(id, callback);
      return id;
    },
    clearTimeout: (id: number): void => {
      callbacks.delete(id);
      cancelled.push(id);
    },
    requestAnimationFrame: (callback: () => void): number => {
      const id = ++nextId;
      callbacks.set(id, callback);
      return id;
    },
    cancelAnimationFrame: (id: number): void => {
      callbacks.delete(id);
      cancelled.push(id);
    },
  } as unknown as Window;
  const scope = new ReaderResourceScope();
  let fired = 0;
  const timer = scope.setTimer(scheduler, () => {
    fired += 1;
  }, 10);
  scope.requestFrame(scheduler, () => {
    fired += 1;
  });
  assert.equal(scope.snapshot().timers, 2);

  callbacks.get(timer)?.();
  callbacks.delete(timer);
  assert.equal(fired, 1);
  assert.equal(scope.snapshot().timers, 1);
  scope.clear();
  assert.equal(scope.snapshot().timers, 0);
  assert.equal(cancelled.length, 1);
});
