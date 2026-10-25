import assert from "node:assert/strict";
import test from "node:test";

import {
  ReaderWorkCoordinator,
  runCooperativeWork,
} from "../src/reader/cooperativeWork";

void test("cooperative work adapts batches and preserves output order", async () => {
  const output: number[] = [];
  let clock = 0;
  const diagnostics = await runCooperativeWork(
    20,
    (start, end) => {
      for (let index = start; index < end; index += 1) {
        output.push(index);
        clock += 2;
      }
    },
    {
      initialBatchSize: 4,
      maximumBatchSize: 8,
      timeSliceMs: 5,
      now: () => clock,
      yieldControl: async () => undefined,
    },
  );

  assert.deepEqual(output, Array.from({ length: 20 }, (_, index) => index));
  assert.ok(diagnostics.yieldCount > 0);
  assert.equal(diagnostics.maximumWorkSliceMs, 8);
  assert.equal(diagnostics.cancelled, false);
});

void test("cooperative work cancels without processing another batch", async () => {
  let batches = 0;
  const diagnostics = await runCooperativeWork(
    100,
    () => {
      batches += 1;
    },
    {
      initialBatchSize: 5,
      isCancelled: () => batches === 1,
      yieldControl: async () => undefined,
    },
  );
  assert.equal(batches, 1);
  assert.equal(diagnostics.processedUnits, 5);
  assert.equal(diagnostics.cancelled, true);
});

void test("foreground work preempts yielding background work", async () => {
  const coordinator = new ReaderWorkCoordinator();
  const finishForeground = coordinator.beginForeground();
  let yields = 0;
  const continued = await coordinator.yieldBackground(async () => {
    yields += 1;
    finishForeground();
  });

  assert.equal(continued, true);
  assert.equal(yields, 1);
  assert.deepEqual(coordinator.snapshot(), {
    foregroundTasks: 0,
    backgroundPreemptions: 1,
  });
});
