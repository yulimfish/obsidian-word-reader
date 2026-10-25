import assert from "node:assert/strict";
import test from "node:test";

import { evaluateReaderMemoryPressure } from "../src/reader/memoryPressure";

void test("memory pressure reduces work before stopping preload", () => {
  const normal = evaluateReaderMemoryPressure({
    packageBytes: 50,
    packageLimitBytes: 100,
  }, 4);
  const elevated = evaluateReaderMemoryPressure({
    cacheEntries: 7,
    cacheLimit: 10,
  }, 4);
  const critical = evaluateReaderMemoryPressure({
    heapUsedBytes: 95,
    heapLimitBytes: 100,
  }, 4);

  assert.deepEqual(
    [normal.level, normal.backgroundConcurrency, normal.allowPreload],
    ["normal", 4, true],
  );
  assert.deepEqual(
    [elevated.level, elevated.backgroundConcurrency, elevated.allowPreload],
    ["elevated", 1, true],
  );
  assert.deepEqual(
    [critical.level, critical.backgroundConcurrency, critical.allowPreload],
    ["critical", 0, false],
  );
  assert.deepEqual(critical.sources, ["heap"]);
});
