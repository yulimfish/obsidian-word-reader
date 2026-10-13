import assert from "node:assert/strict";
import test from "node:test";
import {
  DEFAULT_METAFILE_LIMITS,
  MetafileParseError,
  parseMetafileRecords,
} from "../../src/ooxml/metafileRecords";
import {
  createSyntheticEmfFixture,
  createSyntheticWmfFixture,
} from "./fixtureBuilders";

const SCRIPT_SHAPED_BYTES = Uint8Array.from([
  0x3c, 0x73, 0x63, 0x72, 0x69, 0x70, 0x74, 0x3e, 0x78,
  0x3c, 0x2f, 0x73, 0x63, 0x72, 0x69, 0x70, 0x74, 0x3e,
]);

void test("frames standard and placeable WMF records at exact byte offsets", () => {
  const standard = createSyntheticWmfFixture([
    { type: 0x7fff, payload: Uint8Array.from([0xaa, 0xbb]) },
  ]);
  const standardSet = parseMetafileRecords(standard, "wmf");
  assert.equal(standardSet.format, "wmf");
  assert.equal(standardSet.bounds, null);
  assert.deepEqual(
    standardSet.records.map(({ type, offset, size }) => ({ type, offset, size })),
    [
      { type: 0x7fff, offset: 18, size: 8 },
      { type: 0, offset: 26, size: 6 },
    ],
  );
  assert.deepEqual(standardSet.records[0].payload, Uint8Array.from([0xaa, 0xbb]));

  const placeable = createSyntheticWmfFixture([], { placeable: true });
  const placeableSet = parseMetafileRecords(placeable, "wmf");
  assert.deepEqual(placeableSet.bounds, { left: 0, top: 0, right: 100, bottom: 80 });
  assert.deepEqual(
    placeableSet.records.map(({ type, offset, size }) => ({ type, offset, size })),
    [{ type: 0, offset: 40, size: 6 }],
  );
});

void test("frames EMF header, unknown records and EOF at exact byte offsets", () => {
  const bytes = createSyntheticEmfFixture([
    { type: 0x7fffffff, payload: Uint8Array.from([1, 2, 3, 4]) },
  ]);
  const recordSet = parseMetafileRecords(bytes, "emf");
  assert.equal(recordSet.format, "emf");
  assert.deepEqual(recordSet.bounds, { left: 0, top: 0, right: 100, bottom: 80 });
  assert.deepEqual(
    recordSet.records.map(({ type, offset, size }) => ({ type, offset, size })),
    [
      { type: 1, offset: 0, size: 88 },
      { type: 0x7fffffff, offset: 88, size: 12 },
      { type: 14, offset: 100, size: 20 },
    ],
  );
  assert.deepEqual(recordSet.records[1].payload, Uint8Array.from([1, 2, 3, 4]));
});

void test("accepts an empty EMF EOF palette and a bounded one-entry palette", () => {
  const empty = parseMetafileRecords(createSyntheticEmfFixture([]), "emf");
  assert.equal(empty.records.at(-1)?.size, 20);

  const oneEntryBytes = createSyntheticEmfFixture([], {
    eofPaletteEntries: Uint8Array.from([0, 0, 0xff, 0]),
  });
  const oneEntry = parseMetafileRecords(oneEntryBytes, "emf");
  const view = new DataView(
    oneEntryBytes.buffer,
    oneEntryBytes.byteOffset,
    oneEntryBytes.byteLength,
  );
  const eofRecord = oneEntry.records.at(-1)!;
  const headerPaletteCount = view.getUint32(68, true);
  const eofPaletteCount = view.getUint32(eofRecord.offset + 8, true);
  assert.equal(headerPaletteCount, 1);
  assert.equal(eofPaletteCount, 1);
  assert.equal(headerPaletteCount, eofPaletteCount);
  assert.equal(oneEntry.records.at(-1)?.type, 14);
  assert.equal(oneEntry.records.at(-1)?.size, 24);
});

void test("rejects an EMF header palette count that disagrees with the EOF count", () => {
  const bytes = createSyntheticEmfFixture([], {
    eofPaletteEntries: Uint8Array.from([0, 0, 0xff, 0]),
  });
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const eofOffset = bytes.byteLength - 24;
  assert.equal(view.getUint32(68, true), 1);
  assert.equal(view.getUint32(eofOffset + 8, true), 1);
  view.setUint32(68, 0, true);
  assert.throws(
    () => parseMetafileRecords(bytes, "emf"),
    (error: unknown) =>
      error instanceof MetafileParseError &&
      error.code === "invalid-header" &&
      error.message.includes("palette count"),
  );
});

void test("rejects EMF EOF SizeLast that differs from the record size", () => {
  const bytes = createSyntheticEmfFixture([]);
  const damaged = bytes.slice();
  new DataView(damaged.buffer, damaged.byteOffset, damaged.byteLength).setUint32(
    damaged.byteLength - 4,
    0,
    true,
  );
  expectParseError(damaged, "emf", "invalid-record-size");
});

void test("rejects EMF EOF palette counts that exceed the available record range", () => {
  const bytes = createSyntheticEmfFixture([]);
  const damaged = bytes.slice();
  const view = new DataView(damaged.buffer, damaged.byteOffset, damaged.byteLength);
  view.setUint32(68, 1, true);
  view.setUint32(damaged.byteLength - 20 + 8, 1, true);
  view.setUint32(damaged.byteLength - 20 + 12, 16, true);
  expectParseError(damaged, "emf", "invalid-record-size");
});

void test("rejects EMF EOF palette offsets outside its data area", () => {
  for (const paletteOffset of [12, 0xfffffffc]) {
    const bytes = createSyntheticEmfFixture([]);
    const damaged = bytes.slice();
    const view = new DataView(damaged.buffer, damaged.byteOffset, damaged.byteLength);
    const eofOffset = damaged.byteLength - 20;
    view.setUint32(68, 1, true);
    view.setUint32(eofOffset + 8, 1, true);
    view.setUint32(eofOffset + 12, paletteOffset, true);
    expectParseError(damaged, "emf", "invalid-record-size");
  }
});

void test("rejects EMF EOF palette counts that exceed safe record capacity", () => {
  const bytes = createSyntheticEmfFixture([]);
  const damaged = bytes.slice();
  const view = new DataView(damaged.buffer, damaged.byteOffset, damaged.byteLength);
  const eofOffset = damaged.byteLength - 20;
  view.setUint32(68, 0xffffffff, true);
  view.setUint32(eofOffset + 8, 0xffffffff, true);
  view.setUint32(eofOffset + 12, 16, true);
  expectParseError(damaged, "emf", "invalid-record-size");
});

void test("exports frozen conservative defaults for every shared work limit", () => {
  assert.equal(Object.isFrozen(DEFAULT_METAFILE_LIMITS), true);
  assert.deepEqual(DEFAULT_METAFILE_LIMITS, {
    maxBytes: 16 * 1024 * 1024,
    maxRecords: 50_000,
    maxObjects: 4_096,
    maxStack: 256,
    maxSceneOps: 50_000,
  });
});

void test("keeps WMF escape comments and EMF comments as opaque bounded bytes", () => {
  const wmf = parseMetafileRecords(
    createSyntheticWmfFixture([{ type: 0x626, payload: SCRIPT_SHAPED_BYTES }]),
    "wmf",
  );
  const wmfComment = wmf.records[0];
  assert.equal(wmfComment.type, 0x626);
  assert.ok(wmfComment.payload instanceof Uint8Array);
  assert.deepEqual(wmfComment.payload, SCRIPT_SHAPED_BYTES);
  assert.deepEqual(Object.keys(wmfComment).sort(), ["offset", "payload", "size", "type"]);

  const emf = parseMetafileRecords(
    createSyntheticEmfFixture([{ type: 70, payload: SCRIPT_SHAPED_BYTES }]),
    "emf",
  );
  const emfComment = emf.records[1];
  assert.equal(emfComment.type, 70);
  assert.ok(emfComment.payload instanceof Uint8Array);
  assert.deepEqual(
    emfComment.payload.subarray(0, SCRIPT_SHAPED_BYTES.byteLength),
    SCRIPT_SHAPED_BYTES,
  );
  assert.deepEqual(emfComment.payload.subarray(SCRIPT_SHAPED_BYTES.byteLength), new Uint8Array(2));
  assert.deepEqual(Object.keys(emfComment).sort(), ["offset", "payload", "size", "type"]);
});

void test("rejects zero-sized WMF and EMF records", () => {
  const wmf = createSyntheticWmfFixture([{ type: 0x7fff }]);
  new DataView(wmf.buffer, wmf.byteOffset, wmf.byteLength).setUint32(18, 0, true);
  expectParseError(wmf, "wmf", "invalid-record-size");

  const emf = createSyntheticEmfFixture([{ type: 0x7fffffff }]);
  new DataView(emf.buffer, emf.byteOffset, emf.byteLength).setUint32(92, 0, true);
  expectParseError(emf, "emf", "invalid-record-size");
});

void test("rejects truncated headers and records before exposing record payloads", () => {
  expectParseError(new Uint8Array(17), "wmf", "truncated-header");
  expectParseError(new Uint8Array(87), "emf", "truncated-header");

  const wmfOriginal = createSyntheticWmfFixture([
    { type: 0x7fff, payload: Uint8Array.from([1, 2, 3, 4]) },
  ]);
  const wmf = wmfOriginal.slice(0, wmfOriginal.byteLength - 2);
  new DataView(wmf.buffer, wmf.byteOffset, wmf.byteLength).setUint32(
    6,
    wmf.byteLength / 2,
    true,
  );
  expectParseError(wmf, "wmf", "truncated-record");

  const emfOriginal = createSyntheticEmfFixture([
    { type: 0x7fffffff, payload: Uint8Array.from([1, 2, 3, 4]) },
  ]);
  const emf = emfOriginal.slice(0, emfOriginal.byteLength - 4);
  new DataView(emf.buffer, emf.byteOffset, emf.byteLength).setUint32(48, emf.byteLength, true);
  expectParseError(emf, "emf", "truncated-record");
});

void test("rejects oversized record lengths without overflowing or slicing past input", () => {
  const wmf = createSyntheticWmfFixture([{ type: 0x7fff }]);
  new DataView(wmf.buffer, wmf.byteOffset, wmf.byteLength).setUint32(18, 0xffffffff, true);
  expectParseError(wmf, "wmf", "invalid-record-size");

  const emf = createSyntheticEmfFixture([{ type: 0x7fffffff }]);
  new DataView(emf.buffer, emf.byteOffset, emf.byteLength).setUint32(92, 0xfffffffc, true);
  expectParseError(emf, "emf", "truncated-record");
});

void test("rejects malformed declared lengths, counts and trailing bytes", () => {
  const wmfLengthMismatch = createSyntheticWmfFixture([]);
  const wmfLengthView = new DataView(
    wmfLengthMismatch.buffer,
    wmfLengthMismatch.byteOffset,
    wmfLengthMismatch.byteLength,
  );
  wmfLengthView.setUint32(6, wmfLengthView.getUint32(6, true) + 1, true);
  expectParseError(wmfLengthMismatch, "wmf", "declared-size-mismatch");

  const emfLengthMismatch = createSyntheticEmfFixture([]);
  const emfLengthView = new DataView(
    emfLengthMismatch.buffer,
    emfLengthMismatch.byteOffset,
    emfLengthMismatch.byteLength,
  );
  emfLengthView.setUint32(48, emfLengthView.getUint32(48, true) + 4, true);
  expectParseError(emfLengthMismatch, "emf", "declared-size-mismatch");

  const emfCountMismatch = createSyntheticEmfFixture([]);
  new DataView(emfCountMismatch.buffer, emfCountMismatch.byteOffset, emfCountMismatch.byteLength)
    .setUint32(52, 3, true);
  expectParseError(emfCountMismatch, "emf", "record-count-mismatch");

  const executableLikeTail = Uint8Array.from([
    0x3c, 0x73, 0x63, 0x72, 0x69, 0x70, 0x74, 0x3e,
  ]);
  const wmfTrailing = appendBytes(createSyntheticWmfFixture([]), executableLikeTail);
  new DataView(wmfTrailing.buffer, wmfTrailing.byteOffset, wmfTrailing.byteLength)
    .setUint32(6, wmfTrailing.byteLength / 2, true);
  expectParseError(wmfTrailing, "wmf", "trailing-data");

  const emfTrailing = appendBytes(createSyntheticEmfFixture([]), executableLikeTail);
  new DataView(emfTrailing.buffer, emfTrailing.byteOffset, emfTrailing.byteLength)
    .setUint32(48, emfTrailing.byteLength, true);
  expectParseError(emfTrailing, "emf", "trailing-data");
});

void test("requires a terminal EOF record and rejects bad placeable checksums", () => {
  expectParseError(
    createSyntheticWmfFixture([{ type: 0x7fff }], { appendEof: false }),
    "wmf",
    "missing-eof",
  );
  expectParseError(
    createSyntheticEmfFixture(
      [{ type: 0x7fffffff, payload: new Uint8Array(12) }],
      { appendEof: false },
    ),
    "emf",
    "missing-eof",
  );

  const placeable = createSyntheticWmfFixture([], { placeable: true });
  placeable[20] ^= 0x01;
  expectParseError(placeable, "wmf", "invalid-header");
});

void test("rejects duplicate EOF records and invalid EMF signatures", () => {
  const original = createSyntheticEmfFixture([]);
  const duplicate = appendBytes(original, new Uint8Array(20));
  const duplicateView = new DataView(duplicate.buffer, duplicate.byteOffset, duplicate.byteLength);
  const duplicateOffset = original.byteLength;
  duplicateView.setUint32(duplicateOffset, 14, true);
  duplicateView.setUint32(duplicateOffset + 4, 20, true);
  duplicateView.setUint32(duplicateOffset + 8, 0, true);
  duplicateView.setUint32(duplicateOffset + 12, 0, true);
  duplicateView.setUint32(duplicateOffset + 16, 20, true);
  duplicateView.setUint32(48, duplicate.byteLength, true);
  duplicateView.setUint32(52, 3, true);
  expectParseError(duplicate, "emf", "trailing-data");

  const invalidSignature = original.slice();
  new DataView(
    invalidSignature.buffer,
    invalidSignature.byteOffset,
    invalidSignature.byteLength,
  ).setUint32(40, 0, true);
  expectParseError(invalidSignature, "emf", "invalid-header");
});

void test("reads EMF record offsets relative to a nonzero source byteOffset", () => {
  const source = createSyntheticEmfFixture([
    { type: 0x7fffffff, payload: Uint8Array.from([1, 2, 3, 4]) },
  ]);
  const wrapped = new Uint8Array(source.byteLength + 8);
  wrapped.set(source, 4);
  const view = wrapped.subarray(4, 4 + source.byteLength);
  const records = parseMetafileRecords(view, "emf").records;
  assert.deepEqual(
    records.map(({ type, offset, size }) => ({ type, offset, size })),
    [
      { type: 1, offset: 0, size: 88 },
      { type: 0x7fffffff, offset: 88, size: 12 },
      { type: 14, offset: 100, size: 20 },
    ],
  );
  assert.deepEqual(records[1].payload, Uint8Array.from([1, 2, 3, 4]));
});

void test("bounds object references against declared WMF and EMF handle tables", () => {
  for (const type of [0x12d, 0x1f0]) {
    const wmfValid = createSyntheticWmfFixture(
      [{ type, payload: Uint8Array.from([0, 0]) }],
      { objectCount: 1 },
    );
    assert.equal(parseMetafileRecords(wmfValid, "wmf").records.length, 2);

    const wmfInvalid = createSyntheticWmfFixture(
      [{ type, payload: Uint8Array.from([1, 0]) }],
      { objectCount: 1 },
    );
    expectParseError(wmfInvalid, "wmf", "invalid-handle");
  }

  for (const type of [37, 40]) {
    const emfValid = createSyntheticEmfFixture(
      [{ type, payload: Uint8Array.from([1, 0, 0, 0]) }],
      { handleCount: 2 },
    );
    assert.equal(parseMetafileRecords(emfValid, "emf").records.length, 3);

    const emfInvalid = createSyntheticEmfFixture(
      [{ type, payload: Uint8Array.from([2, 0, 0, 0]) }],
      { handleCount: 2 },
    );
    expectParseError(emfInvalid, "emf", "invalid-handle");
  }
});

void test("enforces byte, record, object and limit-shape caps", () => {
  const wmf = createSyntheticWmfFixture([{ type: 0x7fff }], { objectCount: 2 });
  expectParseError(wmf, "wmf", "file-too-large", {
    ...DEFAULT_METAFILE_LIMITS,
    maxBytes: wmf.byteLength - 1,
  });
  expectParseError(wmf, "wmf", "too-many-records", {
    ...DEFAULT_METAFILE_LIMITS,
    maxRecords: 1,
  });
  expectParseError(wmf, "wmf", "too-many-objects", {
    ...DEFAULT_METAFILE_LIMITS,
    maxObjects: 1,
  });

  const emf = createSyntheticEmfFixture([{ type: 0x7fffffff }], { handleCount: 3 });
  expectParseError(emf, "emf", "too-many-records", {
    ...DEFAULT_METAFILE_LIMITS,
    maxRecords: 2,
  });
  expectParseError(emf, "emf", "too-many-objects", {
    ...DEFAULT_METAFILE_LIMITS,
    maxObjects: 1,
  });
  expectParseError(emf, "emf", "invalid-limits", {
    ...DEFAULT_METAFILE_LIMITS,
    maxStack: -1,
  });

  for (const limits of [
    { ...DEFAULT_METAFILE_LIMITS, maxBytes: Number.NaN },
    { ...DEFAULT_METAFILE_LIMITS, maxRecords: Number.POSITIVE_INFINITY },
    { ...DEFAULT_METAFILE_LIMITS, maxObjects: 1.5 },
  ]) {
    expectParseError(wmf, "wmf", "invalid-limits", limits);
  }
});

void test("treats zero as an enforced work cap rather than an invalid limit shape", () => {
  const noObjects = createSyntheticWmfFixture([]);
  const emf = createSyntheticEmfFixture([]);
  expectParseError(noObjects, "wmf", "file-too-large", {
    ...DEFAULT_METAFILE_LIMITS,
    maxBytes: 0,
  });
  expectParseError(noObjects, "wmf", "too-many-records", {
    ...DEFAULT_METAFILE_LIMITS,
    maxRecords: 0,
  });
  expectParseError(emf, "emf", "too-many-records", {
    ...DEFAULT_METAFILE_LIMITS,
    maxRecords: 0,
  });
  expectParseError(createSyntheticWmfFixture([], { objectCount: 1 }), "wmf", "too-many-objects", {
    ...DEFAULT_METAFILE_LIMITS,
    maxObjects: 0,
  });
  expectParseError(createSyntheticEmfFixture([], { handleCount: 2 }), "emf", "too-many-objects", {
    ...DEFAULT_METAFILE_LIMITS,
    maxObjects: 0,
  });
  assert.equal(
    parseMetafileRecords(noObjects, "wmf", {
      ...DEFAULT_METAFILE_LIMITS,
      maxObjects: 0,
      maxStack: 0,
      maxSceneOps: 0,
    }).records.length,
    1,
  );
});

function appendBytes(bytes: Uint8Array, suffix: Uint8Array): Uint8Array {
  const combined = new Uint8Array(bytes.byteLength + suffix.byteLength);
  combined.set(bytes);
  combined.set(suffix, bytes.byteLength);
  return combined;
}

function expectParseError(
  bytes: Uint8Array,
  format: "wmf" | "emf",
  code: MetafileParseError["code"],
  limits?: typeof DEFAULT_METAFILE_LIMITS,
): void {
  assert.throws(
    () => parseMetafileRecords(bytes, format, limits),
    (error: unknown) =>
      error instanceof MetafileParseError && error.code === code,
    `expected ${format} parse error ${code}`,
  );
}
