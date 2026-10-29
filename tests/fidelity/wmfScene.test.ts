import assert from "node:assert/strict";
import test from "node:test";
import { DEFAULT_METAFILE_LIMITS, MetafileParseError, parseMetafileRecords } from "../../src/ooxml/metafileRecords";
import type { MetafileLimits } from "../../src/ooxml/metafileRecords";
import { MetafileDecodeError } from "../../src/ooxml/metafileScene";
import { decodeWmfScene } from "../../src/ooxml/wmfScene";
import { createSyntheticWmfFixture } from "./fixtureBuilders";
import type { SyntheticMetafileRecord } from "./fixtureBuilders";

const options = { isCancelled: () => false, yieldControl: () => Promise.resolve() };
const record = (type: number, payload: Uint8Array = new Uint8Array()): SyntheticMetafileRecord => ({ type, payload });
function words(...values: number[]): Uint8Array {
  const bytes = new Uint8Array(values.length * 2);
  const view = new DataView(bytes.buffer);
  values.forEach((value, index) => view.setUint16(index * 2, value & 0xffff, true));
  return bytes;
}
function font(name: string | number[] = "Times New Roman", charset = 0, height = -20): SyntheticMetafileRecord {
  const bytes = new Uint8Array(50);
  bytes.set(words(height, 0, 0, 0, 400));
  bytes[13] = charset;
  bytes.set(typeof name === "string" ? Array.from(name, (c) => c.charCodeAt(0)) : name, 18);
  return record(0x2fb, bytes);
}
const pen = (red = 0, width = 2, style = 0): SyntheticMetafileRecord => record(0x2fa, words(style, width, 0, red, 0));
function text(bytes: number[] = [65, 66], advances?: number[], flags = 0, x = 30, y = 40): SyntheticMetafileRecord {
  const start = flags & 6 ? 16 : 8;
  const end = start + bytes.length + bytes.length % 2;
  const payload = new Uint8Array(end + (advances?.length ?? 0) * 2);
  payload.set(words(y, x, bytes.length, flags));
  if (flags & 6) payload.set(words(1, 2, 51, 62), 8);
  payload.set(bytes, start);
  if (advances) payload.set(words(...advances), end);
  return record(0xa32, payload);
}
function parse(records: SyntheticMetafileRecord[], capacity = 2, limits: MetafileLimits = DEFAULT_METAFILE_LIMITS) {
  return parseMetafileRecords(createSyntheticWmfFixture(records, { objectCount: capacity }), "wmf", limits);
}
function parsePlaceable(
  records: SyntheticMetafileRecord[],
  bounds: { left: number; top: number; right: number; bottom: number },
) {
  const bytes = createSyntheticWmfFixture(records, { placeable: true });
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  view.setInt16(6, bounds.left, true);
  view.setInt16(8, bounds.top, true);
  view.setInt16(10, bounds.right, true);
  view.setInt16(12, bounds.bottom, true);
  let checksum = 0;
  for (let offset = 0; offset < 20; offset += 2) checksum ^= view.getUint16(offset, true);
  view.setUint16(20, checksum, true);
  return parseMetafileRecords(bytes, "wmf");
}
async function fails(records: SyntheticMetafileRecord[], index: number, code: string, capacity = 2, limits: MetafileLimits = DEFAULT_METAFILE_LIMITS) {
  const parsed = parse(records, capacity, limits);
  await assert.rejects(decodeWmfScene(parsed, options), (error: unknown) =>
    error instanceof MetafileDecodeError && error.code === code &&
    error.recordOffset === parsed.records[index].offset && error.recordType === parsed.records[index].type);
}

void test("accepts zero-padded WMF scalar records emitted by equation editors", async () => {
  const scene = await decodeWmfScene(parse([
    record(0x102, words(1, 0)),
  ]), options);
  assert.deepEqual(scene.operations, []);
});

void test("rejects unverified mixed symbol runs, zero font sizes, and invalid brush colors", async () => {
  const mixed = await decodeWmfScene(parse([font("MT Extra", 2), record(0x12d, words(0)), text([0x36, 0x61], [4, 5])]), options);
  assert.equal(mixed.operations[0]?.kind, "text");
  if (mixed.operations[0]?.kind === "text") assert.equal(mixed.operations[0].text, null);
  assert.ok(mixed.warnings.some((warning) => warning.kind === "font-glyph-gap"));
  await fails([font("Times New Roman", 161), record(0x12d, words(0)), text([0x80], [4])], 2, "unsupported-charset");
  await fails([font("Symbol", 2, 0), record(0x12d, words(0)), text([0x3d], [4])], 2, "font-metrics-required");
  await fails([record(0x2fc, Uint8Array.from([0, 0, 17, 34, 51, 2, 0, 0]))], 0, "unsupported-color");
});

void test("maps logical window coordinates and snapshots selected solid pen and current position", async () => {
  const scene = await decodeWmfScene(parse([
    record(0x20b, words(10, 20)), record(0x20c, words(80, 100)),
    pen(0x22), record(0x12d, words(0)), record(0x214, words(15, 25)),
    record(0x213, words(30, 40)), pen(0x44, 4), record(0x12d, words(1)),
    record(0x20b, words(0, 0)), record(0x213, words(50, 60)),
  ]), options);
  assert.deepEqual(scene.viewBox, { x: 0, y: 0, width: 100, height: 80 });
  const [first, second] = scene.operations;
  assert.equal(first.kind, "path"); assert.equal(second.kind, "path");
  if (first.kind !== "path" || second.kind !== "path") return;
  assert.deepEqual(first.commands, [{ type: "moveTo", point: { x: 25, y: 15 } }, { type: "lineTo", point: { x: 40, y: 30 } }]);
  assert.deepEqual(first.transform, { a: 1, b: 0, c: 0, d: 1, e: -20, f: -10 });
  assert.equal(first.stroke?.color.red, 34); assert.equal(first.stroke?.width, 2);
  assert.equal(second.stroke?.color.red, 68); assert.equal(second.stroke?.width, 4);
  assert.deepEqual(second.commands[0], { type: "moveTo", point: { x: 40, y: 30 } });
  assert.deepEqual(second.transform, { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 });
});

void test("preserves signed mirrored extents and placeable viewport scale", async () => {
  const records = parseMetafileRecords(createSyntheticWmfFixture([
    record(0x20c, words(-40, 50)), record(0x213, words(10, 20)),
  ], { placeable: true }), "wmf");
  const scene = await decodeWmfScene(records, options);
  assert.deepEqual(scene.viewBox, { x: 0, y: 0, width: 100, height: 80 });
  assert.deepEqual(scene.operations[0].kind === "path" && scene.operations[0].transform, { a: 2, b: 0, c: 0, d: -2, e: 0, f: 0 });
});

void test("normalizes nonzero placeable bounds origin into the scene viewport", async () => {
  const records = parsePlaceable([record(0x213, words(10, 20))], { left: 20, top: 10, right: 120, bottom: 90 });
  const scene = await decodeWmfScene(records, options);
  assert.deepEqual(scene.viewBox, { x: 0, y: 0, width: 100, height: 80 });
  const op = scene.operations[0];
  assert.equal(op.kind, "path");
  if (op.kind !== "path") return;
  assert.deepEqual(op.transform, { a: 1, b: 0, c: 0, d: 1, e: -20, f: -10 });
  assert.deepEqual(op.commands[1], { type: "lineTo", point: { x: 20, y: 10 } });
  assert.equal(op.commands[1].point.x * op.transform.a + op.transform.e, 0);
  assert.equal(op.commands[1].point.y * op.transform.d + op.transform.f, 0);
});

void test("preserves negative font height, reference alignment, colors, background, exact advances and RTL clipping", async () => {
  const scene = await decodeWmfScene(parse([
    font(), record(0x12d, words(0)), record(0x102, words(1)),
    record(0x201, words(0x2211, 0x33)), record(0x209, words(0x5544, 0x66)),
    record(0x12e, words(0x11e)), text([65, 66], [7, -3], 0x86),
  ]), options);
  assert.deepEqual(scene.requiredFonts, ["Times New Roman"]);
  const op = scene.operations[0]; assert.equal(op.kind, "text");
  if (op.kind !== "text") return;
  assert.equal(op.text, "AB"); assert.equal(op.x, 30); assert.equal(op.y, 40);
  assert.equal(op.fontSize, 20); assert.equal(op.logicalFont?.height, -20);
  assert.equal(op.logicalFont?.charset, 0);
  assert.deepEqual(op.advances, [7, -3]); assert.deepEqual(op.sourceAdvances, [7, -3]);
  assert.equal(op.direction, "rtl"); assert.equal(op.textAlign, 0x11e);
  assert.equal(op.backgroundMode, "transparent");
  assert.deepEqual(op.backgroundColor, { red: 17, green: 34, blue: 51, alpha: 1 });
  assert.deepEqual(op.color, { red: 68, green: 85, blue: 102, alpha: 1 });
  assert.deepEqual(op.clipRect, { x: 1, y: 2, width: 50, height: 60 });
  assert.deepEqual(op.opaqueRect, op.clipRect);
  assert.equal(op.textOptions, 0x86);
});

void test("decodes GB2312 charset 134 using codepage 936 for Chinese face names and DBCS text", async () => {
  const scene = await decodeWmfScene(parse([
    font([0xcb, 0xce, 0xcc, 0xe5], 134), record(0x12d, words(0)),
    text([0xd6, 0xd0, 0xce, 0xc4], [3, 5, 7, 2]),
  ]), options);
  assert.deepEqual(scene.requiredFonts, ["宋体"]);
  const op = scene.operations[0]; assert.equal(op.kind, "text");
  if (op.kind !== "text") return;
  assert.equal(op.fontFamily, "宋体"); assert.equal(op.text, "中文");
  assert.equal(op.codePage, 936); assert.deepEqual(op.sourceBytes, [214, 208, 206, 196]);
  assert.deepEqual(op.sourceAdvances, [3, 5, 7, 2]); assert.deepEqual(op.advances, [8, 9]);
});

for (const family of ["Symbol", "MT Extra"]) {
  void test(`preserves ${family} codepage 42 bytes without guessed Unicode equations and reports exact glyph dependency`, async () => {
    const scene = await decodeWmfScene(parse([font(family, 2), record(0x12d, words(0)), text([0x61, 0xb5, 0xf2], [4, 5, 6])]), options);
    const op = scene.operations[0]; assert.equal(op.kind, "text");
    if (op.kind !== "text") return;
    assert.equal(op.text, null); assert.equal(op.encoding, "font-specific");
    assert.equal(op.codePage, 42); assert.deepEqual(op.sourceBytes, [97, 181, 242]);
    assert.equal(op.logicalFont?.charset, 2); assert.deepEqual(op.advances, [4, 5, 6]);
    assert.deepEqual(scene.requiredFonts, [family]);
    assert.ok(scene.warnings.some((warning) => warning.kind === "font-glyph-gap" && warning.fontName === family && warning.recordType === 0xa32 && warning.byteCodes.join(",") === "97,181,242"));
  });
}

void test("retains LOGFONT cell-height and style fields instead of silently losing font semantics", async () => {
  const object = font("Arial", 0, 24);
  object.payload!.set(words(24, 8, 900, 900, 700));
  object.payload!.set([1, 1, 1], 10);
  const scene = await decodeWmfScene(parse([object, record(0x12d, words(0)), text()]), options);
  const op = scene.operations[0]; assert.equal(op.kind, "text");
  if (op.kind !== "text") return;
  assert.equal(op.logicalFont?.height, 24); assert.equal(op.logicalFont?.width, 8);
  assert.equal(op.logicalFont?.escapement, 900); assert.equal(op.logicalFont?.orientation, 900);
  assert.equal(op.logicalFont?.weight, 700); assert.equal(op.logicalFont?.italic, true);
  assert.equal(op.logicalFont?.underline, true); assert.equal(op.logicalFont?.strikeOut, true);
  assert.ok(scene.warnings.some((warning) => warning.kind === "font-metrics-required"));
});

void test("enforces capacity zero/one and reuses the lowest freed object slot", async () => {
  await fails([pen()], 0, "invalid-create", 0);
  await fails([pen(), pen()], 1, "invalid-create", 1);
  const scene = await decodeWmfScene(parse([
    pen(), record(0x1f0, words(0)), pen(0x77), record(0x12d, words(0)), record(0x213, words(3, 4)),
  ], 1, { ...DEFAULT_METAFILE_LIMITS, maxObjects: 1 }), options);
  assert.equal(scene.operations.length, 1);
  const op = scene.operations[0]; assert.equal(op.kind, "path");
  assert.equal(op.kind === "path" && op.stroke?.color.red, 119);
});

void test("rejects deletion of selected objects rather than inventing a default-object replacement", async () => {
  await fails([pen(), record(0x12d, words(0)), record(0x1f0, words(0))], 2, "invalid-delete", 1);
  await fails([font(), record(0x12d, words(0)), record(0x1f0, words(0))], 2, "invalid-delete", 1);
  await fails([
    pen(), record(0x12d, words(0)), record(0x1e), pen(0x22),
    record(0x12d, words(1)), record(0x1f0, words(0)),
  ], 5, "invalid-delete", 2);
});

for (const [type, code] of [[0x12d, "invalid-selection"], [0x1f0, "invalid-delete"]] as const) {
  void test(`rejects uncreated and deleted in-range handles as ${code}`, async () => {
    await fails([record(type, words(0))], 0, code, 1);
    await fails([pen(), record(0x1f0, words(0)), record(type, words(0))], 2, code, 1);
  });
}

void test("leaves out-of-range handles and over-cap headers to the framing parser", () => {
  assert.throws(() => parse([record(0x12d, words(1))], 1), (e: unknown) => e instanceof MetafileParseError && e.code === "invalid-handle");
  assert.throws(() => parse([], 1, { ...DEFAULT_METAFILE_LIMITS, maxObjects: 0 }), (e: unknown) => e instanceof MetafileParseError && e.code === "too-many-objects");
});

void test("bounds saved-state depth excluding initial state and validates restore targets", async () => {
  await fails([record(0x1e)], 0, "max-stack", 0, { ...DEFAULT_METAFILE_LIMITS, maxStack: 0 });
  await fails([record(0x1e), record(0x1e)], 1, "max-stack", 0, { ...DEFAULT_METAFILE_LIMITS, maxStack: 1 });
  await fails([record(0x127, words(-1))], 0, "invalid-restore");
  await fails([record(0x1e), record(0x127, words(2))], 1, "invalid-restore");
  const scene = await decodeWmfScene(parse([record(0x1e), record(0x127, words(-1))], 0, { ...DEFAULT_METAFILE_LIMITS, maxStack: 1 }), options);
  assert.equal(scene.operations.length, 0);
});

void test("restores selected objects, transforms, colors and current position without rolling back the object table", async () => {
  const scene = await decodeWmfScene(parse([
    pen(0x11), record(0x12d, words(0)), record(0x214, words(4, 3)), record(0x1e),
    pen(0x22), record(0x12d, words(1)), record(0x20b, words(7, 8)), record(0x213, words(10, 12)),
    record(0x127, words(1)), record(0x213, words(20, 30)), record(0x12d, words(1)), record(0x213, words(40, 50)),
  ]), options);
  const restored = scene.operations[1]; assert.equal(restored.kind, "path");
  if (restored.kind !== "path") return;
  assert.deepEqual(restored.commands[0], { type: "moveTo", point: { x: 3, y: 4 } });
  assert.equal(restored.stroke?.color.red, 17);
  assert.deepEqual(restored.transform, { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 });
  assert.equal(scene.operations[2].kind === "path" && scene.operations[2].stroke?.color.red, 34);
});

void test("enforces zero/one scene operation caps without counting comments and non-drawing state", async () => {
  const line = record(0x213, words(1, 2));
  await fails([line], 0, "max-scene-ops", 0, { ...DEFAULT_METAFILE_LIMITS, maxSceneOps: 0 });
  await fails([line, line], 1, "max-scene-ops", 0, { ...DEFAULT_METAFILE_LIMITS, maxSceneOps: 1 });
  const scene = await decodeWmfScene(parse([record(0x214, words(0, 0)), record(0x626, words(15, 0)), line], 0, { ...DEFAULT_METAFILE_LIMITS, maxSceneOps: 1 }), options);
  assert.equal(scene.operations.length, 1);
  const empty = await decodeWmfScene(parse([record(0x626, words(15, 0))], 0, { ...DEFAULT_METAFILE_LIMITS, maxSceneOps: 0 }), options);
  assert.equal(empty.operations.length, 0);
});

void test("counts clipped text once toward maxSceneOps and reports exact zero-cap record", async () => {
  const limits = { ...DEFAULT_METAFILE_LIMITS, maxSceneOps: 0 };
  await fails([font(), record(0x12d, words(0)), text([65, 66], undefined, 4)], 2, "max-scene-ops", 1, limits);
  const scene = await decodeWmfScene(parse(
    [font(), record(0x12d, words(0)), text([65, 66], undefined, 4)],
    1,
    { ...DEFAULT_METAFILE_LIMITS, maxSceneOps: 1 },
  ), options);
  assert.equal(scene.operations.length, 1);
  const op = scene.operations[0];
  assert.equal(op.kind, "text");
  assert.deepEqual(op.kind === "text" && op.clipRect, { x: 1, y: 2, width: 50, height: 60 });
});

void test("retains frozen parser caps even after caller limits are weakened", async () => {
  const limits = { ...DEFAULT_METAFILE_LIMITS, maxObjects: 1, maxStack: 0, maxSceneOps: 0 };
  const push = parse([record(0x1e)], 1, limits);
  const line = parse([record(0x213, words(1, 2))], 1, limits);
  limits.maxObjects = 4096; limits.maxStack = 256; limits.maxSceneOps = 50000;
  assert.equal(Object.isFrozen(push.limits), true); assert.equal(push.declaredObjectTableSize, 1);
  await assert.rejects(decodeWmfScene(push, options), (e: unknown) => e instanceof MetafileDecodeError && e.code === "max-stack");
  await assert.rejects(decodeWmfScene(line, options), (e: unknown) => e instanceof MetafileDecodeError && e.code === "max-scene-ops");
});

void test("updates text current position only with known advances and preserves subsequent geometry", async () => {
  const scene = await decodeWmfScene(parse([
    font(), record(0x12d, words(0)), record(0x12e, words(1)), record(0x214, words(6, 5)),
    text([65, 66], [7, 8]), record(0x213, words(10, 30)),
  ]), options);
  const op = scene.operations[0]; assert.equal(op.kind, "text");
  if (op.kind !== "text") return;
  assert.equal(op.x, 5); assert.equal(op.y, 6);
  const path = scene.operations[1]; assert.equal(path.kind, "path");
  assert.deepEqual(path.kind === "path" && path.commands[0], { type: "moveTo", point: { x: 20, y: 6 } });
  await fails([font(), record(0x12d, words(0)), record(0x12e, words(1)), text()], 3, "font-metrics-required");
});

void test("preserves a length-validated opaque escape 15 comment as a diagnostic, never markup or drawing", async () => {
  const payload = new Uint8Array(12); payload.set(words(15, 8)); payload.set([60, 115, 99, 114, 105, 112, 116, 62], 4);
  const parsed = parse([record(0x626, payload)]);
  const scene = await decodeWmfScene(parsed, options);
  assert.equal(scene.operations.length, 0);
  assert.deepEqual(scene.warnings, [{ kind: "opaque-comment", recordOffset: 18, recordType: 0x626, byteLength: 8 }]);
  await fails([record(0x626, words(15, 10))], 0, "invalid-record");
  await fails([record(0x626, words(19, 0))], 0, "unsupported-escape");
  await fails([record(0x626, words(0x1001, 0))], 0, "unsupported-escape");
});

void test("rejects unknown opcodes and unsupported pen/text subtypes at exact offsets", async () => {
  await fails([record(0x7fff)], 0, "unsupported-record");
  await fails([pen(0, 2, 1)], 0, "unsupported-pen");
  await fails([font(), record(0x12d, words(0)), text([65], [2], 0x10)], 2, "font-glyph-gap");
  await fails([font(), record(0x12d, words(0)), text([65], [2], 0x400)], 2, "unsupported-text-options");
  await fails([record(0x102, words(3))], 0, "unsupported-background-mode");
  await fails([record(0x12e, words(4))], 0, "unsupported-text-alignment");
  await fails([record(0x20c, words(0, 1))], 0, "invalid-window");
});

void test("rejects oversized text counts, short strings and mismatched advance arrays before allocations", async () => {
  for (const count of [8193, 0xffff]) {
    await fails([font(), record(0x12d, words(0)), record(0xa32, words(0, 0, count, 0))], 2, "invalid-text-length");
  }
  await fails([font(), record(0x12d, words(0)), record(0xa32, words(0, 0, 4, 0))], 2, "invalid-record");
  await fails([font(), record(0x12d, words(0)), text([65, 66], [1])], 2, "invalid-advances");
  await fails([font(), record(0x12d, words(0)), text([65], [1, 2])], 2, "invalid-advances");
});

void test("rejects malformed encoding and unavailable/locale-dependent code pages without mojibake", async () => {
  await fails([font([0xcb], 134)], 0, "invalid-encoding");
  await fails([font("宋", 134)], 0, "invalid-encoding");
  await fails([font("Arial", 134), record(0x12d, words(0)), text([0x81])], 2, "invalid-encoding");
  await fails([font("Arial", 1), record(0x12d, words(0)), text([0xe9])], 2, "unsupported-charset");
  await fails([font("Arial", 99), record(0x12d, words(0)), text()], 2, "unsupported-charset");
  await fails([record(0xa32, words(0, 0, 0, 0))], 0, "font-metrics-required");
});

void test("rejects truncated font/pen/rectangle payloads and unterminated font names", async () => {
  await fails([record(0x2fb, new Uint8Array(18))], 0, "invalid-record");
  await fails([record(0x2fa, words(0))], 0, "invalid-record");
  const unterminated = font(); unterminated.payload!.fill(65, 18);
  await fails([unterminated], 0, "invalid-font");
  await fails([font(), record(0x12d, words(0)), record(0xa32, words(0, 0, 1, 4))], 2, "invalid-record");
});

void test("cooperatively yields through state-only work and aborts at start and immediately after yielding", async () => {
  const parsed = parse(Array.from({ length: 600 }, () => record(0x214, words(0, 0))), 0);
  let yields = 0;
  const scene = await decodeWmfScene(parsed, { isCancelled: () => false, yieldControl: () => { yields += 1; return Promise.resolve(); } });
  assert.equal(scene.operations.length, 0); assert.ok(yields >= 4);
  await assert.rejects(decodeWmfScene(parsed, { isCancelled: () => true, yieldControl: options.yieldControl }), (e: unknown) => e instanceof MetafileDecodeError && e.code === "cancelled" && e.recordOffset === 18);
  let cancelled = false;
  await assert.rejects(decodeWmfScene(parsed, { isCancelled: () => cancelled, yieldControl: () => { cancelled = true; return Promise.resolve(); } }), (e: unknown) => e instanceof MetafileDecodeError && e.code === "cancelled");
});
