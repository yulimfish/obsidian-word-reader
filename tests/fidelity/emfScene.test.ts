import assert from "node:assert/strict";
import test from "node:test";
import { DEFAULT_METAFILE_LIMITS, MetafileParseError, parseMetafileRecords } from "../../src/ooxml/metafileRecords";
import type { MetafileLimits } from "../../src/ooxml/metafileRecords";
import { MetafileDecodeError } from "../../src/ooxml/metafileScene";
import { decodeEmfScene } from "../../src/ooxml/emfScene";
import { createSyntheticEmfFixture } from "./fixtureBuilders";
import type { SyntheticMetafileRecord } from "./fixtureBuilders";

const options = { isCancelled: () => false, yieldControl: () => Promise.resolve() };
const record = (type: number, payload: Uint8Array = new Uint8Array()): SyntheticMetafileRecord => ({ type, payload });
function dwords(...values: number[]): Uint8Array {
  const bytes = new Uint8Array(values.length * 4);
  const view = new DataView(bytes.buffer);
  values.forEach((value, index) => view.setUint32(index * 4, value >>> 0, true));
  return bytes;
}
function utf16(value: string): Uint8Array {
  const bytes = new Uint8Array(value.length * 2);
  const view = new DataView(bytes.buffer);
  for (let index = 0; index < value.length; index += 1) view.setUint16(index * 2, value.charCodeAt(index), true);
  return bytes;
}
function xform(values: number[], mode?: number): SyntheticMetafileRecord {
  const bytes = new Uint8Array(mode === undefined ? 24 : 28);
  const view = new DataView(bytes.buffer);
  values.forEach((value, index) => view.setFloat32(index * 4, value, true));
  if (mode !== undefined) view.setUint32(24, mode, true);
  return record(mode === undefined ? 35 : 36, bytes);
}
const pen = (handle = 1, red = 0, style = 0) => record(38, dwords(handle, style, 2, 0, red));
const brush = (handle = 2, style = 0) => record(39, dwords(handle, style, 0x332211, 0));
const select = (handle: number) => record(37, dwords(handle));
const line = () => record(54, dwords(30, 40));
function font(family = "宋体", charset = 134): SyntheticMetafileRecord {
  // 320-byte LOGFONTPANOSE, not a hand-filled decoder object.
  const bytes = new Uint8Array(324);
  bytes.set(dwords(1, -20, 0, 0, 0, 400));
  bytes[27] = charset;
  bytes.set(utf16(family), 32);
  return record(82, bytes);
}
interface Run { text: string; dx?: number[]; flags?: number; x?: number; y?: number }
function text(runs: Run[] = [{ text: "中文", dx: [7, 9] }], poly = false): SyntheticMetafileRecord {
  // Fixed EMRTEXT descriptors: reference/count/offString/options/rect/offDx.
  const fixed = (poly ? 32 : 28) + runs.length * 40;
  let size = fixed;
  for (const run of runs) size = Math.ceil((size + run.text.length * 2) / 4) * 4 + (run.dx?.length ?? 0) * 4;
  const bytes = new Uint8Array(size);
  const view = new DataView(bytes.buffer);
  view.setUint32(16, 2, true); // GM_ADVANCED, scales ignored.
  view.setFloat32(20, 1, true); view.setFloat32(24, 1, true);
  if (poly) view.setUint32(28, runs.length, true);
  let cursor = fixed;
  runs.forEach((run, index) => {
    const start = (poly ? 32 : 28) + index * 40;
    bytes.set(dwords(run.x ?? 30, run.y ?? 40, run.text.length, cursor + 8, run.flags ?? 0, 1, 2, 51, 62), start);
    bytes.set(utf16(run.text), cursor);
    cursor = Math.ceil((cursor + run.text.length * 2) / 4) * 4;
    if (run.dx) {
      view.setUint32(start + 36, cursor + 8, true);
      bytes.set(dwords(...run.dx), cursor); cursor += run.dx.length * 4;
    }
  });
  return record(poly ? 97 : 84, bytes);
}
function clip(mode = 5, rects = [[1, 2, 51, 62]]): SyntheticMetafileRecord {
  if (!rects.length) return record(75, dwords(0, mode));
  const bytes = new Uint8Array(40 + rects.length * 16);
  bytes.set(dwords(32 + rects.length * 16, mode, 32, 1, rects.length, rects.length * 16, 1, 2, 51, 62));
  rects.forEach((rect, index) => bytes.set(dwords(...rect), 40 + index * 16));
  return record(75, bytes);
}
function group(begin = true, description = "<script>"): SyntheticMetafileRecord {
  const body = begin ? new Uint8Array(28 + (description.length + 1) * 2) : dwords(0x43494447, 3);
  if (begin) { body.set(dwords(0x43494447, 2, 1, 2, 51, 62, description.length + 1)); body.set(utf16(description), 28); }
  const bytes = new Uint8Array(4 + body.length);
  bytes.set(dwords(body.length)); bytes.set(body, 4);
  return record(70, bytes);
}
function parse(records: SyntheticMetafileRecord[], capacity = 3, limits: MetafileLimits = DEFAULT_METAFILE_LIMITS) {
  return parseMetafileRecords(createSyntheticEmfFixture(records, { handleCount: capacity }), "emf", limits);
}
async function fails(records: SyntheticMetafileRecord[], index: number, code: string, capacity = 3, limits: MetafileLimits = DEFAULT_METAFILE_LIMITS) {
  const parsed = parse(records, capacity, limits);
  await assert.rejects(decodeEmfScene(parsed, options), (error: unknown) =>
    error instanceof MetafileDecodeError && error.code === code &&
    error.recordOffset === parsed.records[index + 1].offset && error.recordType === parsed.records[index + 1].type);
}

void test("decodes actual framed header/EOF without drawing and preserves nonzero bounds", async () => {
  const bytes = createSyntheticEmfFixture([]);
  const view = new DataView(bytes.buffer);
  view.setInt32(8, 20, true); view.setInt32(12, 10, true);
  view.setInt32(16, 120, true); view.setInt32(20, 90, true);
  const scene = await decodeEmfScene(parseMetafileRecords(bytes, "emf"), options);
  assert.deepEqual(scene.viewBox, { x: 20, y: 10, width: 101, height: 81 });
  assert.deepEqual(scene.operations, []);
});

void test("snapshots nested world transforms and restores position, pen and unsigned miter state 7/3/7", async () => {
  const scene = await decodeEmfScene(parse([
    pen(1, 17), select(1), record(27, dwords(5, 6)), record(58, dwords(7)),
    xform([2, 0, 0, 3, 10, 20]), record(33),
    xform([1, 0, 0, 1, 4, 5], 2), record(58, dwords(3)), line(),
    record(33), xform([1, 0, 0, 1, 4, 5], 3), line(), record(34, dwords(-1)), line(),
    record(34, dwords(1)), line(),
  ]), options);
  const paths = scene.operations.filter((op) => op.kind === "path");
  assert.equal(paths.length, 4);
  // GDI row-vector multiplication: LEFT applies supplied translation first.
  assert.deepEqual(paths[0].transform, { a: 2, b: 0, c: 0, d: 3, e: 18, f: 35 });
  assert.deepEqual(paths[1].transform, { a: 2, b: 0, c: 0, d: 3, e: 22, f: 40 });
  assert.deepEqual(paths[2].transform, paths[0].transform);
  assert.deepEqual(paths[3].transform, { a: 2, b: 0, c: 0, d: 3, e: 10, f: 20 });
  assert.deepEqual(paths.map((op) => op.stroke?.miterLimit), [3, 3, 3, 7]);
  assert.equal(paths[3].stroke?.color.red, 17);
  assert.deepEqual(paths[3].commands[0], { type: "moveTo", point: { x: 5, y: 6 } });
});

void test("supports MWT_IDENTITY and MWT_SET and rejects malformed/nonfinite transforms", async () => {
  const scene = await decodeEmfScene(parse([xform([2, 1, 4, 3, 5, 6], 4), line(), xform([0, 0, 0, 0, 0, 0], 1), line()]), options);
  assert.deepEqual(scene.operations[0].kind === "path" && scene.operations[0].transform, { a: 2, b: 1, c: 4, d: 3, e: 5, f: 6 });
  assert.deepEqual(scene.operations[1].kind === "path" && scene.operations[1].transform, { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 });
  await fails([xform([Infinity, 0, 0, 1, 0, 0])], 0, "invalid-transform");
  await fails([xform([1, 0, 0, 1, 0, 0], 9)], 0, "unsupported-transform-mode");
  await fails([record(35, dwords(1))], 0, "invalid-record");
  await fails([record(58, dwords(0))], 0, "invalid-miter-limit");
});

void test("preserves UTF16, signed font height, text style, exact code-unit advances and RTL rectangles", async () => {
  const scene = await decodeEmfScene(parse([
    font(), select(1), record(13, dwords(8, 9)), record(18, dwords(1)), record(19, dwords(2)),
    record(20, dwords(13)), record(21, dwords(4)), record(22, dwords(0x11e)),
    record(24, dwords(0x665544)), record(25, dwords(0x332211)),
    text([{ text: "中文😀", dx: [7, 9, 5, 0], flags: 0x86 }]),
  ]), options);
  assert.deepEqual(scene.requiredFonts, ["宋体"]);
  const op = scene.operations[0]; assert.equal(op.kind, "text");
  if (op.kind !== "text") return;
  assert.equal(op.text, "中文😀"); assert.equal(op.fontSize, 20); assert.equal(op.logicalFont?.height, -20);
  assert.deepEqual(op.sourceAdvances, [7, 9, 5, 0]); assert.deepEqual(op.advances, [7, 9, 5]);
  assert.equal(op.direction, "rtl"); assert.equal(op.textAlign, 0x11e);
  assert.equal(op.backgroundMode, "transparent");
  assert.deepEqual(op.color, { red: 68, green: 85, blue: 102, alpha: 1 });
  assert.deepEqual(op.backgroundColor, { red: 17, green: 34, blue: 51, alpha: 1 });
  assert.deepEqual(op.clipRect, { x: 1, y: 2, width: 50, height: 60 });
  assert.deepEqual(op.opaqueRect, op.clipRect); assert.equal(op.x, 30); assert.equal(op.y, 40);
});

void test("polytext emits each official type97 run in order and updates current position from known advances", async () => {
  const scene = await decodeEmfScene(parse([
    font(), select(1), record(22, dwords(1)), record(27, dwords(5, 6)),
    text([{ text: "AB", dx: [7, 8] }, { text: "中", dx: [9] }], true), line(),
  ]), options);
  assert.equal(scene.operations.length, 3);
  const [a, b, path] = scene.operations;
  assert.equal(a.kind === "text" && a.text, "AB"); assert.equal(b.kind === "text" && b.text, "中");
  assert.equal(a.kind === "text" && a.x, 5); assert.equal(b.kind === "text" && b.x, 20);
  assert.deepEqual(path.kind === "path" && path.commands[0], { type: "moveTo", point: { x: 29, y: 6 } });
  await fails([font(), select(1), record(22, dwords(1)), text([{ text: "A" }])], 3, "font-metrics-required");
});

void test("retains absent advances as explicit font metrics warning", async () => {
  const scene = await decodeEmfScene(parse([font(), select(1), text([{ text: "A" }])]), options);
  assert.ok(scene.warnings.some((warning) => warning.kind === "font-metrics-required" && warning.recordType === 84));
});

void test("clips preserve logical geometry and captured transform, offset and restore without transforming twice", async () => {
  const scene = await decodeEmfScene(parse([
    xform([2, 0, 0, 3, 10, 20]), clip(), record(33), record(26, dwords(4, 5)),
    clip(1), line(), record(34, dwords(-1)), line(), clip(5, []),
  ]), options);
  const clips = scene.operations.filter((op) => op.kind === "clip");
  assert.equal(clips.length, 5);
  assert.equal(clips[0].action, "region");
  if (clips[0].action !== "region") return;
  assert.deepEqual(clips[0].region, { kind: "rectangles", rectangles: [{ x: 1, y: 2, width: 50, height: 60 }], transform: { a: 2, b: 0, c: 0, d: 3, e: 10, f: 20 } });
  assert.deepEqual(clips[1].action === "region" && clips[1].region, { kind: "offset", region: clips[0].region, x: 8, y: 15 });
  assert.equal(clips[2].action === "region" && clips[2].region?.kind, "combine");
  assert.deepEqual(clips[3], clips[0]);
  assert.deepEqual(clips[4], { kind: "clip", action: "region", region: null });
});

void test("maps a normal clip offset through the current transform as an exact device vector", async () => {
  const scene = await decodeEmfScene(parse([
    clip(), xform([2, 0.5, -0.25, 3, 100, -200]), record(26, dwords(4, -2)),
    xform([1, 0, 0, 1, 0, 0]), line(),
  ]), options);
  const offset = scene.operations[1];
  assert.equal(offset.kind, "clip");
  assert.equal(offset.kind === "clip" && offset.action, "region");
  assert.deepEqual(offset.kind === "clip" && offset.action === "region" && offset.region?.kind === "offset" &&
    { x: offset.region.x, y: offset.region.y }, { x: 8.5, y: -4 });
});

void test("accepts an off-grid transient coefficient when its exact clip offset is on the device grid", async () => {
  const scene = await decodeEmfScene(parse([
    clip(), xform([2 ** -17, 0, 0, 1, 0, 0]), record(26, dwords(128, 0)),
    xform([1, 0, 0, 1, 0, 0]), line(),
  ]), options);
  const offset = scene.operations[1];
  assert.equal(offset.kind, "clip");
  assert.equal(offset.kind === "clip" && offset.action, "region");
  assert.deepEqual(offset.kind === "clip" && offset.action === "region" && offset.region?.kind === "offset" &&
    { x: offset.region.x, y: offset.region.y }, { x: 1 / 1024, y: 0 });
});

void test("preserves exact OFFSETCLIPRGN results at signed Int32 displacement limits", async () => {
  const scene = await decodeEmfScene(parse([clip(), record(26, dwords(0x7fffffff, -0x80000000))]), options);
  const offset = scene.operations[1];
  assert.equal(offset.kind, "clip");
  assert.deepEqual(offset.kind === "clip" && offset.action === "region" && offset.region?.kind === "offset" &&
    { x: offset.region.x, y: offset.region.y }, { x: 0x7fffffff, y: -0x80000000 });
});

void test("fails closed when large OFFSETCLIPRGN products cancel below Number precision", async () => {
  const a = 2 ** 30;
  const records = [
    clip(),
    xform([a, 2 ** -17, -(a + 128), -(2 ** -17), 0, 0]),
    record(26, dwords(a + 129, a + 1)),
    xform([1, 0, 0, 1, 0, 0]),
    line(),
  ];
  // Exact device displacement is (-128, 1/1024); binary64 products around 2^60
  // lose the x residue unless the decoder proves the mapped value exact.
  await fails(records, 2, "unsupported-clip-offset");
});

void test("fails closed when an OFFSETCLIPRGN result is off the downstream exact grid", async () => {
  await fails([
    clip(), xform([2 ** -17, 0, 0, 1, 0, 0]), record(26, dwords(1, 0)),
    xform([1, 0, 0, 1, 0, 0]), line(),
  ], 2, "unsupported-clip-offset");
});

void test("clip validates region counts, bounds, rectangle ordering, modes and bounded expression depth", async () => {
  const damaged = clip(); new DataView(damaged.payload!.buffer).setUint32(16, 0xffffffff, true);
  await fails([damaged], 0, "invalid-region");
  await fails([clip(0)], 0, "unsupported-clip-mode");
  await fails([clip(1, [])], 0, "invalid-region");
  await fails([clip(5, [[60, 2, 51, 62]])], 0, "invalid-region");
  await fails([clip(), ...Array.from({ length: 129 }, () => record(26, dwords(1, 1)))], 128, "max-clip-depth");
});

void test("handles reserve slot0 and permit explicit freed handle reuse across pen/brush/font", async () => {
  await fails([pen()], 0, "invalid-create", 1, { ...DEFAULT_METAFILE_LIMITS, maxObjects: 0 });
  await fails([pen(0)], 0, "invalid-create");
  await fails([pen(), pen()], 1, "invalid-create", 2);
  const scene = await decodeEmfScene(parse([
    pen(), record(40, dwords(1)), brush(1), select(1), select(0x80000000), record(40, dwords(1)),
    font(), select(1), text(), select(0x8000000d), record(40, dwords(1)), pen(1, 119), select(1), line(),
  ], 2, { ...DEFAULT_METAFILE_LIMITS, maxObjects: 1 }), options);
  assert.equal(scene.operations.length, 2);
  assert.equal(scene.operations[1].kind === "path" && scene.operations[1].stroke?.color.red, 119);
  const empty = await decodeEmfScene(parse([], 1, { ...DEFAULT_METAFILE_LIMITS, maxObjects: 0 }), options);
  assert.equal(empty.operations.length, 0);
});

void test("rejects uncreated/deleted selections and deleting objects held by current or saved DC", async () => {
  await fails([select(1)], 0, "invalid-selection");
  await fails([record(40, dwords(1))], 0, "invalid-delete");
  await fails([pen(), record(40, dwords(1)), select(1)], 2, "invalid-selection");
  await fails([pen(), select(1), record(40, dwords(1))], 2, "invalid-delete");
  await fails([pen(), select(1), record(33), select(0x80000007), record(40, dwords(1))], 4, "invalid-delete");
  await fails([brush(), select(2), record(40, dwords(2))], 2, "invalid-delete");
  await fails([record(40, dwords(0x80000007))], 0, "invalid-delete");
});

void test("parser retains numeric handle and header cap errors rather than decoder relabeling", () => {
  assert.throws(() => parse([select(0)]), (e: unknown) => e instanceof MetafileParseError && e.code === "invalid-handle");
  assert.throws(() => parse([select(3)]), (e: unknown) => e instanceof MetafileParseError && e.code === "invalid-handle");
  assert.throws(() => parse([], 2, { ...DEFAULT_METAFILE_LIMITS, maxObjects: 0 }), (e: unknown) => e instanceof MetafileParseError && e.code === "too-many-objects");
});

void test("stack limits zero/one and invalid positive/negative restore targets report exact record", async () => {
  await fails([record(33)], 0, "max-stack", 1, { ...DEFAULT_METAFILE_LIMITS, maxStack: 0 });
  await fails([record(33), record(33)], 1, "max-stack", 1, { ...DEFAULT_METAFILE_LIMITS, maxStack: 1 });
  for (const target of [0, -1, 1]) await fails([record(34, dwords(target))], 0, "invalid-restore");
  await fails([record(33), record(34, dwords(2))], 1, "invalid-restore");
  const scene = await decodeEmfScene(parse([record(33), record(34, dwords(-1))], 1, { ...DEFAULT_METAFILE_LIMITS, maxStack: 1, maxSceneOps: 0 }), options);
  assert.equal(scene.operations.length, 0);
});

void test("scene caps count each line, clip, text/polytext run and restored clip before emission", async () => {
  const zero = { ...DEFAULT_METAFILE_LIMITS, maxSceneOps: 0 };
  const one = { ...DEFAULT_METAFILE_LIMITS, maxSceneOps: 1 };
  await fails([line()], 0, "max-scene-ops", 1, zero);
  await fails([line(), line()], 1, "max-scene-ops", 1, one);
  await fails([clip()], 0, "max-scene-ops", 1, zero);
  await fails([font(), select(1), text()], 2, "max-scene-ops", 2, zero);
  await fails([font(), select(1), text([{ text: "A", dx: [2] }, { text: "B", dx: [3] }], true)], 2, "max-scene-ops", 2, one);
  await fails([record(33), clip(), record(34, dwords(-1))], 2, "max-scene-ops", 1, one);
  const scene = await decodeEmfScene(parse([record(13, dwords(0, 0)), group(), group(false), line()], 1, one), options);
  assert.equal(scene.operations.length, 1);
});

void test("uses frozen parser limit snapshot even when caller weakens it", async () => {
  const limits = { ...DEFAULT_METAFILE_LIMITS, maxStack: 0, maxSceneOps: 0 };
  const parsed = parse([record(33)], 1, limits);
  limits.maxStack = 99; limits.maxSceneOps = 99;
  await assert.rejects(decodeEmfScene(parsed, options), (e: unknown) => e instanceof MetafileDecodeError && e.code === "max-stack");
});

void test("GDIC nested groups are validated inert metadata, not drawing or executable descriptions", async () => {
  const scene = await decodeEmfScene(parse([group(), group(true, "中文"), group(false), group(false)]), options);
  assert.equal(scene.operations.length, 0);
  assert.equal(scene.warnings.filter((warning) => warning.kind === "group-metadata").length, 4);
  const opaque = record(70, dwords(8, 0x11223344, 2));
  const privateScene = await decodeEmfScene(parse([opaque]), options);
  assert.equal(privateScene.operations.length, 0); assert.equal(privateScene.warnings[0].kind, "opaque-comment");
});

void test("forged groups cannot escape data size, description, rectangle, nesting or matching bounds", async () => {
  const badLength = group(); new DataView(badLength.payload!.buffer).setUint32(0, 0xffffffff, true);
  await fails([badLength], 0, "invalid-comment");
  const badDescription = group(); new DataView(badDescription.payload!.buffer).setUint32(28, 0xffffffff, true);
  await fails([badDescription], 0, "invalid-comment");
  const inverted = group(); new DataView(inverted.payload!.buffer).setInt32(12, 99, true);
  await fails([inverted], 0, "invalid-comment");
  await fails([group(false)], 0, "invalid-group");
  await fails([group()], 0, "invalid-group");
  await fails([group(), group()], 1, "max-stack", 1, { ...DEFAULT_METAFILE_LIMITS, maxStack: 1 });
  await fails([record(70, dwords(4, 0x43494447))], 0, "invalid-comment");
});

void test("rejects glyph-index and Symbol/MT Extra mapping gaps without pretending Unicode glyph parity", async () => {
  await fails([font(), select(1), text([{ text: "A", dx: [2], flags: 0x10 }])], 2, "font-glyph-gap");
  for (const family of ["Symbol", "MT Extra"]) await fails([font(family, 2), select(1), text()], 2, "font-glyph-gap");
});

void test("rejects unpaired UTF16, unaligned/out-of-record/descriptor-overlap offsets and short dx", async () => {
  await fails([font(), select(1), text([{ text: "\ud800", dx: [2] }])], 2, "invalid-encoding");
  for (const offset of [77, 0xfffffff0, 8]) {
    const bad = text(); new DataView(bad.payload!.buffer).setUint32(40, offset, true);
    await fails([font(), select(1), bad], 2, "invalid-text-offset");
  }
  const badDx = text(); new DataView(badDx.payload!.buffer).setUint32(64, 0xfffffff0, true);
  await fails([font(), select(1), badDx], 2, "invalid-advances");
  const overlap = text(); new DataView(overlap.payload!.buffer).setUint32(64, 76, true);
  await fails([font(), select(1), overlap], 2, "invalid-advances");
  const oversized = text(); new DataView(oversized.payload!.buffer).setUint32(36, 8193, true);
  await fails([font(), select(1), oversized], 2, "invalid-text-length");
  await fails([text([], true)], 0, "invalid-text-length");
});

void test("unknown color/mode/glyph opcodes are named failures, especially color-match121", async () => {
  await fails([line(), record(121, dwords(1, 0, 0, 0))], 1, "unsupported-color-management");
  await fails([record(23, new Uint8Array(24))], 0, "unsupported-color-management");
  await fails([record(24, dwords(0x010000ff))], 0, "unsupported-color");
  await fails([record(18, dwords(3))], 0, "unsupported-background-mode");
  await fails([record(19, dwords(3))], 0, "unsupported-fill-mode");
  await fails([record(20, dwords(7))], 0, "unsupported-rop2");
  await fails([record(21, dwords(9))], 0, "unsupported-stretch-mode");
  await fails([record(22, dwords(4))], 0, "unsupported-text-alignment");
  for (const type of [28, 98, 115, 0x7fff]) await fails([record(type)], 0, "unsupported-record");
  await fails([brush(1, 2)], 0, "unsupported-brush");
  await fails([pen(1, 0, 1)], 0, "unsupported-pen");
  await fails([font(), select(1), text([{ text: "A", dx: [1], flags: 0x400 }])], 2, "unsupported-text-options");
});

void test("cooperative cancellation at start and after yield names exact EMF record", async () => {
  const parsed = parse(Array.from({ length: 600 }, () => record(27, dwords(0, 0))), 1);
  let yields = 0;
  const scene = await decodeEmfScene(parsed, { isCancelled: () => false, yieldControl: () => { yields += 1; return Promise.resolve(); } });
  assert.equal(scene.operations.length, 0); assert.ok(yields >= 4);
  await assert.rejects(decodeEmfScene(parsed, { isCancelled: () => true, yieldControl: options.yieldControl }), (e: unknown) => e instanceof MetafileDecodeError && e.code === "cancelled" && e.recordOffset === 0 && e.recordType === 1);
  let cancelled = false;
  await assert.rejects(decodeEmfScene(parsed, { isCancelled: () => cancelled, yieldControl: () => { cancelled = true; return Promise.resolve(); } }), (e: unknown) => e instanceof MetafileDecodeError && e.code === "cancelled" && e.recordOffset === parsed.records[128].offset && e.recordType === 27);
});

void test("polytext cannot amplify reused string offsets beyond bounded total text units", async () => {
  const payload = text([{ text: "A".repeat(4097) }, { text: "B".repeat(4097) }], true);
  const view = new DataView(payload.payload!.buffer);
  view.setUint32(84, view.getUint32(44, true), true);
  await fails([font(), select(1), payload], 2, "invalid-text-length");
});

void test("region union/xor/exclude preserve exact Boolean mode rather than substituting intersection", async () => {
  for (const [mode, want] of [[2, "union"], [3, "xor"], [4, "exclude"]] as const) {
    const scene = await decodeEmfScene(parse([clip(), clip(mode)]), options);
    const op = scene.operations[1]; assert.equal(op.kind, "clip");
    assert.equal(op.kind === "clip" && op.action === "region" && op.region?.kind === "combine" && op.region.mode, want);
  }
});

void test("font descriptors preserve Panose/style metadata and reject unterminated/truncated faces", async () => {
  const object = font("Arial", 0);
  object.payload!.set(utf16("Arial Regular"), 96); object.payload!.set(utf16("Regular"), 224);
  object.payload![312] = 2;
  const scene = await decodeEmfScene(parse([object, select(1), text()]), options);
  const op = scene.operations[0]; assert.equal(op.kind, "text");
  if (op.kind !== "text") return;
  assert.deepEqual(op.fontMetadata, { fullName: "Arial Regular", style: "Regular", styleSize: 0, panose: [2, 0, 0, 0, 0, 0, 0, 0, 0, 0] });
  assert.ok(scene.warnings.some((warning) => warning.kind === "font-metrics-required"));
  const broken = font(); broken.payload!.fill(65, 32, 96);
  await fails([broken], 0, "invalid-font");
  await fails([record(82, dwords(1))], 0, "invalid-record");
});
