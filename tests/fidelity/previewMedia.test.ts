import assert from "node:assert/strict";
import test from "node:test";
import { childElements, parseXml } from "../../src/pptx/xml";
import {
  PreviewMediaError,
  createPreviewMediaResolver,
  resolvePreviewMedia,
  checkPreviewFonts,
} from "../../src/ooxml/previewMedia";
import { MetafileSerializeError } from "../../src/ooxml/metafileSvg";
import { createSyntheticEmfFixture, createSyntheticWmfFixture } from "./fixtureBuilders";
import type { SyntheticMetafileRecord } from "./fixtureBuilders";

const options = { isCancelled: () => false, yieldControl: () => Promise.resolve() };

function words(...values: number[]): Uint8Array {
  const bytes = new Uint8Array(values.length * 2);
  const view = new DataView(bytes.buffer);
  values.forEach((value, index) => view.setInt16(index * 2, value, true));
  return bytes;
}

function wmfRecord(type: number, payload: Uint8Array = new Uint8Array()): SyntheticMetafileRecord {
  return { type, payload };
}

function wmfFont(options: { readonly escapement?: number; readonly orientation?: number } = {}): SyntheticMetafileRecord {
  const payload = new Uint8Array(50);
  const view = new DataView(payload.buffer);
  view.setInt16(0, -20, true);
  view.setInt16(4, options.escapement ?? 0, true);
  view.setInt16(6, options.orientation ?? 0, true);
  view.setInt16(8, 400, true);
  payload.set(new TextEncoder().encode("Times New Roman"), 18);
  return wmfRecord(0x2fb, payload);
}

function wmfText(
  value: string,
  advances?: readonly number[],
  flags = 0,
  x = 30,
  y = 40,
): SyntheticMetafileRecord {
  const text = new TextEncoder().encode(value);
  const start = flags & 6 ? 16 : 8;
  const end = start + text.byteLength + text.byteLength % 2;
  const payload = new Uint8Array(end + (advances?.length ?? 0) * 2);
  const view = new DataView(payload.buffer);
  view.setInt16(0, y, true);
  view.setInt16(2, x, true);
  view.setInt16(4, text.byteLength, true);
  view.setUint16(6, flags, true);
  if (flags & 6) {
    view.setInt16(8, 1, true); view.setInt16(10, 2, true);
    view.setInt16(12, 51, true); view.setInt16(14, 62, true);
  }
  payload.set(text, start);
  if (advances) payload.set(words(...advances), end);
  return wmfRecord(0xa32, payload);
}

function wmfTextDocument(bytes: Uint8Array, path: string): Document {
  const svg = new TextDecoder().decode(bytes);
  return parseXml(svg, path);
}

function png(): Uint8Array {
  return Uint8Array.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 1, 2]);
}

function jpeg(): Uint8Array {
  return Uint8Array.from([0xff, 0xd8, 0xff, 0xe0, 0, 4, 0xff, 0xd9]);
}

function bmp(): Uint8Array {
  return Uint8Array.from([0x42, 0x4d, 14, 0, 0, 0, 0, 0, 0, 0, 6, 0, 0, 0]);
}

function dwords(...values: number[]): Uint8Array {
  const bytes = new Uint8Array(values.length * 4);
  const view = new DataView(bytes.buffer);
  values.forEach((value, index) => view.setUint32(index * 4, value >>> 0, true));
  return bytes;
}

function regionRecord(): SyntheticMetafileRecord {
  const bytes = new Uint8Array(56);
  bytes.set(dwords(48, 5, 32, 1, 1, 16, 0, 0, 8, 8, 0, 0, 8, 8));
  return { type: 75, payload: bytes };
}

function xform(values: readonly number[]): SyntheticMetafileRecord {
  const payload = new Uint8Array(24);
  const view = new DataView(payload.buffer);
  values.forEach((value, index) => view.setFloat32(index * 4, value, true));
  return { type: 35, payload };
}

function line(): SyntheticMetafileRecord {
  return { type: 54, payload: dwords(4, 4) };
}

function emfRecord(type: number, payload: Uint8Array = new Uint8Array()): SyntheticMetafileRecord {
  return { type, payload };
}

function emfFont(): SyntheticMetafileRecord {
  const payload = new Uint8Array(324);
  const view = new DataView(payload.buffer);
  view.setUint32(0, 1, true);
  view.setInt32(4, -20, true);
  view.setInt32(20, 400, true);
  const family = "Times New Roman";
  for (let index = 0; index < family.length; index += 1) view.setUint16(32 + index * 2, family.charCodeAt(index), true);
  return emfRecord(82, payload);
}

function emfText(value: string, flags = 0): SyntheticMetafileRecord {
  const text = new Uint8Array(value.length * 2);
  const textView = new DataView(text.buffer);
  for (let index = 0; index < value.length; index += 1) textView.setUint16(index * 2, value.charCodeAt(index), true);
  const textOffset = 68;
  const advancesOffset = Math.ceil((textOffset + text.byteLength) / 4) * 4;
  const payload = new Uint8Array(advancesOffset + value.length * 4);
  const view = new DataView(payload.buffer);
  view.setUint32(16, 2, true);
  view.setFloat32(20, 1, true);
  view.setFloat32(24, 1, true);
  payload.set(dwords(30, 40, value.length, textOffset + 8, flags, 1, 2, 51, 62), 28);
  payload.set(text, textOffset);
  payload.set(dwords(...Array.from(value, (_, index) => index === 0 ? 3 : 7)), advancesOffset);
  view.setUint32(64, advancesOffset + 8, true);
  return emfRecord(84, payload);
}

function assertUnsupportedText(error: unknown, detail: RegExp): boolean {
  return error instanceof MetafileSerializeError && error.code === "unsupported-text" &&
    error.operationIndex === 0 && detail.test(error.message);
}

async function assertFailurePreservesCachedMedia(
  path: string,
  format: "wmf" | "emf",
  failureBytes: Uint8Array,
  detail: RegExp,
): Promise<void> {
  const resolver = createPreviewMediaResolver();
  let yields = 0;
  const countedOptions = {
    isCancelled: () => false,
    yieldControl: async () => { yields += 1; },
  };
  const validBytes = format === "wmf" ? createSyntheticWmfFixture([]) : createSyntheticEmfFixture([]);
  const uncachedPath = `${path.slice(0, path.lastIndexOf("."))}.uncached.${format}`;
  try {
    await assert.rejects(resolver.resolve(uncachedPath, failureBytes, countedOptions), (error: unknown) =>
      assertUnsupportedText(error, detail));
    assert.deepEqual(resolver.getCacheDiagnostics(), { entries: 0, bytes: 0 }, `${format}: failed initial resolution must not publish media`);

    const original = await resolver.resolve(path, validBytes, countedOptions);
    const before = resolver.getCacheDiagnostics();
    await assert.rejects(resolver.resolve(path, failureBytes, countedOptions), (error: unknown) =>
      assertUnsupportedText(error, detail));
    assert.deepEqual(resolver.getCacheDiagnostics(), before, `${format}: failed replacement must not alter cached entry count or bytes`);
    const yieldsBeforeCacheHit = yields;
    const cached = await resolver.resolve(path, validBytes, countedOptions);
    assert.deepEqual(cached.bytes, original.bytes, `${format}: original cached SVG remains available`);
    assert.equal(yields, yieldsBeforeCacheHit, `${format}: original source still resolves from cache`);
    assert.deepEqual(resolver.getCacheDiagnostics(), before);
  } finally {
    resolver.dispose();
  }
}

function offGridEmf(): Uint8Array {
  return createSyntheticEmfFixture([xform([0.1, 0, 0, 1, 0, 0]), regionRecord(), line()]);
}

void test("preserves raster bytes exactly and checks extension against magic", async () => {
  for (const [path, source, mimeType] of [
    ["media/a.png", png(), "image/png"],
    ["media/a.jpg", jpeg(), "image/jpeg"],
    ["media/a.bmp", bmp(), "image/bmp"],
  ] as const) {
    const original = source.slice();
    const resolved = await resolvePreviewMedia(path, source, options);
    assert.equal(resolved.mimeType, mimeType);
    assert.deepEqual(resolved.bytes, original);
    assert.deepEqual(source, original);
    assert.deepEqual(resolved.requiredFonts, []);
    assert.deepEqual(resolved.issues, []);
  }
  await assert.rejects(resolvePreviewMedia("media/wrong.jpg", png(), options), (error: unknown) =>
    error instanceof PreviewMediaError && error.code === "invalid-media");
  await assert.rejects(resolvePreviewMedia("media/wrong.png", jpeg(), options), (error: unknown) =>
    error instanceof PreviewMediaError && error.code === "invalid-media");
  await assert.rejects(resolvePreviewMedia("media/wrong.bmp", png(), options), (error: unknown) =>
    error instanceof PreviewMediaError && error.code === "invalid-media");
});

void test("decodes framed local WMF/EMF into static SVG without changing originals", async () => {
  const wmf = createSyntheticWmfFixture([]);
  const emf = createSyntheticEmfFixture([]);
  for (const [path, source] of [["media/preview.wmf", wmf], ["media/preview.emf", emf]] as const) {
    const original = source.slice();
    const result = await resolvePreviewMedia(path, source, options);
    assert.equal(result.mimeType, "image/svg+xml");
    const svg = new TextDecoder().decode(result.bytes);
    const document = parseXml(svg, path);
    assert.equal(document.documentElement.localName, "svg");
    assert.deepEqual(source, original);
    assert.deepEqual(result.issues, []);
  }
});

void test("does not admit arbitrary SVG, unsupported extensions or malformed vector data", async () => {
  const rawSvg = new TextEncoder().encode('<svg xmlns="http://www.w3.org/2000/svg"><script/></svg>');
  await assert.rejects(resolvePreviewMedia("media/user.svg", rawSvg, options), (error: unknown) =>
    error instanceof PreviewMediaError && error.code === "unsupported-media");
  await assert.rejects(resolvePreviewMedia("media/file.gif", png(), options), (error: unknown) =>
    error instanceof PreviewMediaError && error.code === "unsupported-media");
  await assert.rejects(resolvePreviewMedia("media/corrupt.wmf", Uint8Array.from([1, 2, 3]), options), (error: unknown) =>
    error instanceof Error);
});

void test("propagates typed serializer failure before package cache publication", async () => {
  const resolver = createPreviewMediaResolver();
  await assert.rejects(resolver.resolve("media/off-grid.emf", offGridEmf(), options), (error: unknown) =>
    error instanceof MetafileSerializeError && error.code === "unsupported-clip-expression");
  assert.deepEqual(resolver.getCacheDiagnostics(), { entries: 0, bytes: 0 });
  resolver.dispose();
});

void test("rejects parsed WMF/EMF top, bottom, right and center alignments instead of assuming a baseline", async () => {
  const alignments = [
    { name: "top", value: 0 },
    { name: "bottom", value: 8 },
    { name: "right", value: 2 },
    { name: "center", value: 6 },
  ] as const;
  for (const entry of alignments) {
    const wmf = createSyntheticWmfFixture([
      wmfFont(), wmfRecord(0x12d, words(0)), wmfRecord(0x12e, words(entry.value)), wmfText("AB", [3, 7]),
    ], { objectCount: 1 });
    const emf = createSyntheticEmfFixture([
      emfFont(), emfRecord(37, dwords(1)), emfRecord(22, dwords(entry.value)), emfText("AB"),
    ], { handleCount: 2 });
    await assertFailurePreservesCachedMedia(`media/${entry.name}.wmf`, "wmf", wmf, /alignment/i);
    await assertFailurePreservesCachedMedia(`media/${entry.name}.emf`, "emf", emf, /alignment/i);
  }
});

void test("rejects parsed opaque-background text when no exact ETO_OPAQUE rectangle exists", async () => {
  const wmf = createSyntheticWmfFixture([
    wmfFont(), wmfRecord(0x12d, words(0)), wmfRecord(0x12e, words(0x18)),
    wmfRecord(0x102, words(2)), wmfRecord(0x201, Uint8Array.from([17, 34, 51, 0])), wmfText("AB", [3, 7]),
  ], { objectCount: 1 });
  const emf = createSyntheticEmfFixture([
    emfFont(), emfRecord(37, dwords(1)), emfRecord(22, dwords(0x18)), emfRecord(18, dwords(2)),
    emfRecord(25, dwords(0x332211)), emfText("AB"),
  ], { handleCount: 2 });
  await assertFailurePreservesCachedMedia("media/opaque-no-rect.wmf", "wmf", wmf, /opaque|background/i);
  await assertFailurePreservesCachedMedia("media/opaque-no-rect.emf", "emf", emf, /opaque|background/i);
});

void test("cancels before parsing and after cooperative yield without publishing media", async () => {
  const resolver = createPreviewMediaResolver();
  await assert.rejects(resolver.resolve("media/a.png", png(), { isCancelled: () => true, yieldControl: async () => undefined }), (error: unknown) =>
    error instanceof PreviewMediaError && error.code === "cancelled");
  assert.deepEqual(resolver.getCacheDiagnostics(), { entries: 0, bytes: 0 });
  let yields = 0;
  let cancelled = false;
  const records = Array.from({ length: 130 }, () => ({ type: 27, payload: dwords(1, 2) }));
  const emf = createSyntheticEmfFixture(records);
  await assert.rejects(resolver.resolve("media/yield.emf", emf, {
    isCancelled: () => cancelled,
    yieldControl: async () => { yields += 1; cancelled = true; },
  }), (error: unknown) => error instanceof PreviewMediaError && error.code === "cancelled");
  assert.ok(yields >= 1);
  assert.deepEqual(resolver.getCacheDiagnostics(), { entries: 0, bytes: 0 });
});

void test("bounds cache entries and bytes per package, copies owned results, and releases on disposal", async () => {
  const resolver = createPreviewMediaResolver();
  const source = png();
  const first = await resolver.resolve("media/shared.png", source, options);
  const second = await resolver.resolve("media/shared.png", source, options);
  assert.deepEqual(second.bytes, first.bytes);
  assert.notEqual(second.bytes, first.bytes, "callers receive separately owned byte arrays");
  first.bytes[0] = 0;
  const third = await resolver.resolve("media/shared.png", source, options);
  assert.deepEqual(third.bytes, png(), "mutating one caller's bytes does not mutate the cached result");
  for (let index = 0; index < 30; index += 1) {
    await resolver.resolve(`media/${index}.png`, png(), options);
  }
  assert.equal(resolver.getCacheDiagnostics().entries, 24);
  const large = png();
  const budgetResolver = createPreviewMediaResolver();
  const bigA = new Uint8Array(12 * 1024 * 1024);
  bigA.set(png());
  const bigB = new Uint8Array(4 * 1024 * 1024);
  bigB.set(png());
  await budgetResolver.resolve("a".repeat(100) + ".png", bigA, options);
  const bigBPath = "b".repeat(100) + ".png";
  await budgetResolver.resolve(bigBPath, bigB, options);
  assert.deepEqual(budgetResolver.getCacheDiagnostics(), {
    entries: 1,
    bytes: 2 * bigB.byteLength + new TextEncoder().encode(bigBPath).byteLength,
  });
  budgetResolver.dispose();
  assert.deepEqual(budgetResolver.getCacheDiagnostics(), { entries: 0, bytes: 0 });
  assert.ok(large.byteLength > 0);
  resolver.dispose();
  assert.deepEqual(resolver.getCacheDiagnostics(), { entries: 0, bytes: 0 });
});

void test("uses least-recently-used order when the per-package entry cap is reached", async () => {
  const resolver = createPreviewMediaResolver();
  let yields = 0;
  const countedOptions = {
    isCancelled: () => false,
    yieldControl: async () => { yields += 1; },
  };
  const bytes = png();
  for (let index = 0; index < 24; index += 1) {
    await resolver.resolve(`media/lru-${index}.png`, bytes, countedOptions);
  }
  assert.equal(yields, 24);
  await resolver.resolve("media/lru-0.png", bytes, countedOptions);
  assert.equal(yields, 24, "a cache hit does not re-resolve the media");
  await resolver.resolve("media/lru-24.png", bytes, countedOptions);
  assert.equal(yields, 25);
  await resolver.resolve("media/lru-1.png", bytes, countedOptions);
  assert.equal(yields, 26, "the untouched least-recently-used key was evicted");
  assert.equal(resolver.getCacheDiagnostics().entries, 24);
  resolver.dispose();
});

void test("keeps package cache accounting bounded across concurrent same-key resolutions", async () => {
  const resolver = createPreviewMediaResolver();
  const path = "media/concurrent.png";
  const source = png();
  const [first, second] = await Promise.all([
    resolver.resolve(path, source, options),
    resolver.resolve(path, source, options),
  ]);
  assert.deepEqual(first.bytes, second.bytes);
  assert.notEqual(first.bytes, second.bytes);
  assert.deepEqual(resolver.getCacheDiagnostics(), {
    entries: 1,
    bytes: new TextEncoder().encode(path).byteLength + source.byteLength + first.bytes.byteLength,
  });
  resolver.dispose();
});

void test("does not repopulate a disposed package cache from an in-flight resolution", async () => {
  const resolver = createPreviewMediaResolver();
  let resume: (() => void) | undefined;
  const blocked = new Promise<void>((resolve) => { resume = resolve; });
  const pending = resolver.resolve("media/pending.png", png(), {
    isCancelled: () => false,
    yieldControl: () => blocked,
  });
  resolver.dispose();
  assert.ok(resume);
  resume();
  await assert.rejects(pending, (error: unknown) =>
    error instanceof PreviewMediaError && error.code === "disposed");
  assert.deepEqual(resolver.getCacheDiagnostics(), { entries: 0, bytes: 0 });
});

void test("reports missing or unconfirmed supplied fonts without claiming visual parity", async () => {
  const calls: string[] = [];
  const ownerDocument = {
    fonts: {
      check: (font: string) => {
        calls.push(font);
        return !font.includes("Missing");
      },
    },
  } as unknown as Document;
  assert.deepEqual(await checkPreviewFonts(ownerDocument, ["Times New Roman", "Missing Font", "Times New Roman"]), ["Missing Font"]);
  assert.equal(calls.length, 2);
  assert.ok(calls.every((font) => font.includes("12px")));
});

void test("returns inert decoder warnings as occurrence-scoped preview issues", async () => {
  const payload = new Uint8Array(12);
  const view = new DataView(payload.buffer);
  view.setUint16(0, 15, true);
  view.setUint16(2, 8, true);
  payload.set(new TextEncoder().encode("<script>"), 4);
  const path = "media/comment.wmf";
  const result = await resolvePreviewMedia(path, createSyntheticWmfFixture([{ type: 0x626, payload }]), options);
  assert.equal(result.mimeType, "image/svg+xml");
  assert.equal(result.issues.length, 1);
  const issue = result.issues[0];
  assert.ok(issue);
  assert.equal(issue.partPath, path);
  assert.equal(issue.occurrenceId, path);
  assert.equal(issue.code, "opaque-comment");
  assert.match(issue.detail, /byteLength=8/);
  assert.doesNotMatch(issue.detail, /<script>/);
});

void test("surfaces explicit legacy glyph gaps and never draws font-specific bytes as guessed Unicode", async () => {
  const bytes = new Uint8Array(50);
  bytes[0] = 0xec; bytes[1] = 0xff; // height = -20
  bytes[13] = 2; // Symbol charset
  bytes.set(new TextEncoder().encode("MT Extra"), 18);
  const font: SyntheticMetafileRecord = { type: 0x2fb, payload: bytes };
  const select: SyntheticMetafileRecord = { type: 0x12d, payload: Uint8Array.from([0, 0]) };
  const textPayload = Uint8Array.from([40, 0, 30, 0, 1, 0, 0, 0, 0x61, 0, 0, 0]);
  const text: SyntheticMetafileRecord = { type: 0xa32, payload: textPayload };
  const wmf = createSyntheticWmfFixture([
    font, select, wmfRecord(0x12e, words(0x18)), wmfRecord(0x102, words(1)), text,
  ], { objectCount: 1 });
  await assert.rejects(resolvePreviewMedia("media/glyph.wmf", wmf, options), (error: unknown) =>
    error instanceof MetafileSerializeError && error.operationIndex === 0 &&
    error.code === "font-glyph-gap");
});

void test("serializes parsed WMF current-to-next advances at per-glyph LTR and RTL positions", async () => {
  const cases = [
    { name: "ltr", alignment: 0x18, flags: 0, direction: "ltr", advances: [3, 7], x: "30 33" },
    { name: "rtl-option", alignment: 0x18, flags: 0x80, direction: "rtl", advances: [-3, -7], x: "30 27" },
    { name: "rtl-alignment", alignment: 0x118, flags: 0, direction: "rtl", advances: [-3, -7], x: "30 27" },
  ] as const;
  for (const entry of cases) {
    const records = [
      wmfFont(), wmfRecord(0x12d, words(0)),
      ...(entry.alignment ? [wmfRecord(0x12e, words(entry.alignment))] : []),
      wmfRecord(0x102, words(1)), wmfText("AB", entry.advances, entry.flags),
    ];
    const source = createSyntheticWmfFixture(records, { objectCount: 1 });
    const resolved = await resolvePreviewMedia(`media/${entry.name}.wmf`, source, options);
    const document = wmfTextDocument(resolved.bytes, `${entry.name}.svg`);
    const text = document.getElementsByTagName("text")[0];
    assert.ok(text);
    assert.equal(text.getAttribute("direction"), entry.direction);
    assert.equal(text.getAttribute("x"), entry.x, `${entry.name}: x positions follow source advances, not dx glyph shifts`);
    assert.equal(text.hasAttribute("dx"), false);
  }
});

void test("retains isolated internal spaces through real WMF/EMF chains", async () => {
  const sources = [
    ["wmf", createSyntheticWmfFixture([
      wmfFont(), wmfRecord(0x12d, words(0)), wmfRecord(0x12e, words(0x18)),
      wmfRecord(0x102, words(1)), wmfText("A B", [3, 7, 5]),
    ], { objectCount: 1 }), "30 33 40"],
    ["emf", createSyntheticEmfFixture([
      emfFont(), emfRecord(37, dwords(1)), emfRecord(22, dwords(0x18)),
      emfRecord(18, dwords(1)), emfText("A B"),
    ], { handleCount: 2 }), "30 33 40"],
  ] as const;
  for (const [format, source, positions] of sources) {
    const original = source.slice();
    const result = await resolvePreviewMedia(`media/internal-space.${format}`, source, options);
    const text = wmfTextDocument(result.bytes, "internal-space.svg").getElementsByTagName("text")[0];
    assert.ok(text);
    assert.equal(text.textContent, "A B");
    assert.equal(text.getAttribute("x"), positions);
    assert.deepEqual(source, original);
  }
});

void test("rejects SVG-collapsible whitespace through real WMF/EMF parse-decode-serialize chains", async () => {
  const cases: readonly { readonly name: string; readonly text: string; readonly formats: readonly ("wmf" | "emf")[] }[] = [
    { name: "leading-space", text: " A", formats: ["wmf", "emf"] },
    { name: "trailing-space", text: "A ", formats: ["wmf", "emf"] },
    { name: "only-space", text: " ", formats: ["wmf", "emf"] },
    { name: "review-repro", text: " A  B ", formats: ["wmf", "emf"] },
    { name: "repeated-spaces", text: "A  B", formats: ["wmf", "emf"] },
    { name: "tab", text: "A\tB", formats: ["emf"] },
    { name: "line-feed", text: "A\nB", formats: ["emf"] },
    { name: "carriage-return", text: "A\rB", formats: ["emf"] },
  ];
  for (const entry of cases) {
    for (const format of entry.formats) {
      const bytes = format === "wmf"
        ? createSyntheticWmfFixture([
            wmfFont(), wmfRecord(0x12d, words(0)), wmfRecord(0x12e, words(0x18)),
            wmfRecord(0x102, words(1)),
            wmfText(entry.text, Array.from(entry.text, () => 3)),
          ], { objectCount: 1 })
        : createSyntheticEmfFixture([
            emfFont(), emfRecord(37, dwords(1)), emfRecord(22, dwords(0x18)),
            emfRecord(18, dwords(1)), emfText(entry.text),
          ], { handleCount: 2 });
      await assertFailurePreservesCachedMedia(
        `media/whitespace-${entry.name}.${format}`,
        format,
        bytes,
        /whitespace/i,
      );
    }
  }
});

void test("fails closed for unsupported WMF text alignment and nonzero font escapement/orientation", async () => {
  const cases = [
    { name: "top alignment", font: wmfFont(), alignment: 0, message: /alignment/i },
    { name: "bottom alignment", font: wmfFont(), alignment: 8, message: /alignment/i },
    { name: "right alignment", font: wmfFont(), alignment: 2, message: /alignment/i },
    { name: "center alignment", font: wmfFont(), alignment: 6, message: /alignment/i },
    { name: "escapement", font: wmfFont({ escapement: 900 }), alignment: 0x18, message: /escapement|orientation/i },
    { name: "orientation", font: wmfFont({ orientation: 900 }), alignment: 0x18, message: /escapement|orientation/i },
  ] as const;
  for (const entry of cases) {
    const records = [
      entry.font, wmfRecord(0x12d, words(0)),
      ...(entry.alignment ? [wmfRecord(0x12e, words(entry.alignment))] : []),
      wmfRecord(0x102, words(1)), wmfText("AB", [3, 7]),
    ];
    await assert.rejects(
      resolvePreviewMedia(`media/${entry.name}.wmf`, createSyntheticWmfFixture(records, { objectCount: 1 }), options),
      (error: unknown) => error instanceof MetafileSerializeError && error.code === "unsupported-text" &&
        error.operationIndex === 0 && entry.message.test(error.message),
    );
  }
});

void test("does not paint parsed WMF text when source font metrics are unresolved", async () => {
  const source = createSyntheticWmfFixture([
    wmfFont(), wmfRecord(0x12d, words(0)), wmfRecord(0x12e, words(0x18)),
    wmfRecord(0x102, words(1)), wmfText("AB"),
  ], { objectCount: 1 });
  await assert.rejects(resolvePreviewMedia("media/unknown-metrics.wmf", source, options), (error: unknown) =>
    error instanceof MetafileSerializeError && error.code === "font-metrics-required");
});

void test("paints parsed WMF ETO_OPAQUE background independently of transparent background mode", async () => {
  const backgroundColor = Uint8Array.from([17, 34, 51, 0]);
  const source = createSyntheticWmfFixture([
    wmfFont(), wmfRecord(0x12d, words(0)), wmfRecord(0x12e, words(0x18)),
    wmfRecord(0x102, words(1)), wmfRecord(0x201, backgroundColor),
    wmfText("AB", [3, 7], 6),
  ], { objectCount: 1 });
  const resolved = await resolvePreviewMedia("media/opaque-transparent.wmf", source, options);
  const document = wmfTextDocument(resolved.bytes, "opaque-transparent.svg");
  const drawing = Array.from(document.getElementsByTagName("g"))[0];
  assert.ok(drawing);
  const children = childElements(drawing);
  assert.equal(children[0]?.localName, "path", "ETO_OPAQUE paints before text");
  assert.equal(children[0]?.getAttribute("fill"), "rgb(17,34,51)");
  assert.equal(children[0]?.getAttribute("d"), "M1 2L51 2L51 62L1 62Z");
  assert.equal(children[1]?.localName, "g", "ETO_CLIPPED scopes only the glyphs");
  assert.equal(childElements(children[1])[0]?.localName, "text");
});

void test("paints parsed EMF ETO_OPAQUE with transparent background mode before clipped glyphs", async () => {
  const source = createSyntheticEmfFixture([
    emfFont(), emfRecord(37, dwords(1)), emfRecord(22, dwords(0x18)),
    emfRecord(18, dwords(1)), emfRecord(25, dwords(0x332211)), emfText("AB", 6),
  ], { handleCount: 2 });
  const resolved = await resolvePreviewMedia("media/opaque-transparent.emf", source, options);
  const document = parseXml(new TextDecoder().decode(resolved.bytes), "opaque-transparent.emf.svg");
  const drawing = Array.from(document.getElementsByTagName("g"))[0];
  assert.ok(drawing);
  const children = childElements(drawing);
  assert.equal(children[0]?.localName, "path");
  assert.equal(children[0]?.getAttribute("fill"), "rgb(17,34,51)");
  assert.equal(children[0]?.getAttribute("d"), "M1 2L51 2L51 62L1 62Z");
  assert.equal(children[0]?.hasAttribute("clip-path"), false);
  assert.equal(children[1]?.localName, "g", "ETO_CLIPPED applies to glyphs after the opaque fill");
  assert.equal(childElements(children[1])[0]?.localName, "text");
});
