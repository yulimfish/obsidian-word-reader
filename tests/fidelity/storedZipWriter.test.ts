import assert from "node:assert/strict";
import test from "node:test";
import { createJiti } from "jiti";
import path from "node:path";

import { loadOoxmlZip } from "../../src/ooxml/jszipReader";
import {
  DEFAULT_DOCX_ZIP_LIMITS,
  validateZipSafety,
} from "../../src/ooxml/packageSafety";

const jiti = createJiti(__filename);
const OUTPUT_LIMIT = 64 * 1024 * 1024;

async function getWriter(): Promise<{
  writeStoredZip: (
    entries: readonly { path: string; bytes: Uint8Array }[],
    options: { isCancelled: () => boolean; yieldControl: () => Promise<void> },
  ) => Promise<ArrayBuffer>;
}> {
  let module: Record<string, unknown>;
  try {
    module = await jiti.import(
      path.resolve(process.cwd(), "src/ooxml/storedZipWriter.ts"),
    );
  } catch (error) {
    assert.fail(
      `writeStoredZip implementation is missing: ${error instanceof Error ? error.message : String(error)}`,
    );
  }
  assert.equal(typeof module.writeStoredZip, "function", "writeStoredZip export is required");
  return module as unknown as Awaited<ReturnType<typeof getWriter>>;
}

function crc32(bytes: Uint8Array): number {
  let crc = 0xffffffff;
  for (const byte of bytes) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit += 1) {
      crc = crc & 1 ? (crc >>> 1) ^ 0xedb88320 : crc >>> 1;
    }
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function inspectStoredZip(
  buffer: ArrayBuffer,
  expected: readonly { path: string; bytes: Uint8Array }[],
): void {
  const bytes = new Uint8Array(buffer);
  const view = new DataView(buffer);
  const encoder = new TextEncoder();
  const decoder = new TextDecoder("utf-8", { fatal: true });
  const localOffsets: number[] = [];
  let offset = 0;

  for (const entry of expected) {
    localOffsets.push(offset);
    assert.equal(view.getUint32(offset, true), 0x04034b50);
    assert.equal(view.getUint16(offset + 6, true), 0x0800);
    assert.equal(view.getUint16(offset + 8, true), 0);
    assert.equal(view.getUint16(offset + 10, true), 0);
    assert.equal(view.getUint16(offset + 12, true), 0x0021);
    assert.equal(view.getUint32(offset + 14, true), crc32(entry.bytes));
    assert.equal(view.getUint32(offset + 18, true), entry.bytes.byteLength);
    assert.equal(view.getUint32(offset + 22, true), entry.bytes.byteLength);
    const nameLength = view.getUint16(offset + 26, true);
    assert.equal(view.getUint16(offset + 28, true), 0);
    const name = decoder.decode(bytes.subarray(offset + 30, offset + 30 + nameLength));
    assert.equal(name, entry.path);
    const dataStart = offset + 30 + nameLength;
    assert.deepEqual(bytes.subarray(dataStart, dataStart + entry.bytes.byteLength), entry.bytes);
    offset = dataStart + entry.bytes.byteLength;
  }

  const centralOffset = offset;
  for (let index = 0; index < expected.length; index += 1) {
    const entry = expected[index];
    assert.equal(view.getUint32(offset, true), 0x02014b50);
    assert.equal(view.getUint16(offset + 8, true), 0x0800);
    assert.equal(view.getUint16(offset + 10, true), 0);
    assert.equal(view.getUint16(offset + 12, true), 0);
    assert.equal(view.getUint16(offset + 14, true), 0x0021);
    assert.equal(view.getUint32(offset + 16, true), crc32(entry.bytes));
    assert.equal(view.getUint32(offset + 20, true), entry.bytes.byteLength);
    assert.equal(view.getUint32(offset + 24, true), entry.bytes.byteLength);
    const nameLength = view.getUint16(offset + 28, true);
    assert.equal(view.getUint16(offset + 30, true), 0);
    assert.equal(view.getUint16(offset + 32, true), 0);
    assert.equal(view.getUint16(offset + 34, true), 0);
    assert.equal(view.getUint32(offset + 38, true), 0);
    assert.equal(view.getUint32(offset + 42, true), localOffsets[index]);
    assert.equal(decoder.decode(bytes.subarray(offset + 46, offset + 46 + nameLength)), entry.path);
    offset += 46 + nameLength;
  }

  assert.equal(view.getUint32(offset, true), 0x06054b50);
  assert.equal(view.getUint16(offset + 4, true), 0);
  assert.equal(view.getUint16(offset + 6, true), 0);
  assert.equal(view.getUint16(offset + 8, true), expected.length);
  assert.equal(view.getUint16(offset + 10, true), expected.length);
  assert.equal(view.getUint32(offset + 12, true), offset - centralOffset);
  assert.equal(view.getUint32(offset + 16, true), centralOffset);
  assert.equal(view.getUint16(offset + 20, true), 0);
  assert.equal(offset + 22, bytes.byteLength);
  assert.ok(encoder.encode(expected[0]?.path ?? "").byteLength > 0 || expected.length === 0);
}

void test("stored ZIP writer emits independently verifiable ZIP32 STORE records", async () => {
  const { writeStoredZip } = await getWriter();
  const entries = [
    { path: "word/123456789.txt", bytes: new TextEncoder().encode("123456789") },
    { path: "word/你好.xml", bytes: new TextEncoder().encode("<document>ok</document>") },
  ] as const;
  let yields = 0;
  const output = await writeStoredZip(entries, {
    isCancelled: () => false,
    yieldControl: async () => { yields += 1; },
  });

  assert.equal(new DataView(output).getUint32(14, true), 0xcbf43926);
  assert.ok(yields > 0);
  inspectStoredZip(output, entries);
  const safety = validateZipSafety(output, DEFAULT_DOCX_ZIP_LIMITS);
  assert.equal(safety.fileCount, entries.length);
  assert.equal(safety.entries.every((entry) => entry.compressionMethod === 0), true);

  const zip = await loadOoxmlZip(output);
  for (const entry of entries) {
    assert.deepEqual(
      await zip.file(entry.path)!.async("uint8array"),
      entry.bytes,
    );
  }
});

void test("stored ZIP writer rejects unsafe names, duplicate paths and ZIP32 field overflow", async () => {
  const { writeStoredZip } = await getWriter();
  const options = { isCancelled: () => false, yieldControl: async () => {} };
  const byte = new Uint8Array([1]);

  await assert.rejects(() => writeStoredZip([
    { path: "word/a.xml", bytes: byte },
    { path: "word/a.xml", bytes: byte },
  ], options));
  for (const invalidPath of ["", "../escape", "word//file.xml", "word/./file.xml", "word/", "/absolute.xml", "C:/drive.xml"]) {
    await assert.rejects(() => writeStoredZip([{ path: invalidPath, bytes: byte }], options));
  }
  await assert.rejects(
    () => writeStoredZip([{ path: "a".repeat(65_536), bytes: byte }], options),
  );
  await assert.rejects(
    () => writeStoredZip(
      Array.from({ length: 2_001 }, (_, index) => ({ path: `part-${index}.xml`, bytes: new Uint8Array() })),
      options,
    ),
  );

  const oversized = new Proxy(new Uint8Array(), {
    get(target, property, receiver) {
      if (property === "byteLength") return OUTPUT_LIMIT + 1;
      const value: unknown = Reflect.get(target, property, receiver);
      return value;
    },
  }) as Uint8Array;
  await assert.rejects(
    () => writeStoredZip([{ path: "large.bin", bytes: oversized }], options),
  );
});

void test("stored ZIP writer observes cancellation during chunked CRC and returns no partial archive", async () => {
  const { writeStoredZip } = await getWriter();
  const bytes = new Uint8Array(128 * 1024).fill(0x5a);
  let cancelled = false;
  let yields = 0;
  await assert.rejects(
    () => writeStoredZip([{ path: "large.bin", bytes }], {
      isCancelled: () => cancelled,
      yieldControl: async () => {
        yields += 1;
        cancelled = yields === 2;
      },
    }),
    /cancel/i,
  );
  assert.equal(yields, 2);
});
