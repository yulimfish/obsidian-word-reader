import {
  DEFAULT_DOCX_ZIP_LIMITS,
  validateZipSafety,
} from "./packageSafety";

export interface StoredZipEntry {
  readonly path: string;
  readonly bytes: Uint8Array;
}

export interface StoredZipOptions {
  readonly isCancelled: () => boolean;
  readonly yieldControl: () => Promise<void>;
}

interface PreparedEntry {
  readonly path: string;
  readonly bytes: Uint8Array;
  readonly nameLength: number;
  readonly byteLength: number;
  crc32: number;
  localOffset: number;
}

const MAX_ENTRIES = 2_000;
const MAX_ENTRY_BYTES = 64 * 1024 * 1024;
const MAX_TOTAL_BYTES = 256 * 1024 * 1024;
const MAX_ARCHIVE_BYTES = 272 * 1024 * 1024;
const MAX_WORKING_BYTES = 544 * 1024 * 1024;
const CRC_CHUNK_BYTES = 64 * 1024;
const UINT16_MAX = 0xffff;
const UINT32_MAX = 0xffffffff;
const UTF8_FLAG = 0x0800;
const DOS_DATE = 0x0021;
const LOCAL_HEADER_BYTES = 30;
const CENTRAL_HEADER_BYTES = 46;
const END_RECORD_BYTES = 22;

const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let index = 0; index < table.length; index += 1) {
    let value = index;
    for (let bit = 0; bit < 8; bit += 1) {
      value = value & 1 ? (value >>> 1) ^ 0xedb88320 : value >>> 1;
    }
    table[index] = value >>> 0;
  }
  return table;
})();

class StoredZipError extends Error {
  constructor(readonly kind: "invalid" | "limit-exceeded" | "cancelled", message: string) {
    super(message);
    this.name = "StoredZipError";
  }
}

function fail(kind: StoredZipError["kind"], message: string): never {
  throw new StoredZipError(kind, message);
}

function assertNotCancelled(options: StoredZipOptions): void {
  if (options.isCancelled()) {
    fail("cancelled", "Stored ZIP writing was cancelled.");
  }
}

function checkedAdd(total: number, increment: number, ceiling: number, label: string): number {
  if (!Number.isSafeInteger(total) || !Number.isSafeInteger(increment) || increment < 0 || total > ceiling - increment) {
    fail("limit-exceeded", `${label} exceeds the supported ZIP32 limit.`);
  }
  return total + increment;
}

function utf8ByteLength(value: string): number {
  let length = 0;
  for (let index = 0; index < value.length; index += 1) {
    const code = value.charCodeAt(index);
    if (code <= 0x7f) {
      length += 1;
    } else if (code <= 0x7ff) {
      length += 2;
    } else if (code >= 0xd800 && code <= 0xdbff) {
      const next = value.charCodeAt(index + 1);
      if (next < 0xdc00 || next > 0xdfff) {
        fail("invalid", "ZIP entry paths must contain valid UTF-16 text.");
      }
      length += 4;
      index += 1;
    } else if (code >= 0xdc00 && code <= 0xdfff) {
      fail("invalid", "ZIP entry paths must contain valid UTF-16 text.");
    } else {
      length += 3;
    }
    if (length > UINT16_MAX) {
      fail("limit-exceeded", "A ZIP entry path exceeds the ZIP32 filename field.");
    }
  }
  return length;
}

function validatePath(path: string): void {
  if (
    path.length === 0 ||
    path.startsWith("/") ||
    /^[a-z]:/i.test(path) ||
    /[\\%:#?\s]/.test(path) ||
    Array.from(path).some((character) => {
      const code = character.charCodeAt(0);
      return code < 0x20 || code === 0x7f;
    }) ||
    path.split("/").some((part) => part.length === 0 || part === "." || part === "..")
  ) {
    fail("invalid", "ZIP entry paths must be canonical file paths.");
  }
}

function preflight(entries: readonly StoredZipEntry[]): {
  prepared: PreparedEntry[];
  totalPayloadBytes: number;
  archiveBytes: number;
} {
  if (!Array.isArray(entries)) {
    fail("invalid", "ZIP entries must be provided as an array.");
  }
  if (entries.length > MAX_ENTRIES) {
    fail("limit-exceeded", `The archive exceeds the ${MAX_ENTRIES} entry limit.`);
  }

  const paths = new Set<string>();
  const prepared: PreparedEntry[] = [];
  let totalPayloadBytes = 0;
  let archiveBytes = END_RECORD_BYTES;

  for (const candidate of entries as readonly unknown[]) {
    if (!candidate || typeof candidate !== "object") {
      fail("invalid", "Each ZIP entry requires a canonical path and Uint8Array bytes.");
    }
    const record = candidate as { readonly path?: unknown; readonly bytes?: unknown };
    const path = record.path;
    const bytes = record.bytes;
    if (typeof path !== "string" || !(bytes instanceof Uint8Array) || !ArrayBuffer.isView(bytes)) {
      fail("invalid", "Each ZIP entry requires a canonical path and Uint8Array bytes.");
    }
    validatePath(path);
    if (paths.has(path)) {
      fail("invalid", `The archive contains a duplicate path: ${path}`);
    }
    paths.add(path);

    const nameLength = utf8ByteLength(path);
    if (nameLength === 0 || nameLength > UINT16_MAX) {
      fail("limit-exceeded", "A ZIP entry path exceeds the ZIP32 filename field.");
    }
    const byteLength = bytes.byteLength;
    if (!Number.isSafeInteger(byteLength) || byteLength < 0 || byteLength > MAX_ENTRY_BYTES || byteLength > UINT32_MAX) {
      fail("limit-exceeded", `ZIP entry ${path} exceeds the ${MAX_ENTRY_BYTES} byte limit.`);
    }
    totalPayloadBytes = checkedAdd(totalPayloadBytes, byteLength, MAX_TOTAL_BYTES, "Total ZIP payload size");
    archiveBytes = checkedAdd(
      archiveBytes,
      LOCAL_HEADER_BYTES + nameLength + byteLength,
      MAX_ARCHIVE_BYTES,
      "ZIP archive size",
    );
    archiveBytes = checkedAdd(
      archiveBytes,
      CENTRAL_HEADER_BYTES + nameLength,
      MAX_ARCHIVE_BYTES,
      "ZIP archive size",
    );
    if (archiveBytes > UINT32_MAX) {
      fail("limit-exceeded", "The archive exceeds the ZIP32 size fields.");
    }
    prepared.push({
      path,
      bytes,
      nameLength,
      byteLength,
      crc32: 0,
      localOffset: 0,
    });
  }

  if (totalPayloadBytes > MAX_TOTAL_BYTES || archiveBytes > MAX_ARCHIVE_BYTES) {
    fail("limit-exceeded", "The archive exceeds its bounded ZIP32 size limits.");
  }
  const workingBytes = checkedAdd(totalPayloadBytes, archiveBytes, MAX_WORKING_BYTES, "Combined ZIP working buffers");
  if (workingBytes > MAX_WORKING_BYTES) {
    fail("limit-exceeded", "Combined ZIP input and output buffers exceed the working limit.");
  }
  if (prepared.length > UINT16_MAX) {
    fail("limit-exceeded", "The archive exceeds the ZIP32 entry-count field.");
  }
  return { prepared, totalPayloadBytes, archiveBytes };
}

async function writeEntryData(
  output: Uint8Array,
  outputOffset: number,
  entry: PreparedEntry,
  options: StoredZipOptions,
): Promise<number> {
  let crc = 0xffffffff;
  let copied = 0;
  while (copied < entry.byteLength) {
    assertNotCancelled(options);
    if (entry.bytes.byteLength !== entry.byteLength) {
      fail("invalid", `ZIP entry ${entry.path} changed size after preflight.`);
    }
    const chunkLength = Math.min(CRC_CHUNK_BYTES, entry.byteLength - copied);
    const sourceChunk = entry.bytes.subarray(copied, copied + chunkLength);
    output.set(sourceChunk, outputOffset + copied);
    for (let index = 0; index < chunkLength; index += 1) {
      crc = (crc >>> 8) ^ CRC_TABLE[(crc ^ output[outputOffset + copied + index]) & 0xff];
    }
    copied += chunkLength;
    await options.yieldControl();
    assertNotCancelled(options);
    if (entry.bytes.byteLength !== entry.byteLength) {
      fail("invalid", `ZIP entry ${entry.path} changed size after preflight.`);
    }
  }
  entry.crc32 = (crc ^ 0xffffffff) >>> 0;
  return copied;
}

function writeName(output: Uint8Array, offset: number, name: string, expectedBytes: number): number {
  const target = output.subarray(offset, offset + expectedBytes);
  const result = new TextEncoder().encodeInto(name, target);
  if (result.read !== name.length || result.written !== expectedBytes) {
    fail("invalid", "A ZIP entry path could not be encoded as canonical UTF-8.");
  }
  return result.written;
}

function writeLocalHeader(
  view: DataView,
  offset: number,
  nameLength: number,
  byteLength: number,
): void {
  view.setUint32(offset, 0x04034b50, true);
  view.setUint16(offset + 4, 20, true);
  view.setUint16(offset + 6, UTF8_FLAG, true);
  view.setUint16(offset + 8, 0, true);
  view.setUint16(offset + 10, 0, true);
  view.setUint16(offset + 12, DOS_DATE, true);
  view.setUint32(offset + 14, 0, true);
  view.setUint32(offset + 18, byteLength, true);
  view.setUint32(offset + 22, byteLength, true);
  view.setUint16(offset + 26, nameLength, true);
  view.setUint16(offset + 28, 0, true);
}

function writeCentralHeader(view: DataView, offset: number, entry: PreparedEntry): void {
  view.setUint32(offset, 0x02014b50, true);
  view.setUint16(offset + 4, 20, true);
  view.setUint16(offset + 6, 20, true);
  view.setUint16(offset + 8, UTF8_FLAG, true);
  view.setUint16(offset + 10, 0, true);
  view.setUint16(offset + 12, 0, true);
  view.setUint16(offset + 14, DOS_DATE, true);
  view.setUint32(offset + 16, entry.crc32, true);
  view.setUint32(offset + 20, entry.bytes.byteLength, true);
  view.setUint32(offset + 24, entry.bytes.byteLength, true);
  view.setUint16(offset + 28, entry.nameLength, true);
  view.setUint16(offset + 30, 0, true);
  view.setUint16(offset + 32, 0, true);
  view.setUint16(offset + 34, 0, true);
  view.setUint16(offset + 36, 0, true);
  view.setUint32(offset + 38, 0, true);
  view.setUint32(offset + 42, entry.localOffset, true);
}

function writeEndRecord(view: DataView, offset: number, entryCount: number, centralOffset: number, centralSize: number): void {
  view.setUint32(offset, 0x06054b50, true);
  view.setUint16(offset + 4, 0, true);
  view.setUint16(offset + 6, 0, true);
  view.setUint16(offset + 8, entryCount, true);
  view.setUint16(offset + 10, entryCount, true);
  view.setUint32(offset + 12, centralSize, true);
  view.setUint32(offset + 16, centralOffset, true);
  view.setUint16(offset + 20, 0, true);
}

export async function writeStoredZip(
  entries: readonly StoredZipEntry[],
  options: StoredZipOptions,
): Promise<ArrayBuffer> {
  if (!options || typeof options.isCancelled !== "function" || typeof options.yieldControl !== "function") {
    throw new TypeError("Stored ZIP cancellation and yield callbacks are required.");
  }

  let prepared: PreparedEntry[] = [];
  let output: Uint8Array | null = null;
  try {
    assertNotCancelled(options);
    const plan = preflight(entries);
    prepared = plan.prepared;
    for (const entry of prepared) {
      assertNotCancelled(options);
      if (entry.bytes.byteLength > MAX_ENTRY_BYTES) {
        fail("limit-exceeded", `ZIP entry ${entry.path} exceeds the ${MAX_ENTRY_BYTES} byte limit.`);
      }
      if (entry.bytes.byteLength !== entry.byteLength) {
        fail("invalid", `ZIP entry ${entry.path} changed size after preflight.`);
      }
      // Confirm the preflight's UTF-8 length without retaining an encoded-name copy.
      const nameLength = utf8ByteLength(entry.path);
      if (nameLength !== entry.nameLength) {
        fail("invalid", "A ZIP entry path changed after preflight.");
      }
      await options.yieldControl();
      assertNotCancelled(options);
      if (entry.bytes.byteLength !== entry.byteLength) {
        fail("invalid", `ZIP entry ${entry.path} changed size after preflight.`);
      }
    }

    output = new Uint8Array(plan.archiveBytes);
    const view = new DataView(output.buffer, output.byteOffset, output.byteLength);
    let offset = 0;
    for (const entry of prepared) {
      assertNotCancelled(options);
      entry.localOffset = offset;
      writeLocalHeader(view, offset, entry.nameLength, entry.byteLength);
      const nameOffset = offset + LOCAL_HEADER_BYTES;
      const writtenNameBytes = writeName(output, nameOffset, entry.path, entry.nameLength);
      const dataOffset = nameOffset + writtenNameBytes;
      const copied = await writeEntryData(output, dataOffset, entry, options);
      view.setUint32(offset + 14, entry.crc32, true);
      offset = dataOffset + copied;
      await options.yieldControl();
      assertNotCancelled(options);
    }

    const centralOffset = offset;
    for (const entry of prepared) {
      assertNotCancelled(options);
      writeCentralHeader(view, offset, entry);
      const nameOffset = offset + CENTRAL_HEADER_BYTES;
      offset += CENTRAL_HEADER_BYTES + writeName(output, nameOffset, entry.path, entry.nameLength);
      await options.yieldControl();
      assertNotCancelled(options);
    }
    const centralSize = offset - centralOffset;
    writeEndRecord(view, offset, prepared.length, centralOffset, centralSize);
    offset += END_RECORD_BYTES;
    if (offset !== output.byteLength || plan.totalPayloadBytes > MAX_TOTAL_BYTES) {
      fail("invalid", "The ZIP writer did not match its exact preflight size.");
    }
    const result = output.buffer as ArrayBuffer;
    validateZipSafety(result, DEFAULT_DOCX_ZIP_LIMITS);
    prepared.length = 0;
    entries = [];
    output = null;
    return result;
  } catch (error) {
    output?.fill(0);
    output = null;
    prepared.length = 0;
    entries = [];
    throw error;
  }
}
