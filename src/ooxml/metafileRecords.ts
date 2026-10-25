export interface MetafileLimits {
  maxBytes: number;
  maxRecords: number;
  maxObjects: number;
  maxStack: number;
  maxSceneOps: number;
}

export interface MetafileBounds {
  readonly left: number;
  readonly top: number;
  readonly right: number;
  readonly bottom: number;
}

export interface MetafileRecord {
  readonly type: number;
  readonly offset: number;
  readonly size: number;
  readonly payload: Uint8Array;
}

export interface MetafileRecordSet {
  readonly format: "wmf" | "emf";
  readonly bounds: MetafileBounds | null;
  readonly records: readonly MetafileRecord[];
  readonly declaredObjectTableSize: number;
  readonly limits: Readonly<MetafileLimits>;
}

export type MetafileParseErrorCode =
  | "invalid-limits"
  | "invalid-format"
  | "file-too-large"
  | "too-many-records"
  | "too-many-objects"
  | "truncated-header"
  | "invalid-header"
  | "declared-size-mismatch"
  | "record-count-mismatch"
  | "invalid-record-size"
  | "truncated-record"
  | "invalid-handle"
  | "missing-eof"
  | "trailing-data";

export class MetafileParseError extends Error {
  constructor(
    readonly code: MetafileParseErrorCode,
    message: string,
  ) {
    super(message);
    this.name = "MetafileParseError";
  }
}

export const DEFAULT_METAFILE_LIMITS: Readonly<MetafileLimits> = Object.freeze({
  maxBytes: 16 * 1024 * 1024,
  maxRecords: 50_000,
  maxObjects: 4_096,
  maxStack: 256,
  maxSceneOps: 50_000,
});

const WMF_PLACEABLE_SIGNATURE = 0x9ac6cdd7;
const EMF_SIGNATURE = 0x464d4520;
const EMF_HEADER_SIZE = 88;
const EMF_STOCK_OBJECT_FLAG = 0x80000000;
const EMF_LAST_STOCK_OBJECT = 19;

interface ValidatedLimits {
  readonly maxBytes: number;
  readonly maxRecords: number;
  readonly maxObjects: number;
  readonly maxStack: number;
  readonly maxSceneOps: number;
}

export function parseMetafileRecords(
  bytes: Uint8Array,
  format: "wmf" | "emf",
  limits: MetafileLimits = DEFAULT_METAFILE_LIMITS,
): MetafileRecordSet {
  const checkedLimits = validateLimits(limits);
  if (format !== "wmf" && format !== "emf") {
    throw new MetafileParseError("invalid-format", "Unsupported metafile format.");
  }
  if (bytes.byteLength > checkedLimits.maxBytes) {
    throw new MetafileParseError("file-too-large", "Metafile exceeds the byte limit.");
  }

  return format === "wmf"
    ? parseWmfRecords(bytes, checkedLimits)
    : parseEmfRecords(bytes, checkedLimits);
}

function validateLimits(limits: MetafileLimits): ValidatedLimits {
  const snapshot = {
    maxBytes: limits.maxBytes,
    maxRecords: limits.maxRecords,
    maxObjects: limits.maxObjects,
    maxStack: limits.maxStack,
    maxSceneOps: limits.maxSceneOps,
  };
  const entries: ReadonlyArray<[keyof MetafileLimits, number]> = [
    ["maxBytes", snapshot.maxBytes],
    ["maxRecords", snapshot.maxRecords],
    ["maxObjects", snapshot.maxObjects],
    ["maxStack", snapshot.maxStack],
    ["maxSceneOps", snapshot.maxSceneOps],
  ];
  for (const [name, value] of entries) {
    if (!Number.isSafeInteger(value) || value < 0) {
      throw new MetafileParseError(
        "invalid-limits",
        `Metafile limit ${name} must be a non-negative safe integer.`,
      );
    }
  }
  return Object.freeze(snapshot);
}

function parseWmfRecords(
  bytes: Uint8Array,
  limits: ValidatedLimits,
): MetafileRecordSet {
  if (bytes.byteLength < 4) {
    throw new MetafileParseError("truncated-header", "WMF header is truncated.");
  }
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const hasPlaceableHeader = view.getUint32(0, true) === WMF_PLACEABLE_SIGNATURE;
  const standardHeaderOffset = hasPlaceableHeader ? 22 : 0;

  if (hasPlaceableHeader && bytes.byteLength < standardHeaderOffset) {
    throw new MetafileParseError(
      "truncated-header",
      "Placeable WMF header is truncated.",
    );
  }
  if (bytes.byteLength - standardHeaderOffset < 18) {
    throw new MetafileParseError("truncated-header", "WMF header is truncated.");
  }

  const bounds = hasPlaceableHeader ? readPlaceableBounds(view) : null;
  const fileType = view.getUint16(standardHeaderOffset, true);
  const headerWords = view.getUint16(standardHeaderOffset + 2, true);
  const version = view.getUint16(standardHeaderOffset + 4, true);
  const declaredWords = view.getUint32(standardHeaderOffset + 6, true);
  const objectCount = view.getUint16(standardHeaderOffset + 10, true);
  const maxRecordWords = view.getUint32(standardHeaderOffset + 12, true);
  const parameterCount = view.getUint16(standardHeaderOffset + 16, true);

  if (
    (fileType !== 1 && fileType !== 2) ||
    headerWords !== 9 ||
    (version !== 0x0100 && version !== 0x0300) ||
    parameterCount !== 0
  ) {
    throw new MetafileParseError("invalid-header", "WMF standard header is invalid.");
  }
  if (objectCount > limits.maxObjects) {
    throw new MetafileParseError("too-many-objects", "WMF object table exceeds the limit.");
  }

  const declaredByteLength = checkedMultiply(declaredWords, 2, "WMF file size");
  if (declaredByteLength > limits.maxBytes) {
    throw new MetafileParseError("file-too-large", "Declared WMF size exceeds the byte limit.");
  }
  if (declaredByteLength < 24 || declaredByteLength !== bytes.byteLength - standardHeaderOffset) {
    throw new MetafileParseError(
      "declared-size-mismatch",
      "WMF declared size does not match the supplied bytes.",
    );
  }
  if (maxRecordWords < 3 || maxRecordWords > declaredWords - headerWords) {
    throw new MetafileParseError("invalid-header", "WMF maximum record size is invalid.");
  }

  const declaredEnd = checkedAdd(standardHeaderOffset, declaredByteLength, "WMF end offset");
  const records: MetafileRecord[] = [];
  let offset = standardHeaderOffset + 18;
  let foundEof = false;
  const maxRecordBytes = checkedMultiply(maxRecordWords, 2, "WMF maximum record size");

  while (offset < declaredEnd) {
    ensureRecordCapacity(records, limits.maxRecords);
    if (declaredEnd - offset < 6) {
      throw new MetafileParseError("truncated-record", "WMF record header is truncated.");
    }

    const recordWords = view.getUint32(offset, true);
    if (recordWords < 3) {
      throw new MetafileParseError("invalid-record-size", "WMF record size is too small.");
    }
    const size = checkedMultiply(recordWords, 2, "WMF record size");
    if (size > maxRecordBytes) {
      throw new MetafileParseError("invalid-record-size", "WMF record exceeds its declared maximum.");
    }
    if (size > declaredEnd - offset) {
      throw new MetafileParseError("truncated-record", "WMF record exceeds the declared file end.");
    }

    const type = view.getUint16(offset + 4, true);
    const nextOffset = checkedAdd(offset, size, "WMF record end");
    validateWmfHandleReference(view, offset, size, type, objectCount);
    const record = createRecord(bytes, offset, size, type, 6);
    records.push(record);

    if (type === 0) {
      if (size !== 6) {
        throw new MetafileParseError("invalid-record-size", "WMF EOF record must be six bytes.");
      }
      if (nextOffset !== declaredEnd) {
        throw new MetafileParseError("trailing-data", "Bytes follow the WMF EOF record.");
      }
      foundEof = true;
      offset = nextOffset;
      break;
    }
    offset = nextOffset;
  }

  if (!foundEof) {
    throw new MetafileParseError("missing-eof", "WMF has no terminal EOF record.");
  }
  return createRecordSet("wmf", bounds, records, objectCount, limits);
}

function readPlaceableBounds(view: DataView): MetafileBounds {
  const handle = view.getUint16(4, true);
  const left = view.getInt16(6, true);
  const top = view.getInt16(8, true);
  const right = view.getInt16(10, true);
  const bottom = view.getInt16(12, true);
  const unitsPerInch = view.getUint16(14, true);
  const reserved = view.getUint32(16, true);

  let checksum = 0;
  for (let offset = 0; offset < 20; offset += 2) {
    checksum ^= view.getUint16(offset, true);
  }
  if (
    handle !== 0 ||
    unitsPerInch === 0 ||
    reserved !== 0 ||
    checksum !== view.getUint16(20, true) ||
    right < left ||
    bottom < top
  ) {
    throw new MetafileParseError("invalid-header", "Placeable WMF header is invalid.");
  }
  return Object.freeze({ left, top, right, bottom });
}

function validateWmfHandleReference(
  view: DataView,
  offset: number,
  size: number,
  type: number,
  objectCount: number,
): void {
  if (type !== 0x012d && type !== 0x01f0) {
    return;
  }
  if (size !== 8) {
    throw new MetafileParseError(
      "invalid-record-size",
      "WMF object-handle record must contain one WORD handle.",
    );
  }
  const handle = view.getUint16(offset + 6, true);
  if (handle >= objectCount) {
    throw new MetafileParseError("invalid-handle", "WMF object handle is outside its declared table.");
  }
}

function parseEmfRecords(
  bytes: Uint8Array,
  limits: ValidatedLimits,
): MetafileRecordSet {
  if (bytes.byteLength < EMF_HEADER_SIZE) {
    throw new MetafileParseError("truncated-header", "EMF header is truncated.");
  }
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const headerType = view.getUint32(0, true);
  const headerSize = view.getUint32(4, true);
  if (headerType !== 1 || headerSize < EMF_HEADER_SIZE || headerSize % 4 !== 0) {
    throw new MetafileParseError("invalid-header", "EMF header record is invalid.");
  }
  if (headerSize > bytes.byteLength) {
    throw new MetafileParseError("truncated-header", "EMF header record exceeds the supplied bytes.");
  }

  const bounds = readEmfBounds(view);
  const signature = view.getUint32(40, true);
  const version = view.getUint32(44, true);
  const declaredByteLength = view.getUint32(48, true);
  const declaredRecordCount = view.getUint32(52, true);
  const handleCount = view.getUint16(56, true);
  const reserved = view.getUint16(58, true);
  const headerPaletteCount = view.getUint32(68, true);
  if (signature !== EMF_SIGNATURE || version !== 0x00010000 || reserved !== 0) {
    throw new MetafileParseError("invalid-header", "EMF signature or fixed header fields are invalid.");
  }
  if (declaredByteLength > limits.maxBytes) {
    throw new MetafileParseError("file-too-large", "Declared EMF size exceeds the byte limit.");
  }
  if (declaredByteLength !== bytes.byteLength || declaredByteLength < headerSize + 20) {
    throw new MetafileParseError(
      "declared-size-mismatch",
      "EMF declared size does not match the supplied bytes.",
    );
  }
  if (declaredRecordCount < 2) {
    throw new MetafileParseError("invalid-header", "EMF must declare a header and an EOF record.");
  }
  if (declaredRecordCount > limits.maxRecords) {
    throw new MetafileParseError("too-many-records", "EMF record count exceeds the limit.");
  }
  if (handleCount === 0 || handleCount - 1 > limits.maxObjects) {
    throw new MetafileParseError("too-many-objects", "EMF handle table exceeds the limit.");
  }
  validateEmfDescription(view, headerSize);

  const records: MetafileRecord[] = [createRecord(bytes, 0, headerSize, headerType, 8)];
  let offset = headerSize;
  let foundEof = false;

  while (offset < declaredByteLength) {
    ensureRecordCapacity(records, limits.maxRecords);
    if (declaredByteLength - offset < 8) {
      throw new MetafileParseError("truncated-record", "EMF record header is truncated.");
    }

    const type = view.getUint32(offset, true);
    const size = view.getUint32(offset + 4, true);
    if (size < 8 || size % 4 !== 0) {
      throw new MetafileParseError("invalid-record-size", "EMF record size is invalid.");
    }
    if (size > declaredByteLength - offset) {
      throw new MetafileParseError("truncated-record", "EMF record exceeds the declared file end.");
    }

    const nextOffset = checkedAdd(offset, size, "EMF record end");
    validateEmfHandleReference(view, offset, size, type, handleCount);
    if (type === 1) {
      throw new MetafileParseError("invalid-record-size", "EMF header may appear only as the first record.");
    }
    if (type === 14) {
      validateEmfEofRecord(view, offset, size, headerPaletteCount);
      if (nextOffset !== declaredByteLength) {
        throw new MetafileParseError("trailing-data", "Bytes follow the EMF EOF record.");
      }
    }

    records.push(createRecord(bytes, offset, size, type, 8));
    if (type === 14) {
      foundEof = true;
      offset = nextOffset;
      break;
    }
    offset = nextOffset;
  }

  if (!foundEof) {
    throw new MetafileParseError("missing-eof", "EMF has no terminal EOF record.");
  }
  if (records.length !== declaredRecordCount) {
    throw new MetafileParseError(
      "record-count-mismatch",
      "EMF declared record count does not match the framed records.",
    );
  }
  return createRecordSet("emf", bounds, records, handleCount, limits);
}

function validateEmfEofRecord(
  view: DataView,
  offset: number,
  size: number,
  headerPaletteCount: number,
): void {
  if (size < 20) {
    throw new MetafileParseError("invalid-record-size", "EMF EOF record is truncated.");
  }

  const paletteCount = view.getUint32(offset + 8, true);
  const paletteOffset = view.getUint32(offset + 12, true);
  if (paletteCount !== headerPaletteCount) {
    throw new MetafileParseError(
      "invalid-header",
      "EMF header palette count does not match the EOF palette count.",
    );
  }
  const sizeLastOffset = size - 4;
  const sizeLastByteOffset = checkedAdd(offset, sizeLastOffset, "EMF EOF SizeLast offset");
  if (view.getUint32(sizeLastByteOffset, true) !== size) {
    throw new MetafileParseError("invalid-record-size", "EMF EOF SizeLast must equal its record size.");
  }

  if (paletteOffset === 0 && paletteCount === 0) {
    return;
  }
  if (paletteOffset < 16 || paletteOffset > sizeLastOffset) {
    throw new MetafileParseError("invalid-record-size", "EMF EOF palette offset is outside its data area.");
  }
  const paletteByteLength = checkedMultiply(paletteCount, 4, "EMF EOF palette size");
  if (paletteByteLength > sizeLastOffset - paletteOffset) {
    throw new MetafileParseError("invalid-record-size", "EMF EOF palette exceeds its record bounds.");
  }
}

function readEmfBounds(view: DataView): MetafileBounds {
  const left = view.getInt32(8, true);
  const top = view.getInt32(12, true);
  const right = view.getInt32(16, true);
  const bottom = view.getInt32(20, true);
  if (right < left || bottom < top) {
    throw new MetafileParseError("invalid-header", "EMF bounds are inverted.");
  }
  return Object.freeze({ left, top, right, bottom });
}

function validateEmfDescription(view: DataView, headerSize: number): void {
  const characterCount = view.getUint32(60, true);
  const descriptionOffset = view.getUint32(64, true);
  if (characterCount === 0) {
    return;
  }
  const descriptionSize = checkedMultiply(characterCount, 2, "EMF description size");
  if (
    descriptionOffset < EMF_HEADER_SIZE ||
    descriptionOffset % 2 !== 0 ||
    descriptionOffset > headerSize ||
    descriptionSize > headerSize - descriptionOffset
  ) {
    throw new MetafileParseError("invalid-header", "EMF description is outside its header record.");
  }
}

function validateEmfHandleReference(
  view: DataView,
  offset: number,
  size: number,
  type: number,
  handleCount: number,
): void {
  if (type !== 37 && type !== 40) {
    return;
  }
  if (size !== 12) {
    throw new MetafileParseError(
      "invalid-record-size",
      "EMF object-handle record must contain one DWORD handle.",
    );
  }
  const handle = view.getUint32(offset + 8, true);
  if ((handle & EMF_STOCK_OBJECT_FLAG) !== 0) {
    if ((handle & ~EMF_STOCK_OBJECT_FLAG) > EMF_LAST_STOCK_OBJECT) {
      throw new MetafileParseError("invalid-handle", "EMF stock object handle is invalid.");
    }
    return;
  }
  if (handle === 0 || handle >= handleCount) {
    throw new MetafileParseError("invalid-handle", "EMF object handle is outside its declared table.");
  }
}

function createRecord(
  bytes: Uint8Array,
  offset: number,
  size: number,
  type: number,
  headerSize: number,
): MetafileRecord {
  const payloadOffset = checkedAdd(offset, headerSize, "metafile payload offset");
  const recordEnd = checkedAdd(offset, size, "metafile record end");
  return Object.freeze({
    type,
    offset,
    size,
    payload: bytes.subarray(payloadOffset, recordEnd),
  });
}

function createRecordSet(
  format: "wmf" | "emf",
  bounds: MetafileBounds | null,
  records: MetafileRecord[],
  declaredObjectTableSize: number,
  limits: ValidatedLimits,
): MetafileRecordSet {
  return Object.freeze({
    format,
    bounds,
    records: Object.freeze(records),
    declaredObjectTableSize,
    limits,
  });
}

function ensureRecordCapacity(records: readonly MetafileRecord[], maxRecords: number): void {
  if (records.length >= maxRecords) {
    throw new MetafileParseError("too-many-records", "Metafile record count exceeds the limit.");
  }
}

function checkedMultiply(left: number, right: number, subject: string): number {
  if (
    !Number.isSafeInteger(left) ||
    !Number.isSafeInteger(right) ||
    left < 0 ||
    right < 0 ||
    (right !== 0 && left > Number.MAX_SAFE_INTEGER / right)
  ) {
    throw new MetafileParseError("invalid-record-size", `${subject} exceeds safe integer bounds.`);
  }
  const result = left * right;
  if (!Number.isSafeInteger(result)) {
    throw new MetafileParseError("invalid-record-size", `${subject} exceeds safe integer bounds.`);
  }
  return result;
}

function checkedAdd(left: number, right: number, subject: string): number {
  if (
    !Number.isSafeInteger(left) ||
    !Number.isSafeInteger(right) ||
    left < 0 ||
    right < 0 ||
    left > Number.MAX_SAFE_INTEGER - right
  ) {
    throw new MetafileParseError("invalid-record-size", `${subject} exceeds safe integer bounds.`);
  }
  return left + right;
}
