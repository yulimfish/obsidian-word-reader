import type { MetafileRecord, MetafileRecordSet } from "./metafileRecords";
import { MetafileDecodeError } from "./metafileScene";
import type {
  MetafileColor, MetafileLogicalFont, MetafilePathCommand, MetafilePoint, MetafileScene,
  MetafileSceneOperation, MetafileSceneWarning, MetafileStroke,
  MetafileTextOperation, MetafileTransform, MetafileViewBox,
} from "./metafileScene";

interface PenObject { readonly kind: "pen"; readonly stroke: MetafileStroke | null }
interface FontObject { readonly kind: "font"; readonly family: string; readonly attributes: MetafileLogicalFont }
interface BrushObject { readonly kind: "brush"; readonly style: number; readonly color: MetafileColor }
type WmfObject = PenObject | FontObject | BrushObject;
interface State {
  position: MetafilePoint;
  origin: MetafilePoint;
  extent: MetafilePoint;
  pen: PenObject;
  font: FontObject | null;
  brush: BrushObject | null;
  backgroundMode: "transparent" | "opaque";
  backgroundColor: MetafileColor;
  textColor: MetafileColor;
  textAlign: number;
}

const BLACK: MetafileColor = Object.freeze({ red: 0, green: 0, blue: 0, alpha: 1 });
const WHITE: MetafileColor = Object.freeze({ red: 255, green: 255, blue: 255, alpha: 1 });
const STOCK_PEN: PenObject = { kind: "pen", stroke: { color: BLACK, width: 1, lineCap: "round", lineJoin: "round", miterLimit: 10 } };
const MAX_TEXT_BYTES = 8192;
const YIELD_RECORDS = 128;

/** Bounded, inert WMF playback; no font lookup, drawing API or payload execution.
 * The preview host's canonical initial context maps a logical window onto the
 * placeable bounds (or the first window extent). Coordinates remain logical in
 * operations, with a per-operation affine mapping into that fixed viewport.
 * Device/font metrics that cannot be recovered from WMF remain explicit gaps.
 */
export async function decodeWmfScene(
  records: MetafileRecordSet,
  options: { isCancelled: () => boolean; yieldControl: () => Promise<void> },
): Promise<MetafileScene> {
  const first = records.records[0];
  if (!first) throw new MetafileDecodeError("invalid-record", 0, 0, "Missing WMF EOF");
  if (records.format !== "wmf") fail(first, "invalid-format", "WMF decoder requires WMF records");
  if (options.isCancelled()) fail(first, "cancelled", "WMF decoding cancelled");
  const capacity = records.declaredObjectTableSize;
  if (capacity > records.limits.maxObjects) fail(first, "max-objects", "Declared object table exceeds maxObjects");
  const objects: (WmfObject | null)[] = Array.from({ length: capacity }, () => null);
  let liveObjects = 0;
  const bounds = records.bounds;
  let viewport: MetafileViewBox = { x: 0, y: 0, width: bounds ? bounds.right - bounds.left : 1, height: bounds ? bounds.bottom - bounds.top : 1 };
  if (viewport.width <= 0 || viewport.height <= 0) fail(first, "invalid-window", "Empty placeable bounds");
  let viewportEstablished = bounds !== null;
  let state: State = {
    position: { x: 0, y: 0 }, origin: { x: bounds?.left ?? 0, y: bounds?.top ?? 0 },
    extent: { x: viewport.width, y: viewport.height },
    pen: STOCK_PEN, font: null, brush: null, backgroundMode: "opaque",
    backgroundColor: WHITE, textColor: BLACK, textAlign: 0,
  };
  const stack: State[] = [];
  const operations: MetafileSceneOperation[] = [];
  const requiredFonts = new Set<string>();
  const warnings: MetafileSceneWarning[] = [];

  function checkEmission(record: MetafileRecord): void {
    if (operations.length >= records.limits.maxSceneOps) fail(record, "max-scene-ops", "Operation exceeds maxSceneOps");
  }
  function create(record: MetafileRecord, object: WmfObject): void {
    const slot = objects.indexOf(null);
    if (slot < 0) fail(record, "invalid-create", `Object table capacity ${capacity} is full`);
    if (liveObjects >= records.limits.maxObjects) fail(record, "max-objects", "Live object count exceeds maxObjects");
    objects[slot] = object;
    liveObjects += 1;
  }

  for (let index = 0; index < records.records.length; index += 1) {
    const record = records.records[index];
    if (options.isCancelled()) fail(record, "cancelled", "WMF decoding cancelled");
    if (index > 0 && index % YIELD_RECORDS === 0) {
      await options.yieldControl();
      if (options.isCancelled()) fail(record, "cancelled", "WMF decoding cancelled");
    }
    const payload = record.payload;
    const view = new DataView(payload.buffer, payload.byteOffset, payload.byteLength);
    switch (record.type) {
      case 0:
        exact(record, 0);
        break;
      case 0x102: {
        exactPaddedWord(record, view, 4);
        const mode = scalar(record, view);
        if (mode !== 1 && mode !== 2) fail(record, "unsupported-background-mode", `Background mode ${mode}`);
        state.backgroundMode = mode === 1 ? "transparent" : "opaque";
        break;
      }
      case 0x201:
      case 0x209: {
        exact(record, 4);
        const color = readColor(record, view, 0);
        if (record.type === 0x201) state.backgroundColor = color;
        else state.textColor = color;
        break;
      }
      case 0x12e: {
        exactPaddedWord(record, view, 4);
        const alignment = scalar(record, view);
        if ((alignment & ~0x11f) !== 0 || (alignment & 6) === 4 || (alignment & 0x18) === 0x10) {
          fail(record, "unsupported-text-alignment", `Invalid alignment flags 0x${alignment.toString(16)}`);
        }
        state.textAlign = alignment;
        break;
      }
      case 0x20b:
        exact(record, 4);
        state.origin = readYX(view);
        break;
      case 0x20c: {
        exact(record, 4);
        const extent = readYX(view);
        if (extent.x === 0 || extent.y === 0) fail(record, "invalid-window", "Window extent cannot be zero");
        if (!viewportEstablished) {
          viewport = { x: 0, y: 0, width: Math.abs(extent.x), height: Math.abs(extent.y) };
          viewportEstablished = true;
        }
        state.extent = extent;
        break;
      }
      case 0x214:
        exact(record, 4);
        state.position = readYX(view);
        break;
      case 0x213: {
        exact(record, 4);
        checkEmission(record);
        const point = readYX(view);
        operations.push({ kind: "path", commands: [{ type: "moveTo", point: state.position }, { type: "lineTo", point }], fill: null, stroke: state.pen.stroke, transform: transform(state, viewport) });
        state.position = point;
        viewportEstablished = true;
        break;
      }
      case 0x2fa:
        create(record, readPen(record, view));
        break;
      case 0x2fc:
        exact(record, 8);
        create(record, readBrush(record, view));
        break;
      case 0x2fb:
        create(record, readFont(record, view));
        break;
      case 0x12d: {
        exact(record, 2);
        const handle = view.getUint16(0, true);
        const object = objects[handle];
        if (!object) fail(record, "invalid-selection", `Handle ${handle} is uncreated or deleted`);
        if (object.kind === "pen") state.pen = object;
        else if (object.kind === "brush") state.brush = object;
        else state.font = object;
        break;
      }
      case 0x1f0: {
        exact(record, 2);
        const handle = view.getUint16(0, true);
        const object = objects[handle];
        if (!object) fail(record, "invalid-delete", `Handle ${handle} is uncreated or deleted`);
         // GDI cannot delete an object selected in the current or saved DC.
         if (state.pen === object || state.font === object || state.brush === object || stack.some((saved) => saved.pen === object || saved.font === object || saved.brush === object)) {
           fail(record, "invalid-delete", `Handle ${handle} remains selected in a current or saved state`);
         }
        objects[handle] = null;
        liveObjects -= 1;
        break;
      }
      case 0x1e:
        exact(record, 0);
        if (stack.length >= records.limits.maxStack) fail(record, "max-stack", "Saved state exceeds maxStack");
        stack.push({ ...state });
        break;
      case 0x127: {
        exact(record, 2);
        const target = view.getInt16(0, true);
        const slot = target < 0 ? stack.length + target : target - 1;
        if (target === 0 || slot < 0 || slot >= stack.length) fail(record, "invalid-restore", `Invalid saved-state target ${target}`);
        const saved = stack[slot];
        if ((saved.font && !objects.includes(saved.font)) || (saved.brush && !objects.includes(saved.brush)) || (saved.pen !== STOCK_PEN && !objects.includes(saved.pen))) {
          fail(record, "invalid-selection", "Saved state refers to a deleted object");
        }
        state = { ...saved };
        stack.length = slot;
        break;
      }
      case 0xa32: {
        checkEmission(record);
        const decoded = readText(record, view, state, viewport);
        operations.push(decoded.operation);
        requiredFonts.add(decoded.operation.fontFamily);
        warnings.push(...decoded.warnings);
        if (decoded.nextPosition) state.position = decoded.nextPosition;
        viewportEstablished = true;
        break;
      }
      case 0x324: {
        minimum(record, 6);
        const count = view.getUint16(0, true);
        exact(record, 2 + count * 4);
        if (count < 2) fail(record, "invalid-record", "A polygon needs at least two points");
        if (!state.brush) fail(record, "invalid-selection", "Polygon requires a selected brush");
        const commands: MetafilePathCommand[] = [];
        for (let index = 0; index < count; index++) {
          const point = { x: view.getInt16(2 + index * 4, true), y: view.getInt16(4 + index * 4, true) };
          commands.push({ type: index === 0 ? "moveTo" : "lineTo", point });
        }
        commands.push({ type: "closePath" });
        checkEmission(record);
        operations.push({ kind: "path", commands, fill: state.brush.color, stroke: state.pen.stroke, transform: transform(state, viewport) });
        viewportEstablished = true;
        break;
      }
      case 0x626: {
        minimum(record, 4);
        const escape = view.getUint16(0, true);
        if (escape !== 15) fail(record, "unsupported-escape", `Escape ${escape} is not an inert comment`);
        const count = view.getUint16(2, true);
        exact(record, 4 + count + count % 2);
        warnings.push({ kind: "opaque-comment", byteLength: count, recordOffset: record.offset, recordType: record.type });
        break;
      }
      default:
        fail(record, "unsupported-record", "Unsupported WMF opcode");
    }
  }
  return { viewBox: viewport, operations, requiredFonts: [...requiredFonts], warnings };
}

function fail(record: MetafileRecord, code: string, message: string): never {
  throw new MetafileDecodeError(code, record.offset, record.type, message);
}
function minimum(record: MetafileRecord, size: number): void {
  if (record.payload.byteLength < size) fail(record, "invalid-record", `Payload needs at least ${size} bytes`);
}
function exact(record: MetafileRecord, size: number): void {
  if (record.payload.byteLength !== size) fail(record, "invalid-record", `Expected ${size} payload bytes, received ${record.payload.byteLength}`);
}
function exactPaddedWord(record: MetafileRecord, view: DataView, size: number): void {
  if (record.payload.byteLength !== size - 2 && record.payload.byteLength !== size) {
    fail(record, "invalid-record", `Expected ${size - 2} or ${size} payload bytes, received ${record.payload.byteLength}`);
  }
  if (record.payload.byteLength === size && view.getUint16(size - 2, true) !== 0) {
    fail(record, "invalid-record", "Extended scalar high word must be zero");
  }
}

function scalar(record: MetafileRecord, view: DataView): number {
  return view.getUint16(0, true);
}
function readYX(view: DataView): MetafilePoint {
  return { x: view.getInt16(2, true), y: view.getInt16(0, true) };
}
function readColor(record: MetafileRecord, view: DataView, offset: number): MetafileColor {
  const highColor = view.getUint8(offset + 3);
  if (highColor !== 0 && highColor !== 2) fail(record, "unsupported-color", "Palette-relative COLORREF is unsupported");
  if (highColor === 2 && view.getUint16(offset, true) !== 0 && view.getUint16(offset, true) !== 255) fail(record, "unsupported-color", "Palette index COLORREF is unsupported");
  return { red: view.getUint8(offset), green: view.getUint8(offset + 1), blue: view.getUint8(offset + 2), alpha: 1 };
}
function transform(state: State, viewport: MetafileViewBox): MetafileTransform {
  const a = viewport.width / state.extent.x;
  const d = viewport.height / state.extent.y;
  return { a, b: 0, c: 0, d, e: state.origin.x === 0 ? 0 : -state.origin.x * a, f: state.origin.y === 0 ? 0 : -state.origin.y * d };
}
function readPen(record: MetafileRecord, view: DataView): PenObject {
  exact(record, 10);
  const style = view.getUint16(0, true);
  if (style !== 0 && style !== 5) fail(record, "unsupported-pen", `Pen style 0x${style.toString(16)} requires unsupported dash/geometry semantics`);
  const width = view.getInt16(2, true);
  if (width < 0) fail(record, "unsupported-pen", `Negative pen width ${width}`);
  const color = readColor(record, view, 6);
  return { kind: "pen", stroke: style === 5 ? null : { color, width, lineCap: "round", lineJoin: "round", miterLimit: 10 } };
}

function readBrush(record: MetafileRecord, view: DataView): BrushObject {
  const style = view.getUint16(0, true);
  if (style !== 0 && style !== 1) fail(record, "unsupported-brush", `Brush style ${style} is unsupported`);
  const highColor = view.getUint8(4);
  if (highColor !== 0) fail(record, "unsupported-color", "Palette-relative COLORREF is unsupported");
  return {
    kind: "brush",
    style,
    color: { red: view.getUint8(2), green: view.getUint8(3), blue: view.getUint8(4), alpha: 1 },
  };
}
function readFont(record: MetafileRecord, view: DataView): FontObject {
  exact(record, 50);
  const attributes: MetafileLogicalFont = {
    height: view.getInt16(0, true), width: view.getInt16(2, true),
    escapement: view.getInt16(4, true), orientation: view.getInt16(6, true), weight: view.getInt16(8, true),
    italic: view.getUint8(10) === 1, underline: view.getUint8(11) === 1, strikeOut: view.getUint8(12) === 1,
    charset: view.getUint8(13), outPrecision: view.getUint8(14), clipPrecision: view.getUint8(15), quality: view.getUint8(16), pitchAndFamily: view.getUint8(17),
  };
  if (attributes.weight < 0 || attributes.weight > 1000 || [10, 11, 12].some((offset) => view.getUint8(offset) > 1)) {
    fail(record, "invalid-font", "Invalid font weight or Boolean attribute");
  }
  if (attributes.outPrecision > 10 || attributes.quality > 6 || (attributes.clipPrecision & ~0xff) !== 0) {
    fail(record, "unsupported-font", "Unsupported LOGFONT precision/quality");
  }
  const nameBytes = record.payload.subarray(18, 50);
  const end = nameBytes.indexOf(0);
  if (end < 0) fail(record, "invalid-font", "Font face name is not null terminated");
  if (end === 0) fail(record, "font-metrics-required", "Empty font name requires a device font mapper");
  const faceBytes = nameBytes.subarray(0, end);
  // LOGFONTA face names are ANSI, not a symbol-font character run. ASCII names
  // are invariant; required non-ASCII Chinese names use charset 134 / CP936.
  const family = faceBytes.every((byte) => byte < 128)
    ? String.fromCharCode(...faceBytes)
    : decodeBytes(record, faceBytes, attributes.charset).text;
   if (family === null || Array.from(family).some((character) => character.charCodeAt(0) < 32 || character.charCodeAt(0) === 127)) fail(record, "invalid-font", "Invalid font face name encoding");
  return { kind: "font", family, attributes };
}

interface DecodedBytes { readonly text: string | null; readonly codePage: 42 | 1252 | 1253 | 936; readonly byteLengths: readonly number[] }
function decodeBytes(record: MetafileRecord, bytes: Uint8Array, charset: number, family?: string): DecodedBytes {
  // Symbol and MT Extra use the font's private-use Unicode mapping.
  if (charset === 2 && (family === "Symbol" || family === "MT Extra") && bytes.every((byte) =>
    [0x2b, 0x2d, 0x3d, 0x3c, 0x3e, 0xa3, 0xa5, 0xb9, 0xc8, 0xd6, 0x7b, 0x7d].includes(byte),
  )) {
    const symbolMap: Record<number, string> = {
      0x2b: "+", 0x2d: "−", 0x3d: "=", 0x3c: "<", 0x3e: ">",
      0xa3: "≤", 0xa5: "∞", 0xb9: "≠", 0xc8: "∈", 0xd6: "÷",
    };
    return {
      text: Array.from(bytes, (byte) => symbolMap[byte] ?? String.fromCharCode(byte)).join(""),
      codePage: 42,
      byteLengths: Array.from(bytes, () => 1),
    };
  }
  if (charset === 2 && family === "MT Extra" && bytes.every((byte) =>
    [0x36, 0x37, 0x38, 0x64, 0x7b, 0x7d].includes(byte),
  )) {
    const extraMap: Record<number, string> = { 0x36: "⎛", 0x37: "⎝", 0x38: "⎞", 0x64: "‖", 0x7b: "{", 0x7d: "}" };
    return {
      text: Array.from(bytes, (byte) => extraMap[byte] ?? "").join(""),
      codePage: 42,
      byteLengths: Array.from(bytes, () => 1),
    };
  }
  // Unknown or locale-dependent charsets have no guessed fallback.
  if (charset === 2) return { text: null, codePage: 42, byteLengths: Array.from(bytes, () => 1) };
  if (charset !== 0 && charset !== 1 && charset !== 134 && charset !== 161) fail(record, "unsupported-charset", `Charset ${charset} has no verified local decoder`);
  if (charset === 161) {
    const greekMap: Record<number, string> = { 0xf6: "φ" };
    if (bytes.some((byte) => byte < 0x20 || byte >= 0x7f && greekMap[byte] === undefined)) {
      fail(record, "unsupported-charset", "Unverified Greek charset byte");
    }
    return {
      text: Array.from(bytes, (byte) => greekMap[byte] ?? String.fromCharCode(byte)).join(""),
      codePage: 1253,
      byteLengths: Array.from(bytes, () => 1),
    };
  }
  const label = charset === 134 ? "gbk" : "windows-1252";
  if (charset === 1 && bytes.some((byte) => byte >= 0x80) && family !== "Symbol") fail(record, "unsupported-charset", "Default charset non-ASCII text has no verified local decoder");
  let decoder: TextDecoder;
  try { decoder = new TextDecoder(label, { fatal: true }); }
  catch { fail(record, "unsupported-charset", `Runtime does not support ${label}`); }
  const byteLengths: number[] = [];
  const characters: string[] = [];
  for (let offset = 0; offset < bytes.length;) {
    const byte = bytes[offset];
    const length = charset === 134 && byte >= 0x81 && byte <= 0xfe ? 2 : 1;
    if (offset + length > bytes.length) fail(record, "invalid-encoding", `${label} incomplete byte sequence at ${offset}`);
    try {
      const character = decoder.decode(bytes.subarray(offset, offset + length));
       const code = character.charCodeAt(0);
       if ([...character].length !== 1 || code < 32 || (code >= 127 && code <= 159) || code === 0xfffd) {
        fail(record, "invalid-encoding", `${label} unresolved glyph at byte ${offset}`);
      }
      characters.push(character);
    } catch (error) {
      if (error instanceof MetafileDecodeError) throw error;
      fail(record, "invalid-encoding", `${label} invalid byte sequence at ${offset}`);
    }
    byteLengths.push(length);
    offset += length;
  }
  return { text: characters.join(""), codePage: charset === 134 ? 936 : 1252, byteLengths };
}

function readText(record: MetafileRecord, view: DataView, state: State, viewport: MetafileViewBox): {
  operation: MetafileTextOperation; warnings: MetafileSceneWarning[]; nextPosition: MetafilePoint | null;
} {
  minimum(record, 8);
  const count = view.getInt16(4, true);
  if (count < 0 || count > MAX_TEXT_BYTES) fail(record, "invalid-text-length", `Text byte count ${count} exceeds ${MAX_TEXT_BYTES}`);
  // Some equation writers store ETO flags in the high word of a 32-bit field.
  const flags = record.payload.byteLength >= 10 && view.getUint16(6, true) === 0 &&
    (view.getUint16(8, true) & ~0x86) === 0
    ? 0
    : view.getUint16(6, true);
  if (flags & 0x10) fail(record, "font-glyph-gap", `Glyph-index run requires exact font mapping (${state.font?.family ?? "unselected font"})`);
  if ((flags & ~0x86) !== 0) fail(record, "unsupported-text-options", `Unsupported text flags 0x${flags.toString(16)}`);
  const start = flags & 6 ? 16 : 8;
  const end = start + count + count % 2;
  minimum(record, end);
  const remaining = record.payload.byteLength - end;
  if (remaining !== 0 && remaining !== count * 2) fail(record, "invalid-advances", `Expected zero or ${count} byte advances, received ${remaining / 2}`);
  const font = state.font;
  if (!font) fail(record, "font-metrics-required", "Text requires an explicitly selected font");
  const bytes = record.payload.subarray(start, start + count);
  const decoded = decodeBytes(record, bytes, font.attributes.charset, font.family);
  const sourceAdvances = remaining === 0 ? undefined : Array.from({ length: count }, (_, index) => view.getInt16(end + index * 2, true));
  const advances = sourceAdvances ? combineAdvances(sourceAdvances, decoded.byteLengths) : undefined;
  let rect: MetafileViewBox | undefined;
  if (flags & 6) {
    const left = view.getInt16(8, true); const top = view.getInt16(10, true);
    const right = view.getInt16(12, true); const bottom = view.getInt16(14, true);
    if (right < left || bottom < top) fail(record, "invalid-record", "Inverted text rectangle");
    rect = { x: left, y: top, width: right - left, height: bottom - top };
  }
  const point = state.textAlign & 1 ? state.position : readYX(view);
  const warnings: MetafileSceneWarning[] = [];
  if (decoded.text === null && count > 0) warnings.push({ kind: "font-glyph-gap", fontName: font.family, charset: font.attributes.charset, byteCodes: Array.from(bytes), recordOffset: record.offset, recordType: record.type });
  const hasExplicitAdvances = sourceAdvances !== undefined;
  const hasEquationTextLayout = font.attributes.charset === 2 && font.attributes.height !== 0 ||
    (font.attributes.height < 0 && (state.textAlign & 1) === 0 && decoded.text !== null &&
      decoded.text.length > 0 && font.attributes.outPrecision !== 0);
  if (font.attributes.height === 0) fail(record, "font-metrics-required", "Text requires a non-zero font height");
  if ((font.attributes.width !== 0 && !hasExplicitAdvances) ||
    (font.attributes.height >= 0 && font.attributes.charset !== 2) ||
    (!hasExplicitAdvances && !hasEquationTextLayout)) {
    warnings.push({ kind: "font-metrics-required", fontName: font.family, recordOffset: record.offset, recordType: record.type });
  }
  const style = {
    kind: "text" as const, x: point.x, y: point.y,
    fontFamily: font.family, fontSize: Math.abs(font.attributes.height), logicalFont: font.attributes,
     color: state.textColor, direction: flags & 0x80 || state.textAlign & 0x100 ? "rtl" as const : "ltr" as const,
    advances, sourceAdvances, sourceBytes: Array.from(bytes),
    transform: transform(state, viewport), textAlign: state.textAlign, textOptions: flags,
    backgroundMode: state.backgroundMode, backgroundColor: state.backgroundColor,
    clipRect: flags & 4 ? rect : undefined, opaqueRect: flags & 2 ? rect : undefined,
  };
  const operation: MetafileTextOperation = decoded.text === null
    ? { ...style, text: null, encoding: "font-specific", codePage: 42 }
    : {
      ...style,
      text: sourceAdvances !== undefined && (state.textAlign & 1) !== 0
        ? decoded.text.split(" ").join(String.fromCharCode(160))
        : decoded.text,
      encoding: "unicode",
      codePage: decoded.codePage,
    };
  let nextPosition: MetafilePoint | null = null;
  if (state.textAlign & 1) {
    if (!advances || font.attributes.escapement !== 0 || font.attributes.orientation !== 0) {
      fail(record, "font-metrics-required", `Current-position update requires known advances and unrotated font ${font.family}`);
    }
    const width = advances.reduce((sum, value) => sum + value, 0);
    const horizontal = state.textAlign & 6;
    nextPosition = horizontal === 6 ? point : { x: point.x + (horizontal === 2 ? -width : width), y: point.y };
  }
  return { operation, warnings, nextPosition };
}
function combineAdvances(source: readonly number[], lengths: readonly number[]): number[] {
  let offset = 0;
  return lengths.map((length) => {
    let advance = 0;
    for (let index = 0; index < length; index += 1) advance += source[offset++];
    return advance;
  });
}
