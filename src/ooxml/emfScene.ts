import type { MetafileRecord, MetafileRecordSet } from "./metafileRecords";
import { MetafileDecodeError } from "./metafileScene";
import type {
  MetafileClipRegion, MetafileColor, MetafileLogicalFont, MetafilePoint,
  MetafileScene, MetafileSceneOperation, MetafileSceneWarning, MetafileStroke,
  MetafileTextOperation, MetafileTransform, MetafileViewBox,
} from "./metafileScene";

interface Pen { readonly kind: "pen"; readonly stroke: MetafileStroke | null }
interface Brush { readonly kind: "brush"; readonly color: MetafileColor | null }
interface Font {
  readonly kind: "font";
  readonly family: string;
  readonly attributes: MetafileLogicalFont;
  readonly metadata?: MetafileTextOperation["fontMetadata"];
}
type EmfObject = Pen | Brush | Font;
interface State {
  position: MetafilePoint;
  brushOrigin: MetafilePoint;
  pen: Pen;
  brush: Brush;
  font: Font | null;
  transform: MetafileTransform;
  miterLimit: number;
  backgroundMode: "transparent" | "opaque";
  backgroundColor: MetafileColor;
  textColor: MetafileColor;
  textAlign: number;
  fillMode: "evenOdd" | "nonZero";
  stretchMode: number;
  clip: MetafileClipRegion | null;
  clipDepth: number;
}
const BLACK: MetafileColor = { red: 0, green: 0, blue: 0, alpha: 1 };
const WHITE: MetafileColor = { red: 255, green: 255, blue: 255, alpha: 1 };
const IDENTITY: MetafileTransform = { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 };
const BLACK_PEN: Pen = { kind: "pen", stroke: { color: BLACK, width: 1, lineCap: "round", lineJoin: "round", miterLimit: 10 } };
const WHITE_BRUSH: Brush = { kind: "brush", color: WHITE };
const MAX_TEXT_UNITS = 8192;
const MAX_REGION_RECTS = 4096;
const MAX_CLIP_DEPTH = 128;

/** Inert MS-EMF scene decoding only. No drawing/font/device/profile APIs.
 * Layouts: MS-EMF 2.2.5, 2.2.16, 2.2.24, 2.2.28, 2.3.2, 2.3.3,
 * 2.3.5.8/33 and 2.3.11/12. Unknown playback semantics fail closed.
 */
export async function decodeEmfScene(
  records: MetafileRecordSet,
  options: { isCancelled: () => boolean; yieldControl: () => Promise<void> },
): Promise<MetafileScene> {
  const first = records.records[0];
  if (!first) throw new MetafileDecodeError("invalid-record", 0, 0, "Missing EMF header");
  if (records.format !== "emf") fail(first, "invalid-format", "EMF decoder requires EMF records");
  if (options.isCancelled()) fail(first, "cancelled", "EMF decoding cancelled");
  const capacity = records.declaredObjectTableSize;
  // nHandles includes reserved slot 0; maxObjects excludes it (T4 contract).
  if (capacity < 1 || capacity - 1 > records.limits.maxObjects) fail(first, "max-objects", "Declared handle table exceeds maxObjects");
  const objects: (EmfObject | null)[] = Array.from({ length: capacity }, () => null);
  let liveObjects = 0;
  const bounds = records.bounds;
  if (!bounds) fail(first, "invalid-window", "Missing EMF bounds");
  // EMF Header Bounds are inclusive-inclusive device pixels, unlike region
  // rectangles (exclusive right/bottom). Keep device origin in the viewBox.
  const viewBox = { x: bounds.left, y: bounds.top, width: bounds.right - bounds.left + 1, height: bounds.bottom - bounds.top + 1 };
  let state: State = {
    position: { x: 0, y: 0 }, brushOrigin: { x: 0, y: 0 }, pen: BLACK_PEN,
    brush: WHITE_BRUSH, font: null, transform: IDENTITY, miterLimit: 10,
    backgroundMode: "opaque", backgroundColor: WHITE, textColor: BLACK,
    textAlign: 0, fillMode: "evenOdd", stretchMode: 1, clip: null, clipDepth: 0,
  };
  const stack: State[] = [];
  const groups: MetafileRecord[] = [];
  const operations: MetafileSceneOperation[] = [];
  const warnings: MetafileSceneWarning[] = [];
  const requiredFonts = new Set<string>();

  function checkEmission(record: MetafileRecord, count = 1): void {
    if (count > records.limits.maxSceneOps - operations.length) fail(record, "max-scene-ops", "Operation exceeds maxSceneOps");
  }
  function create(record: MetafileRecord, handle: number, object: EmfObject): void {
    if (handle === 0 || handle >= capacity || objects[handle]) fail(record, "invalid-create", `Handle ${handle} is reserved, occupied or outside table ${capacity}`);
    if (liveObjects >= records.limits.maxObjects) fail(record, "max-objects", "Live object count exceeds maxObjects");
    objects[handle] = object; liveObjects += 1;
  }
  function setClip(record: MetafileRecord, region: MetafileClipRegion | null, depth: number): void {
    if (depth > MAX_CLIP_DEPTH) fail(record, "max-clip-depth", `Clip expression exceeds ${MAX_CLIP_DEPTH} levels`);
    checkEmission(record);
    operations.push({ kind: "clip", action: "region", region });
    state.clip = region; state.clipDepth = depth;
  }

  for (let index = 0; index < records.records.length; index += 1) {
    const record = records.records[index];
    if (options.isCancelled()) fail(record, "cancelled", "EMF decoding cancelled");
    if (index > 0 && index % 128 === 0) {
      await options.yieldControl();
      if (options.isCancelled()) fail(record, "cancelled", "EMF decoding cancelled");
    }
    const view = new DataView(record.payload.buffer, record.payload.byteOffset, record.payload.byteLength);
    switch (record.type) {
      case 1:
        minimum(record, 80);
        if (view.getUint32(60, true) !== 0) fail(record, "unsupported-color-management", "Header palette requires device color mapping");
        break;
      case 14:
        if (groups.length) fail(groups[groups.length - 1], "invalid-group", "Begin group has no matching end group");
        break;
      case 13:
      case 27:
        exact(record, 8);
        if (record.type === 13) state.brushOrigin = point(view, 0);
        else state.position = point(view, 0);
        break;
      case 18: {
        exact(record, 4); const mode = view.getUint32(0, true);
        if (mode !== 1 && mode !== 2) fail(record, "unsupported-background-mode", `Background mode ${mode}`);
        state.backgroundMode = mode === 1 ? "transparent" : "opaque";
        break;
      }
      case 19: {
        exact(record, 4); const mode = view.getUint32(0, true);
        if (mode !== 1 && mode !== 2) fail(record, "unsupported-fill-mode", `Polygon fill mode ${mode}`);
        state.fillMode = mode === 1 ? "evenOdd" : "nonZero";
        break;
      }
      case 20:
        exact(record, 4);
        if (view.getUint32(0, true) !== 13) fail(record, "unsupported-rop2", "Non-copy raster operation requires device pixel semantics");
        break;
      case 21: {
        exact(record, 4); const mode = view.getUint32(0, true);
        if (mode < 1 || mode > 4) fail(record, "unsupported-stretch-mode", `Stretch mode ${mode}`);
        state.stretchMode = mode;
        break;
      }
      case 22: {
        exact(record, 4); const align = view.getUint32(0, true);
        if ((align & ~0x11f) !== 0 || (align & 6) === 4 || (align & 0x18) === 0x10) fail(record, "unsupported-text-alignment", `Text alignment ${align}`);
        state.textAlign = align;
        break;
      }
      case 23:
      case 121:
        throw new MetafileDecodeError("unsupported-color-management", record.offset, record.type, "Color adjustment/profile matching has no verified local device color mapping");
      case 24:
      case 25: {
        exact(record, 4); const color = readColor(record, view, 0);
        if (record.type === 24) state.textColor = color;
        else state.backgroundColor = color;
        break;
      }
      case 26: {
        exact(record, 8);
        if (!state.clip) break; // The default clip has no application region to move.
        if (state.clipDepth >= MAX_CLIP_DEPTH) fail(record, "max-clip-depth", "Clip offset exceeds bounded expression depth");
        checkEmission(record);
        const offset = point(view, 0); const t = state.transform;
        const deviceOffset = exactClipOffset(record, t, offset);
        setClip(record, { kind: "offset", region: state.clip, ...deviceOffset }, state.clipDepth + 1);
        break;
      }
      case 33:
        exact(record, 0);
        if (stack.length >= records.limits.maxStack) fail(record, "max-stack", "Saved DC exceeds maxStack");
        stack.push({ ...state });
        break;
      case 34: {
        exact(record, 4); const target = view.getInt32(0, true);
        const slot = target < 0 ? stack.length + target : target - 1;
        if (target === 0 || slot < 0 || slot >= stack.length) fail(record, "invalid-restore", `Invalid saved-state target ${target}`);
        const saved = stack[slot];
        if (saved.clip !== state.clip) {
          checkEmission(record);
          operations.push({ kind: "clip", action: "region", region: saved.clip });
        }
        state = { ...saved }; stack.length = slot;
        break;
      }
      case 35:
      case 36: {
        exact(record, record.type === 35 ? 24 : 28);
        const mode = record.type === 35 ? 4 : view.getUint32(24, true);
        if (mode < 1 || mode > 4) fail(record, "unsupported-transform-mode", `World transform mode ${mode}`);
        const supplied = mode === 1 ? IDENTITY : readTransform(record, view);
        const result = mode === 1 ? IDENTITY : mode === 2 ? multiply(state.transform, supplied) : mode === 3 ? multiply(supplied, state.transform) : supplied;
        if (!Object.values(result).every(Number.isFinite)) fail(record, "invalid-transform", "World transform product is nonfinite");
        state.transform = result;
        break;
      }
      case 37: {
        exact(record, 4); const handle = view.getUint32(0, true);
        if (handle & 0x80000000) {
          const stock = handle & 0x7fffffff;
          if (stock <= 5) state.brush = stockBrush(record, stock);
          else if (stock >= 6 && stock <= 8) state.pen = stock === 7 ? BLACK_PEN : { kind: "pen", stroke: stock === 8 ? null : { ...BLACK_PEN.stroke!, color: WHITE } };
          else if ([10, 11, 12, 13, 14, 16, 17].includes(stock)) state.font = null; // Device font mapper remains unresolved.
          else fail(record, "unsupported-stock-object", `Stock handle ${stock} needs unsupported device semantics`);
          break;
        }
        const object = objects[handle];
        if (!object) fail(record, "invalid-selection", `Handle ${handle} is uncreated or deleted`);
        if (object.kind === "pen") state.pen = object;
        else if (object.kind === "brush") state.brush = object;
        else state.font = object;
        break;
      }
      case 38:
        create(record, readHandle(record, view), readPen(record, view));
        break;
      case 39:
        create(record, readHandle(record, view), readBrush(record, view));
        break;
      case 40: {
        exact(record, 4); const handle = view.getUint32(0, true); const object = objects[handle];
        if (!object) fail(record, "invalid-delete", `Handle ${handle} is stock, uncreated or deleted`);
        const selected = (dc: State) => dc.pen === object || dc.brush === object || dc.font === object;
        if (selected(state) || stack.some(selected)) fail(record, "invalid-delete", `Handle ${handle} remains selected in a current or saved DC`);
        objects[handle] = null; liveObjects -= 1;
        break;
      }
      case 54: {
        exact(record, 8); checkEmission(record); const next = point(view, 0);
        operations.push({ kind: "path", commands: [{ type: "moveTo", point: state.position }, { type: "lineTo", point: next }], fill: null,
          stroke: state.pen.stroke ? { ...state.pen.stroke, miterLimit: state.miterLimit } : null, transform: state.transform });
        state.position = next;
        break;
      }
      case 58: {
        exact(record, 4); const limit = view.getUint32(0, true);
        if (limit < 1) fail(record, "invalid-miter-limit", "Miter limit must be at least one");
        state.miterLimit = limit;
        break;
      }
      case 70: {
        minimum(record, 4); const size = view.getUint32(0, true);
        if (Math.ceil((4 + size) / 4) * 4 !== record.payload.length) fail(record, "invalid-comment", "Comment DataSize exceeds or disagrees with envelope");
        const identifier = size >= 4 ? view.getUint32(4, true) : 0;
        if (identifier !== 0x43494447) {
          warnings.push({ kind: "opaque-comment", byteLength: size, recordOffset: record.offset, recordType: record.type });
          break;
        }
        if (size < 8) fail(record, "invalid-comment", "Truncated GDIC subtype");
        const subtype = view.getUint32(8, true);
        if (subtype === 2) {
          if (size < 28) fail(record, "invalid-comment", "Truncated begin-group metadata");
          const rectangle = readRect(record, view, 12, "invalid-comment");
          const count = view.getUint32(28, true);
          if (count > MAX_TEXT_UNITS || size !== 28 + count * 2) fail(record, "invalid-comment", "Group description exceeds declared DataSize");
          let description = "";
          if (count) {
            if (view.getUint16(30 + count * 2, true) !== 0) fail(record, "invalid-comment", "Group description is not terminated");
            description = unicode(record, record.payload.subarray(32, 30 + count * 2));
          }
          if (groups.length >= records.limits.maxStack) fail(record, "max-stack", "Group nesting exceeds maxStack");
          groups.push(record);
          warnings.push({ kind: "group-metadata", action: "begin", depth: groups.length, rectangle, description, recordOffset: record.offset, recordType: record.type });
        } else if (subtype === 3) {
          if (size !== 8) fail(record, "invalid-comment", "End-group metadata has unexpected fields");
          if (!groups.length) fail(record, "invalid-group", "End group has no matching begin group");
          warnings.push({ kind: "group-metadata", action: "end", depth: groups.length, recordOffset: record.offset, recordType: record.type });
          groups.pop();
        } else warnings.push({ kind: "opaque-comment", byteLength: size, recordOffset: record.offset, recordType: record.type });
        break;
      }
      case 75: {
        minimum(record, 8); checkEmission(record);
        const size = view.getUint32(0, true); const mode = view.getUint32(4, true);
        if (mode < 1 || mode > 5) fail(record, "unsupported-clip-mode", `Region mode ${mode}`);
        exact(record, 8 + size);
        if (size === 0) {
          if (mode !== 5) fail(record, "invalid-region", "Only RGN_COPY permits omitted region data");
          setClip(record, null, 0); break;
        }
        const depth = mode === 5 ? 1 : state.clipDepth + 1;
        if (depth > MAX_CLIP_DEPTH) fail(record, "max-clip-depth", "Clip combine exceeds bounded depth");
        const right: MetafileClipRegion = { kind: "rectangles", rectangles: readRegion(record, view, size), transform: state.transform };
        const modes = ["intersect", "union", "xor", "exclude"] as const;
        setClip(record, mode === 5 ? right : { kind: "combine", mode: modes[mode - 1], left: state.clip, right }, depth);
        break;
      }
      case 82:
        create(record, readHandle(record, view), readFont(record, view));
        break;
      case 84:
      case 97: {
        minimum(record, record.type === 84 ? 68 : 32);
        const mode = view.getUint32(16, true);
        if (mode !== 2) fail(record, "unsupported-graphics-mode", `Text graphics mode ${mode} requires unverified device metrics`);
        const count = record.type === 84 ? 1 : view.getUint32(28, true);
        const start = record.type === 84 ? 28 : 32;
        if (count === 0 || count > MAX_TEXT_UNITS) fail(record, "invalid-text-length", `Polytext count ${count} exceeds bounded run count`);
        checkEmission(record, count);
        minimum(record, start + count * 40);
        const dataStart = start + count * 40;
        let totalUnits = 0;
        for (let run = 0; run < count; run += 1) {
          totalUnits += view.getUint32(start + run * 40 + 8, true);
          if (totalUnits > MAX_TEXT_UNITS) fail(record, "invalid-text-length", `Combined text exceeds ${MAX_TEXT_UNITS} UTF16 units`);
        }
        for (let run = 0; run < count; run += 1) {
          if (options.isCancelled()) fail(record, "cancelled", "EMF text decoding cancelled");
          if (run > 0 && run % 128 === 0) {
            await options.yieldControl();
            if (options.isCancelled()) fail(record, "cancelled", "EMF text decoding cancelled");
          }
          const decoded = readText(record, view, start + run * 40, dataStart, state);
          operations.push(decoded.operation);
          requiredFonts.add(decoded.operation.fontFamily);
          warnings.push(...decoded.warnings);
          if (decoded.nextPosition) state.position = decoded.nextPosition;
        }
        break;
      }
      default:
        fail(record, "unsupported-record", "Unsupported EMF playback opcode");
    }
  }
  return { viewBox, operations, requiredFonts: [...requiredFonts], warnings };
}

function fail(record: MetafileRecord, code: string, message: string): never {
  throw new MetafileDecodeError(code, record.offset, record.type, message);
}
function minimum(record: MetafileRecord, size: number): void {
  if (record.payload.length < size) fail(record, "invalid-record", `Payload needs at least ${size} bytes`);
}
function exact(record: MetafileRecord, size: number): void {
  if (record.payload.length !== size) fail(record, "invalid-record", `Expected ${size} payload bytes, got ${record.payload.length}`);
}
function point(view: DataView, offset: number): MetafilePoint {
  return { x: view.getInt32(offset, true), y: view.getInt32(offset + 4, true) };
}
function exactClipOffset(record: MetafileRecord, transform: MetafileTransform, offset: MetafilePoint): MetafilePoint {
  const x = exactScaledClipOffset(record, transform.a, offset.x, transform.c, offset.y);
  const y = exactScaledClipOffset(record, transform.b, offset.x, transform.d, offset.y);
  return { x: x / 1024, y: y / 1024 };
}

// Compute the exact scaled sum as a short floating-point expansion. The
// error-free transforms preserve product/sum low bits; an offset-time matrix
// coefficient itself need not lie on the downstream 1/1024 grid.
function exactScaledClipOffset(
  record: MetafileRecord,
  firstCoefficient: number,
  firstDisplacement: number,
  secondCoefficient: number,
  secondDisplacement: number,
): number {
  let expansion: number[] = [];
  for (const [low, high] of [
    scaledClipOffsetProduct(record, firstCoefficient, firstDisplacement),
    scaledClipOffsetProduct(record, secondCoefficient, secondDisplacement),
  ]) {
    expansion = growExpansion(expansion, low);
    expansion = growExpansion(expansion, high);
  }
  const nonzero = expansion.filter((component) => component !== 0);
  if (nonzero.length === 0) return 0;
  if (nonzero.length !== 1 || !Number.isSafeInteger(nonzero[0])) {
    fail(record, "unsupported-clip-offset", "Mapped clip offset is not an exact 1/1024 device vector");
  }
  return nonzero[0];
}

const MAX_CLIP_OFFSET_TERM = Number.MAX_SAFE_INTEGER;
const MAX_CLIP_OFFSET_COEFFICIENT = 2 ** 43;
const MIN_CLIP_OFFSET_PRODUCT = 2 ** -900;

function scaledClipOffsetProduct(record: MetafileRecord, coefficient: number, displacement: number): readonly [number, number] {
  if (!Number.isFinite(coefficient) || !Number.isSafeInteger(displacement) || Math.abs(displacement) > 0x80000000) {
    fail(record, "unsupported-clip-offset", "Clip offset input is outside the finite Int32 arithmetic bounds");
  }
  if (coefficient === 0 || displacement === 0) return [0, 0];

  // A nonzero signed Int32 offset has magnitude >= 1, so coefficients at or
  // above 2^43 cannot produce a Number-safe 1/1024-scaled product. This bound
  // also keeps Dekker splitting well inside finite binary64 range.
  if (Math.abs(coefficient) >= MAX_CLIP_OFFSET_COEFFICIENT) {
    fail(record, "unsupported-clip-offset", "Clip offset product exceeds exact integer bounds");
  }
  const scaledDisplacement = displacement * 1024;
  const roundedProduct = coefficient * scaledDisplacement;
  // Dekker's residual is error-free when its intermediate products do not
  // underflow. Values below this conservative floor are not proven here.
  if (!Number.isFinite(roundedProduct) || Math.abs(roundedProduct) < MIN_CLIP_OFFSET_PRODUCT) {
    fail(record, "unsupported-clip-offset", "Clip offset product is outside the proven binary arithmetic range");
  }
  const [low, high] = twoProduct(coefficient, scaledDisplacement);
  if (!Number.isFinite(low) || !Number.isFinite(high) ||
      high > MAX_CLIP_OFFSET_TERM || high < -MAX_CLIP_OFFSET_TERM ||
      (high === MAX_CLIP_OFFSET_TERM && low > 0) ||
      (high === -MAX_CLIP_OFFSET_TERM && low < 0)) {
    fail(record, "unsupported-clip-offset", "Clip offset product exceeds exact integer bounds");
  }
  return [low, high];
}

// Dekker two-product: high + low is the exact product for the bounded finite
// operands above, without relying on decimal string conversion or BigInt.
function twoProduct(first: number, second: number): readonly [number, number] {
  const product = first * second;
  const splitter = 134217729; // 2^27 + 1, for binary64's 53-bit significand.
  const firstSplit = splitter * first;
  const firstHigh = firstSplit - (firstSplit - first);
  const firstLow = first - firstHigh;
  const secondSplit = splitter * second;
  const secondHigh = secondSplit - (secondSplit - second);
  const secondLow = second - secondHigh;
  const error1 = product - firstHigh * secondHigh;
  const error2 = error1 - firstLow * secondHigh;
  const error3 = error2 - firstHigh * secondLow;
  return [firstLow * secondLow - error3, product];
}

// Knuth TwoSum returns (sum, error) with sum + error exactly equal to a + b.
function twoSum(first: number, second: number): readonly [number, number] {
  const sum = first + second;
  const secondVirtual = sum - first;
  const firstVirtual = sum - secondVirtual;
  const secondRoundoff = second - secondVirtual;
  const firstRoundoff = first - firstVirtual;
  return [sum, firstRoundoff + secondRoundoff];
}

// Grow a non-overlapping expansion by one value, retaining every exact low
// part. This routine is used only for four bounded products per EMR record.
function growExpansion(expansion: readonly number[], value: number): number[] {
  let accumulator = value;
  const result: number[] = [];
  for (const component of expansion) {
    const [sum, error] = twoSum(accumulator, component);
    if (error !== 0) result.push(error);
    accumulator = sum;
  }
  if (accumulator !== 0 || result.length === 0) result.push(accumulator);
  return result;
}
function readHandle(record: MetafileRecord, view: DataView): number {
  minimum(record, 4); return view.getUint32(0, true);
}
function readColor(record: MetafileRecord, view: DataView, offset: number): MetafileColor {
  if (view.getUint8(offset + 3) !== 0) fail(record, "unsupported-color", "Palette-relative COLORREF needs device color mapping");
  return { red: view.getUint8(offset), green: view.getUint8(offset + 1), blue: view.getUint8(offset + 2), alpha: 1 };
}
function readTransform(record: MetafileRecord, view: DataView): MetafileTransform {
  const values = Array.from({ length: 6 }, (_, index) => view.getFloat32(index * 4, true));
  if (!values.every(Number.isFinite)) fail(record, "invalid-transform", "Nonfinite XForm component");
  const [a, b, c, d, e, f] = values;
  return { a, b, c, d, e, f };
}
// Column-form affine composition. GDI LEFTMULTIPLY uses row-form matrices,
// hence current * supplied here; RIGHTMULTIPLY reverses these operands.
function multiply(left: MetafileTransform, right: MetafileTransform): MetafileTransform {
  return {
    a: left.a * right.a + left.c * right.b, b: left.b * right.a + left.d * right.b,
    c: left.a * right.c + left.c * right.d, d: left.b * right.c + left.d * right.d,
    e: left.a * right.e + left.c * right.f + left.e, f: left.b * right.e + left.d * right.f + left.f,
  };
}
function readPen(record: MetafileRecord, view: DataView): Pen {
  exact(record, 20);
  const style = view.getUint32(4, true); const width = view.getInt32(8, true);
  if ((style !== 0 && style !== 5) || width < 0) fail(record, "unsupported-pen", `Pen style ${style}/width ${width} needs unsupported geometry or dash semantics`);
  const color = readColor(record, view, 16);
  return { kind: "pen", stroke: style === 5 ? null : { color, width, lineCap: "round", lineJoin: "round", miterLimit: 10 } };
}
function readBrush(record: MetafileRecord, view: DataView): Brush {
  exact(record, 16); const style = view.getUint32(4, true);
  if (style !== 0 && style !== 1) fail(record, "unsupported-brush", `Brush style ${style} needs unsupported pattern semantics`);
  return { kind: "brush", color: style === 1 ? null : readColor(record, view, 8) };
}
function stockBrush(record: MetafileRecord, index: number): Brush {
  if (index === 0) return WHITE_BRUSH;
  if (index === 4) return { kind: "brush", color: BLACK };
  if (index === 5) return { kind: "brush", color: null };
  fail(record, "unsupported-stock-object", `Stock brush ${index} has unverified device dither semantics`);
}
function unicode(record: MetafileRecord, bytes: Uint8Array): string {
  try { return new TextDecoder("utf-16le", { fatal: true, ignoreBOM: true }).decode(bytes); }
  catch { fail(record, "invalid-encoding", "Invalid UTF16-LE sequence"); }
}
function fixedString(record: MetafileRecord, view: DataView, start: number, units: number, requireTerminator = false): string {
  let count = 0;
  while (count < units && view.getUint16(start + count * 2, true) !== 0) count += 1;
  if (requireTerminator && count === units) fail(record, "invalid-font", "Font face is not terminated");
  return unicode(record, record.payload.subarray(start, start + count * 2));
}
function readFont(record: MetafileRecord, view: DataView): Font {
  minimum(record, 96);
  if (record.payload.length !== 96 && record.payload.length !== 324) fail(record, "unsupported-font", "Variable font design-vector layout is not supported");
  const attributes: MetafileLogicalFont = {
    height: view.getInt32(4, true), width: view.getInt32(8, true), escapement: view.getInt32(12, true), orientation: view.getInt32(16, true), weight: view.getInt32(20, true),
    italic: view.getUint8(24) === 1, underline: view.getUint8(25) === 1, strikeOut: view.getUint8(26) === 1,
    charset: view.getUint8(27), outPrecision: view.getUint8(28), clipPrecision: view.getUint8(29), quality: view.getUint8(30), pitchAndFamily: view.getUint8(31),
  };
  if (attributes.weight < 0 || attributes.weight > 1000 || [24, 25, 26].some((offset) => view.getUint8(offset) > 1)) fail(record, "invalid-font", "Invalid font weight/Boolean attributes");
  const family = fixedString(record, view, 32, 32, true);
  if (!family || Array.from(family).some((c) => c.charCodeAt(0) < 32 || c.charCodeAt(0) === 127)) fail(record, "invalid-font", "Invalid/empty font face name");
  const metadata = record.payload.length === 324 ? {
    fullName: fixedString(record, view, 96, 64), style: fixedString(record, view, 224, 32),
    styleSize: view.getUint32(292, true), panose: Array.from(record.payload.subarray(312, 322)),
  } : undefined;
  return { kind: "font", family, attributes, metadata };
}
function readRect(record: MetafileRecord, view: DataView, offset: number, code: string): MetafileViewBox {
  const left = view.getInt32(offset, true); const top = view.getInt32(offset + 4, true);
  const right = view.getInt32(offset + 8, true); const bottom = view.getInt32(offset + 12, true);
  if (right < left || bottom < top) fail(record, code, "Inverted rectangle");
  return { x: left, y: top, width: right - left, height: bottom - top };
}
function readRegion(record: MetafileRecord, view: DataView, size: number): MetafileViewBox[] {
  if (size < 32) fail(record, "invalid-region", "Truncated RegionData header");
  const count = view.getUint32(16, true);
  if (view.getUint32(8, true) !== 32 || view.getUint32(12, true) !== 1 || count > MAX_REGION_RECTS ||
      size !== 32 + count * 16 || view.getUint32(20, true) !== count * 16) fail(record, "invalid-region", "Invalid RegionData size/type/count");
  const bounds = readRect(record, view, 24, "invalid-region");
  const rectangles: MetafileViewBox[] = [];
  for (let index = 0; index < count; index += 1) {
    const rect = readRect(record, view, 40 + index * 16, "invalid-region");
    if (rect.width === 0 || rect.height === 0 || rect.x < bounds.x || rect.y < bounds.y || rect.x + rect.width > bounds.x + bounds.width || rect.y + rect.height > bounds.y + bounds.height) fail(record, "invalid-region", "Rectangle escapes RegionData bounds");
    const previous = rectangles[index - 1];
    if (previous && (rect.y < previous.y || (rect.y === previous.y ? rect.height !== previous.height || rect.x < previous.x + previous.width : rect.y < previous.y + previous.height))) fail(record, "invalid-region", "Region rectangles overlap or are not ordered in bands");
    rectangles.push(rect);
  }
  return rectangles;
}

function readText(record: MetafileRecord, view: DataView, start: number, dataStart: number, state: State): {
  operation: MetafileTextOperation; warnings: MetafileSceneWarning[]; nextPosition: MetafilePoint | null;
} {
  const count = view.getUint32(start + 8, true); const flags = view.getUint32(start + 16, true);
  if (count > MAX_TEXT_UNITS) fail(record, "invalid-text-length", `Text count ${count} exceeds ${MAX_TEXT_UNITS}`);
  if (flags & 0x10) fail(record, "font-glyph-gap", `Glyph indices require exact font mapping (${state.font?.family ?? "device font"})`);
  if ((flags & ~0x86) !== 0) fail(record, "unsupported-text-options", `Text options 0x${flags.toString(16)}`);
  const font = state.font;
  if (!font) fail(record, "font-metrics-required", "Text requires an explicitly selected font");
  if (font.attributes.charset === 2) fail(record, "font-glyph-gap", `${font.family} Symbol charset requires exact local glyph mapping`);
  const stringOffset = view.getUint32(start + 12, true) - 8;
  if (stringOffset < dataStart || stringOffset % 2 !== 0 || stringOffset > record.payload.length || count * 2 > record.payload.length - stringOffset) fail(record, "invalid-text-offset", "UTF16 string escapes or overlaps record descriptor area");
  const bytes = record.payload.subarray(stringOffset, stringOffset + count * 2);
  const text = unicode(record, bytes);
  const dx = view.getUint32(start + 36, true);
  let sourceAdvances: number[] | undefined;
  if (dx !== 0) {
    const offset = dx - 8;
    if (offset < dataStart || dx % 4 !== 0 || offset > record.payload.length || count * 4 > record.payload.length - offset ||
        (offset < stringOffset + count * 2 && stringOffset < offset + count * 4)) fail(record, "invalid-advances", "Advance array escapes or overlaps string/descriptor area");
    sourceAdvances = Array.from({ length: count }, (_, index) => view.getUint32(offset + index * 4, true));
  }
  let cursor = 0;
  const advances = sourceAdvances ? Array.from(text, (character) => {
    let value = 0;
    for (let unit = 0; unit < character.length; unit += 1) value += sourceAdvances[cursor++];
    return value;
  }) : undefined;
  const rectangle = flags & 6 ? readRect(record, view, start + 20, "invalid-record") : undefined;
  const position = state.textAlign & 1 ? state.position : point(view, start);
  const operation: MetafileTextOperation = {
    kind: "text", text, encoding: "unicode", x: position.x, y: position.y,
    fontFamily: font.family, fontSize: Math.abs(font.attributes.height), logicalFont: font.attributes, fontMetadata: font.metadata,
    color: state.textColor, direction: flags & 0x80 || state.textAlign & 0x100 ? "rtl" : "ltr", advances, sourceAdvances,
    sourceBytes: Array.from(bytes), transform: state.transform, textAlign: state.textAlign, textOptions: flags,
    backgroundMode: state.backgroundMode, backgroundColor: state.backgroundColor,
    clipRect: flags & 4 ? rectangle : undefined, opaqueRect: flags & 2 ? rectangle : undefined,
  };
  const warnings: MetafileSceneWarning[] = [];
  if (font.attributes.height >= 0 || font.attributes.width !== 0 || !advances || font.metadata?.styleSize || font.metadata?.panose.some((value) => value !== 0)) warnings.push({ kind: "font-metrics-required", fontName: font.family, recordOffset: record.offset, recordType: record.type });
  let nextPosition: MetafilePoint | null = null;
  if (state.textAlign & 1) {
    if (!advances || font.attributes.escapement !== 0 || font.attributes.orientation !== 0) fail(record, "font-metrics-required", "Current-position update requires known advances and an unrotated font");
    const width = advances.reduce((sum, value) => sum + value, 0);
    const horizontal = state.textAlign & 6;
    nextPosition = horizontal === 6 ? position : { x: position.x + (horizontal === 2 ? -width : width), y: position.y };
  }
  return { operation, warnings, nextPosition };
}
