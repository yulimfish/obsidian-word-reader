import type { MetafilePathCommand, MetafileScene } from "./metafileScene";

const SVG_NS = "http://www.w3.org/2000/svg";
const SCALE = 1024;
const COORDINATE_DENOMINATOR = 1024 * 1024;
const MAX_COORDINATE = 2 ** 24;
const MAX_SOURCE_CORNER = 2 ** 25;
const MAX_EXACT_TERM = 2 ** 50;
const MAX_EXACT_SUM = 2 ** 51;
const MAX_OFFSETS = 128;
const MAX_OPERATIONS = 50_000;
const MAX_CLIP_DEPTH = 128;
const MAX_CLIP_NODES = 50_000;
const MAX_REGION_RECTS = 4_096;
const MAX_PATH_COMMANDS = 50_000;
const MAX_TEXT_UNITS = 8_192;
const MAX_SVG_NODES = 50_000;
const MAX_OUTPUT_BYTES = 16 * 1024 * 1024;
const MAX_VIEW_BOX_PIXELS = 16_000_000;
const MAX_NUMERIC_OPERATIONS = 1_000_000;

export type MetafileSerializeErrorCode =
  | "invalid-scene"
  | "max-clip-depth"
  | "max-clip-nodes"
  | "max-region-rectangles"
  | "max-svg-nodes"
  | "max-output-bytes"
  | "unsupported-clip-expression"
  | "font-glyph-gap"
  | "font-metrics-required"
  | "unsupported-text";

export class MetafileSerializeError extends Error {
  constructor(
    readonly code: MetafileSerializeErrorCode,
    readonly operationIndex: number,
    message: string,
  ) {
    super(message);
    this.name = "MetafileSerializeError";
  }
}

interface PointI {
  readonly x: number;
  readonly y: number;
}

interface BoxI {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly right: number;
  readonly bottom: number;
}

interface TransformI {
  readonly a: number;
  readonly b: number;
  readonly c: number;
  readonly d: number;
  readonly e: number;
  readonly f: number;
}

interface Quad {
  readonly points: readonly [PointI, PointI, PointI, PointI];
}

interface RawRectanglesRegion {
  readonly kind: "rectangles";
  readonly rectangles: readonly BoxI[];
  readonly transform: TransformI;
}

interface RawCombineRegion {
  readonly kind: "combine";
  readonly mode: "intersect" | "union" | "xor" | "exclude";
  readonly left: RawRegion | null;
  readonly right: RawRegion;
}

interface RawOffsetRegion {
  readonly kind: "offset";
  readonly region: RawRegion;
  readonly x: number;
  readonly y: number;
}

type RawRegion = RawRectanglesRegion | RawCombineRegion | RawOffsetRegion;

interface PreparedRectanglesRegion {
  readonly kind: "rectangles";
  readonly quads: readonly Quad[];
}

interface PreparedCombineRegion {
  readonly kind: "combine";
  readonly mode: "intersect" | "union" | "xor" | "exclude";
  readonly left: PreparedRegion | null;
  readonly right: PreparedRegion;
}

type PreparedRegion = PreparedRectanglesRegion | PreparedCombineRegion;

interface PathData {
  readonly commands: readonly MetafilePathCommand[];
  readonly points: readonly (PointI | null)[];
  readonly fillRule: "evenOdd" | "nonZero";
  readonly operationIndex: number;
}

interface PathPrimitive {
  readonly kind: "path";
  readonly path: PathData;
}

interface QuadPrimitive {
  readonly kind: "quad";
  readonly quad: Quad;
}

type Primitive = PathPrimitive | QuadPrimitive;

interface Literal {
  readonly primitive: Primitive;
  readonly negated: boolean;
  readonly operationIndex: number;
}

type Dnf = readonly (readonly Literal[])[];

interface ExpressionUnbounded {
  readonly kind: "unbounded";
  readonly depth: 0;
}

interface ExpressionPath {
  readonly kind: "path";
  readonly path: PathData;
  readonly depth: 1;
}

interface ExpressionRegion {
  readonly kind: "region";
  readonly region: PreparedRegion;
  readonly operationIndex: number;
  readonly depth: number;
}

interface ExpressionCombine {
  readonly kind: "combine";
  readonly mode: "intersect" | "union" | "xor" | "exclude";
  readonly left: ClipExpression;
  readonly right: ClipExpression;
  readonly operationIndex: number;
  readonly depth: number;
}

type ClipExpression = ExpressionUnbounded | ExpressionPath | ExpressionRegion | ExpressionCombine;

interface NormalizedColor {
  readonly red: number;
  readonly green: number;
  readonly blue: number;
  readonly alpha: number;
}

interface NormalizedStroke {
  readonly color: NormalizedColor;
  readonly width: number;
  readonly lineCap: "butt" | "round" | "square";
  readonly lineJoin: "bevel" | "miter" | "round";
  readonly miterLimit: number;
}

interface NormalizedPathOperation {
  readonly kind: "path";
  readonly commands: PathData;
  readonly fill: NormalizedColor | null;
  readonly stroke: NormalizedStroke | null;
  readonly transform: TransformI;
  readonly operationIndex: number;
}

interface NormalizedTextOperation {
  readonly kind: "text";
  readonly x: number;
  readonly y: number;
  readonly fontFamily: string;
  readonly fontSize: number;
  readonly color: NormalizedColor;
  readonly direction: "ltr" | "rtl";
  readonly text: string | null;
  readonly encoding: "unicode" | "font-specific";
  readonly advances: readonly number[] | null;
  readonly glyphPositions: readonly number[] | null;
  readonly transform: TransformI;
  readonly textAlign: number;
  readonly textOptions: number;
  readonly backgroundMode: "transparent" | "opaque" | null;
  readonly backgroundColor: NormalizedColor | null;
  readonly clipQuad: Quad | null;
  readonly clipSpecified: boolean;
  readonly opaqueQuad: Quad | null;
  readonly logicalFont: Readonly<Record<string, number | boolean>> | null;
  readonly operationIndex: number;
}

interface NormalizedClipOperation {
  readonly kind: "clip";
  readonly action: "set" | "intersect" | "exclude" | "reset" | "region";
  readonly path: PathData | null;
  readonly region: PreparedRegion | null;
  readonly fillRule: "evenOdd" | "nonZero" | null;
  readonly operationIndex: number;
}

type NormalizedOperation = NormalizedPathOperation | NormalizedTextOperation | NormalizedClipOperation;

interface NormalizedScene {
  readonly viewBox: BoxI;
  readonly viewBoxQuad: Quad;
  readonly operations: readonly NormalizedOperation[];
  readonly requiredFonts: readonly string[];
  readonly warnings: readonly NormalizedWarning[];
}

interface NormalizedWarning {
  readonly kind: string;
  readonly fontName?: string;
  readonly detail: string;
}

interface Counters {
  numericOperations: number;
  uniqueClipNodes: number;
  expandedClipNodes: number;
  regionRectangles: number;
  pathCommands: number;
  dnfTerms: number;
  literals: number;
  svgNodes: number;
  nextId: number;
}

interface SvgNode {
  readonly tag: "svg" | "defs" | "clipPath" | "path" | "rect" | "g" | "text" | "tspan";
  readonly attributes: readonly (readonly [string, string])[];
  readonly children: readonly (SvgNode | string)[];
}

const EMPTY_ATTRIBUTES: readonly (readonly [string, string])[] = [];
const UNBOUNDED: ExpressionUnbounded = { kind: "unbounded", depth: 0 };

function fail(
  code: MetafileSerializeErrorCode,
  operationIndex: number,
  message: string,
): never {
  throw new MetafileSerializeError(code, operationIndex, message);
}

function invalid(operationIndex: number, message: string): never {
  return fail("invalid-scene", operationIndex, message);
}

function unsupported(operationIndex: number, message: string): never {
  return fail("unsupported-clip-expression", operationIndex, message);
}

function plainRecord(value: unknown, operationIndex: number, label: string): Record<string, unknown> {
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    return invalid(operationIndex, `${label} must be a plain data object.`);
  }
  let prototype: object | null;
  let keys: (string | symbol)[];
  try {
    prototype = Object.getPrototypeOf(value) as object | null;
    keys = Reflect.ownKeys(value);
  } catch {
    return invalid(operationIndex, `${label} cannot be inspected safely.`);
  }
  if (prototype !== Object.prototype && prototype !== null) {
    return invalid(operationIndex, `${label} must not have a custom prototype.`);
  }
  const result: Record<string, unknown> = Object.create(null) as Record<string, unknown>;
  for (const key of keys) {
    if (typeof key !== "string") return invalid(operationIndex, `${label} contains a symbol field.`);
    let descriptor: PropertyDescriptor | undefined;
    try {
      descriptor = Object.getOwnPropertyDescriptor(value, key);
    } catch {
      return invalid(operationIndex, `${label} field cannot be inspected safely.`);
    }
    if (!descriptor || !("value" in descriptor)) {
      return invalid(operationIndex, `${label} must contain own data fields only.`);
    }
    result[key] = descriptor.value as unknown;
  }
  return result;
}

function exactFields(
  value: unknown,
  required: readonly string[],
  optional: readonly string[],
  operationIndex: number,
  label: string,
): Record<string, unknown> {
  const object = plainRecord(value, operationIndex, label);
  const allowed = new Set([...required, ...optional]);
  for (const key of Reflect.ownKeys(object)) {
    if (typeof key !== "string" || !allowed.has(key)) {
      invalid(operationIndex, `${label} contains an unknown field.`);
    }
  }
  for (const key of required) {
    if (!Object.prototype.hasOwnProperty.call(object, key)) {
      invalid(operationIndex, `${label} is missing a required field.`);
    }
  }
  return object;
}

function arrayValues(
  value: unknown,
  operationIndex: number,
  label: string,
  maximum: number,
  limitCode: MetafileSerializeErrorCode = "max-svg-nodes",
): unknown[] {
  if (!Array.isArray(value)) return invalid(operationIndex, `${label} must be an array.`);
  let length: number;
  try {
    length = value.length;
  } catch {
    return invalid(operationIndex, `${label} cannot be inspected safely.`);
  }
  if (!Number.isSafeInteger(length) || length < 0 || length > maximum) {
    fail(limitCode, operationIndex, `${label} exceeds its bounded length.`);
  }
  let keys: (string | symbol)[];
  try {
    keys = Reflect.ownKeys(value);
  } catch {
    return invalid(operationIndex, `${label} cannot be inspected safely.`);
  }
  if (keys.length !== length + 1 || !keys.includes("length")) {
    return invalid(operationIndex, `${label} must not contain extra properties.`);
  }
  const result: unknown[] = [];
  for (let index = 0; index < length; index += 1) {
    let descriptor: PropertyDescriptor | undefined;
    try {
      descriptor = Object.getOwnPropertyDescriptor(value, String(index));
    } catch {
      return invalid(operationIndex, `${label} item cannot be inspected safely.`);
    }
    if (!descriptor || !("value" in descriptor)) {
      return invalid(operationIndex, `${label} must contain dense own data items.`);
    }
    result.push(descriptor.value as unknown);
  }
  return result;
}

function reserveNumeric(counters: Counters, operationIndex: number, count = 1): void {
  if (!Number.isSafeInteger(count) || count < 0 || counters.numericOperations > MAX_NUMERIC_OPERATIONS - count) {
    unsupported(operationIndex, "Numeric preflight exceeds the fixed work budget.");
  }
  counters.numericOperations += count;
}

function scaleCoordinate(value: unknown, operationIndex: number, counters: Counters): number {
  if (typeof value !== "number" || !Number.isFinite(value)) {
    return invalid(operationIndex, "A coordinate is not finite.");
  }
  reserveNumeric(counters, operationIndex);
  const scaled = value * SCALE;
  if (!Number.isInteger(scaled) || Math.abs(scaled) > MAX_COORDINATE) {
    unsupported(operationIndex, "A coordinate is outside the proven 1/1024 fixed-grid subset.");
  }
  return scaled;
}

function checkedIntegerSum(
  first: number,
  second: number,
  maximum: number,
  operationIndex: number,
  counters: Counters,
  message: string,
): number {
  reserveNumeric(counters, operationIndex);
  if (Math.abs(first) > maximum || Math.abs(second) > maximum) unsupported(operationIndex, message);
  const result = first + second;
  if (!Number.isSafeInteger(result) || Math.abs(result) > maximum) unsupported(operationIndex, message);
  return result;
}

function checkedIntegerProduct(
  first: number,
  second: number,
  maximum: number,
  operationIndex: number,
  counters: Counters,
  message: string,
): number {
  reserveNumeric(counters, operationIndex);
  if (Math.abs(first) > maximum || Math.abs(second) > maximum) unsupported(operationIndex, message);
  const result = first * second;
  if (!Number.isSafeInteger(result) || Math.abs(result) > maximum) unsupported(operationIndex, message);
  return result;
}

function checkedGeometrySum(
  first: number,
  second: number,
  operationIndex: number,
  counters: Counters,
): number {
  reserveNumeric(counters, operationIndex);
  if (Math.abs(first) > MAX_EXACT_TERM || Math.abs(second) > MAX_EXACT_TERM) {
    unsupported(operationIndex, "Transformed geometry exceeds the exact numerator bounds.");
  }
  const result = first + second;
  if (!Number.isSafeInteger(result) || Math.abs(result) > MAX_EXACT_SUM || Math.abs(result) > MAX_EXACT_TERM) {
    unsupported(operationIndex, "Transformed geometry exceeds the exact numerator bounds.");
  }
  return result;
}

function validatedBox(
  value: unknown,
  operationIndex: number,
  counters: Counters,
  label: string,
  positiveExtent: boolean,
): BoxI {
  const object = exactFields(value, ["x", "y", "width", "height"], [], operationIndex, label);
  const x = scaleCoordinate(object.x, operationIndex, counters);
  const y = scaleCoordinate(object.y, operationIndex, counters);
  const width = scaleCoordinate(object.width, operationIndex, counters);
  const height = scaleCoordinate(object.height, operationIndex, counters);
  if (width < 0 || height < 0 || (positiveExtent && (width === 0 || height === 0))) {
    return invalid(operationIndex, `${label} has invalid dimensions.`);
  }
  const right = checkedIntegerSum(x, width, MAX_SOURCE_CORNER, operationIndex, counters, `${label} corner exceeds exact integer bounds.`);
  const bottom = checkedIntegerSum(y, height, MAX_SOURCE_CORNER, operationIndex, counters, `${label} corner exceeds exact integer bounds.`);
  return { x, y, width, height, right, bottom };
}

function validatedTransform(
  value: unknown,
  operationIndex: number,
  counters: Counters,
  label: string,
): TransformI {
  const object = exactFields(value, ["a", "b", "c", "d", "e", "f"], [], operationIndex, label);
  return {
    a: scaleCoordinate(object.a, operationIndex, counters),
    b: scaleCoordinate(object.b, operationIndex, counters),
    c: scaleCoordinate(object.c, operationIndex, counters),
    d: scaleCoordinate(object.d, operationIndex, counters),
    e: scaleCoordinate(object.e, operationIndex, counters),
    f: scaleCoordinate(object.f, operationIndex, counters),
  };
}

function identityTransform(): TransformI {
  return { a: SCALE, b: 0, c: 0, d: SCALE, e: 0, f: 0 };
}

function exactNumber(value: unknown, operationIndex: number, label: string, minimum?: number, maximum?: number): number {
  if (typeof value !== "number" || !Number.isFinite(value) ||
      (minimum !== undefined && value < minimum) || (maximum !== undefined && value > maximum)) {
    return invalid(operationIndex, `${label} is outside its valid finite range.`);
  }
  return value;
}

function validatedColor(value: unknown, operationIndex: number, nullable = false): NormalizedColor | null {
  if (nullable && value === null) return null;
  const object = exactFields(value, ["red", "green", "blue", "alpha"], [], operationIndex, "Color");
  const channels = [object.red, object.green, object.blue];
  if (channels.some((channel) => typeof channel !== "number" || !Number.isInteger(channel) || channel < 0 || channel > 255)) {
    return invalid(operationIndex, "Color channels must be byte integers.");
  }
  const alpha = exactNumber(object.alpha, operationIndex, "Color alpha", 0, 1);
  return { red: object.red as number, green: object.green as number, blue: object.blue as number, alpha };
}

function validXmlString(value: string, operationIndex: number, label: string, maximum: number): string {
  if (value.length > maximum) fail("max-svg-nodes", operationIndex, `${label} exceeds its bounded length.`);
  for (let index = 0; index < value.length; index += 1) {
    const first = value.charCodeAt(index);
    if (first >= 0xd800 && first <= 0xdbff) {
      const second = value.charCodeAt(index + 1);
      if (index + 1 >= value.length || second < 0xdc00 || second > 0xdfff) return invalid(operationIndex, `${label} contains malformed Unicode.`);
      const codePoint = 0x10000 + ((first - 0xd800) << 10) + (second - 0xdc00);
      if (codePoint > 0x10ffff) return invalid(operationIndex, `${label} contains invalid XML text.`);
      index += 1;
      continue;
    }
    if (first >= 0xdc00 && first <= 0xdfff) return invalid(operationIndex, `${label} contains malformed Unicode.`);
    if (!(first === 0x9 || first === 0xa || first === 0xd || (first >= 0x20 && first <= 0xd7ff) || (first >= 0xe000 && first <= 0xfffd))) {
      return invalid(operationIndex, `${label} contains a character forbidden by XML 1.0.`);
    }
  }
  return value;
}

function validatedPathCommands(
  value: unknown,
  operationIndex: number,
  counters: Counters,
): PathData {
  const values = arrayValues(value, operationIndex, "Path commands", MAX_PATH_COMMANDS);
  if (counters.pathCommands > MAX_PATH_COMMANDS - values.length) {
    fail("max-svg-nodes", operationIndex, "Scene path commands exceed their aggregate bound.");
  }
  counters.pathCommands += values.length;
  const commands: MetafilePathCommand[] = [];
  const points: (PointI | null)[] = [];
  let hasCurrentSubpath = false;
  for (const item of values) {
    const discriminant = plainRecord(item, operationIndex, "Path command");
    const type = discriminant.type;
    if (type === "moveTo" || type === "lineTo") {
      const object = exactFields(item, ["type", "point"], [], operationIndex, "Path command");
      const pointObject = exactFields(object.point, ["x", "y"], [], operationIndex, "Path point");
      const point: PointI = {
        x: scaleCoordinate(pointObject.x, operationIndex, counters),
        y: scaleCoordinate(pointObject.y, operationIndex, counters),
      };
      if (type === "lineTo" && !hasCurrentSubpath) invalid(operationIndex, "A line command has no preceding move command.");
      if (type === "moveTo") hasCurrentSubpath = true;
      commands.push({ type, point: { x: point.x / SCALE, y: point.y / SCALE } });
      points.push(point);
    } else if (type === "cubicTo") {
      const object = exactFields(item, ["type", "control1", "control2", "point"], [], operationIndex, "Path command");
      if (!hasCurrentSubpath) invalid(operationIndex, "A cubic command has no preceding move command.");
      const readPoint = (pointValue: unknown): PointI => {
        const pointObject = exactFields(pointValue, ["x", "y"], [], operationIndex, "Path point");
        return {
          x: scaleCoordinate(pointObject.x, operationIndex, counters),
          y: scaleCoordinate(pointObject.y, operationIndex, counters),
        };
      };
      const control1 = readPoint(object.control1);
      const control2 = readPoint(object.control2);
      const point = readPoint(object.point);
      commands.push({
        type,
        control1: { x: control1.x / SCALE, y: control1.y / SCALE },
        control2: { x: control2.x / SCALE, y: control2.y / SCALE },
        point: { x: point.x / SCALE, y: point.y / SCALE },
      });
      points.push(control1, control2, point);
    } else if (type === "closePath") {
      exactFields(item, ["type"], [], operationIndex, "Path command");
      if (!hasCurrentSubpath) invalid(operationIndex, "A close command has no preceding move command.");
      commands.push({ type: "closePath" });
      points.push(null);
      hasCurrentSubpath = false;
    } else {
      invalid(operationIndex, "Unknown path command.");
    }
  }
  return { commands, points, fillRule: "nonZero", operationIndex };
}

function isClosedPath(path: PathData): boolean {
  let open = false;
  for (const command of path.commands) {
    if (command.type === "moveTo") {
      if (open) return false;
      open = true;
    } else if (command.type === "closePath") {
      if (!open) return false;
      open = false;
    }
  }
  return !open;
}

function validateRegion(
  value: unknown,
  operationIndex: number,
  counters: Counters,
  cache: WeakMap<object, RawRegion>,
  active: Set<object>,
  depth: number,
): RawRegion {
  if (depth > MAX_CLIP_DEPTH) fail("max-clip-depth", operationIndex, "Clip expression exceeds 128 levels.");
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    return invalid(operationIndex, "Clip region must be a non-null object.");
  }
  if (active.has(value)) return invalid(operationIndex, "Clip region contains a cycle.");
  const cached = cache.get(value);
  if (cached) return cached;
  counters.uniqueClipNodes += 1;
  if (counters.uniqueClipNodes > MAX_CLIP_NODES) fail("max-clip-nodes", operationIndex, "Clip graph exceeds 50,000 unique nodes.");
  active.add(value);
  const record = plainRecord(value, operationIndex, "Clip region");
  const kind = record.kind;
  let result: RawRegion;
  if (kind === "rectangles") {
    const object = exactFields(value, ["kind", "rectangles", "transform"], [], operationIndex, "Rectangle clip region");
    const transform = validatedTransform(object.transform, operationIndex, counters, "Rectangle clip transform");
    const rawRectangles = arrayValues(object.rectangles, operationIndex, "Region rectangles", MAX_REGION_RECTS, "max-region-rectangles");
    if (counters.regionRectangles > MAX_REGION_RECTS - rawRectangles.length) {
      fail("max-region-rectangles", operationIndex, "Scene region rectangles exceed 4,096 entries.");
    }
    counters.regionRectangles += rawRectangles.length;
    const rectangles: BoxI[] = [];
    for (const rectangle of rawRectangles) {
      rectangles.push(validatedBox(rectangle, operationIndex, counters, "Region rectangle", false));
    }
    result = { kind, rectangles, transform };
  } else if (kind === "combine") {
    const object = exactFields(value, ["kind", "mode", "left", "right"], [], operationIndex, "Combined clip region");
    if (object.mode !== "intersect" && object.mode !== "union" && object.mode !== "xor" && object.mode !== "exclude") {
      return invalid(operationIndex, "Clip combine mode is invalid.");
    }
    const leftValue = object.left;
    const left = leftValue === null ? null : validateRegion(leftValue, operationIndex, counters, cache, active, depth + 1);
    if (object.right === null) return invalid(operationIndex, "A combined clip's right region cannot be null.");
    const right = validateRegion(object.right, operationIndex, counters, cache, active, depth + 1);
    result = { kind, mode: object.mode, left, right };
  } else if (kind === "offset") {
    const object = exactFields(value, ["kind", "region", "x", "y"], [], operationIndex, "Offset clip region");
    if (object.region === null) return invalid(operationIndex, "An offset clip region cannot be null.");
    const region = validateRegion(object.region, operationIndex, counters, cache, active, depth + 1);
    const x = scaleCoordinate(object.x, operationIndex, counters);
    const y = scaleCoordinate(object.y, operationIndex, counters);
    result = { kind, region, x, y };
  } else {
    return invalid(operationIndex, "Unknown clip region discriminant.");
  }
  active.delete(value);
  cache.set(value, result);
  return result;
}

function sourceCorner(value: number, operationIndex: number): void {
  if (Math.abs(value) > MAX_SOURCE_CORNER) unsupported(operationIndex, "Source rectangle corner exceeds the fixed-grid bound.");
}

function transformIsSingular(transform: TransformI, operationIndex: number, counters: Counters): boolean {
  const first = checkedIntegerProduct(transform.a, transform.d, 2 ** 48, operationIndex, counters, "Captured matrix determinant exceeds exact bounds.");
  const second = checkedIntegerProduct(transform.b, transform.c, 2 ** 48, operationIndex, counters, "Captured matrix determinant exceeds exact bounds.");
  reserveNumeric(counters, operationIndex);
  const determinant = first - second;
  if (!Number.isSafeInteger(determinant) || Math.abs(determinant) > 2 ** 49) {
    unsupported(operationIndex, "Captured matrix determinant exceeds exact bounds.");
  }
  return determinant === 0;
}

function mappedCoordinate(
  x: number,
  y: number,
  transform: TransformI,
  offsetX: number,
  offsetY: number,
  operationIndex: number,
  counters: Counters,
): PointI {
  const xTermA = checkedIntegerProduct(transform.a, x, MAX_EXACT_TERM, operationIndex, counters, "Transformed x coordinate exceeds exact bounds.");
  const xTermC = checkedIntegerProduct(transform.c, y, MAX_EXACT_TERM, operationIndex, counters, "Transformed x coordinate exceeds exact bounds.");
  const xTermE = checkedIntegerProduct(transform.e, SCALE, MAX_EXACT_TERM, operationIndex, counters, "Transformed x coordinate exceeds exact bounds.");
  const xTermOffset = checkedIntegerProduct(offsetX, SCALE, MAX_EXACT_TERM, operationIndex, counters, "Transformed x offset exceeds exact bounds.");
  let xNumerator = checkedGeometrySum(xTermA, xTermC, operationIndex, counters);
  xNumerator = checkedGeometrySum(xNumerator, xTermE, operationIndex, counters);
  xNumerator = checkedGeometrySum(xNumerator, xTermOffset, operationIndex, counters);

  const yTermB = checkedIntegerProduct(transform.b, x, MAX_EXACT_TERM, operationIndex, counters, "Transformed y coordinate exceeds exact bounds.");
  const yTermD = checkedIntegerProduct(transform.d, y, MAX_EXACT_TERM, operationIndex, counters, "Transformed y coordinate exceeds exact bounds.");
  const yTermF = checkedIntegerProduct(transform.f, SCALE, MAX_EXACT_TERM, operationIndex, counters, "Transformed y coordinate exceeds exact bounds.");
  const yTermOffset = checkedIntegerProduct(offsetY, SCALE, MAX_EXACT_TERM, operationIndex, counters, "Transformed y offset exceeds exact bounds.");
  let yNumerator = checkedGeometrySum(yTermB, yTermD, operationIndex, counters);
  yNumerator = checkedGeometrySum(yNumerator, yTermF, operationIndex, counters);
  yNumerator = checkedGeometrySum(yNumerator, yTermOffset, operationIndex, counters);

  const deviceX = xNumerator / COORDINATE_DENOMINATOR;
  const deviceY = yNumerator / COORDINATE_DENOMINATOR;
  if (!Number.isFinite(deviceX) || !Number.isFinite(deviceY)) unsupported(operationIndex, "Transformed coordinate is not representable.");
  if (Number(Number(deviceX.toString()).toString()) !== deviceX || Number(Number(deviceY.toString()).toString()) !== deviceY) {
    unsupported(operationIndex, "Transformed coordinate cannot round-trip exactly.");
  }
  return { x: xNumerator, y: yNumerator };
}

function mapRectangle(
  box: BoxI,
  transform: TransformI,
  offsetX: number,
  offsetY: number,
  operationIndex: number,
  counters: Counters,
): Quad | null {
  const corners = [
    { x: box.x, y: box.y },
    { x: box.right, y: box.y },
    { x: box.right, y: box.bottom },
    { x: box.x, y: box.bottom },
  ] as const;
  corners.forEach((corner) => {
    sourceCorner(corner.x, operationIndex);
    sourceCorner(corner.y, operationIndex);
  });
  if (box.width === 0 || box.height === 0 || transformIsSingular(transform, operationIndex, counters)) return null;
  const mapped = corners.map((corner) => mappedCoordinate(corner.x, corner.y, transform, offsetX, offsetY, operationIndex, counters));
  const points = mapped.map((point) => ({ x: point.x / COORDINATE_DENOMINATOR, y: point.y / COORDINATE_DENOMINATOR })) as [PointI, PointI, PointI, PointI];
  if (points.some((point) => !Number.isFinite(point.x) || !Number.isFinite(point.y))) {
    unsupported(operationIndex, "Transformed corner is not finite.");
  }
  return { points };
}

function reserveExpandedClipNode(counters: Counters, operationIndex: number, amount = 1): void {
  if (!Number.isSafeInteger(amount) || amount < 0 || counters.expandedClipNodes > MAX_CLIP_NODES - amount) {
    fail("max-clip-nodes", operationIndex, "Expanded clip work exceeds 50,000 nodes.");
  }
  counters.expandedClipNodes += amount;
}

function prepareRegion(
  region: RawRegion | null,
  operationIndex: number,
  counters: Counters,
  active: Set<RawRegion>,
  depth: number,
  offsetDepth: number,
  offsetX: number,
  offsetY: number,
): PreparedRegion | null {
  if (region === null) return null;
  if (depth > MAX_CLIP_DEPTH) fail("max-clip-depth", operationIndex, "Clip expression exceeds 128 levels.");
  if (active.has(region)) invalid(operationIndex, "Clip region contains a cycle.");
  reserveExpandedClipNode(counters, operationIndex);
  active.add(region);
  let prepared: PreparedRegion;
  if (region.kind === "rectangles") {
    const quads: Quad[] = [];
    for (const rectangle of region.rectangles) {
      reserveExpandedClipNode(counters, operationIndex);
      const quad = mapRectangle(rectangle, region.transform, offsetX, offsetY, operationIndex, counters);
      if (quad) quads.push(quad);
    }
    prepared = { kind: "rectangles", quads };
  } else if (region.kind === "combine") {
    const left = prepareRegion(region.left, operationIndex, counters, active, depth + 1, offsetDepth, offsetX, offsetY);
    const right = prepareRegion(region.right, operationIndex, counters, active, depth + 1, offsetDepth, offsetX, offsetY);
    if (right === null) invalid(operationIndex, "A combined clip's right region cannot be null.");
    prepared = { kind: "combine", mode: region.mode, left, right };
  } else {
    const nextOffsetDepth = offsetDepth + 1;
    if (nextOffsetDepth > MAX_OFFSETS) unsupported(operationIndex, "Clip offset nesting exceeds 128 additions.");
    const nextX = checkedIntegerSum(offsetX, region.x, 2 ** 31, operationIndex, counters, "Accumulated clip offset exceeds exact bounds.");
    const nextY = checkedIntegerSum(offsetY, region.y, 2 ** 31, operationIndex, counters, "Accumulated clip offset exceeds exact bounds.");
    const child = prepareRegion(region.region, operationIndex, counters, active, depth + 1, nextOffsetDepth, nextX, nextY);
    if (child === null) invalid(operationIndex, "An offset clip's region cannot be null.");
    prepared = child;
  }
  active.delete(region);
  return prepared;
}

function validatePathFillRule(value: unknown, operationIndex: number): "evenOdd" | "nonZero" {
  if (value !== "evenOdd" && value !== "nonZero") return invalid(operationIndex, "Path fill rule is invalid.");
  return value;
}

function pathWithRule(path: PathData, fillRule: "evenOdd" | "nonZero"): PathData {
  return { ...path, fillRule };
}

function validateStroke(value: unknown, operationIndex: number, counters: Counters): NormalizedStroke | null {
  if (value === null) return null;
  const object = exactFields(value, ["color", "width", "lineCap", "lineJoin", "miterLimit"], [], operationIndex, "Stroke");
  const color = validatedColor(object.color, operationIndex);
  if (!color) return invalid(operationIndex, "Stroke color is missing.");
  const width = exactNumber(object.width, operationIndex, "Stroke width", 0);
  const miterLimit = exactNumber(object.miterLimit, operationIndex, "Stroke miter limit", 1);
  if (object.lineCap !== "butt" && object.lineCap !== "round" && object.lineCap !== "square") return invalid(operationIndex, "Stroke cap is invalid.");
  if (object.lineJoin !== "bevel" && object.lineJoin !== "miter" && object.lineJoin !== "round") return invalid(operationIndex, "Stroke join is invalid.");
  void counters;
  return { color, width, lineCap: object.lineCap, lineJoin: object.lineJoin, miterLimit };
}

function validateLogicalFont(value: unknown, operationIndex: number): Readonly<Record<string, number | boolean>> | null {
  if (value === undefined) return null;
  const keys = ["height", "width", "escapement", "orientation", "weight", "italic", "underline", "strikeOut", "charset", "outPrecision", "clipPrecision", "quality", "pitchAndFamily"] as const;
  const object = exactFields(value, keys, [], operationIndex, "Logical font");
  const result: Record<string, number | boolean> = Object.create(null) as Record<string, number | boolean>;
  for (const key of keys) {
    const entry = object[key];
    if (key === "italic" || key === "underline" || key === "strikeOut") {
      if (typeof entry !== "boolean") return invalid(operationIndex, "Logical font Boolean field is invalid.");
      result[key] = entry;
    } else if (typeof entry !== "number" || !Number.isSafeInteger(entry)) {
      return invalid(operationIndex, "Logical font numeric field is invalid.");
    } else {
      result[key] = entry;
    }
  }
  if ((result.weight as number) < 0 || (result.weight as number) > 1000) return invalid(operationIndex, "Logical font weight is invalid.");
  return result;
}

function validateTextMetadata(value: unknown, operationIndex: number): void {
  if (value === undefined) return;
  const object = exactFields(value, ["fullName", "style", "styleSize", "panose"], [], operationIndex, "Font metadata");
  if (typeof object.fullName !== "string" || typeof object.style !== "string") return invalid(operationIndex, "Font metadata strings are invalid.");
  validXmlString(object.fullName, operationIndex, "Font full name", MAX_TEXT_UNITS);
  validXmlString(object.style, operationIndex, "Font style", MAX_TEXT_UNITS);
  exactNumber(object.styleSize, operationIndex, "Font style size", 0, 0xffffffff);
  const panose = arrayValues(object.panose, operationIndex, "PANOSE values", 10);
  if (panose.length !== 10 || panose.some((entry) => typeof entry !== "number" || !Number.isInteger(entry) || entry < 0 || entry > 255)) {
    invalid(operationIndex, "PANOSE values must contain ten bytes.");
  }
}

function validateByteArray(value: unknown, operationIndex: number, label: string): readonly number[] | null {
  if (value === undefined) return null;
  const entries = arrayValues(value, operationIndex, label, MAX_TEXT_UNITS);
  if (entries.some((entry) => typeof entry !== "number" || !Number.isInteger(entry) || entry < 0 || entry > 255)) {
    return invalid(operationIndex, `${label} must contain byte integers.`);
  }
  return entries as number[];
}

function validateNumberArray(value: unknown, operationIndex: number, counters: Counters, label: string): readonly number[] | null {
  const scaled = validateScaledNumberArray(value, operationIndex, counters, label);
  return scaled?.map((entry) => entry / SCALE) ?? null;
}

function validateScaledNumberArray(value: unknown, operationIndex: number, counters: Counters, label: string): readonly number[] | null {
  if (value === undefined) return null;
  const entries = arrayValues(value, operationIndex, label, MAX_TEXT_UNITS);
  return entries.map((entry) => scaleCoordinate(entry, operationIndex, counters));
}

function validateTextOperation(
  value: unknown,
  operationIndex: number,
  counters: Counters,
): NormalizedTextOperation {
  const required = ["kind", "x", "y", "fontFamily", "fontSize", "color", "direction"];
  const optional = ["text", "encoding", "advances", "transform", "logicalFont", "codePage", "sourceBytes", "sourceAdvances", "textAlign", "textOptions", "backgroundMode", "backgroundColor", "clipRect", "opaqueRect", "fontMetadata"];
  const object = exactFields(value, required, optional, operationIndex, "Text operation");
  const xScaled = scaleCoordinate(object.x, operationIndex, counters);
  const x = xScaled / SCALE;
  const y = scaleCoordinate(object.y, operationIndex, counters) / SCALE;
  if (typeof object.fontFamily !== "string") return invalid(operationIndex, "Font family must be text.");
  const fontFamily = validXmlString(object.fontFamily, operationIndex, "Font family", MAX_TEXT_UNITS);
  const fontSize = exactNumber(object.fontSize, operationIndex, "Font size", 0);
  const color = validatedColor(object.color, operationIndex);
  if (!color) return invalid(operationIndex, "Text color is missing.");
  if (object.direction !== "ltr" && object.direction !== "rtl") return invalid(operationIndex, "Text direction is invalid.");
  let text: string | null = null;
  if (object.text !== undefined && object.text !== null) {
    if (typeof object.text !== "string") return invalid(operationIndex, "Text content must be a string or null.");
    text = validXmlString(object.text, operationIndex, "Text content", MAX_TEXT_UNITS);
    // C4 forbids whitespace-preservation attributes. Default SVG processing
    // retains isolated internal spaces, but trims/collapses boundary or repeated
    // spaces and normalizes XML control whitespace before applying x positions.
    // Other Unicode whitespace is outside this proven one-to-one subset.
    if (/^ | $| {2}|[^\S ]/u.test(text)) {
      fail("unsupported-text", operationIndex, "Text whitespace has no proven one-to-one SVG glyph mapping.");
    }
  }
  const encoding = object.encoding === undefined ? "unicode" : object.encoding;
  if (encoding !== "unicode" && encoding !== "font-specific") return invalid(operationIndex, "Text encoding is invalid.");
  if ((encoding === "font-specific") !== (text === null)) return invalid(operationIndex, "Font-specific text must not be interpreted as Unicode.");
  const scaledAdvances = validateScaledNumberArray(object.advances, operationIndex, counters, "Text advances");
  const advances = scaledAdvances?.map((entry) => entry / SCALE) ?? null;
  const transform = object.transform === undefined
    ? identityTransform()
    : validatedTransform(object.transform, operationIndex, counters, "Text transform");
  const logicalFont = validateLogicalFont(object.logicalFont, operationIndex);
  const sourceBytes = validateByteArray(object.sourceBytes, operationIndex, "Source text bytes");
  validateNumberArray(object.sourceAdvances, operationIndex, counters, "Source text advances");
  if (object.codePage !== undefined && (typeof object.codePage !== "number" || !Number.isInteger(object.codePage) || object.codePage < 0 || object.codePage > 0xffff)) {
    invalid(operationIndex, "Text code page is invalid.");
  }
  if (object.textAlign !== undefined && (typeof object.textAlign !== "number" || !Number.isInteger(object.textAlign) || object.textAlign < 0 || object.textAlign > 0xffff)) {
    invalid(operationIndex, "Text alignment is invalid.");
  }
  if (object.textOptions !== undefined && (typeof object.textOptions !== "number" || !Number.isInteger(object.textOptions) || object.textOptions < 0 || object.textOptions > 0xffff)) {
    invalid(operationIndex, "Text options are invalid.");
  }
  const textAlign = object.textAlign === undefined ? 0 : object.textAlign;
  const textOptions = object.textOptions === undefined ? 0 : object.textOptions;
  if (textAlign !== 0x18 && !(textAlign === 0x118 && object.direction === "rtl")) {
    fail("unsupported-text", operationIndex, `Text alignment 0x${textAlign.toString(16)} requires unsupported horizontal or vertical positioning.`);
  }
  if ((textOptions & ~0x86) !== 0) {
    fail("unsupported-text", operationIndex, `Text options 0x${textOptions.toString(16)} contain unsupported flags.`);
  }
  if ((textOptions & 0x80) !== 0 && object.direction !== "rtl") {
    fail("unsupported-text", operationIndex, "RTL text option disagrees with the decoded text direction.");
  }
  const backgroundMode = object.backgroundMode === undefined ? null : object.backgroundMode;
  if (backgroundMode !== null && backgroundMode !== "transparent" && backgroundMode !== "opaque") return invalid(operationIndex, "Text background mode is invalid.");
  const backgroundColor = object.backgroundColor === undefined ? null : validatedColor(object.backgroundColor, operationIndex);
  if (object.backgroundColor !== undefined && backgroundColor === null) return invalid(operationIndex, "Text background color is invalid.");
  const clipBox = object.clipRect === undefined ? null : validatedBox(object.clipRect, operationIndex, counters, "Text clip rectangle", false);
  const opaqueBox = object.opaqueRect === undefined ? null : validatedBox(object.opaqueRect, operationIndex, counters, "Opaque text rectangle", false);
  const clipQuad = clipBox === null ? null : mapRectangle(clipBox, transform, 0, 0, operationIndex, counters);
  const opaqueQuad = opaqueBox === null ? null : mapRectangle(opaqueBox, transform, 0, 0, operationIndex, counters);
  if (backgroundMode === "opaque" && opaqueQuad === null) {
    fail("unsupported-text", operationIndex, "Opaque text background requires an explicit representable source rectangle.");
  }
  if ((textOptions & 4) !== 0 && clipBox === null) invalid(operationIndex, "Clipped text options require a clip rectangle.");
  if ((textOptions & 2) !== 0 && opaqueBox === null) invalid(operationIndex, "Opaque text options require an opaque rectangle.");
  if (encoding === "font-specific" && (!sourceBytes || object.codePage !== 42)) {
    return invalid(operationIndex, "Font-specific runs require their original bytes and Symbol code page.");
  }
  if (opaqueQuad && (backgroundMode === "opaque" || (textOptions & 2) !== 0) && !backgroundColor) {
    return invalid(operationIndex, "Opaque text background requires its source color.");
  }
  if (logicalFont && ((logicalFont.escapement as number) !== 0 || (logicalFont.orientation as number) !== 0)) {
    fail("unsupported-text", operationIndex, "Nonzero logical-font escapement or orientation requires unsupported glyph rotation.");
  }
  validateTextMetadata(object.fontMetadata, operationIndex);
  const characters = text === null ? null : Array.from(text);
  if (characters !== null && advances && advances.length !== characters.length) {
    return invalid(operationIndex, "Text advances must correspond one-to-one with Unicode characters.");
  }
  let glyphPositions: readonly number[] | null = null;
  if (characters && characters.length > 0 && scaledAdvances) {
    const positions: number[] = [x];
    let position = xScaled;
    for (let index = 0; index < characters.length - 1; index += 1) {
      const advance = scaledAdvances[index];
      if (advance === undefined) return invalid(operationIndex, "A text advance is missing for a Unicode character.");
      position = checkedIntegerSum(position, advance, MAX_EXACT_TERM, operationIndex, counters, "Text glyph position exceeds exact integer bounds.");
      positions.push(position / SCALE);
    }
    glyphPositions = positions;
  }
  return {
    kind: "text", x, y, fontFamily, fontSize, color, direction: object.direction,
    text, encoding, advances, glyphPositions, transform, textAlign, textOptions, backgroundMode, backgroundColor,
    clipQuad, clipSpecified: clipBox !== null, opaqueQuad, logicalFont, operationIndex,
  };
}

function validateWarning(value: unknown, operationIndex: number, counters: Counters): NormalizedWarning {
  const object = plainRecord(value, operationIndex, "Scene warning");
  const kind = object.kind;
  if (typeof kind !== "string") return invalid(operationIndex, "Scene warning kind is invalid.");
  const fields: Readonly<Record<string, readonly [readonly string[], readonly string[]]>> = {
    "group-metadata": [["kind", "action", "depth", "recordOffset", "recordType"], ["rectangle", "description"]],
    "missing-font": [["kind", "fontName"], []],
    "unconfirmed-font": [["kind", "fontName"], []],
    "font-metrics-required": [["kind", "fontName", "recordOffset", "recordType"], []],
    "font-glyph-gap": [["kind", "fontName", "charset", "byteCodes", "recordOffset", "recordType"], []],
    "opaque-comment": [["kind", "byteLength", "recordOffset", "recordType"], []],
    "unsupported-record": [["kind", "recordOffset", "recordType"], []],
  };
  const shape = fields[kind];
  if (!shape) return invalid(operationIndex, "Unknown scene warning kind.");
  const [required, optional] = shape;
  const allowed = new Set([...required, ...optional]);
  for (const key of Object.keys(object)) if (!allowed.has(key)) invalid(operationIndex, "Scene warning contains an unknown field.");
  for (const key of required) if (!Object.prototype.hasOwnProperty.call(object, key)) invalid(operationIndex, "Scene warning is missing a required field.");
  let fontName: string | undefined;
  if (object.fontName !== undefined) {
    if (typeof object.fontName !== "string") return invalid(operationIndex, "Warning font name is invalid.");
    fontName = validXmlString(object.fontName, operationIndex, "Warning font name", MAX_TEXT_UNITS);
  }
  let detail = kind;
  if (kind === "group-metadata") {
    if ((object.action !== "begin" && object.action !== "end") || typeof object.depth !== "number" || !Number.isSafeInteger(object.depth) || object.depth < 0 ||
        typeof object.recordOffset !== "number" || !Number.isSafeInteger(object.recordOffset) || object.recordOffset < 0 ||
        typeof object.recordType !== "number" || !Number.isInteger(object.recordType) || object.recordType < 0) {
      return invalid(operationIndex, "Group warning fields are invalid.");
    }
    if (object.rectangle !== undefined) validatedBox(object.rectangle, operationIndex, counters, "Group warning rectangle", false);
    if (object.description !== undefined) {
      if (typeof object.description !== "string") return invalid(operationIndex, "Group warning description is invalid.");
      validXmlString(object.description, operationIndex, "Group warning description", MAX_TEXT_UNITS);
    }
    detail = `group-metadata:${object.action}`;
  } else if (kind === "font-glyph-gap") {
    if (!fontName || typeof object.charset !== "number" || !Number.isInteger(object.charset) || object.charset < 0 || object.charset > 255 ||
        typeof object.recordOffset !== "number" || !Number.isSafeInteger(object.recordOffset) || object.recordOffset < 0 ||
        typeof object.recordType !== "number" || !Number.isInteger(object.recordType) || object.recordType < 0) {
      return invalid(operationIndex, "Glyph-gap warning fields are invalid.");
    }
    validateByteArray(object.byteCodes, operationIndex, "Glyph-gap byte codes");
    detail = `font-glyph-gap:${fontName ?? ""}`;
  } else if (kind === "opaque-comment") {
    if (typeof object.byteLength !== "number" || !Number.isSafeInteger(object.byteLength) || object.byteLength < 0 ||
        typeof object.recordOffset !== "number" || !Number.isSafeInteger(object.recordOffset) || object.recordOffset < 0 ||
        typeof object.recordType !== "number" || !Number.isInteger(object.recordType) || object.recordType < 0) return invalid(operationIndex, "Opaque-comment warning fields are invalid.");
  } else if (kind === "unsupported-record") {
    if (typeof object.recordOffset !== "number" || !Number.isSafeInteger(object.recordOffset) || object.recordOffset < 0 ||
        typeof object.recordType !== "number" || !Number.isInteger(object.recordType) || object.recordType < 0) return invalid(operationIndex, "Unsupported-record warning fields are invalid.");
  } else if (kind === "font-metrics-required" || kind === "missing-font" || kind === "unconfirmed-font") {
    if (!fontName) return invalid(operationIndex, "Font warning is missing its name.");
    if (kind === "font-metrics-required" && (typeof object.recordOffset !== "number" || !Number.isSafeInteger(object.recordOffset) || object.recordOffset < 0 ||
        typeof object.recordType !== "number" || !Number.isInteger(object.recordType) || object.recordType < 0)) return invalid(operationIndex, "Font-metrics warning fields are invalid.");
  }
  return { kind, fontName, detail };
}

function validateOperation(
  value: unknown,
  operationIndex: number,
  counters: Counters,
  regionCache: WeakMap<object, RawRegion>,
  activeRegions: Set<object>,
): NormalizedOperation {
  const discriminator = plainRecord(value, operationIndex, "Scene operation");
  if (discriminator.kind === "path") {
    const object = exactFields(value, ["kind", "commands", "fill", "stroke"], ["transform"], operationIndex, "Path operation");
    const commands = validatedPathCommands(object.commands, operationIndex, counters);
    const fill = validatedColor(object.fill, operationIndex, true);
    const stroke = validateStroke(object.stroke, operationIndex, counters);
    const transform = object.transform === undefined ? identityTransform() : validatedTransform(object.transform, operationIndex, counters, "Drawing transform");
    return { kind: "path", commands, fill, stroke, transform, operationIndex };
  }
  if (discriminator.kind === "text") return validateTextOperation(value, operationIndex, counters);
  if (discriminator.kind === "clip") {
    if (discriminator.action === "reset") {
      exactFields(value, ["kind", "action"], [], operationIndex, "Reset clip operation");
      return { kind: "clip", action: "reset", path: null, region: null, fillRule: null, operationIndex };
    }
    if (discriminator.action === "region") {
      const object = exactFields(value, ["kind", "action", "region"], [], operationIndex, "Region clip operation");
      const regionInput = object.region === null ? null : validateRegion(object.region, operationIndex, counters, regionCache, activeRegions, 0);
      return { kind: "clip", action: "region", path: null, region: regionInput as unknown as PreparedRegion | null, fillRule: null, operationIndex };
    }
    if (discriminator.action === "set" || discriminator.action === "intersect" || discriminator.action === "exclude") {
      const object = exactFields(value, ["kind", "action", "commands", "fillRule"], [], operationIndex, "Path clip operation");
      const action = object.action;
      if (action !== "set" && action !== "intersect" && action !== "exclude") {
        return invalid(operationIndex, "Unknown path clip action.");
      }
      const commands = validatedPathCommands(object.commands, operationIndex, counters);
      const fillRule = validatePathFillRule(object.fillRule, operationIndex);
      return { kind: "clip", action, path: pathWithRule(commands, fillRule), region: null, fillRule, operationIndex };
    }
    return invalid(operationIndex, "Unknown clip operation action.");
  }
  return invalid(operationIndex, "Unknown scene operation kind.");
}

function prepareScene(
  value: MetafileScene,
  counters: Counters,
): { readonly scene: NormalizedScene } {
  const top = exactFields(value, ["viewBox", "operations", "requiredFonts", "warnings"], [], -1, "Metafile scene");
  const viewBox = validatedBox(top.viewBox, -1, counters, "ViewBox", true);
  reserveNumeric(counters, -1);
  const pixelProduct = viewBox.width * viewBox.height;
  if (!Number.isSafeInteger(pixelProduct) || pixelProduct > MAX_VIEW_BOX_PIXELS * SCALE * SCALE) {
    invalid(-1, "ViewBox exceeds the 16-million-pixel ceiling.");
  }
  const viewBoxQuad: Quad = {
    points: [
      { x: viewBox.x / SCALE, y: viewBox.y / SCALE },
      { x: viewBox.right / SCALE, y: viewBox.y / SCALE },
      { x: viewBox.right / SCALE, y: viewBox.bottom / SCALE },
      { x: viewBox.x / SCALE, y: viewBox.bottom / SCALE },
    ],
  };
  const rawOperations = arrayValues(top.operations, -1, "Scene operations", MAX_OPERATIONS);
  if (rawOperations.length > MAX_OPERATIONS) fail("max-svg-nodes", -1, "Scene exceeds 50,000 operations.");
  const regionCache = new WeakMap<object, RawRegion>();
  const activeRegions = new Set<object>();
  const operations = rawOperations.map((operation, index) => validateOperation(operation, index, counters, regionCache, activeRegions));
  const preparedOperations: NormalizedOperation[] = [];
  for (const operation of operations) {
    if (operation.kind !== "clip" || operation.action !== "region") {
      preparedOperations.push(operation);
      continue;
    }
    const raw = operation.region as unknown as RawRegion | null;
    const prepared = prepareRegion(raw, operation.operationIndex, counters, new Set(), 0, 0, 0, 0);
    const preparedOperation: NormalizedClipOperation = { ...operation, region: prepared };
    preparedOperations.push(preparedOperation);
  }
  const fonts = arrayValues(top.requiredFonts, -1, "Required fonts", MAX_OPERATIONS);
  const requiredFonts: string[] = [];
  const uniqueFonts = new Set<string>();
  for (const font of fonts) {
    if (typeof font !== "string") invalid(-1, "Required font names must be strings.");
    const normalized = validXmlString(font, -1, "Required font name", MAX_TEXT_UNITS);
    if (!uniqueFonts.has(normalized)) {
      uniqueFonts.add(normalized);
      requiredFonts.push(normalized);
    }
  }
  const rawWarnings = arrayValues(top.warnings, -1, "Scene warnings", MAX_OPERATIONS);
  const warnings = rawWarnings.map((warning) => validateWarning(warning, -1, counters));
  return { scene: { viewBox, viewBoxQuad, operations: preparedOperations, requiredFonts, warnings } };
}

function reserveDnfTerms(counters: Counters, count: number, operationIndex: number): void {
  if (!Number.isSafeInteger(count) || count < 0 || counters.dnfTerms > MAX_CLIP_NODES - count) {
    fail("max-clip-nodes", operationIndex, "DNF term expansion exceeds 50,000 terms.");
  }
  counters.dnfTerms += count;
}

function reserveLiterals(counters: Counters, count: number, operationIndex: number): void {
  if (!Number.isSafeInteger(count) || count < 0 || counters.literals > MAX_CLIP_NODES - count) {
    fail("max-clip-nodes", operationIndex, "DNF literal expansion exceeds 50,000 literals.");
  }
  counters.literals += count;
}

function makeTerms(count: number, counters: Counters, operationIndex: number): Literal[][] {
  reserveDnfTerms(counters, count, operationIndex);
  return new Array<Literal[]>(count);
}

function unionDnf(first: Dnf, second: Dnf, counters: Counters, operationIndex: number): Dnf {
  const count = first.length + second.length;
  let literals = 0;
  for (const term of first) literals += term.length;
  for (const term of second) literals += term.length;
  reserveDnfTerms(counters, count, operationIndex);
  reserveLiterals(counters, literals, operationIndex);
  const result = new Array<Literal[]>(count);
  let index = 0;
  for (const term of first) result[index++] = [...term];
  for (const term of second) result[index++] = [...term];
  return result;
}

function intersectDnf(first: Dnf, second: Dnf, counters: Counters, operationIndex: number): Dnf {
  if (first.length === 0 || second.length === 0) return [];
  if (first.length > Math.floor(MAX_CLIP_NODES / second.length)) {
    fail("max-clip-nodes", operationIndex, "DNF product exceeds 50,000 terms.");
  }
  const count = first.length * second.length;
  if (counters.dnfTerms > MAX_CLIP_NODES - count) fail("max-clip-nodes", operationIndex, "DNF product exceeds 50,000 terms.");
  let literalCount = 0;
  for (const left of first) {
    for (const right of second) {
      if (left.length + right.length > MAX_CLIP_NODES - literalCount) {
        fail("max-clip-nodes", operationIndex, "DNF product exceeds 50,000 literals.");
      }
      literalCount += left.length + right.length;
    }
  }
  reserveDnfTerms(counters, count, operationIndex);
  reserveLiterals(counters, literalCount, operationIndex);
  const result = new Array<Literal[]>(count);
  let index = 0;
  for (const left of first) {
    for (const right of second) result[index++] = [...left, ...right];
  }
  return result;
}

function primitiveLiteral(
  primitive: Primitive,
  negated: boolean,
  operationIndex: number,
  counters: Counters,
): Dnf {
  reserveDnfTerms(counters, 1, operationIndex);
  reserveLiterals(counters, 1, operationIndex);
  return [[{ primitive, negated, operationIndex }]];
}

function normalizePreparedRegion(
  region: PreparedRegion | null,
  negated: boolean,
  operationIndex: number,
  counters: Counters,
  depth: number,
): Dnf {
  if (depth > MAX_CLIP_DEPTH) fail("max-clip-depth", operationIndex, "Clip expression exceeds 128 levels.");
  if (region === null) return negated ? [] : makeTerms(1, counters, operationIndex).fill([]);
  if (region.kind === "rectangles") {
    if (!negated) {
      reserveDnfTerms(counters, region.quads.length, operationIndex);
      reserveLiterals(counters, region.quads.length, operationIndex);
      const result = new Array<Literal[]>(region.quads.length);
      for (let index = 0; index < region.quads.length; index += 1) {
        const quad = region.quads[index];
        result[index] = [{ primitive: { kind: "quad", quad }, negated: false, operationIndex }];
      }
      return result;
    }
    if (region.quads.length === 0) return makeTerms(1, counters, operationIndex).fill([]);
    reserveDnfTerms(counters, 1, operationIndex);
    reserveLiterals(counters, region.quads.length, operationIndex);
    return [region.quads.map((quad) => ({ primitive: { kind: "quad" as const, quad }, negated: true, operationIndex }))];
  }
  const leftPositive = (negative: boolean): Dnf => normalizePreparedRegion(region.left, negative, operationIndex, counters, depth + 1);
  const rightPositive = (negative: boolean): Dnf => normalizePreparedRegion(region.right, negative, operationIndex, counters, depth + 1);
  if (region.mode === "intersect") {
    return negated
      ? unionDnf(leftPositive(true), rightPositive(true), counters, operationIndex)
      : intersectDnf(leftPositive(false), rightPositive(false), counters, operationIndex);
  }
  if (region.mode === "union") {
    return negated
      ? intersectDnf(leftPositive(true), rightPositive(true), counters, operationIndex)
      : unionDnf(leftPositive(false), rightPositive(false), counters, operationIndex);
  }
  if (region.mode === "exclude") {
    return negated
      ? unionDnf(leftPositive(true), rightPositive(false), counters, operationIndex)
      : intersectDnf(leftPositive(false), rightPositive(true), counters, operationIndex);
  }
  const left = leftPositive(false);
  const right = rightPositive(false);
  const leftNot = leftPositive(true);
  const rightNot = rightPositive(true);
  const first = intersectDnf(left, negated ? right : rightNot, counters, operationIndex);
  const second = intersectDnf(leftNot, negated ? rightNot : right, counters, operationIndex);
  return unionDnf(first, second, counters, operationIndex);
}

function normalizeExpression(
  expression: ClipExpression,
  counters: Counters,
  depth = 1,
  negated = false,
): Dnf {
  if (depth > MAX_CLIP_DEPTH) {
    const operationIndex = "operationIndex" in expression ? expression.operationIndex : -1;
    fail("max-clip-depth", operationIndex, "Clip expression exceeds 128 levels.");
  }
  if (expression.kind === "unbounded") return negated ? [] : makeTerms(1, counters, -1).fill([]);
  if (expression.kind === "region") return normalizePreparedRegion(expression.region, negated, expression.operationIndex, counters, depth);
  if (expression.kind === "path") {
    if (negated && expression.path.fillRule !== "evenOdd") {
      unsupported(expression.path.operationIndex, "Nonzero path complementation has no proven exact SVG construction.");
    }
    if (negated && !isClosedPath(expression.path)) {
      unsupported(expression.path.operationIndex, "Only closed evenodd paths have a proven exact complement.");
    }
    return primitiveLiteral({ kind: "path", path: expression.path }, negated, expression.path.operationIndex, counters);
  }
  const operationIndex = expression.operationIndex;
  const left = (negative: boolean): Dnf => normalizeExpression(expression.left, counters, depth + 1, negative);
  const right = (negative: boolean): Dnf => normalizeExpression(expression.right, counters, depth + 1, negative);
  if (expression.mode === "intersect") {
    return negated
      ? unionDnf(left(true), right(true), counters, operationIndex)
      : intersectDnf(left(false), right(false), counters, operationIndex);
  }
  if (expression.mode === "union") {
    return negated
      ? intersectDnf(left(true), right(true), counters, operationIndex)
      : unionDnf(left(false), right(false), counters, operationIndex);
  }
  if (expression.mode === "exclude") {
    return negated
      ? unionDnf(left(true), right(false), counters, operationIndex)
      : intersectDnf(left(false), right(true), counters, operationIndex);
  }
  const leftPositive = left(false);
  const rightPositive = right(false);
  const leftNegative = left(true);
  const rightNegative = right(true);
  const first = intersectDnf(leftPositive, negated ? rightPositive : rightNegative, counters, operationIndex);
  const second = intersectDnf(leftNegative, negated ? rightNegative : rightPositive, counters, operationIndex);
  return unionDnf(first, second, counters, operationIndex);
}

function expressionForClipOperations(
  operations: readonly NormalizedOperation[],
  warnings: readonly NormalizedWarning[],
  counters: Counters,
): readonly (Dnf | null)[] {
  const clips: (Dnf | null)[] = Array.from({ length: operations.length }, () => null);
  let active: ClipExpression = UNBOUNDED;
  for (let index = 0; index < operations.length; index += 1) {
    const operation = operations[index];
    if (operation.kind === "clip") {
      if (operation.action === "reset" || operation.action === "region" && operation.region === null) {
        active = UNBOUNDED;
      } else if (operation.action === "region") {
        const region = operation.region as unknown as PreparedRegion;
        active = { kind: "region", region, operationIndex: operation.operationIndex, depth: regionDepthFromPrepared(region) };
        normalizeExpression(active, counters);
      } else {
        const path = operation.path!;
        const atom: ClipExpression = { kind: "path", path, depth: 1 };
        if (operation.action === "set") active = atom;
        else active = {
          kind: "combine",
          mode: operation.action === "intersect" ? "intersect" : "exclude",
          left: active,
          right: atom,
          operationIndex: operation.operationIndex,
          depth: 1 + Math.max(active.depth, atom.depth),
        };
        normalizeExpression(active, counters);
      }
      continue;
    }
    if (operation.kind === "text" && operation.text === null) {
      fail("font-glyph-gap", operation.operationIndex, "Font-specific bytes have no verified local glyph mapping.");
    }
    if (operation.kind === "text") {
      if (warnings.some((warning) => warning.kind === "font-metrics-required" && warning.fontName === operation.fontFamily)) {
        fail("font-metrics-required", operation.operationIndex, "Text metrics are unresolved.");
      }
    }
    clips[index] = normalizeExpression(active, counters);
  }
  return clips;
}

function regionDepthFromPrepared(region: PreparedRegion | null): number {
  if (region === null || region.kind === "rectangles") return region === null ? 0 : 1;
  return 1 + Math.max(regionDepthFromPrepared(region.left), regionDepthFromPrepared(region.right));
}

function formatNumber(value: number, operationIndex: number): string {
  if (!Number.isFinite(value)) unsupported(operationIndex, "SVG numeric value is not finite.");
  const normalized = Object.is(value, -0) ? 0 : value;
  const text = normalized.toString();
  if (!Number.isFinite(Number(text)) || Number(text) !== normalized) unsupported(operationIndex, "SVG numeric value cannot round-trip exactly.");
  return text;
}

function formatScaled(value: number, operationIndex: number): string {
  return formatNumber(value / SCALE, operationIndex);
}

function quadPathData(quad: Quad, operationIndex: number): string {
  const points = quad.points;
  return `M${formatNumber(points[0].x, operationIndex)} ${formatNumber(points[0].y, operationIndex)}` +
    `L${formatNumber(points[1].x, operationIndex)} ${formatNumber(points[1].y, operationIndex)}` +
    `L${formatNumber(points[2].x, operationIndex)} ${formatNumber(points[2].y, operationIndex)}` +
    `L${formatNumber(points[3].x, operationIndex)} ${formatNumber(points[3].y, operationIndex)}Z`;
}

function viewBoxPath(quad: Quad, operationIndex: number): string {
  return quadPathData(quad, operationIndex);
}

function pathData(path: PathData, operationIndex: number): string {
  const chunks: string[] = [];
  for (const command of path.commands) {
    if (command.type === "moveTo") {
      chunks.push(`M${formatScaled(command.point.x * SCALE, operationIndex)} ${formatScaled(command.point.y * SCALE, operationIndex)}`);
    } else if (command.type === "lineTo") {
      chunks.push(`L${formatScaled(command.point.x * SCALE, operationIndex)} ${formatScaled(command.point.y * SCALE, operationIndex)}`);
    } else if (command.type === "cubicTo") {
      chunks.push(`C${formatScaled(command.control1.x * SCALE, operationIndex)} ${formatScaled(command.control1.y * SCALE, operationIndex)}` +
        ` ${formatScaled(command.control2.x * SCALE, operationIndex)} ${formatScaled(command.control2.y * SCALE, operationIndex)}` +
        ` ${formatScaled(command.point.x * SCALE, operationIndex)} ${formatScaled(command.point.y * SCALE, operationIndex)}`);
    } else {
      chunks.push("Z");
    }
  }
  return chunks.join("");
}

function colorText(color: NormalizedColor): string {
  return `rgb(${color.red},${color.green},${color.blue})`;
}

function attributesForColor(color: NormalizedColor, attribute: "fill" | "stroke"): (readonly [string, string])[] {
  const result: (readonly [string, string])[] = [[attribute, colorText(color)]];
  if (color.alpha !== 1) result.push([attribute === "fill" ? "fill-opacity" : "stroke-opacity", formatNumber(color.alpha, -1)]);
  return result;
}

function node(
  tag: SvgNode["tag"],
  attributes: readonly (readonly [string, string])[],
  children: readonly (SvgNode | string)[],
  counters: Counters,
  operationIndex: number,
): SvgNode {
  if (counters.svgNodes >= MAX_SVG_NODES) fail("max-svg-nodes", operationIndex, "Generated SVG exceeds 50,000 elements.");
  counters.svgNodes += 1;
  return { tag, attributes, children };
}

function newClipPath(
  defs: SvgNode[],
  counters: Counters,
  operationIndex: number,
  children: readonly SvgNode[],
): string {
  counters.nextId += 1;
  const id = `mf-clip-${counters.nextId}`;
  const element = node("clipPath", [["id", id], ["clipPathUnits", "userSpaceOnUse"]], children, counters, operationIndex);
  defs.push(element);
  return id;
}

function definitionPath(
  d: string,
  clipRule: "evenodd" | "nonzero",
  reference: string | null,
  counters: Counters,
  operationIndex: number,
): SvgNode {
  const attributes: (readonly [string, string])[] = [["d", d], ["clip-rule", clipRule]];
  if (reference !== null) attributes.push(["clip-path", `url(#${reference})`]);
  return node("path", attributes, [], counters, operationIndex);
}

function buildClip(
  dnf: Dnf,
  viewBox: Quad,
  defs: SvgNode[],
  counters: Counters,
  operationIndex: number,
): string {
  const basePath = viewBoxPath(viewBox, operationIndex);
  const baseId = newClipPath(defs, counters, operationIndex, [definitionPath(basePath, "nonzero", null, counters, operationIndex)]);
  const terms: string[] = [];
  for (const term of dnf) {
    let conjunction = baseId;
    for (const literal of term) {
      let d: string;
      let rule: "evenodd" | "nonzero";
      if (literal.primitive.kind === "quad") {
        const quadData = quadPathData(literal.primitive.quad, literal.operationIndex);
        if (literal.negated) {
          d = `${basePath}${quadData}`;
          rule = "evenodd";
        } else {
          d = quadData;
          rule = "nonzero";
        }
      } else {
        const source = pathData(literal.primitive.path, literal.operationIndex);
        if (literal.negated) {
          d = `${basePath}${source}`;
          rule = "evenodd";
        } else {
          d = source;
          rule = literal.primitive.path.fillRule === "evenOdd" ? "evenodd" : "nonzero";
        }
      }
      const child = definitionPath(d, rule, conjunction, counters, literal.operationIndex);
      conjunction = newClipPath(defs, counters, literal.operationIndex, [child]);
    }
    terms.push(conjunction);
  }
  const unionChildren: SvgNode[] = [];
  for (const termId of terms) {
    unionChildren.push(definitionPath(basePath, "nonzero", termId === baseId ? null : termId, counters, operationIndex));
  }
  if (terms.length === 0) return newClipPath(defs, counters, operationIndex, []);
  return newClipPath(defs, counters, operationIndex, unionChildren);
}

function matrixTransform(transform: TransformI, operationIndex: number): string {
  return `matrix(${formatScaled(transform.a, operationIndex)} ${formatScaled(transform.b, operationIndex)} ${formatScaled(transform.c, operationIndex)} ${formatScaled(transform.d, operationIndex)} ${formatScaled(transform.e, operationIndex)} ${formatScaled(transform.f, operationIndex)})`;
}

function operationPathNode(
  operation: NormalizedPathOperation,
  counters: Counters,
): SvgNode {
  const attributes: (readonly [string, string])[] = [["d", pathData(operation.commands, operation.operationIndex)]];
  if (operation.transform.a !== SCALE || operation.transform.b !== 0 || operation.transform.c !== 0 || operation.transform.d !== SCALE || operation.transform.e !== 0 || operation.transform.f !== 0) {
    attributes.push(["transform", matrixTransform(operation.transform, operation.operationIndex)]);
  }
  if (operation.fill) attributes.push(...attributesForColor(operation.fill, "fill"));
  else attributes.push(["fill", "none"]);
  if (operation.stroke) {
    attributes.push(...attributesForColor(operation.stroke.color, "stroke"));
    attributes.push(["stroke-width", formatNumber(operation.stroke.width, operation.operationIndex)]);
    attributes.push(["stroke-linecap", operation.stroke.lineCap]);
    attributes.push(["stroke-linejoin", operation.stroke.lineJoin]);
    attributes.push(["stroke-miterlimit", formatNumber(operation.stroke.miterLimit, operation.operationIndex)]);
  } else {
    attributes.push(["stroke", "none"]);
  }
  return node("path", attributes, [], counters, operation.operationIndex);
}

function textAttributes(operation: NormalizedTextOperation): (readonly [string, string])[] {
  const attributes: (readonly [string, string])[] = [
    ["x", operation.glyphPositions?.map((value) => formatNumber(value, operation.operationIndex)).join(" ") ?? formatNumber(operation.x, operation.operationIndex)],
    ["y", formatNumber(operation.y, operation.operationIndex)],
    ["font-family", operation.fontFamily],
    ["font-size", formatNumber(operation.fontSize, operation.operationIndex)],
    ["font-weight", String(operation.logicalFont?.weight ?? 400)],
    ["font-style", operation.logicalFont?.italic ? "italic" : "normal"],
    ["direction", operation.direction],
    ["unicode-bidi", operation.direction === "rtl" ? "embed" : "normal"],
  ];
  const decorations: string[] = [];
  if (operation.logicalFont?.underline) decorations.push("underline");
  if (operation.logicalFont?.strikeOut) decorations.push("line-through");
  if (decorations.length) attributes.push(["text-decoration", decorations.join(" ")]);
  attributes.push(...attributesForColor(operation.color, "fill"));
  if (operation.transform.a !== SCALE || operation.transform.b !== 0 || operation.transform.c !== 0 || operation.transform.d !== SCALE || operation.transform.e !== 0 || operation.transform.f !== 0) {
    attributes.push(["transform", matrixTransform(operation.transform, operation.operationIndex)]);
  }
  return attributes;
}

function buildDrawingNodes(
  operation: NormalizedPathOperation | NormalizedTextOperation,
  clip: Dnf,
  viewBox: Quad,
  defs: SvgNode[],
  counters: Counters,
): SvgNode {
  const operationIndex = operation.operationIndex;
  const children: SvgNode[] = [];
  if (operation.kind === "path") {
    children.push(operationPathNode(operation, counters));
  } else {
    if (operation.encoding === "font-specific" || operation.text === null) {
      fail("font-glyph-gap", operationIndex, "Font-specific bytes have no verified local glyph mapping.");
    }
    if (operation.opaqueQuad && (operation.backgroundMode === "opaque" || (operation.textOptions & 2) !== 0)) {
      if (!operation.backgroundColor) invalid(operationIndex, "Opaque text background has no source color.");
      const background = node("path", [
        ["d", quadPathData(operation.opaqueQuad, operationIndex)],
        ...attributesForColor(operation.backgroundColor, "fill"),
      ], [], counters, operationIndex);
      children.push(background);
    }
    const textNode = node("text", textAttributes(operation), [operation.text], counters, operationIndex);
    if (operation.clipSpecified) {
      const tempDnf = operation.clipQuad
        ? primitiveLiteral({ kind: "quad", quad: operation.clipQuad }, false, operationIndex, counters)
        : [];
      const tempId = buildClip(tempDnf, viewBox, defs, counters, operationIndex);
      children.push(node("g", [["clip-path", `url(#${tempId})`]], [textNode], counters, operationIndex));
    } else {
      children.push(textNode);
    }
  }
  const clipId = buildClip(clip, viewBox, defs, counters, operationIndex);
  return node("g", [["clip-path", `url(#${clipId})`]], children, counters, operationIndex);
}

function buildSvg(scene: NormalizedScene, counters: Counters): SvgNode {
  const defs: SvgNode[] = [];
  const drawings: SvgNode[] = [];
  const clips = expressionForClipOperations(scene.operations, scene.warnings, counters);
  for (let index = 0; index < scene.operations.length; index += 1) {
    const operation = scene.operations[index];
    if (operation.kind === "clip") continue;
    const clip = clips[index] ?? makeTerms(1, counters, operation.operationIndex).fill([]);
    drawings.push(buildDrawingNodes(operation, clip, scene.viewBoxQuad, defs, counters));
  }
  const definitionList = node("defs", EMPTY_ATTRIBUTES, defs, counters, -1);
  const x = formatScaled(scene.viewBox.x, -1);
  const y = formatScaled(scene.viewBox.y, -1);
  const width = formatScaled(scene.viewBox.width, -1);
  const height = formatScaled(scene.viewBox.height, -1);
  const viewBoxText = `${x} ${y} ${width} ${height}`;
  return node("svg", [
    ["xmlns", SVG_NS],
    ["viewBox", viewBoxText],
    ["width", width],
    ["height", height],
  ], [definitionList, ...drawings], counters, -1);
}

const SVG_ATTRIBUTES: Readonly<Record<SvgNode["tag"], ReadonlySet<string>>> = {
  svg: new Set(["xmlns", "viewBox", "width", "height"]),
  defs: new Set(),
  clipPath: new Set(["id", "clipPathUnits"]),
  path: new Set(["d", "transform", "fill", "fill-rule", "clip-rule", "clip-path", "stroke", "stroke-width", "stroke-linecap", "stroke-linejoin", "stroke-miterlimit", "fill-opacity", "stroke-opacity"]),
  rect: new Set(["x", "y", "width", "height", "transform", "fill", "fill-opacity", "clip-path"]),
  g: new Set(["transform", "clip-path"]),
  text: new Set(["x", "y", "dx", "dy", "transform", "font-family", "font-size", "font-weight", "font-style", "text-decoration", "direction", "unicode-bidi", "fill", "fill-opacity", "clip-path"]),
  tspan: new Set(["x", "y", "dx", "dy", "transform", "font-family", "font-size", "font-weight", "font-style", "text-decoration", "direction", "unicode-bidi", "fill", "fill-opacity", "clip-path"]),
};

function validateGeneratedSvg(root: SvgNode): void {
  if (root.tag !== "svg") invalid(-1, "Generated SVG root is invalid.");
  const ids = new Map<string, SvgNode>();
  const references = new Map<string, string[]>();
  const allReferences = new Set<string>();
  const activeNodes = new Set<SvgNode>();

  const visit = (
    element: SvgNode,
    parentTag: SvgNode["tag"] | null,
    clipOwner: string | null,
    transformedAncestor: boolean,
  ): void => {
    if (activeNodes.has(element)) invalid(-1, "Generated SVG contains a structural cycle.");
    activeNodes.add(element);
    const allowed = SVG_ATTRIBUTES[element.tag];
    const attributes = new Map<string, string>();
    for (const [name, value] of element.attributes) {
      if (!allowed.has(name) || attributes.has(name)) invalid(-1, "Generated SVG contains a duplicate or forbidden attribute.");
      attributes.set(name, value);
    }
    if (element.tag === "svg" && parentTag !== null) invalid(-1, "Generated SVG root has a parent.");
    if (element.tag === "defs" && parentTag !== "svg") invalid(-1, "SVG definitions are outside the root definition list.");
    if (element.tag === "clipPath" && parentTag !== "defs") invalid(-1, "Clip path definition is outside defs.");
    if (element.tag === "clipPath") {
      const id = attributes.get("id");
      if (!id || !/^mf-clip-[1-9][0-9]*$/.test(id) || ids.has(id) || attributes.get("clipPathUnits") !== "userSpaceOnUse") {
        invalid(-1, "Generated clip path ID or coordinate units are invalid.");
      }
      ids.set(id, element);
      if (attributes.has("transform") || attributes.has("clip-path")) invalid(-1, "Clip definitions cannot be transformed or reference another clip directly.");
    }
    if (attributes.has("clip-rule")) {
      const clipRule = attributes.get("clip-rule");
      if (element.tag !== "path" || clipOwner === null || (clipRule !== "evenodd" && clipRule !== "nonzero")) {
        invalid(-1, "clip-rule is restricted to valid clip-definition paths.");
      }
    }
    if (clipOwner !== null && element.tag === "path") {
      if (attributes.has("transform") || !attributes.has("clip-rule") || attributes.has("fill-rule")) {
        invalid(-1, "Clip path literals must be untransformed and carry an explicit clip-rule.");
      }
    }
    if (element.tag !== "path" && attributes.has("fill-rule")) invalid(-1, "fill-rule is restricted to drawing paths.");
    if (element.tag === "path" && clipOwner === null && attributes.has("clip-rule")) invalid(-1, "Drawing paths cannot carry clip-rule.");
    if (attributes.has("transform")) {
      if (clipOwner !== null || transformedAncestor || parentTag !== "g" || attributes.has("clip-path") ||
          (element.tag !== "path" && element.tag !== "text" && element.tag !== "tspan" && element.tag !== "rect")) {
        invalid(-1, "Transforms are restricted to drawing children.");
      }
    }
    if (element.tag === "g" && attributes.has("clip-path") && attributes.has("transform")) {
      invalid(-1, "Device-space clipping wrappers cannot be transformed.");
    }
    if (element.tag === "svg" && attributes.get("xmlns") !== SVG_NS) invalid(-1, "SVG namespace is invalid.");
    if (element.tag === "clipPath") {
      for (const child of element.children) if (typeof child === "string" || child.tag !== "path") invalid(-1, "Clip definitions may contain path literals only.");
    }
    if (element.tag === "defs") {
      for (const child of element.children) if (typeof child === "string" || child.tag !== "clipPath") invalid(-1, "Defs may contain generated clip paths only.");
    }
    if (element.tag === "svg") {
      if (element.children.length === 0 || typeof element.children[0] === "string" || element.children[0].tag !== "defs") {
        invalid(-1, "SVG defs must precede drawing nodes.");
      }
      for (const child of element.children.slice(1)) if (typeof child === "string" || child.tag !== "g") invalid(-1, "SVG root may contain drawing groups only after defs.");
    }
    if (element.tag === "text" || element.tag === "tspan") {
      for (const child of element.children) if (typeof child !== "string" && child.tag !== "tspan") invalid(-1, "Text nodes may contain text and tspan children only.");
    } else if (element.tag !== "clipPath" && element.tag !== "defs" && element.tag !== "svg") {
      for (const child of element.children) if (typeof child === "string") invalid(-1, "Text data is only permitted in text elements.");
    }
    if (element.tag === "path" || element.tag === "rect") {
      if (element.children.length !== 0) invalid(-1, "Drawing primitives cannot have child nodes.");
    }
    const clipPath = attributes.get("clip-path");
    if (clipPath !== undefined) {
      const match = /^url\(#(mf-clip-[1-9][0-9]*)\)$/.exec(clipPath);
      if (!match?.[1]) invalid(-1, "Only internal generated clip references are allowed.");
      if (!((element.tag === "g" && clipOwner === null) || (element.tag === "path" && clipOwner !== null))) {
        invalid(-1, "Clip references are restricted to drawing wrappers and definition literals.");
      }
      allReferences.add(match[1]);
      if (clipOwner !== null) {
        const existing = references.get(clipOwner) ?? [];
        existing.push(match[1]);
        references.set(clipOwner, existing);
      }
    }
    const nextClipOwner = element.tag === "clipPath" ? attributes.get("id")! : clipOwner;
    const nextTransformed = transformedAncestor || attributes.has("transform");
    for (const child of element.children) {
      if (typeof child !== "string") visit(child, element.tag, nextClipOwner, nextTransformed);
    }
    activeNodes.delete(element);
  };

  visit(root, null, null, false);
  let expectedId = 1;
  for (const id of ids.keys()) {
    if (id !== `mf-clip-${expectedId++}`) invalid(-1, "Generated clip IDs are not deterministic traversal IDs.");
  }
  for (const target of allReferences) if (!ids.has(target)) invalid(-1, "Generated clip reference is unresolved.");
  const visiting = new Set<string>();
  const visited = new Set<string>();
  const checkReference = (id: string): void => {
    if (visited.has(id)) return;
    if (visiting.has(id)) invalid(-1, "Generated clip references contain a cycle.");
    visiting.add(id);
    for (const target of references.get(id) ?? []) checkReference(target);
    visiting.delete(id);
    visited.add(id);
  };
  for (const id of ids.keys()) checkReference(id);
}

function utf8Length(value: string): number {
  let length = 0;
  for (let index = 0; index < value.length; index += 1) {
    const code = value.charCodeAt(index);
    if (code === 38) length += 5; // &amp;
    else if (code === 60) length += 4; // &lt;
    else if (code === 62) length += 4; // &gt;
    else if (code === 34) length += 6; // &quot;
    else if (code === 39) length += 6; // &apos;
    else if (code <= 0x7f) length += 1;
    else if (code <= 0x7ff) length += 2;
    else if (code >= 0xd800 && code <= 0xdbff) {
      length += 4;
      index += 1;
    } else length += 3;
    if (length > MAX_OUTPUT_BYTES) return length;
  }
  return length;
}

function boundedAdd(first: number, second: number, operationIndex: number): number {
  if (second > MAX_OUTPUT_BYTES || first > MAX_OUTPUT_BYTES - second) {
    fail("max-output-bytes", operationIndex, "Generated SVG exceeds 16 MiB.");
  }
  return first + second;
}

function serializedSize(element: SvgNode, operationIndex: number): number {
  let size = 2 + element.tag.length + 3 + element.tag.length;
  for (const [name, value] of element.attributes) {
    size = boundedAdd(size, 1 + name.length + 2 + utf8Length(value) + 1, operationIndex);
  }
  for (const child of element.children) {
    size = boundedAdd(size, typeof child === "string" ? utf8Length(child) : serializedSize(child, operationIndex), operationIndex);
  }
  return size;
}

function writeUtf8(output: Uint8Array, offset: number, value: string): number {
  const writeAscii = (text: string): number => {
    for (let index = 0; index < text.length; index += 1) output[offset++] = text.charCodeAt(index);
    return offset;
  };
  for (let index = 0; index < value.length; index += 1) {
    const code = value.charCodeAt(index);
    if (code === 38) offset = writeAscii("&amp;");
    else if (code === 60) offset = writeAscii("&lt;");
    else if (code === 62) offset = writeAscii("&gt;");
    else if (code === 34) offset = writeAscii("&quot;");
    else if (code === 39) offset = writeAscii("&apos;");
    else if (code <= 0x7f) output[offset++] = code;
    else if (code <= 0x7ff) {
      output[offset++] = 0xc0 | (code >> 6);
      output[offset++] = 0x80 | (code & 0x3f);
    } else if (code >= 0xd800 && code <= 0xdbff) {
      const low = value.charCodeAt(++index);
      const point = 0x10000 + ((code - 0xd800) << 10) + (low - 0xdc00);
      output[offset++] = 0xf0 | (point >> 18);
      output[offset++] = 0x80 | ((point >> 12) & 0x3f);
      output[offset++] = 0x80 | ((point >> 6) & 0x3f);
      output[offset++] = 0x80 | (point & 0x3f);
    } else {
      output[offset++] = 0xe0 | (code >> 12);
      output[offset++] = 0x80 | ((code >> 6) & 0x3f);
      output[offset++] = 0x80 | (code & 0x3f);
    }
  }
  return offset;
}

function renderNode(output: Uint8Array, initialOffset: number, element: SvgNode): number {
  let offset = initialOffset;
  const writeAscii = (value: string): void => {
    for (let index = 0; index < value.length; index += 1) output[offset++] = value.charCodeAt(index);
  };
  writeAscii(`<${element.tag}`);
  for (const [name, value] of element.attributes) {
    writeAscii(` ${name}="`);
    offset = writeUtf8(output, offset, value);
    output[offset++] = 34;
  }
  output[offset++] = 62;
  for (const child of element.children) {
    if (typeof child === "string") offset = writeUtf8(output, offset, child);
    else offset = renderNode(output, offset, child);
  }
  writeAscii(`</${element.tag}>`);
  return offset;
}

function serialize(scene: MetafileScene): Uint8Array {
  const counters: Counters = {
    numericOperations: 0,
    uniqueClipNodes: 0,
    expandedClipNodes: 0,
    regionRectangles: 0,
    pathCommands: 0,
    dnfTerms: 0,
    literals: 0,
    svgNodes: 0,
    nextId: 0,
  };
  const prepared = prepareScene(scene, counters);
  const tree = buildSvg(prepared.scene, counters);
  validateGeneratedSvg(tree);
  const size = serializedSize(tree, -1);
  if (size > MAX_OUTPUT_BYTES) fail("max-output-bytes", -1, "Generated SVG exceeds 16 MiB.");
  const bytes = new Uint8Array(size);
  const written = renderNode(bytes, 0, tree);
  if (written !== size) fail("invalid-scene", -1, "Generated SVG size accounting disagrees with UTF-8 output.");
  return bytes;
}

export function serializeMetafileScene(scene: MetafileScene): Uint8Array {
  return serialize(scene);
}
