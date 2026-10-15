import assert from "node:assert/strict";
import test from "node:test";
import { childElements, parseXml } from "../../src/pptx/xml";
import { MetafileSerializeError, serializeMetafileScene } from "../../src/ooxml/metafileSvg";
import type {
  MetafileClipRegion,
  MetafileColor,
  MetafilePathCommand,
  MetafileScene,
  MetafileSceneOperation,
  MetafileTextOperation,
  MetafileTransform,
  MetafileViewBox,
} from "../../src/ooxml/metafileScene";

const V: MetafileViewBox = { x: 0, y: 0, width: 12, height: 12 };
const IDENTITY: MetafileTransform = { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 };
const RED: MetafileColor = { red: 255, green: 0, blue: 0, alpha: 1 };

function rect(x: number, y: number, width: number, height: number): MetafileViewBox {
  return { x, y, width, height };
}

function region(
  rectangles: readonly MetafileViewBox[],
  transform: MetafileTransform = IDENTITY,
): MetafileClipRegion {
  return { kind: "rectangles", rectangles, transform };
}

function combined(
  mode: "intersect" | "union" | "xor" | "exclude",
  left: MetafileClipRegion | null,
  right: MetafileClipRegion,
): MetafileClipRegion {
  return { kind: "combine", mode, left, right };
}

function filledPath(
  commands: readonly MetafilePathCommand[] = [
    { type: "moveTo", point: { x: 0, y: 0 } },
    { type: "lineTo", point: { x: 1, y: 0 } },
    { type: "lineTo", point: { x: 1, y: 1 } },
    { type: "closePath" },
  ],
  transform?: MetafileTransform,
): MetafileSceneOperation {
  return { kind: "path", commands, fill: RED, stroke: null, transform };
}

function scene(
  operations: readonly MetafileSceneOperation[],
  viewBox: MetafileViewBox = V,
  requiredFonts: readonly string[] = [],
  warnings: MetafileScene["warnings"] = [],
): MetafileScene {
  return { viewBox, operations, requiredFonts, warnings };
}

function textRun(
  text: string,
  advances?: readonly number[],
  sourceAdvances?: readonly number[],
): MetafileTextOperation {
  return {
    kind: "text", x: 1, y: 2, text, fontFamily: "Test Font", fontSize: 1, color: RED, direction: "ltr", advances, sourceAdvances,
    textAlign: 0x18,
  };
}

function regionOperation(value: MetafileClipRegion): MetafileSceneOperation {
  return { kind: "clip", action: "region", region: value };
}

function balancedEmptyRegionTree(leafCount: number): MetafileClipRegion {
  let level = Array.from({ length: leafCount }, () => region([]));
  while (level.length > 1) {
    const next: MetafileClipRegion[] = [];
    for (let index = 0; index < level.length; index += 2) {
      const left = level[index];
      const right = level[index + 1];
      assert.ok(left);
      next.push(right ? combined("intersect", left, right) : left);
    }
    level = next;
  }
  assert.ok(level[0]);
  return level[0];
}

function balancedUniverseUnionTree(
  leafCount: number,
  decoratedLeaves: number,
): { readonly region: MetafileClipRegion; readonly terms: number; readonly termWork: number } {
  let level = Array.from({ length: leafCount }, (_, index) => {
    const empty = region([]);
    const decorated = index < decoratedLeaves;
    return {
      region: combined(decorated ? "exclude" : "union", null, empty),
      terms: 1,
      termWork: decorated ? 3 : 2,
    };
  });
  while (level.length > 1) {
    const next: typeof level = [];
    for (let index = 0; index < level.length; index += 2) {
      const left = level[index];
      const right = level[index + 1];
      assert.ok(left);
      next.push(right
        ? {
            region: combined("union", left.region, right.region),
            terms: left.terms + right.terms,
            termWork: left.termWork + right.termWork + left.terms + right.terms,
          }
        : left);
    }
    level = next;
  }
  assert.ok(level[0]);
  return level[0];
}

function textForEscapedByteLength(byteLength: number): string {
  const quoteCount = Math.floor(byteLength / 6);
  const remainder = byteLength % 6;
  let suffix: string;
  if (quoteCount + remainder <= 8192) {
    suffix = "x".repeat(remainder);
  } else {
    const reducedQuotes = quoteCount - 1;
    const remainingBytes = byteLength - reducedQuotes * 6;
    const suffixes: Readonly<Record<number, string>> = {
      6: '"', 7: '"x', 8: "<<", 9: "<&", 10: "&&", 11: '"&',
    };
    const replacement = suffixes[remainingBytes];
    assert.ok(replacement !== undefined);
    suffix = replacement;
    const text = '"'.repeat(reducedQuotes) + suffix;
    assert.ok(text.length <= 8192);
    return text;
  }
  const text = '"'.repeat(quoteCount) + suffix;
  assert.ok(text.length <= 8192);
  return text;
}

function clipScene(clip: MetafileClipRegion | null): MetafileScene {
  return scene([
    { kind: "clip", action: "region", region: clip },
    filledPath(),
  ]);
}

function xmlFor(value: MetafileScene): { xml: string; parsed: Document } {
  const xml = new TextDecoder().decode(serializeMetafileScene(value));
  return { xml, parsed: parseXml(xml, "generated.svg") };
}

function groups(document: Document): Element[] {
  return childElements(document.documentElement).filter((element) => element.localName === "g");
}

function clipId(element: Element): string | null {
  const reference = element.getAttribute("clip-path");
  if (!reference) return null;
  const match = /^url\(#([A-Za-z0-9_-]+)\)$/.exec(reference);
  return match?.[1] ?? null;
}

function requiredClipId(element: Element): string {
  const id = clipId(element);
  if (id === null) throw new Error("Expected a generated clip-path reference.");
  return id;
}

function polygonsFromPath(data: string): number[][][] {
  const tokens = data.match(/[MLCZ]|[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?/g) ?? [];
  const polygons: number[][][] = [];
  let polygon: number[][] | null = null;
  for (let index = 0; index < tokens.length;) {
    const command = tokens[index++];
    if (command === "M") {
      polygon = [];
      polygons.push(polygon);
      polygon.push([Number(tokens[index++]), Number(tokens[index++])]);
    } else if (command === "L") {
      assert.ok(polygon, "line command must follow a move command");
      polygon.push([Number(tokens[index++]), Number(tokens[index++])]);
    } else if (command === "Z") {
      assert.ok(polygon, "close command must follow a move command");
      polygon = null;
    } else {
      assert.fail(`unsupported clip polygon command in test evaluator: ${command}`);
    }
  }
  return polygons;
}

function pathSegments(data: string): { readonly type: string; readonly coordinates: readonly number[] }[] {
  const tokens = data.match(/[MLCZ]|[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?/g) ?? [];
  const segments: { type: string; coordinates: number[] }[] = [];
  const arity: Readonly<Record<string, number>> = { M: 2, L: 2, C: 6, Z: 0 };
  for (let index = 0; index < tokens.length;) {
    const type = tokens[index++];
    assert.ok(type && Object.prototype.hasOwnProperty.call(arity, type), `invalid SVG path command ${type}`);
    const count = arity[type];
    assert.ok(count !== undefined && index + count <= tokens.length, `truncated SVG ${type} command`);
    const coordinates = tokens.slice(index, index + count).map((token) => {
      assert.ok(!/^[MLCZ]$/.test(token), `missing coordinate in SVG ${type} command`);
      const value = Number(token);
      assert.ok(Number.isFinite(value), `non-finite SVG ${type} coordinate`);
      return value;
    });
    index += count;
    segments.push({ type, coordinates });
  }
  return segments;
}

function contourWinding(points: readonly number[][], x: number, y: number): number {
  let winding = 0;
  for (let index = 0; index < points.length; index += 1) {
    const first = points[index];
    const second = points[(index + 1) % points.length];
    if (!first || !second) continue;
    const cross = (second[0] - first[0]) * (y - first[1]) -
      (x - first[0]) * (second[1] - first[1]);
    if (Math.abs(cross) < 1e-9 &&
        x >= Math.min(first[0], second[0]) && x <= Math.max(first[0], second[0]) &&
        y >= Math.min(first[1], second[1]) && y <= Math.max(first[1], second[1])) return Number.POSITIVE_INFINITY;
    if (first[1] <= y) {
      if (second[1] > y && cross > 0) winding += 1;
    } else if (second[1] <= y && cross < 0) winding -= 1;
  }
  return winding;
}

function pathContains(path: Element, x: number, y: number): boolean {
  const contours = polygonsFromPath(path.getAttribute("d") ?? "");
  const rule = path.getAttribute("clip-rule") || "nonzero";
  if (rule === "evenodd") {
    let parity = 0;
    for (const contour of contours) {
      const winding = contourWinding(contour, x, y);
      if (winding === Number.POSITIVE_INFINITY) return true;
      if (winding !== 0) parity ^= 1;
    }
    return parity === 1;
  }
  let winding = 0;
  for (const contour of contours) {
    const contribution = contourWinding(contour, x, y);
    if (contribution === Number.POSITIVE_INFINITY) return true;
    winding += contribution;
  }
  return winding !== 0;
}

function clipMembership(document: Document, id: string, x: number, y: number, stack = new Set<string>()): boolean {
  assert.ok(!stack.has(id), `clip reference cycle at ${id}`);
  const clip = Array.from(document.getElementsByTagName("clipPath")).find((node) => node.getAttribute("id") === id);
  assert.ok(clip, `unresolved clip reference ${id}`);
  stack.add(id);
  const result = childElements(clip).some((path) => {
    if (path.localName !== "path" || !pathContains(path, x, y)) return false;
    const nested = clipId(path);
    return nested === null || clipMembership(document, nested, x, y, stack);
  });
  stack.delete(id);
  return result;
}

function drawingClipContains(document: Document, drawingIndex: number, x: number, y: number): boolean {
  const drawing = groups(document)[drawingIndex];
  assert.ok(drawing, `missing drawing group ${drawingIndex}`);
  const id = clipId(drawing);
  return id === null || clipMembership(document, id, x, y);
}

function assertClipPoint(
  value: MetafileScene,
  point: readonly [number, number],
  expected: boolean,
  drawingIndex = 0,
): void {
  const { parsed } = xmlFor(value);
  assert.equal(drawingClipContains(parsed, drawingIndex, point[0], point[1]), expected, `clip membership at ${point.join(",")}`);
}

const allowedAttributes: Record<string, ReadonlySet<string>> = {
  svg: new Set(["xmlns", "viewBox", "width", "height"]),
  defs: new Set(),
  clipPath: new Set(["id", "clipPathUnits"]),
  path: new Set(["d", "transform", "fill", "fill-rule", "clip-rule", "clip-path", "stroke", "stroke-width", "stroke-linecap", "stroke-linejoin", "stroke-miterlimit", "fill-opacity", "stroke-opacity"]),
  rect: new Set(["x", "y", "width", "height", "transform", "fill", "fill-opacity", "clip-path"]),
  g: new Set(["transform", "clip-path"]),
  text: new Set(["x", "y", "dx", "dy", "transform", "font-family", "font-size", "font-weight", "font-style", "text-decoration", "direction", "unicode-bidi", "fill", "fill-opacity", "clip-path"]),
  tspan: new Set(["x", "y", "dx", "dy", "transform", "font-family", "font-size", "font-weight", "font-style", "text-decoration", "direction", "unicode-bidi", "fill", "fill-opacity", "clip-path"]),
};

function assertSafeSvgGrammar(document: Document): void {
  const ids = new Set<string>();
  const clipRefs = new Map<string, string[]>();
  const visit = (element: Element, insideClip: boolean): void => {
    const tag = element.localName;
    assert.ok(tag in allowedAttributes, `unexpected SVG tag ${tag}`);
    for (let index = 0; index < element.attributes.length; index += 1) {
      const attribute = element.attributes.item(index)!;
      assert.ok(allowedAttributes[tag].has(attribute.name), `unexpected ${tag} attribute ${attribute.name}`);
      assert.ok(!/^on/i.test(attribute.name), "event attributes are forbidden");
      if (attribute.name === "clip-path") {
        const target = clipId(element);
        assert.ok(target, "only internal fragment references are allowed");
        assert.ok((tag === "g" && !insideClip) || (tag === "path" && insideClip), "clip references are restricted to wrappers and definition literals");
        const source = insideClip ? element.parentElement?.getAttribute("id") : "drawing";
        if (source) clipRefs.set(source, [...(clipRefs.get(source) ?? []), target]);
      }
    }
    if (tag === "clipPath") {
      assert.equal(element.getAttribute("clipPathUnits"), "userSpaceOnUse");
      assert.equal(element.hasAttribute("transform"), false);
      const id = element.getAttribute("id");
      assert.ok(id && /^mf-clip-[1-9][0-9]*$/.test(id));
      assert.equal(ids.has(id), false, `duplicate id ${id}`);
      ids.add(id);
    }
    if (insideClip && tag === "path") {
      assert.equal(element.hasAttribute("transform"), false, "clip geometry must stay in device space");
      const rule = element.getAttribute("clip-rule");
      assert.ok(rule === "evenodd" || rule === "nonzero", "clip definition path needs an explicit clip-rule");
    }
    if (tag === "g" && element.hasAttribute("clip-path")) {
      assert.equal(element.hasAttribute("transform"), false, "device-space clipping wrapper cannot transform");
    }
    if (element.hasAttribute("transform")) {
      assert.equal(element.hasAttribute("clip-path"), false, "transformed drawing children cannot carry device-space clips");
    }
    for (const child of childElements(element)) visit(child, insideClip || tag === "clipPath");
  };
  visit(document.documentElement, false);
  for (const [source, targets] of clipRefs) {
    assert.ok(ids.has(source) || source === "drawing", `unknown ref source ${source}`);
    for (const target of targets) assert.ok(ids.has(target), `unresolved ref ${target}`);
  }
  const active = new Set<string>();
  const complete = new Set<string>();
  const check = (id: string): void => {
    if (complete.has(id)) return;
    assert.ok(!active.has(id), `cyclic clip reference ${id}`);
    active.add(id);
    for (const target of clipRefs.get(id) ?? []) check(target);
    active.delete(id);
    complete.add(id);
  };
  for (const id of ids) check(id);
}

function serializeFails(value: MetafileScene, code: string, operationIndex?: number): void {
  assert.throws(() => serializeMetafileScene(value), (error: unknown) => {
    assert.ok(error instanceof MetafileSerializeError);
    assert.equal(error.code, code);
    if (operationIndex !== undefined) assert.equal(error.operationIndex, operationIndex);
    return true;
  });
}

void test("serializes allowed SVG and independently preserves Boolean clip membership", () => {
  const first = region([rect(0, 0, 6, 6)]);
  const second = region([rect(3, 0, 6, 6)]);
  const samples: readonly [number, number][] = [[1, 2], [4, 2], [8, 2], [10, 2]];
  const truth: Record<"intersect" | "union" | "xor" | "exclude", readonly boolean[]> = {
    intersect: [false, true, false, false],
    union: [true, true, true, false],
    xor: [true, false, true, false],
    exclude: [true, false, false, false],
  };
  for (const mode of ["intersect", "union", "xor", "exclude"] as const) {
    const { parsed: document } = xmlFor(clipScene(combined(mode, first, second)));
    assertSafeSvgGrammar(document);
    samples.forEach((point, index) => {
      assert.equal(drawingClipContains(document, 0, point[0], point[1]), truth[mode][index], `${mode} at ${point.join(",")}`);
    });
  }
});

void test("handles all U/E identities with null-left and empty rectangle regions", () => {
  const a = region([rect(2, 2, 3, 3)]);
  const empty = region([]);
  const unbounded = combined("union", null, empty);
  const points: readonly [number, number][] = [[3, 3], [9, 9]];
  const cases: readonly ["intersect" | "union" | "xor" | "exclude", MetafileClipRegion | null, MetafileClipRegion, readonly boolean[]][] = [
    ["intersect", null, a, [true, false]], ["union", null, a, [true, true]], ["xor", null, a, [false, true]], ["exclude", null, a, [false, true]],
    ["intersect", empty, a, [false, false]], ["union", empty, a, [true, false]], ["xor", empty, a, [true, false]], ["exclude", empty, a, [false, false]],
    ["intersect", a, empty, [false, false]], ["union", a, empty, [true, false]], ["xor", a, empty, [true, false]], ["exclude", a, empty, [true, false]],
    ["intersect", a, unbounded, [true, false]], ["union", a, unbounded, [true, true]], ["xor", a, unbounded, [false, true]], ["exclude", a, unbounded, [false, false]],
  ];
  for (const [mode, left, right, expected] of cases) {
    const { parsed: document } = xmlFor(clipScene(combined(mode, left, right)));
    points.forEach((point, index) => assert.equal(drawingClipContains(document, 0, point[0], point[1]), expected[index], `${mode} at ${point.join(",")}`));
  }
  assertClipPoint(clipScene(region([])), [3, 3], false);
  assertClipPoint(clipScene(region([])), [9, 9], false);
});

void test("preserves every Boolean operator for disjoint, identical, touching and contained rectangles", () => {
  const cases: readonly {
    readonly name: string;
    readonly first: MetafileViewBox;
    readonly second: MetafileViewBox;
    readonly points: readonly (readonly [number, number])[];
  }[] = [
    { name: "disjoint", first: rect(0, 0, 2, 2), second: rect(5, 0, 2, 2), points: [[1, 1], [6, 1], [3, 1], [9, 9]] },
    { name: "identical", first: rect(0, 0, 4, 4), second: rect(0, 0, 4, 4), points: [[2, 2], [6, 6]] },
    { name: "touching", first: rect(0, 0, 2, 2), second: rect(2, 0, 2, 2), points: [[1, 1], [3, 1], [6, 1]] },
    { name: "contained", first: rect(0, 0, 6, 6), second: rect(2, 2, 2, 2), points: [[1, 1], [3, 3], [5, 5], [9, 9]] },
  ];
  for (const entry of cases) {
    const first = entry.first;
    const second = entry.second;
    for (const mode of ["intersect", "union", "xor", "exclude"] as const) {
      const value = clipScene(combined(mode, region([first]), region([second])));
      for (const point of entry.points) {
        const inFirst = point[0] > first.x && point[0] < first.x + first.width && point[1] > first.y && point[1] < first.y + first.height;
        const inSecond = point[0] > second.x && point[0] < second.x + second.width && point[1] > second.y && point[1] < second.y + second.height;
        const expected = mode === "intersect" ? inFirst && inSecond
          : mode === "union" ? inFirst || inSecond
            : mode === "xor" ? inFirst !== inSecond
              : inFirst && !inSecond;
        assertClipPoint(value, point, expected);
      }
    }
  }
  assertClipPoint(clipScene(combined("exclude", null, region([]))), [2, 2], true);
  assertClipPoint(clipScene(combined("exclude", null, region([]))), [12.1, 2], false);
  const rectangle = region([rect(2, 2, 4, 4)]);
  assertClipPoint(clipScene(combined("exclude", null, combined("exclude", null, rectangle))), [3, 3], true);
  assertClipPoint(clipScene(combined("exclude", null, combined("exclude", null, rectangle))), [9, 9], false);
});

void test("complements a rectangle leaf as a union, not a multi-contour parity path", () => {
  const q1 = rect(0, 0, 6, 6);
  const q2 = rect(3, 0, 6, 6);
  const overlap = region([q1, q2]);
  const outside = combined("exclude", null, overlap);
  assertClipPoint(clipScene(outside), [4, 2], false);
  assertClipPoint(clipScene(outside), [10, 2], true);
  for (const rectangles of [[q1, q1], [q1, q1, q2], [rect(0, 0, 9, 9), rect(2, 2, 2, 2)]] as const) {
    const complement = clipScene(combined("xor", null, region(rectangles)));
    const { parsed: document } = xmlFor(complement);
    assertSafeSvgGrammar(document);
    assert.equal(drawingClipContains(document, 0, 4, 2), false);
  }
  const { parsed: document } = xmlFor(clipScene(combined("exclude", null, overlap)));
  const negativePaths = Array.from(document.getElementsByTagName("path")).filter((path) => path.getAttribute("clip-rule") === "evenodd");
  assert.ok(negativePaths.length >= 2);
  for (const path of negativePaths) assert.equal(polygonsFromPath(path.getAttribute("d") ?? "").length, 2, "each complement literal contains V and exactly one quad");
});

void test("applies sequential path clip state, replacement snapshots, null reset and restore", () => {
  const pathA: readonly MetafilePathCommand[] = [
    { type: "moveTo", point: { x: 0, y: 0 } }, { type: "lineTo", point: { x: 6, y: 0 } },
    { type: "lineTo", point: { x: 6, y: 6 } }, { type: "lineTo", point: { x: 0, y: 6 } }, { type: "closePath" },
  ];
  const pathB: readonly MetafilePathCommand[] = [
    { type: "moveTo", point: { x: 3, y: 0 } }, { type: "lineTo", point: { x: 9, y: 0 } },
    { type: "lineTo", point: { x: 9, y: 6 } }, { type: "lineTo", point: { x: 3, y: 6 } }, { type: "closePath" },
  ];
  const savedRegion = region([rect(8, 8, 2, 2)]);
  const operations: MetafileSceneOperation[] = [
    { kind: "clip", action: "set", commands: pathA, fillRule: "nonZero" }, filledPath(),
    { kind: "clip", action: "intersect", commands: pathB, fillRule: "nonZero" }, filledPath(),
    { kind: "clip", action: "exclude", commands: pathB, fillRule: "evenOdd" }, filledPath(),
    { kind: "clip", action: "region", region: savedRegion }, filledPath(),
    { kind: "clip", action: "region", region: null }, filledPath(),
    { kind: "clip", action: "reset" }, filledPath(),
    { kind: "clip", action: "region", region: savedRegion }, filledPath(),
    { kind: "clip", action: "region", region: null }, filledPath(),
  ];
  const { parsed: document } = xmlFor(scene(operations));
  assert.equal(drawingClipContains(document, 0, 2, 2), true);
  assert.equal(drawingClipContains(document, 1, 2, 2), false);
  assert.equal(drawingClipContains(document, 1, 4, 2), true);
  assert.equal(drawingClipContains(document, 2, 4, 2), false);
  assert.equal(drawingClipContains(document, 3, 9, 9), true);
  assert.equal(drawingClipContains(document, 3, 4, 2), false);
  assert.equal(drawingClipContains(document, 4, 4, 2), true);
  assert.equal(drawingClipContains(document, 5, 4, 2), true);
  assert.equal(drawingClipContains(document, 6, 9, 9), true, "restoring a saved region replaces the prior null snapshot");
  assert.equal(drawingClipContains(document, 7, 9, 9), true, "a later null snapshot restores unbounded clipping");
});

void test("maps transformed device rectangles once and keeps clip fixed under drawing transforms", () => {
  const captured = { a: 0, b: 1, c: -1, d: 0, e: 4, f: 2 };
  const clip = region([rect(1, 2, 4, 3)], captured);
  const drawingTransform = { a: 2, b: 0, c: 0, d: 2, e: 10, f: 10 };
  const value = scene([{ kind: "clip", action: "region", region: clip }, filledPath(undefined, drawingTransform)]);
  const { parsed: document } = xmlFor(value);
  assertSafeSvgGrammar(document);
  // Source rectangle corners map to (2,3), (2,7), (-1,7), (-1,3).
  assert.equal(drawingClipContains(document, 0, 1, 4), true);
  assert.equal(drawingClipContains(document, 0, 3, 4), false);
  const wrapper = groups(document)[0];
  assert.equal(wrapper.hasAttribute("transform"), false);
  const child = childElements(wrapper).find((element) => element.localName === "path");
  assert.ok(child);
  assert.equal(child.getAttribute("transform"), "matrix(2 0 0 2 10 10)");
});

void test("retains reflected and sheared quadrilateral geometry in device space", () => {
  const reflected = clipScene(region([rect(1, 1, 2, 2)], { a: -1, b: 0, c: 0, d: 1, e: 5, f: 0 }));
  assertClipPoint(reflected, [3, 2], true);
  assertClipPoint(reflected, [1, 2], false);
  const sheared = clipScene(region([rect(1, 1, 2, 2)], { a: 1, b: 0, c: 1, d: 1, e: 0, f: 0 }));
  assertClipPoint(sheared, [4, 2], true);
  assertClipPoint(sheared, [2.5, 2.5], false);
});

void test("applies nested device offsets before final viewBox clipping", () => {
  let clip: MetafileClipRegion = region([rect(12, 2, 2, 2)]);
  clip = { kind: "offset", region: clip, x: -4, y: 0 };
  assertClipPoint(clipScene(clip), [9, 3], true);
  assertClipPoint(clipScene(clip), [11, 3], false);
});

void test("serializes escaped text and font names while preserving opaque background order and temporary clip", () => {
  const textOperation: MetafileSceneOperation = {
    kind: "text", x: 2, y: 4, text: `A & <B> "C"`, fontFamily: `Font & <Family> "X"`, fontSize: 2,
    color: RED, direction: "ltr", advances: Array.from(`A & <B> "C"`, () => 1),
    textAlign: 0x18,
    transform: IDENTITY, backgroundMode: "opaque", backgroundColor: { red: 255, green: 255, blue: 0, alpha: 1 },
    clipRect: rect(1, 1, 3, 3), opaqueRect: rect(0, 0, 8, 5),
  };
  const { xml, parsed: document } = xmlFor(scene([textOperation]));
  assertSafeSvgGrammar(document);
  assert.match(xml, /A &amp; &lt;B&gt; &quot;C&quot;/);
  assert.match(xml, /Font &amp; &lt;Family&gt; &quot;X&quot;/);
  const texts = Array.from(document.getElementsByTagName("text"));
  assert.equal(texts.length, 1);
  assert.equal(texts[0].textContent, `A & <B> "C"`);
  const paths = Array.from(document.getElementsByTagName("path"));
  assert.ok(paths.length >= 1, "opaque text background is painted as a static rectangle before glyphs");
  const wrapper = groups(document)[0];
  const renderedChildren = childElements(wrapper);
  const backgroundIndex = renderedChildren.findIndex((element) => element.localName === "rect" || element.localName === "path");
  const textClipIndex = renderedChildren.findIndex((element) => element.localName === "g");
  assert.ok(backgroundIndex >= 0 && textClipIndex > backgroundIndex);
  const background = renderedChildren[backgroundIndex];
  assert.equal(pathContains(background, 7, 4), true, "opaqueRect extends beyond the glyph-only clipRect");
  assert.equal(background.hasAttribute("clip-path"), false, "the temporary glyph clip must not be attached to the background");
  const textClipGroup = renderedChildren[textClipIndex];
  assert.equal(childElements(textClipGroup)[0]?.localName, "text");
  assert.equal(clipMembership(document, requiredClipId(textClipGroup), 2, 2), true);
  assert.equal(clipMembership(document, requiredClipId(textClipGroup), 7, 4), false);
  const { parsed: nextDocument } = xmlFor(scene([textOperation, filledPath()]));
  assert.equal(drawingClipContains(nextDocument, 1, 7, 4), true, "text clipRect must not leak to the next drawing");
});

void test("positions each glyph at its preceding source advance for LTR and RTL text", () => {
  const ltr = { ...textRun("AB", [3, 7]), x: 2, y: 4 };
  const rtl = { ...textRun("AB", [-3, -7]), x: 2, y: 4, direction: "rtl" as const, textAlign: 0x118 };
  const ltrText = Array.from(xmlFor(scene([ltr])).parsed.getElementsByTagName("text"))[0];
  const rtlText = Array.from(xmlFor(scene([rtl])).parsed.getElementsByTagName("text"))[0];
  assert.ok(ltrText && rtlText);
  assert.equal(ltrText.getAttribute("x"), "2 5", "the first glyph stays at the source origin; only the next glyph advances");
  assert.equal(rtlText.getAttribute("x"), "2 -1", "signed RTL displacement positions the next glyph without shifting the current one");
  assert.equal(ltrText.hasAttribute("dx"), false, "source current-to-next displacements are not SVG current-glyph dx offsets");
  assert.equal(rtlText.hasAttribute("dx"), false);
});

void test("rejects whitespace with unproven SVG glyph mapping at the owning operation", () => {
  for (const text of [" A", "A ", " ", "A  B", " A  B ", "A\tB", "A\nB", "A\rB", "A\r\nB", "A\u00a0B", "A\u2003B", "A\u2028B"]) {
    serializeFails(scene([filledPath(), textRun(text, Array.from(text, () => 3))]), "unsupported-text", 1);
    serializeFails(scene([textRun(text)]), "unsupported-text", 0);
  }
});

void test("retains single internal spaces and escaped text with one position per character", () => {
  const text = 'A &<>" B';
  const { parsed } = xmlFor(scene([textRun(text, Array.from(text, () => 3))]));
  assertSafeSvgGrammar(parsed);
  const element = parsed.getElementsByTagName("text")[0];
  assert.ok(element);
  assert.equal(element.textContent, text);
  assert.equal(element.getAttribute("x"), "1 4 7 10 13 16 19 22");
});

void test("supports only exact left TA_BASELINE alignment with optional RTL reading", () => {
  for (const [textAlign, direction] of [[0x18, "ltr"], [0x18, "rtl"], [0x118, "rtl"]] as const) {
    const operation = { ...textRun("AB", [3, 7]), textAlign, direction };
    const rendered = Array.from(xmlFor(scene([operation])).parsed.getElementsByTagName("text"))[0];
    assert.ok(rendered);
    assert.equal(rendered.getAttribute("y"), "2");
    assert.equal(rendered.getAttribute("direction"), direction);
  }
  for (const [textAlign, direction] of [
    [0, "ltr"], [8, "ltr"], [2, "ltr"], [6, "ltr"], [0x118, "ltr"], [0x19, "ltr"],
  ] as const) {
    serializeFails(scene([{ ...textRun("AB", [3, 7]), textAlign, direction }]), "unsupported-text", 0);
  }
  const missingAlignment = { ...textRun("AB", [3, 7]) };
  delete (missingAlignment as { textAlign?: number }).textAlign;
  serializeFails(scene([missingAlignment]), "unsupported-text", 0);
});

void test("preserves ETO_OPAQUE fill when background mode is transparent", () => {
  const operation: MetafileSceneOperation = {
    kind: "text", x: 2, y: 4, text: "AB", fontFamily: "Test Font", fontSize: 1,
    color: RED, direction: "ltr", textAlign: 0x18, textOptions: 6, backgroundMode: "transparent",
    backgroundColor: { red: 17, green: 34, blue: 51, alpha: 1 },
    clipRect: rect(1, 1, 3, 3), opaqueRect: rect(0, 0, 8, 5),
  };
  const { parsed } = xmlFor(scene([operation]));
  const drawing = groups(parsed)[0];
  assert.ok(drawing);
  const children = childElements(drawing);
  assert.equal(children[0]?.localName, "path", "opaque fill is emitted before glyphs");
  assert.equal(children[0]?.getAttribute("fill"), "rgb(17,34,51)");
  assert.equal(children[0]?.hasAttribute("clip-path"), false);
  assert.equal(pathContains(children[0], 7, 4), true, "opaque fill retains its requested extent outside the glyph clip");
  assert.equal(children[1]?.localName, "g", "ETO_CLIPPED applies only to glyphs");
  const textClipId = requiredClipId(children[1]);
  assert.equal(clipMembership(parsed, textClipId, 2, 2), true);
  assert.equal(clipMembership(parsed, textClipId, 7, 4), false);
  assert.equal(childElements(children[1])[0]?.localName, "text");
});

void test("specified zero-area and singular text clips remain empty instead of becoming absent", () => {
  const cases: readonly [string, MetafileSceneOperation][] = [
    ["zero width", { ...textRun("AB", [1, 1]), clipRect: rect(1, 1, 0, 3) }],
    ["zero height", { ...textRun("AB", [1, 1]), clipRect: rect(1, 1, 3, 0) }],
    ["singular transform", {
      ...textRun("AB", [1, 1]), clipRect: rect(1, 1, 3, 3),
      transform: { a: 1, b: 2, c: 2, d: 4, e: 0, f: 0 },
    }],
  ];
  for (const [name, operation] of cases) {
    const { parsed } = xmlFor(scene([operation]));
    const drawing = groups(parsed)[0];
    assert.ok(drawing, `${name}: missing drawing wrapper`);
    const textClip = childElements(drawing)[0];
    assert.ok(textClip?.localName === "g", `${name}: specified clip must wrap the glyphs`);
    const id = requiredClipId(textClip);
    const definition = Array.from(parsed.getElementsByTagName("clipPath")).find((clip) => clip.getAttribute("id") === id);
    assert.ok(definition, `${name}: missing empty clip definition`);
    assert.equal(childElements(definition).length, 0, `${name}: empty clip must contain no drawable region`);
    assert.equal(clipMembership(parsed, id, 2, 2), false, `${name}: glyphs must not escape an empty clip`);
  }
});

void test("serializes cubic path data with valid separators and all six coordinates", () => {
  const commands: readonly MetafilePathCommand[] = [
    { type: "moveTo", point: { x: 1, y: 2 } },
    {
      type: "cubicTo", control1: { x: 3, y: 4 }, control2: { x: 5, y: 6 }, point: { x: 7, y: 8 },
    },
    { type: "closePath" },
  ];
  const { parsed } = xmlFor(scene([filledPath(commands)]));
  const path = Array.from(parsed.getElementsByTagName("path")).find((element) => element.getAttribute("d")?.includes("C"));
  assert.ok(path);
  assert.equal(path.getAttribute("d"), "M1 2C3 4 5 6 7 8Z");
  assert.deepEqual(pathSegments(path.getAttribute("d") ?? ""), [
    { type: "M", coordinates: [1, 2] },
    { type: "C", coordinates: [3, 4, 5, 6, 7, 8] },
    { type: "Z", coordinates: [] },
  ]);
});

void test("retains cubic commands in an evenodd complement definition path", () => {
  const cubicContour: readonly MetafilePathCommand[] = [
    { type: "moveTo", point: { x: 2, y: 2 } },
    {
      type: "cubicTo", control1: { x: 2, y: 6 }, control2: { x: 6, y: 6 }, point: { x: 6, y: 2 },
    },
    { type: "closePath" },
  ];
  const { parsed } = xmlFor(scene([
    { kind: "clip", action: "exclude", commands: cubicContour, fillRule: "evenOdd" },
    filledPath(),
  ]));
  const complement = Array.from(parsed.getElementsByTagName("path")).find((path) =>
    path.getAttribute("clip-rule") === "evenodd" && path.getAttribute("d")?.includes("C"));
  assert.ok(complement, "cubic contour remains in a definition-side evenodd complement path");
  const segments = pathSegments(complement.getAttribute("d") ?? "");
  assert.equal(segments.filter((segment) => segment.type === "M").length, 2, "complement contains V and the source contour");
  assert.equal(segments.filter((segment) => segment.type === "C").length, 1, "source cubic is not flattened or dropped");
  assert.equal(complement.getAttribute("clip-rule"), "evenodd");
});

void test("definition-side evenodd negative controls fail the expected inner-hole membership", () => {
  const makeDocument = (): Document => xmlFor(clipScene(combined("exclude", null, region([rect(2, 2, 4, 4)])))).parsed;
  const cases: readonly [string, (document: Document, definitionPath: Element, consumer: Element) => void][] = [
    ["missing clip-rule", (_document, path) => path.removeAttribute("clip-rule")],
    ["fill-rule only", (_document, path) => { path.removeAttribute("clip-rule"); path.setAttribute("fill-rule", "evenodd"); }],
    ["consumer clip-rule only", (_document, path, consumer) => {
      path.removeAttribute("clip-rule");
      consumer.setAttribute("clip-rule", "evenodd");
    }],
  ];
  for (const [name, mutate] of cases) {
    const document = makeDocument();
    const definitionPath = Array.from(document.getElementsByTagName("path")).find((path) => path.getAttribute("clip-rule") === "evenodd");
    const consumer = groups(document)[0];
    assert.ok(definitionPath && consumer, `${name}: missing positive-control nodes`);
    assert.equal(drawingClipContains(document, 0, 3, 3), false, `${name}: valid positive control excludes the inner point`);
    mutate(document, definitionPath, consumer);
    assert.equal(drawingClipContains(document, 0, 3, 3), true, `${name}: default nonzero incorrectly fills the inner point`);
  }
});

void test("maps temporary text clipping and opaque backgrounds through the captured transform once", () => {
  const operation: MetafileSceneOperation = {
    kind: "text", x: 1, y: 1, text: "x", fontFamily: "Test Font", fontSize: 1, color: RED, direction: "ltr",
    textAlign: 0x18,
    transform: { a: 0, b: 1, c: -1, d: 0, e: 6, f: 0 },
    clipRect: rect(1, 1, 2, 2), opaqueRect: rect(0, 0, 4, 2),
    backgroundMode: "opaque", backgroundColor: RED,
  };
  const { parsed } = xmlFor(scene([operation]));
  assertSafeSvgGrammar(parsed);
  const wrapper = groups(parsed)[0];
  const children = childElements(wrapper);
  const background = children[0];
  const textClip = children[1];
  assert.equal(background.localName, "path");
  assert.equal(pathContains(background, 5, 3), true);
  assert.equal(pathContains(background, 3, 3), false);
  assert.equal(clipMembership(parsed, requiredClipId(textClip), 4.5, 2), true);
  assert.equal(clipMembership(parsed, requiredClipId(textClip), 2, 2), false);
  assert.equal(children[0]?.hasAttribute("transform"), false, "opaque geometry is already in device coordinates");
  assert.equal(childElements(textClip)[0]?.getAttribute("transform"), "matrix(0 1 -1 0 6 0)");
});

void test("requires definition-side evenodd clip-rule for a same-winding complement hole", () => {
  const { parsed: document } = xmlFor(clipScene(combined("exclude", null, region([rect(2, 2, 4, 4)]))));
  assertSafeSvgGrammar(document);
  const complement = Array.from(document.getElementsByTagName("path")).find((path) => path.getAttribute("clip-rule") === "evenodd");
  assert.ok(complement);
  assert.equal(pathContains(complement, 3, 3), false);
  assert.equal(pathContains(complement, 1, 1), true);
});

void test("rejects unproven nonzero path negation and arithmetic-induced coordinate collapse", () => {
  const openPath: readonly MetafilePathCommand[] = [
    { type: "moveTo", point: { x: 1, y: 1 } }, { type: "lineTo", point: { x: 4, y: 1 } },
    { type: "lineTo", point: { x: 4, y: 4 } }, { type: "closePath" },
  ];
  serializeFails(scene([{ kind: "clip", action: "exclude", commands: openPath, fillRule: "nonZero" }, filledPath()]), "unsupported-clip-expression", 0);
  serializeFails(clipScene(region([rect(0, 0, 1, 1)], { ...IDENTITY, e: 2 ** 53 })), "unsupported-clip-expression", 0);
  serializeFails(scene([filledPath(undefined, { ...IDENTITY, e: 2 ** 53 })]), "unsupported-clip-expression", 0);
});

void test("distinguishes true zero-area and singular rectangles from collapsed arithmetic", () => {
  const zeroWidth = clipScene(region([rect(2, 2, 0, 4)]));
  const singular = clipScene(region([rect(2, 2, 4, 4)], { a: 1, b: 2, c: 2, d: 4, e: 0, f: 0 }));
  assertClipPoint(zeroWidth, [2, 2], false);
  assertClipPoint(singular, [3, 3], false);
});

void test("enforces fixed-grid bounds, exact fractions, offset depth and raw numeric validation", () => {
  assertClipPoint(clipScene(region([rect(0, 0, 1 / 1024, 1)])), [1 / 2048, 1 / 2], true);
  serializeFails(clipScene(region([rect(0, 0, 16384 + 1 / 1024, 1)])), "unsupported-clip-expression", 0);
  serializeFails(clipScene(region([rect(Number.NaN, 0, 1, 1)])), "invalid-scene", 0);
  serializeFails(clipScene(region([rect(0, 0, 1, 1)], { ...IDENTITY, a: Number.POSITIVE_INFINITY })), "invalid-scene", 0);
  let allowed: MetafileClipRegion = region([rect(1, 1, 1, 1)]);
  for (let index = 0; index < 128; index += 1) allowed = { kind: "offset", region: allowed, x: 1 / 1024, y: 0 };
  assertClipPoint(clipScene(allowed), [1.125, 1.5], true);
  const exceeded = { kind: "offset", region: allowed, x: 1 / 1024, y: 0 } as MetafileClipRegion;
  serializeFails(clipScene(exceeded), "max-clip-depth", 0);
});

void test("rejects forged cycles, excessive depth and invalid scene discriminants before output", () => {
  const cycle: { kind: "offset"; region: MetafileClipRegion; x: number; y: number } = {
    kind: "offset", region: region([]), x: 0, y: 0,
  };
  cycle.region = cycle;
  serializeFails(clipScene(cycle), "invalid-scene", 0);
  let deep: MetafileClipRegion = region([]);
  for (let index = 0; index < 129; index += 1) deep = { kind: "offset", region: deep, x: 0, y: 0 };
  serializeFails(clipScene(deep), "max-clip-depth", 0);
  serializeFails(clipScene({ kind: "combine", mode: "intersect", left: null, right: null } as unknown as MetafileClipRegion), "invalid-scene", 0);
  serializeFails(clipScene({ kind: "offset", region: null, x: 0, y: 0 } as unknown as MetafileClipRegion), "invalid-scene", 0);
  serializeFails(scene([{ kind: "raw-svg" } as unknown as MetafileSceneOperation]), "invalid-scene", 0);
});

void test("rejects oversized arrays before enumerating their entries", () => {
  let enumerated = false;
  const oversized = new Proxy(new Array<MetafileSceneOperation>(50_001), {
    ownKeys(target) {
      enumerated = true;
      return Reflect.ownKeys(target);
    },
  });
  serializeFails({ ...scene([]), operations: oversized }, "max-svg-nodes", -1);
  assert.equal(enumerated, false);
});

void test("enforces viewBox pixel, rectangle and DNF expansion bounds", () => {
  serializeFails(scene([], { x: 0, y: 0, width: 4096, height: 4096 }), "invalid-scene", -1);
  serializeFails(clipScene(region(Array.from({ length: 4097 }, (_, index) => rect(index, 0, 1, 1)))), "max-region-rectangles", 0);
  const many = region(Array.from({ length: 4096 }, (_, index) => rect(index, 0, 0.5, 0.5)));
  const product = combined("intersect", many, many);
  serializeFails(clipScene(product), "max-clip-nodes", 0);
});

void test("pins exact scene, path and text length boundaries", () => {
  const reset = { kind: "clip", action: "reset" } as const;
  assert.doesNotThrow(() => serializeMetafileScene(scene(Array.from({ length: 50_000 }, () => reset))));
  serializeFails(scene(Array.from({ length: 50_001 }, () => reset)), "max-svg-nodes", -1);

  const commands = Array.from({ length: 50_000 }, () => ({ type: "moveTo", point: { x: 0, y: 0 } } as const));
  assert.doesNotThrow(() => serializeMetafileScene(scene([filledPath(commands)])));
  serializeFails(scene([filledPath([...commands, { type: "moveTo", point: { x: 0, y: 0 } }])]), "max-svg-nodes", 0);

  assert.doesNotThrow(() => serializeMetafileScene(scene([textRun("x".repeat(8_192))])));
  serializeFails(scene([textRun("x".repeat(8_193))]), "max-svg-nodes", 0);
});

void test("pins per-leaf and aggregate rectangle caps plus total clip-node expansion", () => {
  const rects4096 = Array.from({ length: 4_096 }, (_, index) => rect(index, 0, 1, 1));
  assert.doesNotThrow(() => serializeMetafileScene(scene([regionOperation(region(rects4096))])));
  serializeFails(scene([regionOperation(region([...rects4096, rect(4_096, 0, 1, 1)]))]), "max-region-rectangles", 0);

  const firstHalf = region(Array.from({ length: 2_048 }, (_, index) => rect(index, 0, 1, 1)));
  const secondHalf = region(Array.from({ length: 2_048 }, (_, index) => rect(index + 2_048, 0, 1, 1)));
  assert.doesNotThrow(() => serializeMetafileScene(scene([regionOperation(firstHalf), regionOperation(secondHalf)])));
  const overflow = region(Array.from({ length: 2_049 }, (_, index) => rect(index + 2_048, 0, 1, 1)));
  serializeFails(scene([regionOperation(firstHalf), regionOperation(overflow)]), "max-region-rectangles", 1);

  const atLimit = combined("intersect", null, balancedEmptyRegionTree(25_000));
  assert.doesNotThrow(() => serializeMetafileScene(scene([regionOperation(atLimit)])));
  const overLimit = balancedEmptyRegionTree(25_001);
  serializeFails(scene([regionOperation(overLimit)]), "max-clip-nodes", 0);

  let shared: MetafileClipRegion = region([]);
  for (let depth = 0; depth < 15; depth += 1) shared = combined("intersect", shared, shared);
  serializeFails(scene([regionOperation(shared)]), "max-clip-nodes", 0);
});

void test("pins cumulative DNF-term work at the exact cap", () => {
  let leafCount = 3_606;
  let base = balancedUniverseUnionTree(leafCount, 0);
  while (base.termWork > 50_000) {
    leafCount -= 1;
    base = balancedUniverseUnionTree(leafCount, 0);
  }
  const extraWork = 50_000 - base.termWork;
  assert.ok(extraWork >= 0 && extraWork < leafCount, `base DNF term work: ${base.termWork}`);
  const atLimit = balancedUniverseUnionTree(leafCount, extraWork);
  assert.equal(atLimit.termWork, 50_000);
  assert.doesNotThrow(() => serializeMetafileScene(scene([regionOperation(atLimit.region)])));
  const overLimit = balancedUniverseUnionTree(leafCount, extraWork + 1);
  assert.equal(overLimit.termWork, 50_001);
  serializeFails(scene([regionOperation(overLimit.region)]), "max-clip-nodes", 0);
});

void test("pins generated SVG element-count boundary", () => {
  const minimal = filledPath([{ type: "moveTo", point: { x: 0, y: 0 } }]);
  assert.doesNotThrow(() => serializeMetafileScene(scene(Array.from({ length: 8_333 }, () => minimal))));
  serializeFails(scene(Array.from({ length: 8_334 }, () => minimal)), "max-svg-nodes");
});

void test("pins exact viewBox pixel and fixed-grid coordinate boundaries", () => {
  assert.doesNotThrow(() => serializeMetafileScene(scene([], { x: 0, y: 0, width: 4_000, height: 4_000 })));
  serializeFails(scene([], { x: 0, y: 0, width: 4_000 + 1 / 1_024, height: 4_000 }), "invalid-scene", -1);
  assert.doesNotThrow(() => serializeMetafileScene(scene([], { x: 0, y: 0, width: 16_384, height: 1 })));
  serializeFails(scene([], { x: 0, y: 0, width: 16_384 + 1 / 1_024, height: 1 }), "unsupported-clip-expression", -1);
});

void test("pins the scalar numeric-work budget at its exact boundary", () => {
  const large = region(Array.from({ length: 3_840 }, (_, index) => rect(index, 0, 1, 1)));
  const medium = region(Array.from({ length: 89 }, (_, index) => rect(index, 0, 1, 1)));
  const empty = region([]);
  const prefix = [
    regionOperation(large), regionOperation(large), regionOperation(large), regionOperation(large),
    regionOperation(medium), regionOperation(empty),
  ];
  const textLengths = [8_192, 8_192, 8_192, 7_877];
  const atLimit = textLengths.map((length) => textRun("x".repeat(length), Array.from({ length }, () => 1)));
  assert.doesNotThrow(() => serializeMetafileScene(scene([...prefix, ...atLimit])));
  const last = textLengths[textLengths.length - 1];
  const overLimit = [...atLimit.slice(0, -1), textRun("x".repeat(last), Array.from({ length: last }, () => 1), [1])];
  serializeFails(scene([...prefix, ...overLimit]), "unsupported-clip-expression", 4);
});

void test("accounts escaped UTF-8 bytes exactly at the output ceiling", () => {
  const operationCount = 339;
  const fixedCount = operationCount - 1;
  const empty = Array.from({ length: operationCount }, () => textRun(""));
  const baselineBytes = serializeMetafileScene(scene(empty)).byteLength;
  const fixedText = '"'.repeat(8_192);
  const adjustableBytes = 16 * 1024 * 1024 - baselineBytes - fixedCount * 8_192 * 6;
  assert.ok(adjustableBytes > 0 && adjustableBytes <= 8_192 * 6, `adjustable output bytes: ${adjustableBytes}`);
  const finalText = textForEscapedByteLength(adjustableBytes);
  const exact = [...Array.from({ length: fixedCount }, () => textRun(fixedText)), textRun(finalText)];
  assert.equal(serializeMetafileScene(scene(exact)).byteLength, 16 * 1024 * 1024);
  const over = [...exact.slice(0, -1), textRun(`${finalText}x`)];
  serializeFails(scene(over), "max-output-bytes", -1);
});

void test("accounts UTF-8 output bytes, rejects malformed XML text and does not return partial output", () => {
  const valid: MetafileSceneOperation = {
    kind: "text", x: 1, y: 2, text: "文🙂", fontFamily: "宋体", fontSize: 1, color: RED, direction: "ltr",
    textAlign: 0x18,
  };
  const bytes = serializeMetafileScene(scene([valid]));
  assert.equal(new TextDecoder("utf-8", { fatal: true }).decode(bytes), new TextDecoder().decode(bytes));
  serializeFails(scene([{ ...valid, text: `\ud800` }]), "invalid-scene", 0);
  serializeFails(scene([{ ...valid, text: "x".repeat(8193) }]), "max-svg-nodes", 0);
});

void test("validates nested clip references and deterministic IDs", () => {
  const { parsed: document } = xmlFor(clipScene(combined("intersect", region([rect(1, 1, 4, 4)]), region([rect(2, 2, 4, 4)]))));
  assertSafeSvgGrammar(document);
  const ids = Array.from(document.getElementsByTagName("clipPath"), (element) => element.getAttribute("id"));
  assert.equal(new Set(ids).size, ids.length);
  assert.ok(ids.every((id) => /^mf-clip-\d+$/.test(id ?? "")));
});
