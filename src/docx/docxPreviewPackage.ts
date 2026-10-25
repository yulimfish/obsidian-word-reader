import type JSZip from "jszip";

import { childElements, parseXml } from "../pptx/xml";
import {
  isCanonicalPreviewPath,
  readOoxmlPreviewMetadata,
} from "../ooxml/equationPreviews";
import { loadOoxmlZip } from "../ooxml/jszipReader";
import { enforceOoxmlPackagePolicy } from "../ooxml/packagePolicy";
import {
  DEFAULT_DOCX_ZIP_LIMITS,
  validateZipSafety,
  type ZipSafetyEntry,
} from "../ooxml/packageSafety";
import {
  createPreviewMediaResolver,
  type ResolvedPreviewMedia,
} from "../ooxml/previewMedia";
import type {
  PreviewInventory,
  PreviewIssue,
  PreviewOccurrence,
} from "../ooxml/previewTypes";
import { writeStoredZip, type StoredZipEntry } from "../ooxml/storedZipWriter";

const WORD = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const OFFICE = "urn:schemas-microsoft-com:office:office";
const VML = "urn:schemas-microsoft-com:vml";
const DRAWING = "http://schemas.openxmlformats.org/drawingml/2006/main";
const WORD_DRAWING = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const PICTURE = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PACKAGE_REL = "http://schemas.openxmlformats.org/package/2006/relationships";
const CONTENT_TYPES = "http://schemas.openxmlformats.org/package/2006/content-types";
const MARKUP_COMPATIBILITY = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const XMLNS = "http://www.w3.org/2000/xmlns/";
const XML = "http://www.w3.org/XML/1998/namespace";
const EMU_PER_POINT = 12_700;
const EMU_PER_INCH = 914_400;
const EMU_PER_CM = 360_000;
const EMU_PER_MM = 36_000;
const EMU_PER_PIXEL = 9_525;
const MAX_ENTRY_BYTES = DEFAULT_DOCX_ZIP_LIMITS.maxEntryUncompressedBytes;
const MAX_TOTAL_BYTES = DEFAULT_DOCX_ZIP_LIMITS.maxTotalUncompressedBytes;

export interface DocxPreviewPackageOptions {
  readonly isCancelled: () => boolean;
  readonly yieldControl: () => Promise<void>;
}

export interface PreparedDocxPreviewPackage {
  readonly buffer: ArrayBuffer;
  readonly inventory: PreviewInventory;
  readonly issues: readonly PreviewIssue[];
}

export type DocxPreviewPackageErrorCode =
  | "cancelled"
  | "limit-exceeded"
  | "invalid-package"
  | "unsupported-placement";

export class DocxPreviewPackageError extends Error {
  constructor(readonly code: DocxPreviewPackageErrorCode, message: string) {
    super(message);
    this.name = "DocxPreviewPackageError";
  }
}

interface RelationshipRecord {
  readonly id: string;
  readonly type: string;
  readonly target: string;
  readonly external: boolean;
  readonly path: string | null;
}

interface RelationshipDocument {
  readonly path: string;
  readonly source: string;
  readonly xmlDocument: Document;
  readonly records: ReadonlyMap<string, RelationshipRecord>;
}

interface LengthPosition {
  readonly relativeFrom: string;
  readonly align: string | null;
  readonly offset: number;
}

interface VmlPlacement {
  readonly kind: "inline" | "anchor";
  readonly width: number;
  readonly height: number;
  readonly rotation: number;
  readonly crop: { left: number; top: number; right: number; bottom: number };
  readonly horizontal: LengthPosition;
  readonly vertical: LengthPosition;
  readonly relativeHeight: number;
  readonly behindDocument: boolean;
  readonly wrap: "none" | "square" | "topAndBottom";
  readonly distances: { top: number; bottom: number; left: number; right: number };
}

interface LocatedOccurrence {
  readonly occurrence: PreviewOccurrence;
  readonly object: Element;
  readonly placement: VmlPlacement;
}

interface MediaPart {
  readonly path: string;
  readonly media: ResolvedPreviewMedia;
  readonly contentType: string;
}

function fail(code: DocxPreviewPackageErrorCode, message: string): never {
  throw new DocxPreviewPackageError(code, message);
}

function checkCancelled(options: DocxPreviewPackageOptions): void {
  if (options.isCancelled()) {
    fail("cancelled", "DOCX preview preparation was cancelled.");
  }
}

function isEmbeddingPath(path: string): boolean {
  return /(?:^|\/)(?:embeddings|oleobjects)(?:\/|$)/i.test(path);
}

function isRelationshipMetadataPath(path: string): boolean {
  return path === "_rels/.rels" ||
    (isCanonicalPreviewPath(path) && /(?:^|\/)_rels\/[^/]+\.rels$/.test(path));
}

function sourceForRelationshipPath(path: string): string {
  if (path === "_rels/.rels") return "";
  const match = /^(.*\/)?_rels\/([^/]+)\.rels$/.exec(path);
  if (!match) return "";
  return `${match[1] ?? ""}${match[2]}`;
}

function checkedAdd(total: number, value: number, maximum: number, label: string): number {
  if (!Number.isSafeInteger(total) || !Number.isSafeInteger(value) || value < 0 || total > maximum - value) {
    fail("limit-exceeded", `${label} exceeds the supported package limit.`);
  }
  return total + value;
}

function utf8(bytes: Uint8Array, path: string): string {
  try {
    return new TextDecoder("utf-8", { fatal: true }).decode(bytes);
  } catch {
    fail("invalid-package", `The package entry ${path} is not valid UTF-8 XML.`);
  }
}

function serialize(document: Document): string {
  return new XMLSerializer().serializeToString(document);
}

/** Admission stays private: only this T8 route opts into validated legacy equation previews. */
async function validateDocxPreviewSource(zip: JSZip): Promise<PreviewInventory> {
  const inventory = await enforceOoxmlPackagePolicy(zip, "docx", {
    legacyEquationPreviews: true,
  });
  if (!inventory) {
    fail("invalid-package", "DOCX equation preview admission returned no inventory.");
  }
  return inventory;
}

function relationshipDocuments(
  summaryEntries: readonly ZipSafetyEntry[],
  zip: JSZip,
  metadata: Awaited<ReturnType<typeof readOoxmlPreviewMetadata>>,
  omittedPaths: Set<string>,
  options: DocxPreviewPackageOptions,
): Promise<Map<string, RelationshipDocument>> {
  return (async () => {
    const result = new Map<string, RelationshipDocument>();
    for (const entry of summaryEntries) {
      if (!isRelationshipMetadataPath(entry.name)) continue;
      checkCancelled(options);
      const source = sourceForRelationshipPath(entry.name);
      const records = metadata.relationships.get(source);
      const zipEntry = zip.file(entry.name);
      if (!records || !zipEntry) {
        fail("invalid-package", `The validated relationship part ${entry.name} is missing.`);
      }
      const originalXml = await zipEntry.async("string");
      const xmlDocument = parseXml(originalXml, entry.name);
      const root = xmlDocument.documentElement;
      if (root.namespaceURI !== PACKAGE_REL || root.localName !== "Relationships") {
        fail("invalid-package", `The relationship part ${entry.name} is malformed.`);
      }
      const removedIds = new Set<string>();
      for (const relationship of records.values()) {
        const altChunk = /(?:^|\/)afchunk$/i.test(relationship.type);
        const payloadTarget = relationship.path !== null && omittedPaths.has(relationship.path);
        if (relationship.external || altChunk || payloadTarget) {
          removedIds.add(relationship.id);
          if (altChunk && relationship.path) omittedPaths.add(relationship.path);
        }
      }
      for (const element of childElements(root)) {
        const id = element.getAttribute("Id") ?? "";
        if (removedIds.has(id)) root.removeChild(element);
      }
      result.set(entry.name, { path: entry.name, source, xmlDocument, records });
    }
    return result;
  })();
}

function contentTypesDocument(
  zip: JSZip,
  omittedPaths: ReadonlySet<string>,
): Promise<Document> {
  const entry = zip.file("[Content_Types].xml");
  if (!entry) fail("invalid-package", "The DOCX package content-types part is missing.");
  return entry.async("string").then((xml) => {
    const document = parseXml(xml, "[Content_Types].xml");
    const root = document.documentElement;
    if (root.namespaceURI !== CONTENT_TYPES || root.localName !== "Types") {
      fail("invalid-package", "The DOCX content-types part is malformed.");
    }
    for (const element of childElements(root)) {
      if (element.namespaceURI !== CONTENT_TYPES) {
        fail("invalid-package", "The DOCX content-types part has an invalid namespace.");
      }
      if (element.localName === "Override") {
        const partName = element.getAttribute("PartName") ?? "";
        const path = partName.startsWith("/") ? partName.slice(1) : "";
        const type = (element.getAttribute("ContentType") ?? "").toLowerCase();
        if (omittedPaths.has(path) || type.includes("oleobject") || type.includes("ole-object")) {
          root.removeChild(element);
        }
      } else if (element.localName === "Default") {
        const type = (element.getAttribute("ContentType") ?? "").toLowerCase();
        if (type.includes("oleobject") || type.includes("ole-object")) {
          root.removeChild(element);
        }
      }
    }
    return document;
  });
}

function elementPathMap(root: Element): Map<string, Element> {
  const paths = new Map<string, Element>();
  const stack = [{ element: root, path: `/${root.nodeName}[1]`, depth: 0 }];
  let visited = 0;
  while (stack.length > 0) {
    const current = stack.pop();
    if (!current) break;
    if (current.depth > 512 || visited >= 100_000) {
      fail("limit-exceeded", "The DOCX document tree exceeds the occurrence mapping limit.");
    }
    visited += 1;
    paths.set(current.path, current.element);
    const counts = new Map<string, number>();
    const children = childElements(current.element).map((element) => {
      const index = (counts.get(element.nodeName) ?? 0) + 1;
      counts.set(element.nodeName, index);
      return {
        element,
        path: `${current.path}/${element.nodeName}[${index}]`,
        depth: current.depth + 1,
      };
    });
    stack.push(...children.reverse());
  }
  return paths;
}

function ancestor(element: Element, namespace: string, name: string): Element | null {
  let current = element.parentNode;
  while (current?.nodeType === 1) {
    const candidate = current as Element;
    if (candidate.namespaceURI === namespace && candidate.localName === name) return candidate;
    current = current.parentNode;
  }
  return null;
}

function descendants(element: Element, namespace: string, name: string): Element[] {
  return Array.from(element.getElementsByTagNameNS(namespace, name));
}

function styleDeclarations(shape: Element): Map<string, string> {
  const declarations = new Map<string, string>();
  const source = shape.getAttribute("style") ?? "";
  for (const declaration of source.split(";")) {
    const trimmed = declaration.trim();
    if (!trimmed) continue;
    const separator = trimmed.indexOf(":");
    if (separator <= 0) {
      fail("unsupported-placement", "VML equation placement contains an unsupported style declaration.");
    }
    const name = trimmed.slice(0, separator).trim().toLowerCase();
    const value = trimmed.slice(separator + 1).trim();
    if (!/^[a-z][a-z0-9-]*$/.test(name) || !value || declarations.has(name)) {
      fail("unsupported-placement", "VML equation placement contains an ambiguous style declaration.");
    }
    declarations.set(name, value);
  }
  return declarations;
}

function parseLength(value: string | undefined, field: string, allowNegative = false): number {
  if (value === undefined) return 0;
  const match = /^(-?(?:\d+(?:\.\d+)?|\.\d+))(pt|in|cm|mm|px)?$/i.exec(value.trim());
  if (!match) {
    fail("unsupported-placement", `VML equation ${field} uses an unsupported physical unit.`);
  }
  const amount = Number(match[1]);
  const unit = (match[2] ?? "pt").toLowerCase();
  if (!Number.isFinite(amount) || (!allowNegative && amount < 0)) {
    fail("unsupported-placement", `VML equation ${field} is outside its supported range.`);
  }
  const factor = unit === "pt" ? EMU_PER_POINT
    : unit === "in" ? EMU_PER_INCH
      : unit === "cm" ? EMU_PER_CM
        : unit === "mm" ? EMU_PER_MM
          : EMU_PER_PIXEL;
  const emu = Math.round(amount * factor);
  if (!Number.isSafeInteger(emu) || Math.abs(emu) > 0x7fffffff) {
    fail("unsupported-placement", `VML equation ${field} exceeds the supported DrawingML range.`);
  }
  return emu;
}

function parseCrop(value: string | null, field: string): number {
  if (value === null || value === "") return 0;
  const match = /^(-?(?:\d+(?:\.\d*)?|\.\d+))(%)?$/.exec(value.trim());
  if (!match) fail("unsupported-placement", `VML equation ${field} crop is malformed.`);
  const amount = Number(match[1]);
  const fraction = match[2] ? amount / 100 : amount;
  if (!Number.isFinite(fraction) || fraction < 0 || fraction >= 1) {
    fail("unsupported-placement", `VML equation ${field} crop is outside the supported range.`);
  }
  return Math.round(fraction * 100_000);
}

function mappedRelative(value: string | undefined, axis: "horizontal" | "vertical"): string {
  if (value === undefined || value === "") return axis === "horizontal" ? "page" : "paragraph";
  const normalized = value.toLowerCase();
  const allowed = axis === "horizontal"
    ? new Set(["page", "margin", "column", "character", "leftmargin", "rightmargin", "insidemargin", "outsidemargin"])
    : new Set(["page", "margin", "paragraph", "line", "topmargin", "bottommargin", "insidemargin", "outsidemargin"]);
  if (!allowed.has(normalized)) {
    fail("unsupported-placement", `VML equation ${axis} anchor uses an unsupported relative reference.`);
  }
  const names: Record<string, string> = {
    leftmargin: "leftMargin",
    rightmargin: "rightMargin",
    topmargin: "topMargin",
    bottommargin: "bottomMargin",
    insidemargin: "insideMargin",
    outsidemargin: "outsideMargin",
  };
  return names[normalized] ?? normalized;
}

function parseAlignment(value: string | undefined, axis: "horizontal" | "vertical"): string | null {
  if (value === undefined || value === "" || value.toLowerCase() === "absolute") return null;
  const normalized = value.toLowerCase();
  const allowed = axis === "horizontal"
    ? new Set(["left", "center", "right", "inside", "outside"])
    : new Set(["top", "center", "bottom", "inside", "outside"]);
  if (!allowed.has(normalized)) {
    fail("unsupported-placement", `VML equation ${axis} anchor uses an unsupported alignment.`);
  }
  return normalized;
}

function parsePlacement(shape: Element, imageData: Element | undefined): VmlPlacement {
  const declarations = styleDeclarations(shape);
  const width = parseLength(declarations.get("width"), "width");
  const height = parseLength(declarations.get("height"), "height");
  if (width <= 0 || height <= 0) {
    fail("unsupported-placement", "VML equation width and height must be positive physical lengths.");
  }
  const position = declarations.get("position")?.toLowerCase();
  if (position !== undefined && position !== "absolute" && position !== "static") {
    fail("unsupported-placement", "VML equation uses an unsupported positioned layout.");
  }
  const kind = position === "absolute" ? "anchor" : "inline";
  const crop = {
    left: parseCrop(imageData?.getAttribute("cropleft") ?? null, "left"),
    top: parseCrop(imageData?.getAttribute("croptop") ?? null, "top"),
    right: parseCrop(imageData?.getAttribute("cropright") ?? null, "right"),
    bottom: parseCrop(imageData?.getAttribute("cropbottom") ?? null, "bottom"),
  };
  if (crop.left + crop.right >= 100_000 || crop.top + crop.bottom >= 100_000) {
    fail("unsupported-placement", "VML equation crop removes the complete preview image.");
  }
  const rotationValue = declarations.get("rotation") ?? shape.getAttribute("rotation") ?? "0";
  const rotation = Number(rotationValue);
  if (!Number.isFinite(rotation) || Math.abs(rotation) > 360) {
    fail("unsupported-placement", "VML equation rotation is unsupported.");
  }
  const zIndexValue = declarations.get("z-index") ?? "0";
  const zIndex = Number(zIndexValue);
  if (!Number.isInteger(zIndex) || Math.abs(zIndex) > 0x7fffffff) {
    fail("unsupported-placement", "VML equation stacking order is unsupported.");
  }
  const wrapValue = (declarations.get("mso-wrap-style") ?? "square").toLowerCase();
  const wrap = wrapValue === "none" ? "none"
    : wrapValue === "square" ? "square"
      : wrapValue === "top-and-bottom" ? "topAndBottom"
        : null;
  if (!wrap) {
    fail("unsupported-placement", "VML equation wrap style cannot be represented safely.");
  }
  const horizontalAlign = parseAlignment(declarations.get("mso-position-horizontal"), "horizontal");
  const verticalAlign = parseAlignment(declarations.get("mso-position-vertical"), "vertical");
  return {
    kind,
    width,
    height,
    rotation: Math.round(rotation * 60_000),
    crop,
    horizontal: {
      relativeFrom: mappedRelative(declarations.get("mso-position-horizontal-relative"), "horizontal"),
      align: horizontalAlign,
      offset: parseLength(declarations.get("left"), "horizontal offset", true),
    },
    vertical: {
      relativeFrom: mappedRelative(declarations.get("mso-position-vertical-relative"), "vertical"),
      align: verticalAlign,
      offset: parseLength(declarations.get("top"), "vertical offset", true),
    },
    relativeHeight: Math.abs(zIndex),
    behindDocument: zIndex < 0,
    wrap,
    distances: {
      top: parseLength(declarations.get("mso-wrap-distance-top"), "top wrap distance"),
      bottom: parseLength(declarations.get("mso-wrap-distance-bottom"), "bottom wrap distance"),
      left: parseLength(declarations.get("mso-wrap-distance-left"), "left wrap distance"),
      right: parseLength(declarations.get("mso-wrap-distance-right"), "right wrap distance"),
    },
  };
}

function ensurePrefix(root: Element, preferred: string, namespace: string): string {
  if (root.lookupNamespaceURI(preferred) === namespace) return preferred;
  if (root.lookupNamespaceURI(preferred) === null) {
    root.setAttributeNS(XMLNS, `xmlns:${preferred}`, namespace);
    return preferred;
  }
  for (let index = 1; index < 20; index += 1) {
    const candidate = `${preferred}${index}`;
    if (root.lookupNamespaceURI(candidate) === namespace) return candidate;
    if (root.lookupNamespaceURI(candidate) === null) {
      root.setAttributeNS(XMLNS, `xmlns:${candidate}`, namespace);
      return candidate;
    }
  }
  fail("invalid-package", `Unable to declare a safe XML prefix for ${namespace}.`);
}

function createElement(document: Document, namespace: string, prefix: string, name: string): Element {
  return document.createElementNS(namespace, `${prefix}:${name}`);
}

function setLengthAttribute(element: Element, name: string, value: number): void {
  element.setAttribute(name, String(value));
}

function appendPosition(
  document: Document,
  parent: Element,
  prefix: string,
  name: "positionH" | "positionV",
  position: LengthPosition,
): void {
  const node = createElement(document, WORD_DRAWING, prefix, name);
  node.setAttribute("relativeFrom", position.relativeFrom);
  if (position.align) {
    const align = createElement(document, WORD_DRAWING, prefix, "align");
    align.appendChild(document.createTextNode(position.align));
    node.appendChild(align);
  } else {
    const offset = createElement(document, WORD_DRAWING, prefix, "posOffset");
    offset.appendChild(document.createTextNode(String(position.offset)));
    node.appendChild(offset);
  }
  parent.appendChild(node);
}

function createDrawing(
  document: Document,
  placement: VmlPlacement,
  relationshipId: string,
  shapeId: string,
  drawingId: number,
): Element {
  const root = document.documentElement;
  const wordPrefix = ensurePrefix(root, "w", WORD);
  const wordDrawingPrefix = ensurePrefix(root, "wp", WORD_DRAWING);
  const drawingPrefix = ensurePrefix(root, "a", DRAWING);
  const picturePrefix = ensurePrefix(root, "pic", PICTURE);
  const relationshipPrefix = ensurePrefix(root, "r", REL);
  const drawing = createElement(document, WORD, wordPrefix, "drawing");
  const wrapper = createElement(
    document,
    WORD_DRAWING,
    wordDrawingPrefix,
    placement.kind === "inline" ? "inline" : "anchor",
  );
  for (const [name, value] of Object.entries({
    distT: placement.distances.top,
    distB: placement.distances.bottom,
    distL: placement.distances.left,
    distR: placement.distances.right,
  })) setLengthAttribute(wrapper, name, value);
  if (placement.kind === "anchor") {
    wrapper.setAttribute("simplePos", "0");
    wrapper.setAttribute("relativeHeight", String(placement.relativeHeight));
    wrapper.setAttribute("behindDoc", placement.behindDocument ? "1" : "0");
    wrapper.setAttribute("locked", "0");
    wrapper.setAttribute("layoutInCell", "1");
    wrapper.setAttribute("allowOverlap", "1");
    const simplePosition = createElement(document, WORD_DRAWING, wordDrawingPrefix, "simplePos");
    simplePosition.setAttribute("x", "0");
    simplePosition.setAttribute("y", "0");
    wrapper.appendChild(simplePosition);
    appendPosition(document, wrapper, wordDrawingPrefix, "positionH", placement.horizontal);
    appendPosition(document, wrapper, wordDrawingPrefix, "positionV", placement.vertical);
  }
  const extent = createElement(document, WORD_DRAWING, wordDrawingPrefix, "extent");
  setLengthAttribute(extent, "cx", placement.width);
  setLengthAttribute(extent, "cy", placement.height);
  wrapper.appendChild(extent);
  const effectExtent = createElement(document, WORD_DRAWING, wordDrawingPrefix, "effectExtent");
  for (const side of ["l", "t", "r", "b"]) effectExtent.setAttribute(side, "0");
  wrapper.appendChild(effectExtent);
  const docProperties = createElement(document, WORD_DRAWING, wordDrawingPrefix, "docPr");
  docProperties.setAttribute("id", String(drawingId));
  docProperties.setAttribute("name", `Word Reader Equation ${shapeId}`);
  wrapper.appendChild(docProperties);
  const frameProperties = createElement(document, WORD_DRAWING, wordDrawingPrefix, "cNvGraphicFramePr");
  const frameLocks = createElement(document, DRAWING, drawingPrefix, "graphicFrameLocks");
  frameLocks.setAttribute("noChangeAspect", "1");
  frameProperties.appendChild(frameLocks);
  wrapper.appendChild(frameProperties);

  if (placement.kind === "anchor") {
    const wrapName = placement.wrap === "none" ? "wrapNone"
      : placement.wrap === "topAndBottom" ? "wrapTopAndBottom"
        : "wrapSquare";
    const wrap = createElement(document, WORD_DRAWING, wordDrawingPrefix, wrapName);
    if (wrapName === "wrapSquare") wrap.setAttribute("wrapText", "bothSides");
    wrapper.appendChild(wrap);
  }

  const graphic = createElement(document, DRAWING, drawingPrefix, "graphic");
  const graphicData = createElement(document, DRAWING, drawingPrefix, "graphicData");
  graphicData.setAttribute("uri", PICTURE);
  const picture = createElement(document, PICTURE, picturePrefix, "pic");
  const nonVisual = createElement(document, PICTURE, picturePrefix, "nvPicPr");
  const pictureProperties = createElement(document, PICTURE, picturePrefix, "cNvPr");
  pictureProperties.setAttribute("id", String(drawingId));
  pictureProperties.setAttribute("name", `Word Reader Equation ${shapeId}`);
  nonVisual.appendChild(pictureProperties);
  nonVisual.appendChild(createElement(document, PICTURE, picturePrefix, "cNvPicPr"));
  nonVisual.appendChild(createElement(document, PICTURE, picturePrefix, "nvPicPr"));
  picture.appendChild(nonVisual);

  const fill = createElement(document, PICTURE, picturePrefix, "blipFill");
  const blip = createElement(document, DRAWING, drawingPrefix, "blip");
  blip.setAttributeNS(REL, `${relationshipPrefix}:embed`, relationshipId);
  fill.appendChild(blip);
  if (placement.crop.left || placement.crop.top || placement.crop.right || placement.crop.bottom) {
    const sourceRect = createElement(document, DRAWING, drawingPrefix, "srcRect");
    sourceRect.setAttribute("l", String(placement.crop.left));
    sourceRect.setAttribute("t", String(placement.crop.top));
    sourceRect.setAttribute("r", String(placement.crop.right));
    sourceRect.setAttribute("b", String(placement.crop.bottom));
    fill.appendChild(sourceRect);
  }
  const stretch = createElement(document, DRAWING, drawingPrefix, "stretch");
  stretch.appendChild(createElement(document, DRAWING, drawingPrefix, "fillRect"));
  fill.appendChild(stretch);
  picture.appendChild(fill);

  const shapeProperties = createElement(document, PICTURE, picturePrefix, "spPr");
  const transform = createElement(document, DRAWING, drawingPrefix, "xfrm");
  if (placement.rotation !== 0) transform.setAttribute("rot", String(placement.rotation));
  const offset = createElement(document, DRAWING, drawingPrefix, "off");
  offset.setAttribute("x", "0");
  offset.setAttribute("y", "0");
  transform.appendChild(offset);
  const shapeExtent = createElement(document, DRAWING, drawingPrefix, "ext");
  setLengthAttribute(shapeExtent, "cx", placement.width);
  setLengthAttribute(shapeExtent, "cy", placement.height);
  transform.appendChild(shapeExtent);
  shapeProperties.appendChild(transform);
  const geometry = createElement(document, DRAWING, drawingPrefix, "prstGeom");
  geometry.setAttribute("prst", "rect");
  geometry.appendChild(createElement(document, DRAWING, drawingPrefix, "avLst"));
  shapeProperties.appendChild(geometry);
  picture.appendChild(shapeProperties);
  graphicData.appendChild(picture);
  graphic.appendChild(graphicData);
  wrapper.appendChild(graphic);
  drawing.appendChild(wrapper);
  return drawing;
}

function createDiagnosticText(document: Document, detail: string): Element {
  const wordPrefix = ensurePrefix(document.documentElement, "w", WORD);
  const text = createElement(document, WORD, wordPrefix, "t");
  text.setAttributeNS(XML, "xml:space", "preserve");
  text.appendChild(document.createTextNode(detail));
  return text;
}

function unwrapAlternatives(alternatives: Map<Element, Element>): void {
  const depth = (element: Element): number => {
    let count = 0;
    let parent: Node | null = element.parentNode;
    while (parent) {
      count += 1;
      parent = parent.parentNode;
    }
    return count;
  };
  const ordered = [...alternatives.entries()].sort((first, second) => depth(second[0]) - depth(first[0]));
  for (const [container, selected] of ordered) {
    const parent = container.parentNode;
    if (!parent) continue;
    while (selected.firstChild) parent.insertBefore(selected.firstChild, container);
    parent.removeChild(container);
  }
}

function splitRuns(
  document: Document,
  replacements: ReadonlyMap<Element, Element>,
): void {
  const runs = Array.from(document.getElementsByTagNameNS(WORD, "r"));
  for (const run of runs) {
    const children = Array.from(run.childNodes);
    if (!children.some((node) => node.nodeType === 1 && replacements.has(node as Element))) continue;
    const parent = run.parentNode;
    if (!parent) fail("invalid-package", "A Word run containing an equation has no parent.");
    const runProperties = children.filter((node) =>
      node.nodeType === 1 && (node as Element).namespaceURI === WORD && (node as Element).localName === "rPr"
    );
    const pieces: Array<{ nodes: Node[] } | { replacement: Element }> = [];
    let segment: Node[] = [];
    const flush = (): void => {
      if (segment.length > 0) pieces.push({ nodes: segment });
      segment = [];
    };
    for (const child of children) {
      if (runProperties.includes(child)) continue;
      if (child.nodeType === 1 && replacements.has(child as Element)) {
        flush();
        pieces.push({ replacement: replacements.get(child as Element)! });
      } else {
        if (child.nodeType === 1 && (child as Element).namespaceURI === WORD && (child as Element).localName === "object") {
          fail("invalid-package", "An unaccounted Word OLE object remains in the selected document tree.");
        }
        segment.push(child);
      }
    }
    flush();
    for (const piece of pieces) {
      const nextRun = run.cloneNode(false) as Element;
      for (const properties of runProperties) nextRun.appendChild(properties.cloneNode(true));
      if ("replacement" in piece) {
        nextRun.appendChild(piece.replacement);
      } else {
        for (const node of piece.nodes) nextRun.appendChild(node);
      }
      parent.insertBefore(nextRun, run);
    }
    parent.removeChild(run);
  }
}

function removeAltChunks(document: Document): boolean {
  const nodes = Array.from(document.getElementsByTagNameNS(WORD, "altChunk"));
  for (const node of nodes) node.parentNode?.removeChild(node);
  return nodes.length > 0;
}

function appendRelationship(
  relationshipDocument: RelationshipDocument,
  id: string,
  target: string,
): void {
  const root = relationshipDocument.xmlDocument.documentElement;
  const relationship = relationshipDocument.xmlDocument.createElementNS(PACKAGE_REL, "Relationship");
  relationship.setAttribute("Id", id);
  relationship.setAttribute("Type", `${REL}/image`);
  relationship.setAttribute("Target", target);
  root.appendChild(relationship);
}

function appendContentTypeOverride(document: Document, path: string, type: string): void {
  const root = document.documentElement;
  const override = document.createElementNS(CONTENT_TYPES, "Override");
  override.setAttribute("PartName", `/${path}`);
  override.setAttribute("ContentType", type);
  root.appendChild(override);
}

function relativeTarget(sourcePath: string, targetPath: string): string {
  const sourceDirectory = sourcePath.slice(0, sourcePath.lastIndexOf("/") + 1);
  if (targetPath.startsWith(sourceDirectory)) return targetPath.slice(sourceDirectory.length);
  fail("invalid-package", "Generated equation preview path is outside its source part directory.");
}

function nextRelationshipId(source: string, metadata: Awaited<ReturnType<typeof readOoxmlPreviewMetadata>>, used: Map<string, Set<string>>): string {
  let ids = used.get(source);
  if (!ids) {
    ids = new Set(metadata.relationships.get(source)?.keys() ?? []);
    used.set(source, ids);
  }
  for (let index = 1; index <= 10_000; index += 1) {
    const candidate = `rIdWordReaderPreview${index}`;
    if (!ids.has(candidate)) {
      ids.add(candidate);
      return candidate;
    }
  }
  fail("limit-exceeded", "The DOCX relationship identifier space is exhausted.");
}

function nextPreviewPath(
  extension: string,
  usedPaths: Set<string>,
): string {
  for (let index = 1; index <= 10_000; index += 1) {
    const candidate = `word/word-reader-preview-${index}.${extension}`;
    if (!usedPaths.has(candidate)) {
      usedPaths.add(candidate);
      return candidate;
    }
  }
  fail("limit-exceeded", "The DOCX generated-preview path space is exhausted.");
}

function contentTypeForMime(mimeType: ResolvedPreviewMedia["mimeType"]): { extension: string; type: string } {
  if (mimeType === "image/svg+xml") return { extension: "svg", type: "image/svg+xml" };
  if (mimeType === "image/png") return { extension: "png", type: "image/png" };
  if (mimeType === "image/jpeg") return { extension: "jpg", type: "image/jpeg" };
  return { extension: "bmp", type: "image/bmp" };
}

function removeActiveReferencesFromWordPart(document: Document): boolean {
  return removeAltChunks(document);
}

function verifyNoUnsafeReferences(
  zip: JSZip,
  options: DocxPreviewPackageOptions,
): Promise<void> {
  return (async () => {
    const names = Object.keys(zip.files).filter((name) => !zip.files[name].dir);
    if (names.some(isEmbeddingPath)) {
      fail("invalid-package", "The prepared DOCX still contains an embedding part.");
    }
    const contentTypes = zip.file("[Content_Types].xml");
    if (!contentTypes) fail("invalid-package", "The prepared DOCX has no content-types part.");
    const contentTypeXml = await contentTypes.async("string");
    if (/oleObject|ole-object/i.test(contentTypeXml)) {
      fail("invalid-package", "The prepared DOCX still declares an OLE content type.");
    }
    for (const name of names) {
      checkCancelled(options);
      if (isRelationshipMetadataPath(name)) {
        const xml = await zip.file(name)!.async("string");
        if (/TargetMode\s*=\s*["']External["']|oleObject|ole-object|aFChunk/i.test(xml)) {
          fail("invalid-package", `The prepared relationship part ${name} retains an active reference.`);
        }
      } else if (name.startsWith("word/") && name.endsWith(".xml")) {
        const xml = await zip.file(name)!.async("string");
        const document = parseXml(xml, name);
        if (
          document.getElementsByTagNameNS(WORD, "altChunk").length > 0 ||
          document.getElementsByTagNameNS(WORD, "object").length > 0 ||
          document.getElementsByTagNameNS(OFFICE, "OLEObject").length > 0
        ) {
          fail("invalid-package", `The prepared Word part ${name} retains an active object reference.`);
        }
      }
    }
    const strictInventory = await enforceOoxmlPackagePolicy(zip, "docx");
    if (strictInventory !== null) {
      fail("invalid-package", "The prepared DOCX did not pass the strict default OOXML policy.");
    }
  })();
}

export async function prepareDocxPreviewPackage(
  buffer: ArrayBuffer,
  options: DocxPreviewPackageOptions,
): Promise<PreparedDocxPreviewPackage> {
  if (!options || typeof options.isCancelled !== "function" || typeof options.yieldControl !== "function") {
    throw new TypeError("DOCX preview cancellation and yield callbacks are required.");
  }
  checkCancelled(options);
  const safety = validateZipSafety(buffer, DEFAULT_DOCX_ZIP_LIMITS);
  const zip = await loadOoxmlZip(buffer);
  const inventory = await validateDocxPreviewSource(zip);
  checkCancelled(options);
  const metadata = await readOoxmlPreviewMetadata(zip);
  checkCancelled(options);

  const omittedPaths = new Set(metadata.embeddingPaths);
  for (const name of Object.keys(zip.files)) if (isEmbeddingPath(name)) omittedPaths.add(name);
  const relationshipParts = await relationshipDocuments(safety.entries, zip, metadata, omittedPaths, options);
  for (const relationship of metadata.relationships.values()) {
    for (const entry of relationship.values()) {
      if (/(?:^|\/)afchunk$/i.test(entry.type) && entry.path) omittedPaths.add(entry.path);
    }
  }
  const contentTypes = await contentTypesDocument(zip, omittedPaths);
  const relationshipDocumentsBySource = new Map<string, RelationshipDocument>();
  for (const document of relationshipParts.values()) relationshipDocumentsBySource.set(document.source, document);

  const retained = new Map<string, Uint8Array>();
  const xmlDocuments = new Map<string, Document>();
  let retainedBytes = 0;
  for (const entry of safety.entries) {
    checkCancelled(options);
    const zipEntry = zip.files[entry.name];
    if (!zipEntry || zipEntry.dir || entry.name === "[Content_Types].xml" || isRelationshipMetadataPath(entry.name) || omittedPaths.has(entry.name)) {
      continue;
    }
    const bytes = await zipEntry.async("uint8array");
    if (bytes.byteLength !== entry.uncompressedBytes || bytes.byteLength > MAX_ENTRY_BYTES) {
      fail("limit-exceeded", `Expanded length for ${entry.name} differs from its validated ZIP declaration.`);
    }
    retainedBytes = checkedAdd(retainedBytes, bytes.byteLength, MAX_TOTAL_BYTES, "Retained DOCX content");
    if (entry.name.startsWith("word/") && entry.name.endsWith(".xml")) {
      const sourceText = utf8(bytes, entry.name);
      const partOccurrences = inventory.occurrences.filter((occurrence) => occurrence.partPath === entry.name);
      if (partOccurrences.length > 0 || /altChunk\b/.test(sourceText)) {
        const document = parseXml(sourceText, entry.name);
        xmlDocuments.set(entry.name, document);
      } else {
        retained.set(entry.name, bytes);
      }
    } else {
      retained.set(entry.name, bytes);
    }
    await options.yieldControl();
    checkCancelled(options);
  }
  for (const occurrence of inventory.occurrences) {
    if (!xmlDocuments.has(occurrence.partPath)) {
      fail("invalid-package", `Validated equation part ${occurrence.partPath} was not retained for transformation.`);
    }
  }

  const issues: PreviewIssue[] = [...inventory.issues];
  const mediaResolver = createPreviewMediaResolver();
  const mediaBySourcePath = new Map<string, ResolvedPreviewMedia | null>();
  const mediaByOccurrence = new Map<string, { mediaPath: string; relationshipId: string } | null>();
  const mediaParts = new Map<string, MediaPart>();
  const usedPaths = new Set(Object.keys(zip.files));
  const usedRelationshipIds = new Map<string, Set<string>>();
  try {
    let generatedBytes = 0;
    for (const occurrence of inventory.occurrences) {
      checkCancelled(options);
      if (!occurrence.previewPath) {
        mediaByOccurrence.set(occurrence.occurrenceId, null);
        continue;
      }
      let resolved = mediaBySourcePath.get(occurrence.previewPath);
      if (resolved === undefined) {
        const sourceBytes = retained.get(occurrence.previewPath);
        if (!sourceBytes) {
          fail("invalid-package", `Validated preview part ${occurrence.previewPath} was not retained.`);
        }
        try {
          resolved = await mediaResolver.resolve(occurrence.previewPath, sourceBytes, options);
          generatedBytes = checkedAdd(generatedBytes, resolved.bytes.byteLength, MAX_TOTAL_BYTES, "Generated equation previews");
          const contentType = contentTypeForMime(resolved.mimeType);
          const path = nextPreviewPath(contentType.extension, usedPaths);
          mediaParts.set(path, { path, media: resolved, contentType: contentType.type });
          appendContentTypeOverride(contentTypes, path, contentType.type);
          mediaBySourcePath.set(occurrence.previewPath, resolved);
        } catch (error) {
          if (error instanceof Error && /cancel/i.test(`${(error as { code?: string }).code} ${error.message}`)) {
            fail("cancelled", "DOCX preview media resolution was cancelled.");
          }
          resolved = null;
          mediaBySourcePath.set(occurrence.previewPath, null);
        }
      }
      if (!resolved) {
        issues.push({
          partPath: occurrence.partPath,
          occurrenceId: occurrence.occurrenceId,
          code: "preview-unavailable",
          detail: `Preview unavailable for equation object ${occurrence.shapeId}.`,
        });
        mediaByOccurrence.set(occurrence.occurrenceId, null);
        continue;
      }
      const media = mediaParts.get([...mediaParts.keys()].find((path) => mediaParts.get(path)?.media === resolved) ?? "");
      if (!media) fail("invalid-package", `Generated preview bytes for ${occurrence.previewPath} are missing.`);
      mediaByOccurrence.set(occurrence.occurrenceId, {
        mediaPath: media.path,
        relationshipId: "",
      });
      for (const issue of resolved.issues) {
        issues.push({
          partPath: occurrence.partPath,
          occurrenceId: occurrence.occurrenceId,
          code: issue.code,
          detail: issue.detail,
        });
      }
    }

    const transformedDocuments = new Map<string, Uint8Array>();
    for (const [partPath, document] of xmlDocuments) {
      const partOccurrences: LocatedOccurrence[] = [];
      const occurrencesForPart = inventory.occurrences.filter((occurrence) => occurrence.partPath === partPath);
      const pathMap = elementPathMap(document.documentElement);
      for (const occurrence of occurrencesForPart) {
        const source = pathMap.get(occurrence.sourceElementPath);
        if (!source || source.namespaceURI !== OFFICE || source.localName !== "OLEObject") {
          fail("invalid-package", `Equation occurrence ${occurrence.occurrenceId} no longer matches its validated source identity.`);
        }
        const object = ancestor(source, WORD, "object");
        if (!object) fail("invalid-package", `Equation occurrence ${occurrence.occurrenceId} has no Word object container.`);
        const shapes = descendants(object, VML, "shape").filter((shape) => shape.getAttribute("id") === occurrence.shapeId);
        if (shapes.length !== 1) fail("invalid-package", `Equation occurrence ${occurrence.occurrenceId} has an ambiguous VML shape.`);
        const imageData = descendants(shapes[0], VML, "imagedata");
        if (imageData.length > 1) fail("invalid-package", `Equation occurrence ${occurrence.occurrenceId} has multiple image references.`);
        const placement = parsePlacement(shapes[0], imageData[0]);
        partOccurrences.push({ occurrence, object, placement });
      }
      if (partOccurrences.length > 0) {
        const alternatives = new Map<Element, Element>();
        for (const item of partOccurrences) {
          let current: Node | null = pathMap.get(item.occurrence.sourceElementPath) ?? null;
          while (current?.nodeType === 1) {
            const element = current as Element;
            if (element.namespaceURI === MARKUP_COMPATIBILITY && element.localName === "AlternateContent") {
              let branch: Node = pathMap.get(item.occurrence.sourceElementPath)!;
              while (branch.parentNode && branch.parentNode !== element) branch = branch.parentNode;
              const selected = branch as Element;
              if (selected.parentNode !== element || selected.namespaceURI !== MARKUP_COMPATIBILITY ||
                (selected.localName !== "Choice" && selected.localName !== "Fallback")) {
                fail("invalid-package", `Equation occurrence ${item.occurrence.occurrenceId} is outside a selected alternate branch.`);
              }
              const previous = alternatives.get(element);
              if (previous && previous !== selected) fail("invalid-package", "Equation occurrences select conflicting alternate branches.");
              alternatives.set(element, selected);
            }
            current = current.parentNode;
          }
        }
        const replacements = new Map<Element, Element>();
        let drawingId = 1;
        for (const item of partOccurrences) {
          const media = mediaByOccurrence.get(item.occurrence.occurrenceId);
          if (!media) {
            replacements.set(item.object, createDiagnosticText(document, "Preview unavailable"));
            continue;
          }
          const relationDocument = relationshipDocumentsBySource.get(partPath);
          if (!relationDocument) fail("invalid-package", `The image relationship part for ${partPath} is missing.`);
          const relationshipId = nextRelationshipId(partPath, metadata, usedRelationshipIds);
          appendRelationship(relationDocument, relationshipId, relativeTarget(partPath, media.mediaPath));
          const drawing = createDrawing(document, item.placement, relationshipId, item.occurrence.shapeId, drawingId);
          drawingId += 1;
          replacements.set(item.object, drawing);
        }
        unwrapAlternatives(alternatives);
        splitRuns(document, replacements);
      }
      if (removeActiveReferencesFromWordPart(document) || partOccurrences.length > 0) {
        transformedDocuments.set(partPath, new TextEncoder().encode(serialize(document)));
      }
    }

    const outputEntries: StoredZipEntry[] = [];
    for (const entry of safety.entries) {
      checkCancelled(options);
      const name = entry.name;
      const zipEntry = zip.files[name];
      if (!zipEntry || zipEntry.dir || omittedPaths.has(name) || isEmbeddingPath(name)) continue;
      if (name === "[Content_Types].xml") {
        outputEntries.push({ path: name, bytes: new TextEncoder().encode(serialize(contentTypes)) });
      } else if (relationshipParts.has(name)) {
        outputEntries.push({ path: name, bytes: new TextEncoder().encode(serialize(relationshipParts.get(name)!.xmlDocument)) });
      } else {
        const transformed = transformedDocuments.get(name);
        const retainedBytes = retained.get(name);
        if (transformed) outputEntries.push({ path: name, bytes: transformed });
        else if (retainedBytes) outputEntries.push({ path: name, bytes: retainedBytes });
      }
    }
    for (const media of mediaParts.values()) {
      outputEntries.push({ path: media.path, bytes: media.media.bytes });
    }
    checkCancelled(options);
    const preparedBuffer = await writeStoredZip(outputEntries, options);
    const preparedZip = await loadOoxmlZip(preparedBuffer);
    await verifyNoUnsafeReferences(preparedZip, options);
    return { buffer: preparedBuffer, inventory, issues };
  } finally {
    mediaResolver.dispose();
  }
}
