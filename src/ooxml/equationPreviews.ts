import type JSZip from "jszip";
import { childElements, namespacedAttribute, parseXml } from "../pptx/xml";
import { normalizeAlternateContent } from "./alternateContent";
import type { PreviewInventory, PreviewIssue, PreviewOccurrence } from "./previewTypes";

const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PACKAGE_REL = "http://schemas.openxmlformats.org/package/2006/relationships";
const CT = "http://schemas.openxmlformats.org/package/2006/content-types";
const WORD = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const PPT = "http://schemas.openxmlformats.org/presentationml/2006/main";
const DRAWING = "http://schemas.openxmlformats.org/drawingml/2006/main";
const OFFICE = "urn:schemas-microsoft-com:office:office";
const VML = "urn:schemas-microsoft-com:vml";
const IMAGE_TYPES: Readonly<Record<string, readonly string[]>> = {
  png: ["image/png"], jpg: ["image/jpeg"], jpeg: ["image/jpeg"], bmp: ["image/bmp", "image/x-bmp"],
  wmf: ["image/x-wmf", "image/wmf"], emf: ["image/x-emf", "image/emf"],
};

interface Relationship {
  id: string;
  type: string;
  target: string;
  external: boolean;
  path: string | null;
}

export interface OoxmlPreviewMetadata {
  embeddingPaths: ReadonlySet<string>;
  contentTypes: ReadonlyMap<string, string>;
  relationships: ReadonlyMap<string, ReadonlyMap<string, Relationship>>;
}

export function isCanonicalPreviewPath(path: string): boolean {
  return path.length > 0 && !unsafeCharacters(path) &&
    path.split("/").every((segment) => segment !== "" && segment !== "." && segment !== "..");
}

function fail(detail: string): never { throw new Error(`Invalid equation preview metadata: ${detail}`); }

function unsafeCharacters(value: string): boolean {
  return /[\\%:#?\s]/.test(value) || Array.from(value).some((character) => character.charCodeAt(0) < 32);
}

function resolveTarget(source: string, target: string): string {
  if (!target || unsafeCharacters(target) || target.startsWith("/")) fail("unsafe-target");
  const segments = source.split("/");
  segments.pop();
  for (const segment of target.split("/")) {
    if (!segment || segment === ".") fail("target-alias");
    if (segment === "..") {
      if (segments.length === 0) fail("root-escape");
      segments.pop();
    } else segments.push(segment);
  }
  const path = segments.join("/");
  if (!isCanonicalPreviewPath(path)) fail("target-path");
  return path;
}

function embeddingName(path: string): boolean { return /(?:^|\/)(?:embeddings|oleobjects)(?:\/|$)/i.test(path); }
function oleType(type: string): boolean { return /ole-?object/i.test(type); }

/** Fixed metadata roles permit string validation only; collisions never grant admission. */
export async function readOoxmlPreviewMetadata(zip: JSZip): Promise<OoxmlPreviewMetadata> {
  const names = Object.keys(zip.files).filter((name) => !zip.files[name].dir);
  const embeddingPaths = new Set(names.filter(embeddingName));
  const metadataPaths = new Set(names.filter((name) => isCanonicalPreviewPath(name) &&
    (name === "[Content_Types].xml" || name === "_rels/.rels" || /(?:^|\/)_rels\/[^/]+\.rels$/.test(name)),
  ));
  const rejectMetadataPayload = (): void => {
    for (const path of embeddingPaths) if (metadataPaths.has(path)) fail("metadata-payload");
  };
  const contentTypes = new Map<string, string>();
  const defaults = new Map<string, string>();
  const overrides = new Map<string, string>();
  const entry = zip.file("[Content_Types].xml");
  if (entry) {
    const root = parseXml(await entry.async("string"), "[Content_Types].xml").documentElement;
    if (root.namespaceURI !== CT || root.localName !== "Types") fail("types-namespace");
    for (const element of childElements(root)) {
      if (element.namespaceURI !== CT) fail("types-namespace");
      const type = element.getAttribute("ContentType");
      if (!type) fail("missing-type");
      if (element.localName === "Default") {
        const extension = (element.getAttribute("Extension") ?? "").toLowerCase();
        if (!extension || defaults.has(extension)) fail("type-default");
        defaults.set(extension, type);
      } else if (element.localName === "Override") {
        const name = element.getAttribute("PartName") ?? "";
        const path = name.startsWith("/") ? name.slice(1) : "";
        if (!isCanonicalPreviewPath(path) || overrides.has(path)) fail("type-override");
        overrides.set(path, type);
        if (oleType(type)) embeddingPaths.add(path);
      } else fail("type-element");
    }
    for (const name of names) {
      const type = overrides.get(name) ?? defaults.get(name.split(".").at(-1)?.toLowerCase() ?? "");
      if (type) contentTypes.set(name, type);
      if (type && oleType(type)) embeddingPaths.add(name);
    }
  }
  rejectMetadataPayload();
  const relationships = new Map<string, ReadonlyMap<string, Relationship>>();
  for (const name of [...metadataPaths].filter((name) => name !== "[Content_Types].xml").sort()) {
    const root = parseXml(await zip.file(name)!.async("string"), name).documentElement;
    if (root.namespaceURI !== PACKAGE_REL || root.localName !== "Relationships") fail("rels-namespace");
    const source = name === "_rels/.rels" ? "" : name.replace(/_rels\/([^/]+)\.rels$/, "$1");
    const entries = new Map<string, Relationship>();
    for (const element of childElements(root)) {
      if (element.namespaceURI !== PACKAGE_REL || element.localName !== "Relationship") fail("rels-namespace");
      const id = element.getAttribute("Id") ?? "";
      const type = element.getAttribute("Type");
      const target = element.getAttribute("Target");
      const mode = element.getAttribute("TargetMode");
      if (!/^[A-Za-z_][A-Za-z0-9_.-]*$/.test(id) || entries.has(id) || !type || !target || (mode && mode !== "Internal" && mode !== "External")) fail("relationship-id/type/mode");
      const external = mode === "External";
      const path = external ? null : resolveTarget(source, target);
      entries.set(id, { id, type, target, external, path });
      if (oleType(type)) {
        if (type !== `${REL}/oleObject` || external || !path) fail("ole-relationship");
        embeddingPaths.add(path);
        if (metadataPaths.has(path)) fail("metadata-payload");
      }
    }
    relationships.set(source, entries);
  }
  rejectMetadataPayload();
  return { embeddingPaths, contentTypes, relationships };
}

function imagePath(zip: JSZip, metadata: OoxmlPreviewMetadata, relationship: Relationship): string {
  const path = relationship.path;
  if (relationship.type !== `${REL}/image` || relationship.external || !path || metadata.embeddingPaths.has(path) || embeddingName(path)) fail("image-relationship");
  const extension = path.split(".").at(-1)?.toLowerCase() ?? "";
  const type = metadata.contentTypes.get(path);
  if (!zip.file(path) || !type || !IMAGE_TYPES[extension]?.includes(type)) fail("image-part/type");
  return path;
}

function elements(root: Element): Array<{ element: Element; path: string }> {
  const result: Array<{ element: Element; path: string }> = [];
  const stack = [{ element: root, path: `/${root.nodeName}[1]`, depth: 0 }];
  while (stack.length) {
    const item = stack.pop()!;
    if (item.depth > 512 || result.length >= 100_000) fail("element-limit");
    result.push(item);
    const counts = new Map<string, number>();
    const children = childElements(item.element).map((element) => {
      const index = (counts.get(element.nodeName) ?? 0) + 1;
      counts.set(element.nodeName, index);
      return { element, path: `${item.path}/${element.nodeName}[${index}]`, depth: item.depth + 1 };
    });
    stack.push(...children.reverse());
  }
  return result;
}

function matching(root: Element, namespace: string, name: string): Element[] {
  return Array.from(root.getElementsByTagNameNS(namespace, name));
}

function ancestor(element: Element, namespace: string, name: string): Element | null {
  let node = element.parentNode;
  while (node?.nodeType === 1) {
    const current = node as Element;
    if (current.namespaceURI === namespace && current.localName === name) return current;
    node = current.parentNode;
  }
  return null;
}

export async function inspectEquationPreviews(zip: JSZip, format: "docx" | "pptx"): Promise<PreviewInventory> {
  const metadata = await readOoxmlPreviewMetadata(zip);
  const occurrences: PreviewOccurrence[] = [];
  const issues: PreviewIssue[] = [];
  const accounted = new Set<string>();
  const prefix = format === "docx" ? "word/" : "ppt/";
  for (const partPath of Object.keys(zip.files).filter((path) => path.startsWith(prefix) && path.endsWith(".xml") && !zip.files[path].dir)) {
    if (metadata.embeddingPaths.has(partPath)) continue;
    const original = parseXml(await zip.file(partPath)!.async("string"), partPath);
    // No Office/VML extension is advertised as a rendered capability.
    const normalized = normalizeAlternateContent(original, new Set([WORD, PPT, DRAWING]));
    if (normalized.issues.some((issue) => issue.code !== "unsupported-requires")) fail("alternate-content");
    const originalObjects = elements(original.documentElement).filter(({ element, path }) =>
      (element.localName === "OLEObject" || element.localName === "oleObj") &&
      normalized.selections.every((selection) => !path.startsWith(`${selection.sourceElementPath}/`) || path.startsWith(`${selection.selectedElementPath}/`)),
    );
    const rels = metadata.relationships.get(partPath) ?? new Map<string, Relationship>();
    const shapeIds = new Set<string>();
    for (const { element } of elements(normalized.document.documentElement)) {
      const name = element.localName;
      const expectedNamespace = name === "shape" || name === "imagedata" ? VML : name === "blip" ? DRAWING : null;
      if (expectedNamespace && element.namespaceURI !== expectedNamespace) fail("image-namespace");
      if (name !== "OLEObject" && name !== "oleObj") continue;
      if (element.namespaceURI !== (format === "docx" ? OFFICE : PPT)) fail("ole-namespace");
      const progId = element.getAttribute(format === "docx" ? "ProgID" : "progId");
      if (progId !== "Equation.3" && progId !== "Equation.DSMT4") fail("ole-class");
      const id = namespacedAttribute(element, "id", REL);
      const relationship = id ? rels.get(id) : undefined;
      if (!relationship || relationship.type !== `${REL}/oleObject` || relationship.external || !relationship.path || !zip.file(relationship.path)) fail("embedding-relationship");
      const embeddingPath = relationship.path;
      const container = ancestor(element, format === "docx" ? WORD : PPT, format === "docx" ? "object" : "graphicFrame");
      if (!container) fail("ole-container");
      let shapeId: string;
      let images: Element[];
      if (format === "docx") {
        if (element.getAttribute("Type") !== "Embed") fail("linked-ole");
        shapeId = element.getAttribute("ShapeID") ?? "";
        const shapes = matching(container, VML, "shape").filter((shape) => shape.getAttribute("id") === shapeId);
        if (shapes.length !== 1) fail("shape-id");
        images = matching(shapes[0], VML, "imagedata");
      } else {
        const properties = childElements(container).find((child) => child.namespaceURI === PPT && child.localName === "nvGraphicFramePr");
        const ids = properties ? matching(properties, PPT, "cNvPr") : [];
        if (ids.length !== 1 || matching(element, PPT, "embed").length !== 1 || matching(element, PPT, "link").length) fail("ppt-ole-id");
        shapeId = ids[0].getAttribute("id") ?? "";
        const pictures = childElements(element).filter((child) => child.namespaceURI === PPT && child.localName === "pic");
        if (pictures.length > 1) fail("preview-count");
        images = pictures.length ? matching(pictures[0], DRAWING, "blip") : [];
      }
      if (!shapeId || shapeIds.has(shapeId)) fail("shape-id");
      shapeIds.add(shapeId);
      if (images.length > 1) fail("preview-count");
      const previewRelationshipId = images.length ? namespacedAttribute(images[0], format === "docx" ? "id" : "embed", REL) : null;
      if (images.length && !previewRelationshipId) fail("image-ref-id");
      if (images.length && namespacedAttribute(images[0], "link", REL)) fail("external-image");
      const occurrenceId = `${partPath}#${shapeId}:${id}`;
      let previewPath: string | null = null;
      const preview = previewRelationshipId ? rels.get(previewRelationshipId) : undefined;
      if (preview) previewPath = imagePath(zip, metadata, preview);
      else issues.push({ partPath, occurrenceId, code: "missing-preview", detail: "Equation has no internal image preview relationship." });
      const sourceIndex = originalObjects.findIndex((item) => namespacedAttribute(item.element, "id", REL) === id);
      const source = originalObjects.splice(sourceIndex, 1)[0];
      if (!source) fail("source-id");
      occurrences.push({ occurrenceId, partPath, shapeId, embeddingPath, previewPath, previewRelationshipId, progId, sourceElementPath: source.path });
      accounted.add(embeddingPath);
    }
  }
  for (const path of metadata.embeddingPaths) {
    if (!accounted.has(path)) fail("unknown-embedding");
  }
  return { occurrences, issues };
}

/** Inventory is not a grant of access: revalidate actual package metadata. */
export async function collectValidatedPptxImagePaths(zip: JSZip, inventory: PreviewInventory | null): Promise<ReadonlySet<string>> {
  const metadata = await readOoxmlPreviewMetadata(zip);
  const paths = new Set<string>();
  for (const [source, rels] of metadata.relationships) {
    for (const relationship of rels.values()) {
      if (relationship.path && metadata.embeddingPaths.has(relationship.path) && relationship.type !== `${REL}/oleObject`) fail("payload-relationship");
      if (!/^ppt\/(?:slides|slideLayouts|slideMasters)\/[^/]+\.xml$/.test(source)) continue;
      if (relationship.type.endsWith("/image")) paths.add(imagePath(zip, metadata, relationship));
    }
  }
  // Reinspection prevents a caller-forged inventory from enlarging the allowlist.
  if (inventory) {
    const actual = await inspectEquationPreviews(zip, "pptx");
    for (const occurrence of actual.occurrences) if (occurrence.previewPath) paths.add(occurrence.previewPath);
  }
  return paths;
}
