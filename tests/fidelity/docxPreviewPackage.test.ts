/* eslint-disable obsidianmd/no-global-this -- Node-only test shim installs xmldom's serializer on the test global. */
import assert from "node:assert/strict";
import { createJiti } from "jiti";
import test from "node:test";
import path from "node:path";
import JSZip from "jszip";
import { XMLSerializer } from "@xmldom/xmldom";

import { parseXml } from "../../src/pptx/xml";
import {
  createSyntheticEmfFixture,
  createSyntheticWmfFixture,
  createEquationFixture,
} from "./fixtureBuilders";
import { createLargeDocx } from "../performanceFixtures";

const jiti = createJiti(__filename);
const WORD = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const DRAWING = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const PICTURE = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PACKAGE_REL = "http://schemas.openxmlformats.org/package/2006/relationships";

interface PreparedDocx {
  buffer: ArrayBuffer;
  inventory: {
    occurrences: ReadonlyArray<{
      occurrenceId: string;
      shapeId: string;
      previewPath: string | null;
    }>;
    issues: readonly unknown[];
  };
    issues: ReadonlyArray<{
    partPath: string;
    occurrenceId: string;
    code: string;
    detail: string;
  }>;
}

async function prepareDocxPreviewPackage(
  buffer: ArrayBuffer,
  overrides: Partial<{
    isCancelled: () => boolean;
    yieldControl: () => Promise<void>;
  }> = {},
): Promise<PreparedDocx> {
  let module: Record<string, unknown>;
  try {
    module = await jiti.import(
      path.resolve(process.cwd(), "src/docx/docxPreviewPackage.ts"),
    );
  } catch (error) {
    assert.fail(
      `prepareDocxPreviewPackage implementation is missing: ${error instanceof Error ? error.message : String(error)}`,
    );
  }
  assert.equal(typeof module.prepareDocxPreviewPackage, "function");
  const previousSerializer = Object.getOwnPropertyDescriptor(globalThis, "XMLSerializer");
  Object.defineProperty(globalThis, "XMLSerializer", {
    configurable: true,
    writable: true,
    value: XMLSerializer,
  });
  try {
    return await (module.prepareDocxPreviewPackage as (
      buffer: ArrayBuffer,
      options: object,
    ) => Promise<PreparedDocx>)(buffer, {
      isCancelled: () => false,
      yieldControl: async () => {},
      ...overrides,
    });
  } finally {
    if (previousSerializer) {
      Object.defineProperty(globalThis, "XMLSerializer", previousSerializer);
    } else {
      Reflect.deleteProperty(globalThis, "XMLSerializer");
    }
  }
}

async function updateDocument(
  buffer: ArrayBuffer,
  update: (xml: string) => string,
): Promise<ArrayBuffer> {
  const zip = await JSZip.loadAsync(buffer);
  const documentPath = "word/document.xml";
  zip.file(documentPath, update(await zip.file(documentPath)!.async("string")));
  return await zip.generateAsync({ type: "arraybuffer", compression: "STORE" });
}

async function makeEquationFixture(options: {
  duplicate?: boolean;
  missingPreview?: boolean;
  alternative?: boolean;
} = {}): Promise<ArrayBuffer> {
  return await createEquationFixture("docx", {
    progId: "Equation.3",
    previewRelationship: { id: "image", target: "media/preview.png" },
    previewBytes: Uint8Array.from([137, 80, 78, 71, 13, 10, 26, 10, 0]),
    opaqueEmbeddingBytes: new TextEncoder().encode("opaque bytes are not preview data"),
    ...options,
  });
}

async function loadPackage(buffer: ArrayBuffer): Promise<JSZip> {
  return await JSZip.loadAsync(buffer);
}

async function assertNoActiveEquationPayload(zip: JSZip): Promise<string> {
  const names = Object.keys(zip.files).filter((name) => !zip.files[name].dir);
  assert.equal(names.some((name) => /(?:^|\/)(?:embeddings|oleobjects)(?:\/|$)/i.test(name)), false);
  const relationships = await zip.file("word/_rels/document.xml.rels")!.async("string");
  const contentTypes = await zip.file("[Content_Types].xml")!.async("string");
  const document = await zip.file("word/document.xml")!.async("string");
  assert.equal(/oleObject|ole-object|TargetMode="External"/i.test(relationships), false);
  assert.equal(/oleObject|ole-object/i.test(contentTypes), false);
  assert.equal(/<w:object\b|OLEObject|<w:altChunk\b/i.test(document), false);
  return document;
}

function elements(document: Document, namespace: string, name: string): Element[] {
  return Array.from(document.getElementsByTagNameNS(namespace, name));
}

void test("replaces one equation occurrence in run order and preserves baseline and exact EMU dimensions", async () => {
  const source = await updateDocument(await makeEquationFixture(), (xml) => xml
    .replace('<w:r><w:object>', '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t>before</w:t><w:object>')
    .replace('</w:object></w:r></w:p>', '</w:object><w:t>after</w:t></w:r></w:p>')
    .replace('style="width:36pt;height:18pt"', 'style="width:61.9pt;height:30.95pt"'));
  const original = new Uint8Array(source).slice();
  const prepared = await prepareDocxPreviewPackage(source);
  const zip = await loadPackage(prepared.buffer);
  const documentXml = await assertNoActiveEquationPayload(zip);
  const document = parseXml(documentXml, "word/document.xml");
  const paragraph = elements(document, WORD, "p")[0];
  const events: string[] = [];
  for (const run of elements(document, WORD, "r")) {
    if (run.parentNode !== paragraph) continue;
    for (const child of Array.from(run.childNodes)) {
      if (child.nodeType !== 1) continue;
      const element = child as Element;
      if (element.namespaceURI === WORD && element.localName === "t") events.push(element.textContent ?? "");
      if (element.namespaceURI === WORD && element.localName === "drawing") events.push("drawing");
    }
  }
  assert.deepEqual(events, ["before", "drawing", "after"]);
  assert.equal(prepared.inventory.occurrences.length, 1);
  assert.equal(prepared.issues.length, 0);
  const inline = elements(document, DRAWING, "inline")[0];
  assert.ok(inline);
  const extent = elements(document, DRAWING, "extent")[0];
  assert.equal(extent.getAttribute("cx"), "786130");
  assert.equal(extent.getAttribute("cy"), "393065");
  const image = elements(document, PICTURE, "pic")[0];
  assert.ok(image);
  let drawingRun: Node | null = image;
  while (drawingRun && (drawingRun as Element).localName !== "r") {
    drawingRun = drawingRun.parentNode;
  }
  assert.ok(drawingRun && (drawingRun as Element).namespaceURI === WORD);
  assert.equal((drawingRun as Element).getElementsByTagNameNS(WORD, "vertAlign")[0]?.getAttributeNS(WORD, "val"), "superscript");
  assert.deepEqual(new Uint8Array(source), original);
});

void test("preserves a positioned anchor, 1-inch extent, offsets and stacking order", async () => {
  const source = await updateDocument(await makeEquationFixture(), (xml) => xml.replace(
    'style="width:36pt;height:18pt"',
    'style="position:absolute;left:12pt;top:4pt;width:1in;height:30.95pt;mso-position-horizontal-relative:page;mso-position-vertical-relative:paragraph;z-index:7"',
  ));
  const prepared = await prepareDocxPreviewPackage(source);
  const zip = await loadPackage(prepared.buffer);
  const document = parseXml(await assertNoActiveEquationPayload(zip), "word/document.xml");
  const anchor = elements(document, DRAWING, "anchor")[0];
  assert.ok(anchor);
  assert.equal(anchor.getAttribute("relativeHeight"), "7");
  const extent = elements(document, DRAWING, "extent")[0];
  assert.equal(extent.getAttribute("cx"), "914400");
  assert.equal(extent.getAttribute("cy"), "393065");
  const horizontal = elements(anchor.ownerDocument, DRAWING, "positionH").find((element) => element.parentNode === anchor);
  const vertical = elements(anchor.ownerDocument, DRAWING, "positionV").find((element) => element.parentNode === anchor);
  assert.equal(horizontal?.getAttribute("relativeFrom"), "page");
  assert.equal(horizontal && elements(anchor.ownerDocument, DRAWING, "posOffset").find((node) => node.parentNode === horizontal)?.textContent, "152400");
  assert.equal(vertical?.getAttribute("relativeFrom"), "paragraph");
  assert.equal(vertical && elements(anchor.ownerDocument, DRAWING, "posOffset").find((node) => node.parentNode === vertical)?.textContent, "50800");
});

void test("renders two occurrences sharing a source preview as two independent image occurrences", async () => {
  const source = await makeEquationFixture({ duplicate: true });
  const prepared = await prepareDocxPreviewPackage(source);
  const zip = await loadPackage(prepared.buffer);
  const document = parseXml(await assertNoActiveEquationPayload(zip), "word/document.xml");
  const relationships = parseXml(await zip.file("word/_rels/document.xml.rels")!.async("string"), "word/_rels/document.xml.rels");
  const images = elements(document, PICTURE, "pic");
  const imageRelationshipIds = elements(document, "http://schemas.openxmlformats.org/drawingml/2006/main", "blip")
    .map((element) => element.getAttributeNS(REL, "embed"));
  const relationshipsById = new Map(elements(relationships, PACKAGE_REL, "Relationship").map((entry) =>
    [entry.getAttribute("Id"), entry] as const
  ));
  const occurrenceRelationships = imageRelationshipIds.map((id) => relationshipsById.get(id));
  assert.equal(prepared.inventory.occurrences.length, 2);
  assert.equal(images.length, 2);
  assert.equal(occurrenceRelationships.length, 2);
  assert.ok(occurrenceRelationships.every((entry) => entry?.getAttribute("Type") === `${REL}/image`));
  assert.notEqual(imageRelationshipIds[0], imageRelationshipIds[1]);
  assert.equal(occurrenceRelationships[0]?.getAttribute("Target"), occurrenceRelationships[1]?.getAttribute("Target"));
  assert.equal(Object.keys(zip.files).filter((name) => name.startsWith("word/word-reader-preview-")).length, 1);
});

void test("uses object-specific visible diagnostics when the preview relationship is missing", async () => {
  const source = await makeEquationFixture({ missingPreview: true });
  const prepared = await prepareDocxPreviewPackage(source);
  const zip = await loadPackage(prepared.buffer);
  const document = await assertNoActiveEquationPayload(zip);
  assert.equal(prepared.inventory.occurrences.length, 1);
  assert.equal(prepared.inventory.occurrences[0].previewPath, null);
  assert.equal(prepared.issues.length, 1);
  assert.equal(prepared.issues[0].code, "missing-preview");
  assert.equal(prepared.issues[0].occurrenceId, prepared.inventory.occurrences[0].occurrenceId);
  assert.match(document, /Preview unavailable/);
  assert.equal(/<w:drawing\b/.test(document), false);
});

void test("resolves WMF and EMF previews through T7 into readable safe generated images", async () => {
  for (const [extension, mediaBytes] of [
    ["wmf", createSyntheticWmfFixture([])],
    ["emf", createSyntheticEmfFixture([])],
  ] as const) {
    const base = await makeEquationFixture();
    const zip = await JSZip.loadAsync(base);
    zip.remove("word/media/preview.png");
    zip.file(`word/media/preview.${extension}`, mediaBytes);
    const types = await zip.file("[Content_Types].xml")!.async("string");
    zip.file("[Content_Types].xml", types.replace(
      '<Default Extension="png" ContentType="image/png"/>',
      `<Default Extension="${extension}" ContentType="image/x-${extension}"/>`,
    ));
    const relationships = await zip.file("word/_rels/document.xml.rels")!.async("string");
    zip.file("word/_rels/document.xml.rels", relationships.replace("media/preview.png", `media/preview.${extension}`));
    const source = await zip.generateAsync({ type: "arraybuffer", compression: "STORE" });
    const prepared = await prepareDocxPreviewPackage(source);
    const output = await loadPackage(prepared.buffer);
    const generatedName = Object.keys(output.files).find((name) =>
      name.startsWith("word/word-reader-preview-") && name.endsWith(".svg")
    );
    assert.ok(generatedName);
    const generated = await output.file(generatedName)!.async("uint8array");
    assert.match(new TextDecoder().decode(generated), /^<svg\b/);
    const contentTypes = await output.file("[Content_Types].xml")!.async("string");
    assert.match(contentTypes, new RegExp(`<Override PartName="/${generatedName}" ContentType="image/svg\\+xml"`));
    assert.equal(prepared.issues.length, 0);
  }
});

void test("rejects unsupported positioned VML instead of silently converting it to inline", async () => {
  const source = await updateDocument(await makeEquationFixture(), (xml) => xml.replace(
    'style="width:36pt;height:18pt"',
    'style="position:fixed;width:36pt;height:18pt"',
  ));
  await assert.rejects(
    () => prepareDocxPreviewPackage(source),
    (error: unknown) => error instanceof Error && /unsupported-placement/.test(`${(error as { code?: string }).code} ${error.message}`),
  );
});

void test("observes cancellation before reading retained package entries", async () => {
  const source = await makeEquationFixture();
  await assert.rejects(
    () => prepareDocxPreviewPackage(source, { isCancelled: () => true }),
    (error: unknown) => error instanceof Error && /cancel/i.test(`${(error as { code?: string }).code} ${error.message}`),
  );
});

void test("removes namespaced altChunk references and their content with an aliased prefix", async () => {
  const zip = await JSZip.loadAsync(await createLargeDocx(1));
  const target = "word/chunks/part.html";
  const documentXml = await zip.file("word/document.xml")!.async("string");
  zip.file("word/document.xml", documentXml
    .replace(
      'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
      'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:q="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
    )
    .replace("<w:body>", '<w:body><q:altChunk r:id="chunk"/>'));
  zip.file("word/_rels/document.xml.rels", `<Relationships xmlns="${PACKAGE_REL}"><Relationship Id="chunk" Type="${REL}/aFChunk" Target="chunks/part.html"/></Relationships>`);
  const contentTypes = await zip.file("[Content_Types].xml")!.async("string");
  zip.file("[Content_Types].xml", contentTypes.replace("</Types>", `<Override PartName="/${target}" ContentType="text/html"/></Types>`));
  zip.file(target, "<html><script>never execute</script></html>");

  const source = await zip.generateAsync({ type: "arraybuffer", compression: "STORE" });
  const prepared = await prepareDocxPreviewPackage(source);
  const output = await loadPackage(prepared.buffer);
  const document = parseXml(await output.file("word/document.xml")!.async("string"), "word/document.xml");
  const relationships = await output.file("word/_rels/document.xml.rels")!.async("string");
  assert.equal(elements(document, WORD, "altChunk").length, 0);
  assert.equal(output.file(target), null);
  assert.doesNotMatch(relationships, /aFChunk/i);
});
