/* eslint-disable obsidianmd/no-global-this -- Node-only test shim installs xmldom's serializer on the test global. */
import assert from "node:assert/strict";
import { createRequire } from "node:module";
import test, { mock } from "node:test";
import JSZip from "jszip";
import { DOMParser, XMLSerializer } from "@xmldom/xmldom";

import { createDocxRenderer } from "../../src/docx/docxRenderer";
import * as docxEngine from "../../src/docx/docxEngine";
import { createEquationFixture } from "./fixtureBuilders";

const require = createRequire(__filename);
const docxPreview = require("docx-preview") as {
  parseAsync: (...args: unknown[]) => Promise<unknown>;
  renderDocument: (...args: unknown[]) => Promise<void>;
};
const renderDocx = createDocxRenderer(docxEngine);
const ownerDocument: Document = new DOMParser().parseFromString(
  "<html><body></body></html>",
  "text/html",
);
const targetEl = { ownerDocument } as HTMLElement;

async function withXmlSerializer<T>(operation: () => Promise<T>): Promise<T> {
  const previous = Object.getOwnPropertyDescriptor(globalThis, "XMLSerializer");
  Object.defineProperty(globalThis, "XMLSerializer", {
    configurable: true,
    writable: true,
    value: XMLSerializer,
  });
  try {
    return await operation();
  } finally {
    if (previous) Object.defineProperty(globalThis, "XMLSerializer", previous);
    else Reflect.deleteProperty(globalThis, "XMLSerializer");
  }
}

async function makeFixture(options: { missingPreview?: boolean } = {}): Promise<ArrayBuffer> {
  return await createEquationFixture("docx", {
    progId: "Equation.3",
    previewRelationship: { id: "image", target: "media/preview.png" },
    previewBytes: Uint8Array.from([137, 80, 78, 71, 13, 10, 26, 10, 0]),
    opaqueEmbeddingBytes: new TextEncoder().encode("opaque equation payload"),
    ...options,
  });
}

async function withUnsupportedPlacement(buffer: ArrayBuffer): Promise<ArrayBuffer> {
  const zip = await JSZip.loadAsync(buffer);
  const document = zip.file("word/document.xml")!;
  zip.file("word/document.xml", (await document.async("string")).replace(
    'style="width:36pt;height:18pt"',
    'style="position:fixed;width:36pt;height:18pt"',
  ));
  return await zip.generateAsync({ type: "arraybuffer", compression: "STORE" });
}

void test("renders only the distinct inert package and returns object issues", async () => {
  const source = await makeFixture({ missingPreview: true });
  const original = new Uint8Array(source).slice();
  let parsedBuffer: ArrayBuffer | null = null;
  const parseSpy = mock.method(docxPreview, "parseAsync", async (buffer: unknown) => {
    parsedBuffer = buffer as ArrayBuffer;
    return {};
  });
  const renderSpy = mock.method(docxPreview, "renderDocument", async () => undefined);
  try {
    const result = await withXmlSerializer(() => renderDocx(
      source,
      targetEl,
    ));
    assert.equal(parseSpy.mock.callCount(), 1);
    assert.ok(parsedBuffer);
    assert.notEqual(parsedBuffer, source);
    assert.deepEqual(new Uint8Array(source), original);
    const prepared = await JSZip.loadAsync(parsedBuffer);
    assert.equal(Object.keys(prepared.files).some((name) => /(?:^|\/)embeddings\//i.test(name)), false);
    const document = await prepared.file("word/document.xml")!.async("string");
    assert.doesNotMatch(document, /<w:object\b|OLEObject/i);
    assert.match(document, /Preview unavailable/);
    assert.equal(result.inventory.occurrences.length, 1);
    assert.equal(result.issues.length, 1);
    assert.equal(result.issues[0].code, "missing-preview");
    assert.equal(renderSpy.mock.callCount(), 1);
  } finally {
    renderSpy.mock.restore();
    parseSpy.mock.restore();
  }
});

void test("does not call the parser when preparation is cancelled or placement is unsupported", async () => {
  const source = await makeFixture();
  const unsupported = await withUnsupportedPlacement(source);
  const parseSpy = mock.method(docxPreview, "parseAsync", async () => ({}));
  try {
    await assert.rejects(
      () => withXmlSerializer(() => renderDocx(source, targetEl, {
        isCancelled: () => true,
      })),
      /cancel/i,
    );
    await assert.rejects(
      () => withXmlSerializer(() => renderDocx(unsupported, targetEl)),
      (error: unknown) => error instanceof Error &&
        (error as { code?: string }).code === "unsupported-placement",
    );
    assert.equal(parseSpy.mock.callCount(), 0);
  } finally {
    parseSpy.mock.restore();
  }
});
