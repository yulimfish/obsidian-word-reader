import assert from "node:assert/strict";
import test, { mock } from "node:test";
import { createRequire } from "node:module";
import type { App } from "obsidian";
import JSZip from "jszip";
import { DocxAdapter } from "../../src/docx/DocxAdapter";
import { renderDocx } from "../../src/docx/docxRenderer";
import { enforceOoxmlPackagePolicy } from "../../src/ooxml/packagePolicy";
import { createEquationFixture } from "./fixtureBuilders";
import { createLargeDocx } from "../performanceFixtures";
import ReadOnlyZip, { loadOoxmlZip } from "../../src/ooxml/jszipReader";

const require = createRequire(__filename);
const parser = require("docx-preview") as { parseAsync: (...args: unknown[]) => Promise<unknown> };

type EntryRead = { path: string; mode: unknown };
function observeProductionReads(reads: EntryRead[]): () => void {
  const originalLoad = ReadOnlyZip.loadAsync.bind(ReadOnlyZip) as (buffer: ArrayBuffer) => Promise<ReadOnlyZip>;
  const entries: Array<{ mock: { restore: () => void } }> = [];
  const loadSpy = mock.method(ReadOnlyZip, "loadAsync", async (buffer: ArrayBuffer) => {
    const zip = await originalLoad(buffer);
    for (const entry of Object.values(zip.files)) {
      const original = entry.async.bind(entry) as typeof entry.async;
      entries.push(mock.method(entry, "async", (...args: Parameters<typeof entry.async>) => {
        reads.push({ path: entry.name, mode: args[0] });
        return Reflect.apply(original, entry, args);
      }));
    }
    return zip;
  });
  return () => { for (const spy of entries) spy.mock.restore(); loadSpy.mock.restore(); };
}

void test("DOCX original equation source remains rejected before parser through adapter and direct renderer", async () => {
  const buffer = await createEquationFixture("docx", {
    progId: "Equation.3", previewRelationship: { id: "image", target: "media/preview.png" },
    previewBytes: new Uint8Array([137]), opaqueEmbeddingBytes: new TextEncoder().encode("malicious"),
  });
  const before = new Uint8Array(buffer).slice();
  const reads: EntryRead[] = [];
  const restore = observeProductionReads(reads);
  const parseSpy = mock.method(parser, "parseAsync", async () => { throw new Error("parser reached"); });
  try {
    assert.equal((await enforceOoxmlPackagePolicy(await loadOoxmlZip(buffer), "docx", { legacyEquationPreviews: true }))?.occurrences.length, 1);
    const app = { vault: { readBinary: async () => buffer } } as unknown as App;
    await assert.rejects(() => new DocxAdapter().open(app, {} as never), /OLE/);
    await assert.rejects(() => renderDocx(buffer, {} as HTMLElement), /OLE/);
    assert.equal(parseSpy.mock.callCount(), 0);
    assert.equal(reads.filter((read) => read.path.includes("/embeddings/")).length, 0);
    assert.ok(reads.some((read) => read.path === "word/_rels/document.xml.rels" && read.mode === "string"));
    assert.deepEqual(new Uint8Array(buffer), before);
  } finally { restore(); parseSpy.mock.restore(); }
});

void test("DOCX metadata-role collisions reject inventory opt-in and keep both runtime parser routes closed", async (context) => {
  for (const target of ["word/_rels/a.xml.rels", "word/_rels/document.xml.rels", "_rels/.rels", "[Content_Types].xml"]) {
    for (const reverse of [false, true]) {
      const zip = await JSZip.loadAsync(await createEquationFixture("docx", {
        progId: "Equation.3", previewRelationship: { id: "image", target: "media/preview.png" },
        previewBytes: new Uint8Array([137]), opaqueEmbeddingBytes: new Uint8Array([1]),
      }));
      zip.remove("word/embeddings/oleObject1.bin");
      zip.file("[Content_Types].xml", (await zip.file("[Content_Types].xml")!.async("string")).replace(/<Override PartName="\/word\/embeddings\/oleObject1\.bin"[^>]*\/>/, ""));
      const declaration = "word/_rels/document.xml.rels";
      const relative = target.startsWith("word/") ? target.slice(5) : `../${target}`;
      zip.file(declaration, (await zip.file(declaration)!.async("string")).replace("embeddings/oleObject1.bin", relative));
      if (target === "word/_rels/a.xml.rels") zip.file(target, '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>');
      const ordered = new JSZip();
      const files = Object.values(zip.files).filter((entry) => !entry.dir);
      if (reverse) files.reverse();
      for (const file of files) ordered.file(file.name, await file.async("uint8array"), { createFolders: false });
      const buffer = await ordered.generateAsync({ type: "arraybuffer" });
      const reads: EntryRead[] = [];
      const restore = observeProductionReads(reads);
      const parseSpy = mock.method(parser, "parseAsync", async () => { throw new Error("parser reached"); });
      try {
        await assert.rejects(async () => enforceOoxmlPackagePolicy(await loadOoxmlZip(buffer), "docx", { legacyEquationPreviews: true }), /metadata-payload/);
        const app = { vault: { readBinary: async () => buffer } } as unknown as App;
        await assert.rejects(() => new DocxAdapter().open(app, {} as never), /OLE|metadata-payload/);
        await assert.rejects(() => renderDocx(buffer, {} as HTMLElement), /OLE|metadata-payload/);
        assert.equal(parseSpy.mock.callCount(), 0);
        const targetReads = reads.filter((read) => read.path === target);
        assert.ok(targetReads.every((read) => read.mode === "string"));
        assert.equal(reads.filter((read) => read.path === "word/document.xml" || read.path.includes("/embeddings/")).length, 0);
        context.diagnostic(JSON.stringify({ target, reverse, targetReads, parserCalls: 0 }));
      } finally { restore(); parseSpy.mock.restore(); }
    }
  }
});

void test("ordinary strict DOCX still reaches the parser", async () => {
  const spy = mock.method(parser, "parseAsync", async () => { throw new Error("parser reached"); });
  try {
    const buffer = await createLargeDocx(1);
    await assert.rejects(() => renderDocx(buffer, {} as HTMLElement), /parser reached/);
    assert.equal(spy.mock.callCount(), 1);
    assert.equal(spy.mock.calls[0].arguments[0], buffer);
  } finally { spy.mock.restore(); }
});
