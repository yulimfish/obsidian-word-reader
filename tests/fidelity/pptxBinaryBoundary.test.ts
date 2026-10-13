import assert from "node:assert/strict";
import test, { mock } from "node:test";
import JSZip from "jszip";
import { PptxPackage, PptxPackageError } from "../../src/pptx/pptxPackage";
import { collectValidatedPptxImagePaths, inspectEquationPreviews } from "../../src/ooxml/equationPreviews";
import ReadOnlyZip, { loadOoxmlZip } from "../../src/ooxml/jszipReader";
import { enforceOoxmlPackagePolicy } from "../../src/ooxml/packagePolicy";
import { createMinimalPptx } from "../pptxFixture";
import { createEquationFixture } from "./fixtureBuilders";

type EntryRead = { path: string; mode: unknown };

// Observe the actual production host's returned entry instances, not JSZip prototypes.
function observeProductionReads(reads: EntryRead[]): () => void {
  const originalLoad = ReadOnlyZip.loadAsync.bind(ReadOnlyZip) as (buffer: ArrayBuffer) => Promise<ReadOnlyZip>;
  const entrySpies: Array<{ mock: { restore: () => void } }> = [];
  const loadSpy = mock.method(ReadOnlyZip, "loadAsync", async (buffer: ArrayBuffer) => {
    const zip = await originalLoad(buffer);
    for (const entry of Object.values(zip.files)) {
      const original = entry.async.bind(entry) as typeof entry.async;
      entrySpies.push(mock.method(entry, "async", (...args: Parameters<typeof entry.async>) => {
        reads.push({ path: entry.name, mode: args[0] });
        return Reflect.apply(original, entry, args);
      }));
    }
    return zip;
  });
  return () => { for (const spy of entrySpies) spy.mock.restore(); loadSpy.mock.restore(); };
}

for (const variant of ["earlier-metadata", "later-metadata", "self-reference", "cross-cycle", "root", "types", "ole-type", "ole-name"] as const) {
  for (const reverse of [false, true]) {
  void test(`OLE metadata-target ${variant}, reverse=${reverse}, permits only private strings and rejects`, async (context) => {
    const zip = await JSZip.loadAsync(await createEquationFixture("pptx", {
      progId: "Equation.3", previewRelationship: { id: "image", target: "../media/preview.png" },
      previewBytes: new Uint8Array([137, 80, 78, 71]), opaqueEmbeddingBytes: new Uint8Array([1]),
    }));
    const declaration = "ppt/slides/_rels/slide1.xml.rels";
    const target = variant === "self-reference" ? declaration : variant === "later-metadata" ? "ppt/z/_rels/a.xml.rels" : variant === "root" ? "_rels/.rels" : variant === "types" ? "[Content_Types].xml" : variant === "ole-name" ? "ppt/embeddings/_rels/a.xml.rels" : "ppt/_rels/a.xml.rels";
    zip.remove("ppt/embeddings/oleObject1.bin");
    zip.file("[Content_Types].xml", (await zip.file("[Content_Types].xml")!.async("string")).replace(/<Override PartName="\/ppt\/embeddings\/oleObject1\.bin"[^>]*\/>/, ""));
    const relative = variant === "self-reference" ? "_rels/slide1.xml.rels" : variant === "later-metadata" ? "../z/_rels/a.xml.rels" : variant === "root" ? "../../_rels/.rels" : variant === "types" ? "../../[Content_Types].xml" : variant === "ole-name" ? "../embeddings/_rels/a.xml.rels" : "../_rels/a.xml.rels";
    zip.file(declaration, (await zip.file(declaration)!.async("string")).replace("../embeddings/oleObject1.bin", relative));
    if (variant !== "self-reference" && variant !== "root" && variant !== "types") {
      const cross = variant === "cross-cycle" ? '<Relationship Id="back" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="slides/_rels/slide1.xml.rels"/>' : "";
      zip.file(target, `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${cross}</Relationships>`);
    }
    if (variant === "ole-type") zip.file("[Content_Types].xml", (await zip.file("[Content_Types].xml")!.async("string")).replace("</Types>", `<Override PartName="/${target}" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/></Types>`));
    const ordered = new JSZip();
    const entries = Object.values(zip.files).filter((entry) => !entry.dir);
    if (reverse) entries.reverse();
    for (const entry of entries) ordered.file(entry.name, await entry.async("uint8array"), { createFolders: false });
    const buffer = await ordered.generateAsync({ type: "arraybuffer" });
    const designated = new Set([target, ...(variant === "cross-cycle" ? [declaration] : [])]);
    const reads: EntryRead[] = [];
    const restore = observeProductionReads(reads);
    try {
      const operations = {
        inventory: async () => inspectEquationPreviews(await loadOoxmlZip(buffer), "pptx"),
        policy: async () => enforceOoxmlPackagePolicy(await loadOoxmlZip(buffer), "pptx", { legacyEquationPreviews: true }),
        allowlist: async () => collectValidatedPptxImagePaths(await loadOoxmlZip(buffer), null),
        load: async () => PptxPackage.load(buffer),
      };
      for (const [operation, run] of Object.entries(operations)) {
        reads.length = 0;
        let error: unknown;
        try { await run(); } catch (caught) { error = caught; }
        const targetReads = reads.filter((read) => designated.has(read.path));
        context.diagnostic(JSON.stringify({ operation, variant, reverse, targetReads, error: String(error) }));
        assert.ok(targetReads.every((read) => read.mode === "string"), "collision metadata permits no other output modes");
        assert.match(String(error), /metadata-payload/, "collision must reject before inventory/allowlist/package return");
        if (operation === "load") assert.ok(error instanceof PptxPackageError && error.kind === "unsupported");
        assert.equal(reads.filter((read) => read.path === "ppt/media/preview.png" || read.path === "ppt/slides/slide1.xml").length, 0, "collision must stop before image or renderer XML reads");
      }
    } finally { restore(); }
  });
  }
}

void test("PPTX binary boundary allows validated images and denies direct, cached and alias payload reads", async () => {
  const reads: EntryRead[] = [];
  const restore = observeProductionReads(reads);
  try {
    const ordinary = await PptxPackage.load(await createMinimalPptx());
    assert.equal((await ordinary.getBinary("ppt/media/image1.png"))?.[0], 137);
    const buffer = await createEquationFixture("pptx", {
      progId: "Equation.DSMT4", previewRelationship: { id: "image", target: "../media/preview.png" },
      previewBytes: new Uint8Array([137, 80, 78, 71]), opaqueEmbeddingBytes: new TextEncoder().encode("malicious opaque"), alternative: true,
    });
    const pptx = await PptxPackage.load(buffer);
    assert.equal(pptx.previewInventory?.occurrences.length, 1);
    assert.equal((await pptx.getBinary("ppt/media/preview.png"))?.[0], 137);
    // Public lookalike properties and a forged inventory cannot alter private grants.
    Object.assign(pptx, {
      readBoundary: { images: new Set(["ppt/embeddings/oleObject1.bin"]), metadata: { embeddingPaths: new Set() } },
      previewInventory: { occurrences: [{ previewPath: "ppt/embeddings/oleObject1.bin" }], issues: [] },
    });
    for (const path of ["ppt/embeddings/oleObject1.bin", "ppt/embeddings/oleObject1.bin", "ppt/media/../embeddings/oleObject1.bin", "/ppt/embeddings/oleObject1.bin", "ppt/%65mbeddings/oleObject1.bin", "ppt\\embeddings\\oleObject1.bin", "ppt/slides/slide1.xml"]) {
      await assert.rejects(() => pptx.getBinary(path), (error: unknown) => error instanceof PptxPackageError && error.kind === "unsupported");
    }
    pptx.clearCaches();
    await assert.rejects(() => pptx.getBinary("ppt/embeddings/oleObject1.bin"), PptxPackageError);
    assert.equal(reads.filter((read) => read.path.includes("/embeddings/")).length, 0);
    assert.ok(reads.some((read) => read.path === "ppt/media/preview.png" && read.mode === "uint8array"));
  } finally { restore(); }
});

void test("PPTX allowlist validates namespaces, types, content types and rejects forged inventory", async () => {
  const zip = await JSZip.loadAsync(await createMinimalPptx());
  zip.file("ppt/media/unreferenced.png", new Uint8Array([1]));
  const forged = { occurrences: [{ previewPath: "ppt/media/unreferenced.png" }], issues: [] };
  const paths = await collectValidatedPptxImagePaths(zip, forged as never);
  assert.equal(paths.has("ppt/media/unreferenced.png"), false);
  for (const [from, to] of [["relationships/image", "evil/image"], ["package/2006/relationships", "evil/relationships"], ["../media/image1.png", "../media/%69mage1.png"]]) {
    const changed = await JSZip.loadAsync(await createMinimalPptx());
    const path = "ppt/slides/_rels/slide1.xml.rels";
    changed.file(path, (await changed.file(path)!.async("string")).replace(from, to));
    await assert.rejects(() => collectValidatedPptxImagePaths(changed, null));
  }
  const content = await zip.file("[Content_Types].xml")!.async("string");
  zip.file("[Content_Types].xml", content.replace("image/png", "application/vnd.openxmlformats-officedocument.oleObject"));
  await assert.rejects(() => collectValidatedPptxImagePaths(zip, null));
});

void test("validated equation WMF and EMF previews remain readable without payload reads", async () => {
  for (const extension of ["wmf", "emf"]) {
    const zip = await JSZip.loadAsync(await createEquationFixture("pptx", {
      progId: "Equation.3", previewRelationship: { id: "image", target: "../media/preview.png" },
      previewBytes: new Uint8Array([7, 8]), opaqueEmbeddingBytes: new Uint8Array([1]),
    }));
    zip.remove("ppt/media/preview.png");
    zip.file(`ppt/media/preview.${extension}`, new Uint8Array([7, 8]));
    const rels = "ppt/slides/_rels/slide1.xml.rels";
    zip.file(rels, (await zip.file(rels)!.async("string")).replace("preview.png", `preview.${extension}`));
    zip.file("[Content_Types].xml", (await zip.file("[Content_Types].xml")!.async("string")).replace('Extension="png" ContentType="image/png"', `Extension="${extension}" ContentType="image/x-${extension}"`));
    const pptx = await PptxPackage.load(await zip.generateAsync({ type: "arraybuffer" }));
    assert.deepEqual(await pptx.getBinary(`ppt/media/preview.${extension}`), new Uint8Array([7, 8]));
  }
});

void test("OLE payload renamed PNG cannot grant image or XML access and is never read", async () => {
  for (const route of ["direct", "image", "notesSlide"]) {
    const zip = await JSZip.loadAsync(await createEquationFixture("pptx", {
      progId: "Equation.3", previewRelationship: { id: "image", target: "../media/preview.png" },
      previewBytes: new Uint8Array([137]), opaqueEmbeddingBytes: new Uint8Array([1]),
    }));
    zip.remove("ppt/embeddings/oleObject1.bin");
    zip.file("ppt/media/payload.png", "opaque payload");
    const relPath = "ppt/slides/_rels/slide1.xml.rels";
    let rels = (await zip.file(relPath)!.async("string")).replace("../embeddings/oleObject1.bin", "../media/payload.png");
    if (route !== "direct") rels = rels.replace("</Relationships>", `<Relationship Id="forged" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${route}" Target="../media/payload.png"/></Relationships>`);
    zip.file(relPath, rels);
    zip.file("[Content_Types].xml", (await zip.file("[Content_Types].xml")!.async("string")).replace("/ppt/embeddings/oleObject1.bin", "/ppt/media/payload.png"));
    const buffer = await zip.generateAsync({ type: "arraybuffer" });
    const reads: EntryRead[] = [];
    const restore = observeProductionReads(reads);
    try {
      if (route === "direct") {
        const pptx = await PptxPackage.load(buffer);
        await assert.rejects(() => pptx.getBinary("ppt/media/payload.png"), PptxPackageError);
      } else await assert.rejects(() => PptxPackage.load(buffer), PptxPackageError);
      assert.equal(reads.filter((read) => read.path === "ppt/media/payload.png").length, 0);
    } finally { restore(); }
  }
});

void test("forged slide, layout, master and notes targets never dereference equation payload", async () => {
  for (const kind of ["slide", "slideLayout", "slideMaster", "notesSlide"]) {
    const zip = await JSZip.loadAsync(await createEquationFixture("pptx", {
      progId: "Equation.3", previewRelationship: { id: "image", target: "../media/preview.png" },
      previewBytes: new Uint8Array([137]), opaqueEmbeddingBytes: new Uint8Array([1]),
    }));
    const path = kind === "slide" ? "ppt/_rels/presentation.xml.rels" : "ppt/slides/_rels/slide1.xml.rels";
    let rels = await zip.file(path)!.async("string");
    if (kind === "slide") rels = rels.replace("slides/slide1.xml", "embeddings/oleObject1.bin");
    else rels = rels.replace("</Relationships>", `<Relationship Id="forged" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${kind}" Target="../embeddings/oleObject1.bin"/></Relationships>`);
    zip.file(path, rels);
    const reads: EntryRead[] = [];
    const restore = observeProductionReads(reads);
    try {
      await assert.rejects(async () => {
        const pptx = await PptxPackage.load(await zip.generateAsync({ type: "arraybuffer" }));
        await pptx.getSlideContext(0);
        await pptx.getSlideMetadata(0);
      });
      assert.equal(reads.filter((read) => read.path.includes("/embeddings/")).length, 0);
    } finally { restore(); }
  }
});
