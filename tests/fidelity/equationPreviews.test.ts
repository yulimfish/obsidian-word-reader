import assert from "node:assert/strict";
import test from "node:test";
import JSZip from "jszip";
import { inspectEquationPreviews } from "../../src/ooxml/equationPreviews";
import { createEquationFixture } from "./fixtureBuilders";
import { loadOoxmlZip } from "../../src/ooxml/jszipReader";

void test("equation inventory selects one alternative and preserves independent shared previews without payload reads", async () => {
  for (const format of ["docx", "pptx"] as const) {
    for (const progId of ["Equation.3", "Equation.DSMT4"] as const) {
      const zip = await loadOoxmlZip(await createEquationFixture(format, {
        progId, previewRelationship: { id: "image", target: `${format === "docx" ? "" : "../"}media/preview.png` },
        previewBytes: Uint8Array.from([137, 80, 78, 71]), opaqueEmbeddingBytes: new TextEncoder().encode("<script>never execute</script>"),
        alternative: true, duplicate: true,
      }));
      let reads = 0;
      for (const entry of Object.values(zip.files).filter((entry) => entry.name.includes("/embeddings/"))) {
        entry.async = () => { reads++; throw new Error("Payload read"); };
      }
      const inventory = await inspectEquationPreviews(zip, format);
      assert.equal(inventory.occurrences.length, 2);
      assert.equal(inventory.issues.length, 0);
      assert.notEqual(inventory.occurrences[0].occurrenceId, inventory.occurrences[1].occurrenceId);
      assert.notEqual(inventory.occurrences[0].shapeId, inventory.occurrences[1].shapeId);
      assert.equal(inventory.occurrences[0].previewPath, inventory.occurrences[1].previewPath);
      assert.notEqual(inventory.occurrences[0].sourceElementPath, inventory.occurrences[1].sourceElementPath);
      assert.ok(inventory.occurrences[0].sourceElementPath.includes("mc:Fallback"));
      assert.equal(reads, 0);
    }
  }
});

void test("relationship targets cannot expand fixed metadata roles for ordinary opaque XML or RELS payloads", async () => {
  for (const format of ["docx", "pptx"] as const) {
    for (const extension of ["xml", "rels"]) {
      const zip = await JSZip.loadAsync(await createEquationFixture(format, {
        progId: "Equation.3", previewRelationship: { id: "image", target: `${format === "docx" ? "" : "../"}media/preview.png` },
        previewBytes: new Uint8Array([137]), opaqueEmbeddingBytes: new Uint8Array([1]),
      }));
      const prefix = format === "docx" ? "word" : "ppt";
      const payload = `${prefix}/opaque.${extension}`;
      const relPath = format === "docx" ? "word/_rels/document.xml.rels" : "ppt/slides/_rels/slide1.xml.rels";
      zip.remove(`${prefix}/embeddings/oleObject1.bin`);
      zip.file(payload, '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>');
      zip.file(relPath, (await zip.file(relPath)!.async("string")).replace(`${format === "docx" ? "" : "../"}embeddings/oleObject1.bin`, `${format === "docx" ? "" : "../"}opaque.${extension}`));
      zip.file("[Content_Types].xml", (await zip.file("[Content_Types].xml")!.async("string")).replace(`/${prefix}/embeddings/oleObject1.bin`, `/${payload}`));
      const production = await loadOoxmlZip(await zip.generateAsync({ type: "arraybuffer" }));
      const modes: unknown[] = [];
      production.file(payload)!.async = (...args) => { modes.push(args[0]); throw new Error("Opaque target read"); };
      assert.equal((await inspectEquationPreviews(production, format)).occurrences[0].embeddingPath, payload);
      assert.deepEqual(modes, []);
    }
  }
});

void test("missing equation preview is an object-specific issue, never guessed text", async () => {
  const zip = await JSZip.loadAsync(await createEquationFixture("docx", {
    progId: "Equation.3", previewRelationship: { id: "image", target: "media/preview.png" },
    previewBytes: new Uint8Array(), opaqueEmbeddingBytes: new Uint8Array([1]), missingPreview: true,
  }));
  const inventory = await inspectEquationPreviews(zip, "docx");
  assert.equal(inventory.occurrences.length, 1);
  assert.equal(inventory.occurrences[0].previewPath, null);
  assert.equal(inventory.issues[0].code, "missing-preview");
  assert.equal(inventory.issues[0].occurrenceId, inventory.occurrences[0].occurrenceId);
});
