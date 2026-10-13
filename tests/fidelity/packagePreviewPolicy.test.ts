import assert from "node:assert/strict";
import test from "node:test";
import JSZip from "jszip";
import { enforceOoxmlPackagePolicy, OoxmlPolicyError } from "../../src/ooxml/packagePolicy";
import { createEquationFixture } from "./fixtureBuilders";

async function fixture(): Promise<JSZip> {
  return JSZip.loadAsync(await createEquationFixture("docx", {
    progId: "Equation.3", previewRelationship: { id: "image", target: "media/preview.png" },
    previewBytes: Uint8Array.from([137, 80, 78, 71]), opaqueEmbeddingBytes: new Uint8Array([1]),
  }));
}

void test("strict policy rejects equations while internal metadata opt-in validates them", async () => {
  const zip = await fixture();
  await assert.rejects(() => enforceOoxmlPackagePolicy(zip, "docx"), OoxmlPolicyError);
  const result = await enforceOoxmlPackagePolicy(zip, "docx", { legacyEquationPreviews: true });
  assert.equal(result?.occurrences.length, 1);
  zip.file("word/embeddings/unknown.bin", new Uint8Array([1]));
  await assert.rejects(() => enforceOoxmlPackagePolicy(zip, "docx", { legacyEquationPreviews: true }), OoxmlPolicyError);
});

void test("preview opt-in fails closed for malicious relationships, namespaces and parts", async () => {
  const relPath = "word/_rels/document.xml.rels";
  const mutations: Array<(xml: string) => string> = [
    (xml) => xml.replace('Id="rIdOle1"', 'Id="image"'),
    (xml) => xml.replace("package/2006/relationships", "attacker/relationships"),
    (xml) => xml.replace("relationships/image", "relationships/hyperlink"),
    (xml) => xml.replace("media/preview.png", "https://example.invalid/preview.png"),
    (xml) => xml.replace("media/preview.png", "../../../escape.png"),
    (xml) => xml.replace("media/preview.png", "%2e%2e/media/preview.png"),
    (xml) => xml.replace('Target="media/preview.png"', 'Target="media/preview.png" TargetMode="External"'),
    (xml) => xml.replace("media/preview.png", "media/missing.png"),
  ];
  for (const mutate of mutations) {
    const zip = await fixture();
    zip.file(relPath, mutate(await zip.file(relPath)!.async("string")));
    await assert.rejects(() => enforceOoxmlPackagePolicy(zip, "docx", { legacyEquationPreviews: true }), OoxmlPolicyError);
  }
  for (const replacement of [
    ["Equation.3", "Package"], ["urn:schemas-microsoft-com:office:office", "urn:attacker"],
    ['ShapeID="Shape1"', 'ShapeID="wrong"'],
  ]) {
    const zip = await fixture();
    zip.file("word/document.xml", (await zip.file("word/document.xml")!.async("string")).replace(replacement[0], replacement[1]));
    await assert.rejects(() => enforceOoxmlPackagePolicy(zip, "docx", { legacyEquationPreviews: true }), OoxmlPolicyError);
  }
});

void test("opt-in retains macro, ActiveX and script media rejection and cannot relax XLSX", async () => {
  for (const path of ["word/vbaProject.bin", "word/activeX/object.xml", "word/media/active.svg", "word/media/active.js"]) {
    const zip = await fixture();
    zip.file(path, "unsafe");
    await assert.rejects(() => enforceOoxmlPackagePolicy(zip, "docx", { legacyEquationPreviews: true }), OoxmlPolicyError);
  }
  const zip = new JSZip();
  zip.file("xl/embeddings/oleObject1.bin", "opaque");
  await assert.rejects(() => enforceOoxmlPackagePolicy(zip, "xlsx", { legacyEquationPreviews: true }), (error: unknown) => error instanceof OoxmlPolicyError && error.kind === "ole-object");
});
