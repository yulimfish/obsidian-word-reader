import assert from "node:assert/strict";
import test from "node:test";
import JSZip from "jszip";
import { loadOoxmlZip } from "../../src/ooxml/jszipReader";

void test("SVG entries preserve original async reader and assign an image MIME to blobs", async () => {
  const input = await new JSZip().file("word/preview.svg", '<svg xmlns="http://www.w3.org/2000/svg"/>').generateAsync({ type: "arraybuffer" });
  const zip = await loadOoxmlZip(input);
  const entry = zip.file("word/preview.svg")!;
  const original = entry;
  const first = zip.file("word/preview.svg")!;
  const second = zip.file("word/preview.svg")!;
  assert.equal(first, second);
  assert.equal(original, entry);
  const blob = await second.async("blob");
  assert.equal(blob.type, "image/svg+xml");
  assert.equal(await blob.text(), '<svg xmlns="http://www.w3.org/2000/svg"/>');
  assert.equal(await entry.async("string"), '<svg xmlns="http://www.w3.org/2000/svg"/>');
});
