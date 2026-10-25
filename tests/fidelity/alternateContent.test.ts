import assert from "node:assert/strict";
import test from "node:test";
import { XMLSerializer } from "@xmldom/xmldom";
import { parseXml } from "../../src/pptx/xml";
import { normalizeAlternateContent } from "../../src/ooxml/alternateContent";

const mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const serialize = (document: Document): string => new XMLSerializer().serializeToString(document);

void test("selects one branch by URI and retains distinct repeated pictures", () => {
  const document = parseXml(`<root xmlns:mc="${mc}" xmlns:a="urn:supported" xmlns:b="urn:supported" xmlns:u="urn:unknown">
    <mc:AlternateContent><mc:Choice Requires="u"><pic src="wrong"/></mc:Choice><mc:Choice Requires="a b"><pic src="shared"/></mc:Choice><mc:Fallback><pic src="fallback"/></mc:Fallback></mc:AlternateContent>
    <mc:AlternateContent><mc:Choice Requires="u"><pic src="wrong"/></mc:Choice><mc:Fallback><pic src="shared"/></mc:Fallback></mc:AlternateContent>
    <fake:AlternateContent xmlns:fake="urn:attacker"><fake:Fallback/></fake:AlternateContent>
  </root>`, "test.xml");
  const original = serialize(document);
  const result = normalizeAlternateContent(document, new Set(["urn:supported"]));
  assert.equal(serialize(document), original);
  assert.notEqual(result.document, document);
  const pictures = result.document.getElementsByTagName("pic");
  assert.equal(pictures.length, 2);
  assert.equal(pictures.item(0)?.getAttribute("src"), "shared");
  assert.equal(pictures.item(1)?.getAttribute("src"), "shared");
  assert.deepEqual(result.selections.map((item) => item.branch), ["choice", "fallback"]);
  assert.notEqual(result.selections[0].sourceElementPath, result.selections[1].sourceElementPath);
  assert.equal(result.document.getElementsByTagNameNS("urn:attacker", "AlternateContent").length, 1);
  assert.ok(result.issues.some((item) => item.code === "unsupported-requires"));
});

void test("rejects absent fallback, unresolved Requires and excessive nesting", () => {
  const document = parseXml(`<root xmlns:mc="${mc}"><mc:AlternateContent><mc:Choice Requires="missing"><pic/></mc:Choice></mc:AlternateContent></root>`, "test.xml");
  const result = normalizeAlternateContent(document, new Set());
  assert.deepEqual(result.issues.map((item) => item.code), ["unresolved-requires", "missing-fallback"]);
  assert.equal(result.document.getElementsByTagName("pic").length, 0);
  const nested = (count: number): Document => parseXml(`<root xmlns:mc="${mc}">${"<mc:AlternateContent><mc:Fallback>".repeat(count)}<pic/>${"</mc:Fallback></mc:AlternateContent>".repeat(count)}</root>`, "nested.xml");
  assert.equal(normalizeAlternateContent(nested(128), new Set()).issues.length, 0);
  const capped = normalizeAlternateContent(nested(129), new Set());
  assert.ok(capped.issues.some((item) => item.code === "depth-limit"));
  assert.equal(capped.document.getElementsByTagName("pic").length, 0);
});

void test("preserves original nested source paths and enforces a count cap", () => {
  const document = parseXml(`<root xmlns:mc="${mc}"><mc:AlternateContent><mc:Fallback><mc:AlternateContent><mc:Fallback><pic/></mc:Fallback></mc:AlternateContent></mc:Fallback></mc:AlternateContent></root>`, "nested.xml");
  const result = normalizeAlternateContent(document, new Set());
  assert.equal(result.selections[1].sourceElementPath, "/root[1]/mc:AlternateContent[1]/mc:Fallback[1]/mc:AlternateContent[1]");
  const wide = parseXml(`<root xmlns:mc="${mc}">${"<mc:AlternateContent><mc:Fallback><pic/></mc:Fallback></mc:AlternateContent>".repeat(10001)}</root>`, "wide.xml");
  const capped = normalizeAlternateContent(wide, new Set());
  assert.equal(capped.selections.length, 10000);
  assert.ok(capped.issues.some((item) => item.code === "count-limit"));
});
