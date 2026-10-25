import assert from "node:assert/strict";
import test from "node:test";

import { planReaderWorkspaceRecovery } from "../src/reader/workspaceRecovery";

void test("workspace recovery isolates missing, changed, and duplicate views", () => {
  assert.deepEqual(planReaderWorkspaceRecovery([
    { id: "docx-1", format: "docx", path: "a.docx", extension: "docx" },
    { id: "docx-2", format: "docx", path: "a.docx", extension: "docx" },
    { id: "missing", format: "pptx", path: null, extension: null },
    { id: "changed", format: "xlsx", path: "book.xlsm", extension: "xlsm" },
    { id: "valid", format: "xlsx", path: "book.xlsx", extension: "XLSX" },
  ]), [
    { id: "docx-1", reason: null, duplicate: false },
    { id: "docx-2", reason: null, duplicate: true },
    { id: "missing", reason: "missing-file", duplicate: false },
    { id: "changed", reason: "format-changed", duplicate: false },
    { id: "valid", reason: null, duplicate: false },
  ]);
});
