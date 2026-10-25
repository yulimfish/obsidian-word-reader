import assert from "node:assert/strict";
import test from "node:test";

import { recoverPluginData } from "../src/reader/pluginData";
import { DEFAULT_OFFICE_READER_SETTINGS } from "../src/settingsModel";

void test("corrupted and partial plugin data fall back without losing valid fields", () => {
  assert.deepEqual(
    recoverPluginData("{partial-json").settings,
    DEFAULT_OFFICE_READER_SETTINGS,
  );

  const recovered = recoverPluginData({
    schemaVersion: 1,
    common: {
      language: "en",
      defaultZoomPercent: 125,
      largeFileWarningMb: "partial",
    },
    docx: { defaultFitWidth: true },
    pptx: null,
    unknownFutureField: { enabled: true },
    readingStates: [{
      path: "recover.docx",
      mtime: 10,
      format: "docx",
      lastAccessed: 1,
      state: {
        zoom: 1.5,
        fitWidth: true,
        scrollTop: 80,
      },
    }],
  });

  assert.equal(recovered.settings.common.language, "en");
  assert.equal(recovered.settings.common.defaultZoomPercent, 125);
  assert.equal(
    recovered.settings.common.largeFileWarningMb,
    DEFAULT_OFFICE_READER_SETTINGS.common.largeFileWarningMb,
  );
  assert.equal(recovered.settings.docx.defaultFitWidth, true);
  assert.equal(recovered.readingStates.get("recover.docx")?.scrollTop, 80);
  assert.equal("unknownFutureField" in recovered.normalizedData, false);
  assert.equal(recovered.requiresRepair, true);
});

void test("canonical plugin data does not trigger a repair write", () => {
  const normalized = recoverPluginData({
    ...DEFAULT_OFFICE_READER_SETTINGS,
    readingStates: [],
  });
  assert.equal(normalized.requiresRepair, false);
});

void test("legacy and partially written data migrate independently", () => {
  const legacy = recoverPluginData({
    language: "en",
    defaultZoomPercent: 150,
    defaultFitWidth: true,
    showOutlineByDefault: false,
    enableImagePreview: false,
    largeFileWarningMb: 64,
    readingStates: [{
      path: "legacy.pptx",
      mtime: 4,
      format: "pptx",
      state: {
        zoom: 1.25,
        page: 3,
        notesVisible: true,
      },
    }, {
      path: "partial.xlsx",
      position: null,
      zoom: { scale: "partial" },
    }],
  });

  assert.equal(legacy.settings.common.language, "en");
  assert.equal(legacy.settings.common.defaultZoomPercent, 150);
  assert.equal(legacy.settings.docx.defaultFitWidth, true);
  assert.equal(legacy.settings.docx.showOutlineByDefault, false);
  assert.equal(legacy.settings.docx.enableImagePreview, false);
  assert.equal(legacy.settings.common.largeFileWarningMb, 64);
  assert.equal(legacy.readingStates.get("legacy.pptx")?.page, 3);
  assert.equal(legacy.readingStates.get("legacy.pptx")?.notesVisible, true);
  assert.equal(legacy.readingStates.get("partial.xlsx")?.scrollTop, 0);
  assert.equal(legacy.requiresRepair, true);
});
