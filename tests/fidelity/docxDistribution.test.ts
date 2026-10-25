import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import { pathToFileURL } from "node:url";
import test from "node:test";

import { writeStoredZip } from "../../src/ooxml/storedZipWriter";
import { validateZipSafety } from "../../src/ooxml/packageSafety";

async function loadBuildConfig(): Promise<Record<string, unknown>> {
  try {
    // eslint-disable-next-line no-unsanitized/method -- test-only import of the local build configuration module
    return await import(pathToFileURL(
      path.resolve(process.cwd(), "scripts/docx-build-config.mjs"),
    ).href) as Record<string, unknown>;
  } catch (error) {
    assert.fail(
      `DOCX dual-entry build contract is missing: ${error instanceof Error ? error.message : String(error)}`,
    );
  }
}

void test("DOCX distribution exports shared dual-entry build and package contracts", async () => {
  const config = await loadBuildConfig();

  assert.equal(typeof config.createDocxBuildConfig, "function");
  assert.equal(typeof config.buildDocxDistribution, "function");
  assert.deepEqual(config.DOCX_DISTRIBUTION_FILES, [
    "main.js",
    "docx-engine.cjs",
    "manifest.json",
    "styles.css",
  ]);
  assert.equal(config.DOCX_MAIN_LIMIT_BYTES, 512_000);
  assert.equal(config.DOCX_ARCHIVE_LIMIT_BYTES, 8_388_608);
  assert.equal(config.DOCX_ARCHIVE_PAYLOAD_LIMIT_BYTES, 8_388_192);
  assert.equal(config.DOCX_ARCHIVE_OVERHEAD_BYTES, 416);
});

void test("measures and packages the actual dual production outputs from one snapshot", async () => {
  const config = await loadBuildConfig() as {
    buildDocxDistribution: (options: { root: string; production: boolean }) => Promise<{
      buildId: string;
      files: Record<string, Uint8Array>;
      bytes: Record<string, number>;
      hashes: Record<string, string>;
      metafiles: { main: { inputs: Record<string, unknown>; outputs: Record<string, { entryPoint?: string }> }; engine: { inputs: Record<string, unknown>; outputs: Record<string, { entryPoint?: string }> } };
      sharedInputs: string[];
    }>;
    DOCX_DISTRIBUTION_FILES: readonly string[];
    DOCX_MAIN_LIMIT_BYTES: number;
    DOCX_ARCHIVE_LIMIT_BYTES: number;
    DOCX_ARCHIVE_PAYLOAD_LIMIT_BYTES: number;
    DOCX_ARCHIVE_OVERHEAD_BYTES: number;
  };
  const root = process.cwd();
  const snapshot = await config.buildDocxDistribution({ root, production: true });
  const expectedFiles = ["main.js", "docx-engine.cjs", "manifest.json", "styles.css"];
  assert.deepEqual(config.DOCX_DISTRIBUTION_FILES, expectedFiles);
  assert.deepEqual(Object.keys(snapshot.files).sort(), [...expectedFiles].sort());
  assert.equal(Object.keys(snapshot.metafiles.main.outputs).length, 1);
  assert.equal(Object.keys(snapshot.metafiles.engine.outputs).length, 1);
  assert.equal(snapshot.bytes.main, snapshot.files["main.js"].byteLength);
  assert.equal(snapshot.bytes.engine, snapshot.files["docx-engine.cjs"].byteLength);
  assert.equal(snapshot.bytes.manifest, snapshot.files["manifest.json"].byteLength);
  assert.equal(snapshot.bytes.styles, snapshot.files["styles.css"].byteLength);
  assert.ok(snapshot.bytes.main <= config.DOCX_MAIN_LIMIT_BYTES);
  assert.ok(snapshot.bytes.engine <= snapshot.bytes.engineLimit);
  assert.ok(snapshot.sharedInputs.length > 0);

  const mainInputs = Object.keys(snapshot.metafiles.main.inputs).map((input) => input.split("\\").join("/"));
  const engineInputs = Object.keys(snapshot.metafiles.engine.inputs).map((input) => input.split("\\").join("/"));
  assert.equal(mainInputs.some((input) => input.includes("node_modules/docx-preview/")), false);
  assert.equal(mainInputs.some((input) => /src\/(?:docx\/docxPreviewPackage|ooxml\/(?:storedZipWriter|metafileSvg|metafileScene|metafileRecords|wmfScene|emfScene|previewMedia))\.ts$/.test(input)), false);
  assert.ok(engineInputs.some((input) => input.includes("node_modules/docx-preview/")));
  for (const input of ["src/docx/docxPreviewPackage.ts", "src/ooxml/storedZipWriter.ts", "src/ooxml/metafileSvg.ts", "src/ooxml/wmfScene.ts", "src/ooxml/emfScene.ts"]) {
    assert.ok(engineInputs.some((actual) => actual.endsWith(input)), `engine metafile is missing ${input}`);
  }
  for (const metafile of [snapshot.metafiles.main, snapshot.metafiles.engine]) {
    assert.equal(Object.values(metafile.outputs).some((output) => /(?:^|\/)chunk-/.test(output.entryPoint ?? "")), false);
  }

  const ownedDir = fs.mkdtempSync(path.join(__dirname, "docx-distribution-"));
  try {
    const entries = expectedFiles.map((file) => {
      const bytes = snapshot.files[file];
      fs.writeFileSync(path.join(ownedDir, file), bytes);
      const onDisk = fs.readFileSync(path.join(ownedDir, file));
      assert.deepEqual(onDisk, bytes, `${file} must match the measured build snapshot`);
      return { path: file, bytes: onDisk };
    });
    const archive = await writeStoredZip(entries, {
      isCancelled: () => false,
      yieldControl: async () => {},
    });
    const archiveBytes = new Uint8Array(archive);
    const payloadBytes = entries.reduce((total, entry) => total + entry.bytes.byteLength, 0);
    const zipSummary = validateZipSafety(archive, {
      maxFileCount: 4,
      maxEntryUncompressedBytes: config.DOCX_ARCHIVE_PAYLOAD_LIMIT_BYTES,
      maxTotalUncompressedBytes: config.DOCX_ARCHIVE_PAYLOAD_LIMIT_BYTES,
      maxCompressionRatio: 1,
    });
    assert.equal(zipSummary.fileCount, 4);
    assert.equal(zipSummary.totalCompressedBytes, payloadBytes);
    assert.equal(zipSummary.totalUncompressedBytes, payloadBytes);
    assert.equal(zipSummary.entries.every((entry) => entry.compressionMethod === 0), true);
    assert.equal(archiveBytes.byteLength - payloadBytes, config.DOCX_ARCHIVE_OVERHEAD_BYTES);
    assert.ok(payloadBytes <= config.DOCX_ARCHIVE_PAYLOAD_LIMIT_BYTES);
    assert.ok(archiveBytes.byteLength <= config.DOCX_ARCHIVE_LIMIT_BYTES);
    assert.equal(snapshot.bytes.payload, payloadBytes);
    fs.writeFileSync(path.join(ownedDir, "distribution.zip"), archiveBytes);
    assert.deepEqual(new Uint8Array(fs.readFileSync(path.join(ownedDir, "distribution.zip"))), archiveBytes);
    /* diagnostic measurement output retained as a comment for fidelity inspection
    console.log(
      `Measured DOCX package: main=${snapshot.bytes.main} B, engine=${snapshot.bytes.engine} B, ` +
        `manifest=${snapshot.bytes.manifest} B, styles=${snapshot.bytes.styles} B, ` +
        `payload=${payloadBytes} B, STORE ZIP=${archiveBytes.byteLength} B, ` +
        `engine allowance=${snapshot.bytes.engineLimit} B, buildId=${snapshot.buildId}, ` +
        `mainSha256=${snapshot.hashes.main}, engineSha256=${snapshot.hashes.engine}`,
    );
    */
  } finally {
    fs.rmSync(ownedDir, { recursive: true, force: true });
  }
});
