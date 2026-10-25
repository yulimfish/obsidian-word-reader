import { createRequire } from "node:module";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { build } from "esbuild";
import JSZip from "jszip";

import { createReadOnlyZipPlugin } from "./read-only-zip-plugin.mjs";

const rootDir = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const outputDir = path.join(rootDir, ".zip-smoke-dist");
const outputPath = path.join(outputDir, "reader.cjs");

fs.rmSync(outputDir, { recursive: true, force: true });
fs.mkdirSync(outputDir, { recursive: true });

try {
  await build({
    entryPoints: [path.join(rootDir, "src", "ooxml", "jszipReader.ts")],
    outfile: outputPath,
    bundle: true,
    platform: "node",
    format: "cjs",
    target: "node18",
    logLevel: "warning",
    plugins: [createReadOnlyZipPlugin(rootDir)],
  });

  const fixture = new JSZip();
  fixture.file("word/document.xml", "<document>read-only smoke</document>");
  const buffer = await fixture.generateAsync({
    type: "arraybuffer",
    compression: "DEFLATE",
  });
  const require = createRequire(import.meta.url);
  const reader = require(outputPath);
  const loaded = await reader.loadOoxmlZip(buffer);
  const entry = loaded.file("word/document.xml");
  if (!entry) {
    throw new Error("Read-only ZIP smoke fixture entry is missing.");
  }
  const text = await entry.async("string");
  if (text !== "<document>read-only smoke</document>") {
    throw new Error("Read-only ZIP async decompression changed content.");
  }

  const chunks = [];
  await new Promise((resolve, reject) => {
    entry.internalStream("string")
      .on("data", (chunk) => chunks.push(chunk))
      .on("error", reject)
      .on("end", resolve)
      .resume();
  });
  if (chunks.join("") !== text) {
    throw new Error("Read-only ZIP streamed decompression changed content.");
  }
  const loadedFromDefault = await reader.default.loadAsync(buffer);
  if (!loadedFromDefault.file("word/document.xml")) {
    throw new Error("The docx-preview-compatible static loader is unavailable.");
  }
  console.log("Read-only ZIP production smoke passed.");
} finally {
  fs.rmSync(outputDir, { recursive: true, force: true });
}
