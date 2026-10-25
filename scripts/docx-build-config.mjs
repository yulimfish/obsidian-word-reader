import { createHash } from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";
import { builtinModules } from "node:module";
import { build } from "esbuild";

import { createReadOnlyZipPlugin } from "./read-only-zip-plugin.mjs";

export const DOCX_DISTRIBUTION_FILES = Object.freeze([
  "main.js",
  "docx-engine.cjs",
  "manifest.json",
  "styles.css",
]);
export const DOCX_MAIN_LIMIT_BYTES = 512_000;
export const DOCX_ARCHIVE_LIMIT_BYTES = 8_388_608;
export const DOCX_ARCHIVE_PAYLOAD_LIMIT_BYTES = 8_388_192;
export const DOCX_ARCHIVE_OVERHEAD_BYTES = 416;

const WRITER_ROUTE_PATTERNS = Object.freeze([
  /(?<!un)(?<!de)compressWorker/,
  /createWorkerFrom/,
  /FlateWorker\/Deflate/,
  /STORE compression/,
  /generateInternalStream/,
  /generateNodeStream/,
]);

const ROOT_ASSETS = [
  "package.json",
  "package-lock.json",
  "manifest.json",
  "styles.css",
  "esbuild.config.mjs",
  "scripts/docx-build-config.mjs",
  "scripts/read-only-zip-plugin.mjs",
];
const ENGINE_ONLY_INPUTS = [
  "node_modules/docx-preview/",
  "src/docx/docxEngine.ts",
  "src/docx/docxOutline.ts",
  "src/docx/docxPreviewPackage.ts",
  "src/ooxml/storedZipWriter.ts",
  "src/ooxml/metafileSvg.ts",
  "src/ooxml/metafileScene.ts",
  "src/ooxml/metafileRecords.ts",
  "src/ooxml/wmfScene.ts",
  "src/ooxml/emfScene.ts",
  "src/ooxml/previewMedia.ts",
];
const MAIN_EXTERNALS = [
  "obsidian",
  "electron",
  "@codemirror/autocomplete",
  "@codemirror/collab",
  "@codemirror/commands",
  "@codemirror/language",
  "@codemirror/lint",
  "@codemirror/search",
  "@codemirror/state",
  "@codemirror/view",
  "@lezer/common",
  "@lezer/highlight",
  "@lezer/lr",
  ...builtinModules,
  ...builtinModules.map((name) => `node:${name}`),
];

function sha256(bytes) {
  return createHash("sha256").update(bytes).digest("hex");
}

function normalizePath(value) {
  return value.replaceAll("\\", "/");
}

async function walkFiles(directory) {
  const entries = await fs.readdir(directory, { withFileTypes: true });
  const results = [];
  for (const entry of entries.sort((left, right) => left.name.localeCompare(right.name))) {
    const fullPath = path.join(directory, entry.name);
    if (entry.isDirectory()) {
      results.push(...await walkFiles(fullPath));
    } else if (entry.isFile()) {
      results.push(fullPath);
    }
  }
  return results;
}

async function calculateBuildId(root, production) {
  const sourceFiles = (await walkFiles(path.join(root, "src")))
    .filter((file) => /\.(?:ts|tsx|js|mjs)$/.test(file));
  const inputFiles = [...new Set([
    ...ROOT_ASSETS.map((file) => path.join(root, file)),
    ...sourceFiles,
  ])].sort((left, right) => normalizePath(path.relative(root, left)).localeCompare(
    normalizePath(path.relative(root, right)),
  ));
  const inputs = [];
  for (const file of inputFiles) {
    const bytes = await fs.readFile(file);
    inputs.push({ path: normalizePath(path.relative(root, file)), sha256: sha256(bytes) });
  }
  const configuration = {
    format: "cjs",
    target: "es2018",
    production,
    mainExternal: MAIN_EXTERNALS,
    engineExternal: [],
    aliases: ["jszip=src/ooxml/jszipReader.ts", "pako=node_modules/pako/lib/inflate.js"],
    distributionFiles: DOCX_DISTRIBUTION_FILES,
  };
  return {
    buildId: sha256(Buffer.from(JSON.stringify({ inputs, configuration }))),
    inputs,
    configuration,
  };
}

function defineString(value) {
  return JSON.stringify(value);
}

function defineNumber(value) {
  return JSON.stringify(value);
}

function createEvidenceInputPlugin(captureHooks, entryLabel) {
  return {
    name: "docx-evidence-pinned-inputs",
    setup(build) {
      build.onLoad({ filter: /.*/, namespace: "file" }, (args) => {
        const result = captureHooks.onFileLoad(args, entryLabel);
        if (
          !result ||
          !(typeof result.contents === "string" || result.contents instanceof Uint8Array) ||
          typeof result.loader !== "string" ||
          typeof result.resolveDir !== "string"
        ) {
          throw new Error(`Evidence input hook did not return pinned bytes for ${args.path}.`);
        }
        return result;
      });
    },
  };
}

function serializeBuildOptions(options) {
  return JSON.stringify(Object.fromEntries(
    Object.entries(options).filter(([key]) => key !== "plugins"),
  ));
}

function createCommonOptions(root, production, captureHooks, captureEntry) {
  const plugins = [];
  if (typeof captureHooks?.onFileLoad === "function") {
    plugins.push(createEvidenceInputPlugin(captureHooks, captureEntry));
  }
  plugins.push(createReadOnlyZipPlugin(root, {
    onVirtualLoad: typeof captureHooks?.onVirtualLoad === "function"
      ? (input) => captureHooks.onVirtualLoad({ ...input, entry: captureEntry })
      : undefined,
  }));
  return {
    bundle: true,
    format: "cjs",
    target: "es2018",
    platform: "browser",
    minify: production,
    treeShaking: true,
    sourcemap: false,
    legalComments: "inline",
    write: false,
    metafile: true,
    logLevel: "silent",
    plugins,
  };
}

export function createDocxBuildConfig(options) {
  const root = path.resolve(options.root ?? process.cwd());
  const production = options.production ?? true;
  const buildId = options.buildId;
  const pluginVersion = options.pluginVersion;
  const manifest = options.manifest;
  const manifestBytes = options.manifestBytes;
  const stylesBytes = options.stylesBytes;
  const manifestSha256 = options.manifestSha256;
  const stylesSha256 = options.stylesSha256;
  const engineSha256 = options.engineSha256;
  const engineBytes = options.engineBytes;
  const captureHooks = options.captureHooks;
  const captureEntry = options.captureEntry;
  if (
    typeof buildId !== "string" ||
    typeof pluginVersion !== "string" ||
    typeof manifest?.minAppVersion !== "string" ||
    !(manifestBytes instanceof Uint8Array) ||
    !(stylesBytes instanceof Uint8Array) ||
    typeof manifestSha256 !== "string" ||
    typeof stylesSha256 !== "string"
  ) {
    throw new TypeError("DOCX build configuration requires one immutable asset snapshot.");
  }
  const common = createCommonOptions(root, production, captureHooks, captureEntry);
  const baseDefines = {
    __DEV__: JSON.stringify(!production),
    __DOCX_PLUGIN_VERSION__: defineString(pluginVersion),
    __DOCX_BUILD_ID__: defineString(buildId),
  };
  const engine = {
    ...common,
    entryPoints: [path.join(root, "src/docx/docxEngine.ts")],
    outfile: path.join(root, "dist/docx-engine.cjs"),
    external: [],
    define: baseDefines,
    banner: {
      js: `/* Office Reader DOCX engine ${pluginVersion} ${buildId} */`,
    },
  };
  const main = {
    ...common,
    entryPoints: [path.join(root, "src/main.ts")],
    outfile: path.join(root, "dist/main.js"),
    external: MAIN_EXTERNALS,
    define: {
      ...baseDefines,
      __DOCX_MANIFEST_MIN_APP_VERSION__: defineString(manifest.minAppVersion),
      __DOCX_MANIFEST_SHA256__: defineString(manifestSha256),
      __DOCX_MANIFEST_BYTES__: defineNumber(manifestBytes.byteLength),
      __DOCX_STYLES_SHA256__: defineString(stylesSha256),
      __DOCX_STYLES_BYTES__: defineNumber(stylesBytes.byteLength),
      __DOCX_ENGINE_SHA256__: defineString(engineSha256 ?? ""),
      __DOCX_ENGINE_BYTES__: defineNumber(engineBytes ?? 0),
    },
    banner: {
      js: `/* Office Reader */\n/* WORD_READER_BUILD_ID:${buildId} */`,
    },
  };
  return { engine, main };
}

function getOutputForEntry(metafile, entryPoint, root, label) {
  const expectedEntry = normalizePath(path.relative(root, entryPoint));
  const outputs = Object.entries(metafile.outputs ?? {});
  const matches = outputs.filter(([, output]) =>
    normalizePath(path.relative(root, path.resolve(root, output.entryPoint ?? ""))) === expectedEntry,
  );
  if (outputs.length !== 1 || matches.length !== 1) {
    throw new Error(`${label} build emitted an unexpected output/chunk inventory.`);
  }
  return matches[0];
}

function assertNoForbiddenRuntimeImports(metafile, outputKey, label) {
  const output = metafile.outputs[outputKey];
  const externals = (output.imports ?? []).filter((entry) => entry.external);
  if (externals.length > 0) {
    throw new Error(`${label} build has runtime externals: ${externals.map((entry) => entry.path).join(", ")}`);
  }
}

function normalizedInputSet(metafile, root) {
  return new Set(Object.keys(metafile.inputs ?? {}).map((input) => {
    const absolute = path.isAbsolute(input) ? input : path.resolve(root, input);
    return normalizePath(path.relative(root, absolute));
  }));
}

function assertOutputOwnership(mainMetafile, engineMetafile, root) {
  const mainInputs = normalizedInputSet(mainMetafile, root);
  const engineInputs = normalizedInputSet(engineMetafile, root);
  for (const forbidden of ENGINE_ONLY_INPUTS) {
    if (mainInputs.has(forbidden) || [...mainInputs].some((input) => input.startsWith(forbidden))) {
      throw new Error(`DOCX engine-owned input leaked into main.js: ${forbidden}`);
    }
  }
  for (const required of ENGINE_ONLY_INPUTS.slice(0, 2).concat([
    "src/docx/docxPreviewPackage.ts",
    "src/ooxml/storedZipWriter.ts",
    "src/ooxml/metafileSvg.ts",
    "src/ooxml/wmfScene.ts",
    "src/ooxml/emfScene.ts",
  ])) {
    if (![...engineInputs].some((input) => input === required || input.startsWith(required))) {
      throw new Error(`DOCX engine metafile is missing required closure input ${required}.`);
    }
  }
  return [...mainInputs].filter((input) => engineInputs.has(input)).sort();
}

function assertNoWriterRoutes(bytes, label) {
  const text = Buffer.from(bytes).toString("utf8");
  for (const pattern of WRITER_ROUTE_PATTERNS) {
    if (pattern.test(text)) {
      throw new Error(`A JSZip writer route is present in the ${label} runtime closure: ${pattern.source}`);
    }
  }
}

async function buildOne(config, entryPoint, root, label, captureHooks) {
  await captureHooks?.beforeBuild?.(label === "DOCX engine" ? "engine" : "main");
  const result = await build(config);
  if (result.outputFiles?.length !== 1) {
    throw new Error(`${label} build did not produce exactly one in-memory JavaScript output.`);
  }
  const [outputKey] = getOutputForEntry(result.metafile, entryPoint, root, label);
  return {
    bytes: Buffer.from(result.outputFiles[0].contents),
    metafile: result.metafile,
    outputKey,
    serializedOptions: serializeBuildOptions(config),
  };
}

function ensureMainBudget(bytes) {
  if (bytes.byteLength > DOCX_MAIN_LIMIT_BYTES) {
    throw new Error(`dist/main.js is ${bytes.byteLength} bytes; budget is ${DOCX_MAIN_LIMIT_BYTES} bytes.`);
  }
}

async function publishDistSnapshot(root, snapshot) {
  const distDir = path.join(root, "dist");
  const stagingDir = path.join(root, `.dist-docx-stage-${process.pid}-${Date.now()}`);
  const backupDir = path.join(root, `.dist-docx-backup-${process.pid}-${Date.now()}`);
  let movedExisting = false;
  try {
    await fs.mkdir(stagingDir);
    for (const file of DOCX_DISTRIBUTION_FILES) {
      await fs.writeFile(path.join(stagingDir, file), snapshot.files[file]);
    }
    const stagedNames = (await fs.readdir(stagingDir)).sort();
    if (JSON.stringify(stagedNames) !== JSON.stringify([...DOCX_DISTRIBUTION_FILES].sort())) {
      throw new Error("Staged DOCX distribution has an unexpected file inventory.");
    }
    const existing = await fs.lstat(distDir).catch((error) => {
      if (error && error.code === "ENOENT") return null;
      throw error;
    });
    if (existing) {
      if (existing.isSymbolicLink() || !existing.isDirectory()) {
        throw new Error("dist must be a real directory before publishing the build snapshot.");
      }
      const names = (await fs.readdir(distDir)).sort();
      const allowedOldFiles = new Set([...DOCX_DISTRIBUTION_FILES, "main.js.LEGAL.txt"]);
      const unexpected = names.filter((name) => !allowedOldFiles.has(name));
      if (unexpected.length > 0) {
        throw new Error(`dist contains unowned files; refusing to replace: ${unexpected.join(", ")}`);
      }
      await fs.rename(distDir, backupDir);
      movedExisting = true;
    }
    await fs.rename(stagingDir, distDir);
    if (movedExisting) {
      await fs.rm(backupDir, { recursive: true, force: true });
    }
  } catch (error) {
    if (movedExisting) {
      const distExists = await fs.lstat(distDir).then(() => true, () => false);
      if (!distExists) {
        await fs.rename(backupDir, distDir).catch(() => {});
      }
    }
    throw error;
  } finally {
    await fs.rm(stagingDir, { recursive: true, force: true });
  }
}

export async function buildDocxDistribution(options = {}) {
  const root = path.resolve(options.root ?? process.cwd());
  const production = options.production ?? true;
  const manifestBytes = await fs.readFile(path.join(root, "manifest.json"));
  const stylesBytes = await fs.readFile(path.join(root, "styles.css"));
  const packageJson = JSON.parse(await fs.readFile(path.join(root, "package.json"), "utf8"));
  const manifest = JSON.parse(manifestBytes.toString("utf8"));
  if (packageJson.version !== manifest.version) {
    throw new Error("package.json and manifest.json versions do not match.");
  }
  const buildInputs = await calculateBuildId(root, production);
  const commonOptions = {
    root,
    production,
    buildId: buildInputs.buildId,
    pluginVersion: packageJson.version,
    manifest,
    manifestBytes,
    stylesBytes,
    manifestSha256: sha256(manifestBytes),
    stylesSha256: sha256(stylesBytes),
  };
  const firstConfig = createDocxBuildConfig({
    ...commonOptions,
    captureHooks: options.captureHooks,
    captureEntry: "engine",
  });
  const engineBuild = await buildOne(
    firstConfig.engine,
    path.join(root, "src/docx/docxEngine.ts"),
    root,
    "DOCX engine",
    options.captureHooks,
  );
  assertNoForbiddenRuntimeImports(engineBuild.metafile, engineBuild.outputKey, "DOCX engine");
  assertNoWriterRoutes(engineBuild.bytes, "DOCX engine");
  const engineSha256 = sha256(engineBuild.bytes);
  const config = createDocxBuildConfig({
    ...commonOptions,
    engineSha256,
    engineBytes: engineBuild.bytes.byteLength,
    captureHooks: options.captureHooks,
    captureEntry: "main",
  });
  const mainBuild = await buildOne(
    config.main,
    path.join(root, "src/main.ts"),
    root,
    "main",
    options.captureHooks,
  );
  if (typeof options.captureHooks?.beforeOwnershipGuard === "function") {
    const postBuildInputs = await calculateBuildId(root, production);
    await options.captureHooks.beforeOwnershipGuard({
      root,
      production,
      buildId: buildInputs.buildId,
      buildIdInputs: buildInputs.inputs,
      buildIdConfiguration: buildInputs.configuration,
      postBuildId: postBuildInputs.buildId,
      postBuildIdInputs: postBuildInputs.inputs,
      version: packageJson.version,
      manifestBytes: Buffer.from(manifestBytes),
      stylesBytes: Buffer.from(stylesBytes),
      engine: engineBuild,
      main: mainBuild,
    });
  }
  const sharedInputs = assertOutputOwnership(
    mainBuild.metafile,
    engineBuild.metafile,
    root,
  );
  assertNoWriterRoutes(mainBuild.bytes, "main");
  ensureMainBudget(mainBuild.bytes);
  const mainSha256 = sha256(mainBuild.bytes);
  const fixedPayloadBytes =
    mainBuild.bytes.byteLength +
    engineBuild.bytes.byteLength +
    manifestBytes.byteLength +
    stylesBytes.byteLength;
  const engineLimitBytes =
    DOCX_ARCHIVE_LIMIT_BYTES -
    mainBuild.bytes.byteLength -
    manifestBytes.byteLength -
    stylesBytes.byteLength -
    DOCX_ARCHIVE_OVERHEAD_BYTES;
  if (!Number.isSafeInteger(engineLimitBytes) || engineLimitBytes <= 0) {
    throw new Error("The measured main/assets leave no safe DOCX engine package allowance.");
  }
  const snapshot = {
    buildId: buildInputs.buildId,
    version: packageJson.version,
    files: Object.freeze({
      "main.js": mainBuild.bytes,
      "docx-engine.cjs": engineBuild.bytes,
      "manifest.json": Buffer.from(manifestBytes),
      "styles.css": Buffer.from(stylesBytes),
    }),
    bytes: {
      main: mainBuild.bytes.byteLength,
      engine: engineBuild.bytes.byteLength,
      manifest: manifestBytes.byteLength,
      styles: stylesBytes.byteLength,
      payload: fixedPayloadBytes,
      engineLimit: engineLimitBytes,
    },
    hashes: {
      main: mainSha256,
      engine: engineSha256,
      manifest: sha256(manifestBytes),
      styles: sha256(stylesBytes),
      buildId: buildInputs.buildId,
    },
    metafiles: {
      main: mainBuild.metafile,
      engine: engineBuild.metafile,
    },
    sharedInputs,
    buildInputs: buildInputs.inputs,
  };
  if (snapshot.bytes.engine > engineLimitBytes) {
    throw new Error(`Measured DOCX engine is ${snapshot.bytes.engine} bytes; package allowance is ${engineLimitBytes} bytes.`);
  }
  if (options.write === true) {
    await publishDistSnapshot(root, snapshot);
  }
  return snapshot;
}
