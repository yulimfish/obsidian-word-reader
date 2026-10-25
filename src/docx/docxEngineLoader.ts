import type {
  DocxEngine,
  DocxEngineHost,
  DocxRenderResult,
} from "./docxEngineContract";
import type { PreviewInventory, PreviewIssue } from "../ooxml/previewTypes";
import { loadOoxmlZip } from "../ooxml/jszipReader";
import { enforceOoxmlPackagePolicy } from "../ooxml/packagePolicy";
import {
  DEFAULT_DOCX_ZIP_LIMITS,
  validateZipSafety,
} from "../ooxml/packageSafety";

const ENGINE_FILE = "docx-engine.cjs";
const MANIFEST_FILE = "manifest.json";
const STYLES_FILE = "styles.css";
const MAIN_FILE = "main.js";
const ENGINE_EXPORTS = ["abiVersion", "buildId", "pluginVersion", "renderDocx"];

export interface DocxEnginePluginContext {
  readonly app: {
    readonly vault: {
      readonly adapter: unknown;
    };
  };
  readonly manifest: {
    readonly id: string;
    readonly version: string;
    readonly dir?: string;
  };
}

export interface DocxEngineAssetExpectation {
  readonly bytes: number;
  readonly sha256: string;
}

export interface DocxEngineBuildExpectation {
  readonly pluginId: string;
  readonly pluginVersion: string;
  readonly buildId: string;
  readonly abiVersion: 1;
  readonly minAppVersion: string;
  readonly mainMaxBytes: number;
  readonly assets: Readonly<Record<
    typeof MANIFEST_FILE | typeof STYLES_FILE | typeof ENGINE_FILE,
    DocxEngineAssetExpectation
  >>;
}

interface RuntimeStats {
  readonly size: number;
  isFile(): boolean;
  isDirectory(): boolean;
  isSymbolicLink(): boolean;
}

export interface DocxEngineNodeRuntime {
  readonly fs: {
    lstat(path: string): Promise<RuntimeStats>;
    stat(path: string): Promise<RuntimeStats>;
    realpath(path: string): Promise<string>;
    readFile(path: string): Promise<Uint8Array>;
  };
  readonly path: {
    resolve(...paths: string[]): string;
    join(...paths: string[]): string;
    relative(from: string, to: string): string;
    isAbsolute(path: string): boolean;
    basename(path: string): string;
    readonly sep: string;
  };
  readonly crypto: {
    createHash(algorithm: string): {
      update(data: Uint8Array): unknown;
      digest(encoding: "hex"): string;
    };
  };
  moduleRequire(canonicalVerifiedChild: string): unknown;
}

interface FileSystemAdapterLike {
  getFullPath(vaultPath: string): string;
}

interface FileSystemAdapterConstructor {
  new (...args: never[]): FileSystemAdapterLike;
}

export interface DocxEngineLoaderDependencies {
  readonly isDesktopApp: () => boolean;
  readonly FileSystemAdapter: FileSystemAdapterConstructor;
  readonly getRuntime: () => DocxEngineNodeRuntime;
}

export interface DocxEngineLoader {
  load(buffer?: ArrayBuffer): Promise<DocxEngine>;
  renderDocx(
    buffer: ArrayBuffer,
    targetEl: HTMLElement,
    host: DocxEngineHost,
  ): Promise<DocxRenderResult>;
  dispose(): void;
}

export class DocxEngineLoadError extends Error {
  readonly code: string;
  readonly issues: readonly PreviewIssue[];

  constructor(code: string, detail: string, issues: readonly PreviewIssue[] = []) {
    super(detail);
    this.name = "DocxEngineLoadError";
    this.code = code;
    this.issues = issues;
  }

  withIssues(issues: readonly PreviewIssue[]): DocxEngineLoadError {
    return new DocxEngineLoadError(this.code, this.message, issues);
  }
}

interface VerifiedFile {
  readonly path: string;
  readonly bytes: Uint8Array;
  readonly sha256: string | null;
}

interface ModuleIdentity {
  readonly directory: string;
  readonly pluginVersion: string;
  readonly buildId: string;
  readonly manifestSha256: string;
  readonly stylesSha256: string;
  readonly engineSha256: string;
}

interface ActiveRender {
  readonly generation: number;
  readonly host: DocxEngineHost;
  cancelled: boolean;
}

function abortError(): Error {
  const error = new Error("DOCX rendering was cancelled.");
  error.name = "AbortError";
  return error;
}

function isContainedPath(
  pathModule: DocxEngineNodeRuntime["path"],
  directory: string,
  candidate: string,
  expectedName: string,
): boolean {
  const relative = pathModule.relative(directory, candidate);
  return pathModule.basename(candidate) === expectedName &&
    relative === expectedName &&
    !relative.startsWith(`..${pathModule.sep}`) &&
    relative !== ".." &&
    !pathModule.isAbsolute(relative);
}

function hashBytes(
  runtime: DocxEngineNodeRuntime,
  bytes: Uint8Array,
): string {
  const hash = runtime.crypto.createHash("sha256");
  hash.update(bytes);
  return hash.digest("hex");
}

function getOwnDataValue(value: unknown, key: string): unknown {
  if ((typeof value !== "object" && typeof value !== "function") || value === null) {
    return undefined;
  }
  const descriptor = Object.getOwnPropertyDescriptor(value, key);
  return descriptor && "value" in descriptor ? descriptor.value : undefined;
}

function assertExpectedAssets(expected: DocxEngineBuildExpectation): void {
  for (const file of [MANIFEST_FILE, STYLES_FILE, ENGINE_FILE] as const) {
    const asset = expected.assets[file];
    if (
      !asset ||
      !Number.isSafeInteger(asset.bytes) ||
      asset.bytes < 0 ||
      !/^[a-f\d]{64}$/i.test(asset.sha256)
    ) {
      throw new DocxEngineLoadError(
        "asset-mismatch",
        `The build snapshot has invalid expectations for ${file}.`,
      );
    }
  }
}

async function createUnavailableIssues(
  buffer: ArrayBuffer,
  reason: string,
): Promise<readonly PreviewIssue[]> {
  let inventory: PreviewInventory | null = null;
  try {
    validateZipSafety(buffer, DEFAULT_DOCX_ZIP_LIMITS);
    const zip = await loadOoxmlZip(buffer);
    inventory = await enforceOoxmlPackagePolicy(zip, "docx", {
      legacyEquationPreviews: true,
    });
  } catch {
    inventory = null;
  }
  if (inventory && inventory.occurrences.length > 0) {
    return inventory.occurrences.map((occurrence) => ({
      partPath: occurrence.partPath,
      occurrenceId: occurrence.occurrenceId,
      code: "docx-engine-unavailable",
      detail: reason,
    }));
  }
  return [{
    partPath: "word/document.xml",
    occurrenceId: "document",
    code: "docx-engine-unavailable",
    detail: reason,
  }];
}

export function createDocxEngineLoader(
  plugin: DocxEnginePluginContext,
  expected: DocxEngineBuildExpectation,
  dependencies: DocxEngineLoaderDependencies,
): DocxEngineLoader {
  let generation = 0;
  let disposed = false;
  let executingRequire = false;
  let inFlightLoad: Promise<DocxEngine> | null = null;
  let cachedEngine: { readonly identity: ModuleIdentity; readonly engine: DocxEngine } | null = null;
  const activeRenders = new Set<ActiveRender>();

  const fail = (code: string, detail: string): never => {
    throw new DocxEngineLoadError(code, detail);
  };

  const verifyFile = async (
    runtime: DocxEngineNodeRuntime,
    directory: string,
    fileName: typeof MANIFEST_FILE | typeof STYLES_FILE | typeof ENGINE_FILE | typeof MAIN_FILE,
    expectedAsset?: DocxEngineAssetExpectation,
  ): Promise<VerifiedFile> => {
    const candidate = runtime.path.join(directory, fileName);
    if (!isContainedPath(runtime.path, directory, candidate, fileName)) {
      fail("invalid-path", `The packaged ${fileName} path is not canonical.`);
    }
    let stats: RuntimeStats;
    try {
      stats = await runtime.fs.lstat(candidate);
    } catch {
      return fail("asset-mismatch", `The packaged ${fileName} file is missing.`);
    }
    if (stats.isSymbolicLink() || !stats.isFile()) {
      fail("invalid-path", `The packaged ${fileName} must be a regular non-symlink file.`);
    }
    const maximumBytes = expectedAsset?.bytes ?? expected.mainMaxBytes;
    if (!Number.isSafeInteger(stats.size) || stats.size < 0 || stats.size > maximumBytes) {
      fail("asset-mismatch", `The packaged ${fileName} exceeds its baked size bound.`);
    }
    if (expectedAsset && stats.size !== expectedAsset.bytes) {
      fail("asset-mismatch", `The packaged ${fileName} size does not match this build.`);
    }
    let canonicalFile: string;
    try {
      canonicalFile = await runtime.fs.realpath(candidate);
    } catch {
      return fail("invalid-path", `The packaged ${fileName} could not be canonicalized.`);
    }
    if (
      canonicalFile !== candidate ||
      !isContainedPath(runtime.path, directory, canonicalFile, fileName)
    ) {
      fail("invalid-path", `The packaged ${fileName} resolves outside its canonical sibling path.`);
    }
    let bytes: Uint8Array;
    try {
      bytes = await runtime.fs.readFile(canonicalFile);
    } catch {
      return fail("asset-mismatch", `The packaged ${fileName} could not be read.`);
    }
    if (bytes.byteLength !== stats.size) {
      fail("asset-mismatch", `The packaged ${fileName} changed while it was read.`);
    }
    const digest = expectedAsset ? hashBytes(runtime, bytes) : null;
    if (expectedAsset && digest !== expectedAsset.sha256) {
      fail("asset-mismatch", `The packaged ${fileName} hash does not match this build.`);
    }
    return { path: canonicalFile, bytes, sha256: digest };
  };

  const getPluginDirectory = (): string => {
    const pluginDir = plugin.manifest.dir;
    if (typeof pluginDir !== "string" || pluginDir.trim().length === 0) {
      return fail("invalid-path", "The plugin manifest has no install directory.");
    }
    const adapter = plugin.app.vault.adapter;
    if (!(adapter instanceof dependencies.FileSystemAdapter)) {
      return fail("unsupported-adapter", "DOCX engine loading requires Obsidian's FileSystemAdapter.");
    }
    let resolved: string;
    try {
      resolved = adapter.getFullPath(pluginDir);
    } catch {
      return fail("invalid-path", "The plugin directory could not be resolved by the vault adapter.");
    }
    if (typeof resolved !== "string" || resolved.length === 0) {
      fail("invalid-path", "The vault adapter returned an invalid plugin directory.");
    }
    return resolved;
  };

  const getCanonicalDirectory = async (
    runtime: DocxEngineNodeRuntime,
    pluginDirectory: string,
  ): Promise<string> => {
    const resolved = runtime.path.resolve(pluginDirectory);
    let canonicalDirectory: string;
    try {
      canonicalDirectory = await runtime.fs.realpath(resolved);
      const stats = await runtime.fs.stat(canonicalDirectory);
      if (!stats.isDirectory()) {
        return fail("invalid-path", "The canonical plugin directory is not a directory.");
      }
    } catch (error) {
      if (error instanceof DocxEngineLoadError) {
        throw error;
      }
      return fail("invalid-path", "The plugin directory could not be canonicalized.");
    }
    return canonicalDirectory;
  };

  const verifySnapshot = async (
    runtime: DocxEngineNodeRuntime,
    directory: string,
  ): Promise<{
    readonly identity: ModuleIdentity;
    readonly engineFile: VerifiedFile;
  }> => {
    const [main, manifest, styles, engineFile] = await Promise.all([
      verifyFile(runtime, directory, MAIN_FILE),
      verifyFile(runtime, directory, MANIFEST_FILE, expected.assets[MANIFEST_FILE]),
      verifyFile(runtime, directory, STYLES_FILE, expected.assets[STYLES_FILE]),
      verifyFile(runtime, directory, ENGINE_FILE, expected.assets[ENGINE_FILE]),
    ]);
    const mainText = new TextDecoder().decode(main.bytes);
    if (!mainText.includes(`WORD_READER_BUILD_ID:${expected.buildId}`)) {
      fail("build-mismatch", "The on-disk main.js does not match the executing build marker.");
    }
    let manifestData: unknown;
    try {
      manifestData = JSON.parse(new TextDecoder().decode(manifest.bytes));
    } catch {
      fail("asset-mismatch", "The packaged manifest.json is malformed.");
    }
    if (
      typeof manifestData !== "object" ||
      manifestData === null ||
      (manifestData as Record<string, unknown>).id !== expected.pluginId ||
      (manifestData as Record<string, unknown>).version !== expected.pluginVersion ||
      (manifestData as Record<string, unknown>).isDesktopOnly !== true ||
      (manifestData as Record<string, unknown>).minAppVersion !== expected.minAppVersion
    ) {
      fail("manifest-mismatch", "The packaged manifest.json does not match this desktop build.");
    }
    const identity: ModuleIdentity = {
      directory,
      pluginVersion: expected.pluginVersion,
      buildId: expected.buildId,
      manifestSha256: manifest.sha256 ?? "",
      stylesSha256: styles.sha256 ?? "",
      engineSha256: engineFile.sha256 ?? "",
    };
    return { identity, engineFile };
  };

  const loadEngine = async (): Promise<DocxEngine> => {
    if (disposed) {
      fail("disposed", "The DOCX engine loader has been disposed.");
    }
    if (!dependencies.isDesktopApp()) {
      fail("desktop-only", "DOCX previews are available only in the desktop application.");
    }
    if (
      plugin.manifest.id !== expected.pluginId ||
      plugin.manifest.version !== expected.pluginVersion
    ) {
      fail("manifest-mismatch", "The running plugin manifest does not match the baked build.");
    }
    const pluginDirectory = getPluginDirectory();
    let runtime: DocxEngineNodeRuntime;
    try {
      runtime = dependencies.getRuntime();
    } catch {
      return fail("desktop-only", "The desktop DOCX runtime is unavailable.");
    }
    assertExpectedAssets(expected);
    const loadGeneration = generation;
    const directory = await getCanonicalDirectory(runtime, pluginDirectory);
    const snapshot = await verifySnapshot(runtime, directory);
    if (disposed || loadGeneration !== generation) {
      fail("disposed", "The DOCX engine loader was invalidated while loading.");
    }
    if (cachedEngine) {
      const cached = cachedEngine.identity;
      const actual = snapshot.identity;
      if (
        cached.directory !== actual.directory ||
        cached.pluginVersion !== actual.pluginVersion ||
        cached.buildId !== actual.buildId ||
        cached.manifestSha256 !== actual.manifestSha256 ||
        cached.stylesSha256 !== actual.stylesSha256 ||
        cached.engineSha256 !== actual.engineSha256
      ) {
        fail("stale-cache", "A cached DOCX engine identity differs from the verified package.");
      }
      return cachedEngine.engine;
    }

    let rawEngine: unknown;
    try {
      executingRequire = true;
      rawEngine = runtime.moduleRequire(snapshot.engineFile.path);
    } catch {
      return fail("module-load-failed", "The verified local DOCX engine could not be loaded.");
    } finally {
      executingRequire = false;
    }
    const keys = typeof rawEngine === "object" && rawEngine !== null
      ? Object.keys(rawEngine).sort()
      : [];
    if (
      keys.length !== ENGINE_EXPORTS.length ||
      keys.some((key, index) => key !== ENGINE_EXPORTS[index]) ||
      getOwnDataValue(rawEngine, "abiVersion") !== expected.abiVersion ||
      getOwnDataValue(rawEngine, "pluginVersion") !== expected.pluginVersion ||
      getOwnDataValue(rawEngine, "buildId") !== expected.buildId ||
      typeof getOwnDataValue(rawEngine, "renderDocx") !== "function"
    ) {
      fail("engine-mismatch", "The loaded DOCX engine exports do not match the expected ABI, version, and build.");
    }
    const render = getOwnDataValue(rawEngine, "renderDocx") as DocxEngine["renderDocx"];
    const engine: DocxEngine = Object.freeze({
      abiVersion: expected.abiVersion,
      pluginVersion: expected.pluginVersion,
      buildId: expected.buildId,
      renderDocx: async (
        buffer: ArrayBuffer,
        targetEl: HTMLElement,
        host: DocxEngineHost,
      ): Promise<DocxRenderResult> => {
        const renderGeneration = generation;
        const active: ActiveRender = {
          generation: renderGeneration,
          host,
          cancelled: false,
        };
        activeRenders.add(active);
        const guardedHost: DocxEngineHost = {
          ownerDocument: host.ownerDocument,
          isCancelled: () =>
            active.cancelled ||
            disposed ||
            renderGeneration !== generation ||
            host.isCancelled(),
          yieldControl: host.yieldControl,
        };
        try {
          if (guardedHost.isCancelled()) {
            throw abortError();
          }
          const result = await render(buffer, targetEl, guardedHost);
          if (guardedHost.isCancelled()) {
            throw abortError();
          }
          return result;
        } finally {
          activeRenders.delete(active);
        }
      },
    });
    const postRequireSnapshot = await verifySnapshot(runtime, directory);
    if (
      disposed ||
      loadGeneration !== generation ||
      postRequireSnapshot.identity.engineSha256 !== snapshot.identity.engineSha256 ||
      postRequireSnapshot.identity.manifestSha256 !== snapshot.identity.manifestSha256 ||
      postRequireSnapshot.identity.stylesSha256 !== snapshot.identity.stylesSha256
    ) {
      fail("asset-mismatch", "The DOCX package changed while its engine was being loaded.");
    }
    cachedEngine = { identity: snapshot.identity, engine };
    return engine;
  };

  const withFailureIssues = async (
    promise: Promise<DocxEngine>,
    buffer: ArrayBuffer | undefined,
  ): Promise<DocxEngine> => {
    try {
      return await promise;
    } catch (error) {
      if (!buffer) {
        throw error;
      }
      const loadError = error instanceof DocxEngineLoadError
        ? error
        : new DocxEngineLoadError("load-failed", error instanceof Error ? error.message : String(error));
      const issues = await createUnavailableIssues(buffer, loadError.message);
      throw loadError.withIssues(issues);
    }
  };

  const loader: DocxEngineLoader = {
    load(buffer?: ArrayBuffer): Promise<DocxEngine> {
      if (disposed) {
        return Promise.reject(new DocxEngineLoadError("disposed", "The DOCX engine loader has been disposed."));
      }
      if (executingRequire) {
        return Promise.reject(new DocxEngineLoadError("recursive-load", "Recursive DOCX engine loading is not allowed."));
      }
      if (!inFlightLoad) {
        const current = loadEngine();
        inFlightLoad = current;
        void current.then(
          () => {
            if (inFlightLoad === current) {
              inFlightLoad = null;
            }
          },
          () => {
            if (inFlightLoad === current) {
              inFlightLoad = null;
            }
          },
        );
      }
      return withFailureIssues(inFlightLoad, buffer);
    },

    async renderDocx(
      buffer: ArrayBuffer,
      targetEl: HTMLElement,
      host: DocxEngineHost,
    ): Promise<DocxRenderResult> {
      if (disposed) {
        throw new DocxEngineLoadError("disposed", "The DOCX engine loader has been disposed.");
      }
      const renderGeneration = generation;
      const active: ActiveRender = {
        generation: renderGeneration,
        host,
        cancelled: false,
      };
      activeRenders.add(active);
      const isCancelled = (): boolean =>
        active.cancelled ||
        disposed ||
        renderGeneration !== generation ||
        host.isCancelled();
      try {
        if (isCancelled()) {
          throw abortError();
        }
        const engine = await loader.load(buffer);
        if (isCancelled()) {
          throw abortError();
        }
        return await engine.renderDocx(buffer, targetEl, {
          ownerDocument: host.ownerDocument,
          isCancelled,
          yieldControl: host.yieldControl,
        });
      } finally {
        activeRenders.delete(active);
      }
    },

    dispose(): void {
      if (disposed) {
        return;
      }
      disposed = true;
      generation += 1;
      cachedEngine = null;
      inFlightLoad = null;
      for (const active of activeRenders) {
        active.cancelled = true;
      }
      activeRenders.clear();
    },
  };

  return loader;
}
