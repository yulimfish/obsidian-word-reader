import assert from "node:assert/strict";
import crypto from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";
import test from "node:test";
import { createJiti } from "jiti";
import { createEquationFixture } from "./fixtureBuilders";

const jiti = createJiti(__filename);
const BUILD_ID = "build-abc123";
const PLUGIN_VERSION = "3.5.0";
const PLUGIN_ID = "word-reader";
const MAIN_MARKER = `/* WORD_READER_BUILD_ID:${BUILD_ID} */`;
const MANIFEST_PATH = "manifest.json";
const STYLES_PATH = "styles.css";
const ENGINE_PATH = "docx-engine.cjs";
const MAIN_PATH = "main.js";

interface LoaderModule {
  readonly createDocxEngineLoader: (...args: unknown[]) => Loader;
  readonly DocxEngineLoadError: new (...args: never[]) => Error & {
    readonly code: string;
  };
}

interface Loader {
  load(buffer?: ArrayBuffer): Promise<{
    readonly abiVersion: number;
    readonly pluginVersion: string;
    readonly buildId: string;
    readonly renderDocx: (...args: unknown[]) => Promise<unknown>;
  }>;
  renderDocx(
    buffer: ArrayBuffer,
    targetEl: HTMLElement,
    host: {
      readonly ownerDocument: Document;
      readonly isCancelled: () => boolean;
      readonly yieldControl: () => Promise<void>;
    },
  ): Promise<unknown>;
  dispose(): void;
}

interface Fixture {
  readonly directory: string;
  readonly module: LoaderModule;
  readonly loader: Loader;
  readonly counters: {
    runtimeCalls: number;
    requireCalls: string[];
    engineCalls: number;
  };
  readonly files: {
    readonly main: Buffer;
    readonly manifest: Buffer;
    readonly styles: Buffer;
    readonly engine: Buffer;
  };
  cleanup: () => Promise<void>;
}

function sha256(bytes: Uint8Array): string {
  return crypto.createHash("sha256").update(bytes).digest("hex");
}

function createEngine(
  counters: Fixture["counters"],
  overrides: Record<string, unknown> = {},
): Record<string, unknown> {
  return {
    abiVersion: 1,
    pluginVersion: PLUGIN_VERSION,
    buildId: BUILD_ID,
    renderDocx: async () => {
      counters.engineCalls += 1;
      return { inventory: { occurrences: [], issues: [] }, issues: [] };
    },
    ...overrides,
  };
}

async function createFixture(options: {
  readonly isDesktopApp?: boolean;
  readonly adapter?: unknown;
  readonly moduleExports?: Record<string, unknown>;
  readonly beforeRead?: () => Promise<void>;
  readonly moduleRequire?: (filename: string) => unknown;
  readonly mainBytes?: Uint8Array;
  readonly manifestBytes?: Uint8Array;
  readonly manifestData?: Record<string, unknown>;
  readonly pluginManifest?: { id: string; version: string; dir?: string };
} = {}): Promise<Fixture> {
  // eslint-disable-next-line @typescript-eslint/no-unnecessary-type-assertion -- dynamic import is typed as unknown, assertion narrows to LoaderModule
  const module = await jiti.import(
    path.resolve(process.cwd(), "src/docx/docxEngineLoader.ts"),
  ) as unknown as LoaderModule;
  const parentDir = path.join(process.cwd(), ".test-dist");
  const directory = await fs.mkdtemp(path.join(parentDir, "docx-loader-"));
  const main = Buffer.from(options.mainBytes ?? `${MAIN_MARKER}\n/* Office Reader */\n`);
  const manifest = Buffer.from(options.manifestBytes ?? JSON.stringify(options.manifestData ?? {
    id: PLUGIN_ID,
    version: PLUGIN_VERSION,
    isDesktopOnly: true,
    minAppVersion: "1.12.7",
  }));
  const styles = Buffer.from(".docx { display: block; }\n");
  const engine = Buffer.from("verified engine payload");
  const files = { main, manifest, styles, engine };
  await Promise.all([
    fs.writeFile(path.join(directory, MAIN_PATH), main),
    fs.writeFile(path.join(directory, MANIFEST_PATH), manifest),
    fs.writeFile(path.join(directory, STYLES_PATH), styles),
    fs.writeFile(path.join(directory, ENGINE_PATH), engine),
  ]);

  class FixtureFileSystemAdapter {
    getFullPath(_vaultPath: string): string {
      return directory;
    }
  }
  const adapter = options.adapter ?? new FixtureFileSystemAdapter();
  const counters = { runtimeCalls: 0, requireCalls: [] as string[], engineCalls: 0 };
  const readFile = async (filePath: string): Promise<Buffer> => {
    await options.beforeRead?.();
    return fs.readFile(filePath);
  };
  const expected = {
    pluginId: PLUGIN_ID,
    pluginVersion: PLUGIN_VERSION,
    buildId: BUILD_ID,
    abiVersion: 1,
    minAppVersion: "1.12.7",
    mainMaxBytes: 512_000,
    assets: {
      [MANIFEST_PATH]: { bytes: manifest.byteLength, sha256: sha256(manifest) },
      [STYLES_PATH]: { bytes: styles.byteLength, sha256: sha256(styles) },
      [ENGINE_PATH]: { bytes: engine.byteLength, sha256: sha256(engine) },
    },
  };
  const loader = module.createDocxEngineLoader(
    {
      app: { vault: { adapter } },
      manifest: options.pluginManifest ?? {
        id: PLUGIN_ID,
        version: PLUGIN_VERSION,
        dir: "plugins/word-reader",
      },
    },
    expected,
    {
      isDesktopApp: () => options.isDesktopApp ?? true,
      FileSystemAdapter: FixtureFileSystemAdapter,
      getRuntime: () => {
        counters.runtimeCalls += 1;
        return {
          fs: { lstat: fs.lstat, stat: fs.stat, realpath: fs.realpath, readFile },
          path,
          crypto,
          moduleRequire: (filename: string) => {
            counters.requireCalls.push(filename);
            if (options.moduleRequire) {
              return options.moduleRequire(filename);
            }
            return options.moduleExports ?? createEngine(counters);
          },
        };
      },
    },
  );

  return {
    directory,
    module,
    loader,
    counters,
    files,
    cleanup: () => fs.rm(directory, { recursive: true, force: true }),
  };
}

async function expectLoaderCode(
  operation: () => Promise<unknown>,
  code: string,
): Promise<void> {
  await assert.rejects(operation, (error: unknown) => {
    assert.equal(typeof error, "object");
    assert.equal((error as { code?: unknown }).code, code);
    return true;
  });
}

void test("DOCX engine loader validates and freezes the exact four-export ABI", async (t) => {
  const fixture = await createFixture();
  t.after(fixture.cleanup);

  const engine = await fixture.loader.load();

  assert.deepEqual(
    Object.keys(engine).sort(),
    ["abiVersion", "buildId", "pluginVersion", "renderDocx"],
  );
  assert.equal(engine.abiVersion, 1);
  assert.equal(engine.pluginVersion, PLUGIN_VERSION);
  assert.equal(engine.buildId, BUILD_ID);
  assert.equal(Object.isFrozen(engine), true);
  assert.deepEqual(fixture.counters.requireCalls, [
    path.join(fixture.directory, ENGINE_PATH),
  ]);
});

void test("mobile and non-filesystem adapters reject before Node runtime access", async (t) => {
  const mobile = await createFixture({ isDesktopApp: false });
  t.after(mobile.cleanup);
  await expectLoaderCode(() => mobile.loader.load(), "desktop-only");
  assert.equal(mobile.counters.runtimeCalls, 0);
  assert.equal(mobile.counters.requireCalls.length, 0);

  const wrongAdapter = await createFixture({ adapter: { getFullPath: () => "/unused" } });
  t.after(wrongAdapter.cleanup);
  await expectLoaderCode(() => wrongAdapter.loader.load(), "unsupported-adapter");
  assert.equal(wrongAdapter.counters.runtimeCalls, 0);
  assert.equal(wrongAdapter.counters.requireCalls.length, 0);

  const wrongPluginManifest = await createFixture({
    pluginManifest: { id: PLUGIN_ID, version: "3.4.9", dir: "plugins/word-reader" },
  });
  t.after(wrongPluginManifest.cleanup);
  await expectLoaderCode(() => wrongPluginManifest.loader.load(), "manifest-mismatch");
  assert.equal(wrongPluginManifest.counters.runtimeCalls, 0);
});

void test("loader resolves only canonical contained regular package siblings", async (t) => {
  const fixture = await createFixture();
  t.after(fixture.cleanup);
  const outside = path.join(fixture.directory, "..", "outside-engine.cjs");
  await fs.writeFile(outside, fixture.files.engine);
  await fs.rm(path.join(fixture.directory, ENGINE_PATH));
  await fs.symlink(outside, path.join(fixture.directory, ENGINE_PATH));

  await expectLoaderCode(() => fixture.loader.load(), "invalid-path");
  assert.equal(fixture.counters.requireCalls.length, 0);
  await fs.rm(outside, { force: true });
});

void test("missing, non-file, oversized, corrupt, and hash-mismatched assets fail before require", async (t) => {
  for (const failure of ["missing", "non-file", "oversized", "corrupt"] as const) {
    const fixture = await createFixture();
    t.after(fixture.cleanup);
    const enginePath = path.join(fixture.directory, ENGINE_PATH);
    if (failure === "missing") {
      await fs.rm(enginePath);
    } else if (failure === "non-file") {
      await fs.rm(enginePath);
      await fs.mkdir(enginePath);
    } else if (failure === "oversized") {
      await fs.appendFile(enginePath, Buffer.alloc(64));
    } else {
      await fs.writeFile(enginePath, Buffer.from("tampered engine payload"));
    }

    await expectLoaderCode(
      () => fixture.loader.load(),
      failure === "non-file" ? "invalid-path" : "asset-mismatch",
    );
    assert.equal(fixture.counters.requireCalls.length, 0, failure);
  }
});

void test("manifest identity, version, desktop flag, minimum version, and main build marker bind before require", async (t) => {
  const manifestBase = {
    id: PLUGIN_ID,
    version: PLUGIN_VERSION,
    isDesktopOnly: true,
    minAppVersion: "1.12.7",
  };
  for (const [label, patch] of [
    ["id", { id: "other-plugin" }],
    ["version", { version: "3.4.9" }],
    ["desktop-only", { isDesktopOnly: false }],
    ["minimum-version", { minAppVersion: "1.11.0" }],
  ] as const) {
    const fixture = await createFixture();
    t.after(fixture.cleanup);
    await fs.writeFile(
      path.join(fixture.directory, MANIFEST_PATH),
      JSON.stringify({ ...manifestBase, ...patch }),
    );
    await expectLoaderCode(() => fixture.loader.load(), "asset-mismatch");
    assert.equal(fixture.counters.requireCalls.length, 0, label);
  }

  for (const [label, patch] of [
    ["id", { id: "other-plugin" }],
    ["version", { version: "3.4.9" }],
    ["desktop-only", { isDesktopOnly: false }],
    ["minimum-version", { minAppVersion: "1.11.0" }],
  ] as const) {
    const fixture = await createFixture({ manifestData: { ...manifestBase, ...patch } });
    t.after(fixture.cleanup);
    await expectLoaderCode(() => fixture.loader.load(), "manifest-mismatch");
    assert.equal(fixture.counters.requireCalls.length, 0, `baked ${label}`);
  }

  const mismatchedMain = await createFixture();
  t.after(mismatchedMain.cleanup);
  await fs.writeFile(path.join(mismatchedMain.directory, MAIN_PATH), "/* wrong build */");
  await expectLoaderCode(() => mismatchedMain.loader.load(), "build-mismatch");
  assert.equal(mismatchedMain.counters.requireCalls.length, 0);

  const malformedManifest = await createFixture({ manifestBytes: Buffer.from("{") });
  t.after(malformedManifest.cleanup);
  await expectLoaderCode(() => malformedManifest.loader.load(), "asset-mismatch");
  assert.equal(malformedManifest.counters.requireCalls.length, 0);
});

void test("manifest, styles, and engine hashes are checked against the baked build snapshot", async (t) => {
  for (const filename of [MANIFEST_PATH, STYLES_PATH, ENGINE_PATH]) {
    const fixture = await createFixture();
    t.after(fixture.cleanup);
    await fs.writeFile(path.join(fixture.directory, filename), "mixed build bytes");
    await expectLoaderCode(() => fixture.loader.load(), "asset-mismatch");
    assert.equal(fixture.counters.requireCalls.length, 0, filename);
  }
});

void test("every required sibling is mandatory, and main.js remains independently bounded", async (t) => {
  for (const filename of [MAIN_PATH, MANIFEST_PATH, STYLES_PATH, ENGINE_PATH]) {
    const fixture = await createFixture();
    t.after(fixture.cleanup);
    await fs.rm(path.join(fixture.directory, filename));
    await expectLoaderCode(() => fixture.loader.load(), "asset-mismatch");
    assert.equal(fixture.counters.requireCalls.length, 0, filename);
  }

  const oversizedMain = await createFixture({
    mainBytes: Buffer.concat([
      Buffer.from(MAIN_MARKER),
      Buffer.alloc(512_001),
    ]),
  });
  t.after(oversizedMain.cleanup);
  await expectLoaderCode(() => oversizedMain.loader.load(), "asset-mismatch");
  assert.equal(oversizedMain.counters.requireCalls.length, 0);
});

void test("stale ABI, plugin version, and build ID exports are rejected after require", async (t) => {
  for (const [label, mismatch] of [
    ["abi", { abiVersion: 2 }],
    ["version", { pluginVersion: "3.4.9" }],
    ["build-id", { buildId: "old-build" }],
  ] as const) {
    const counters = { runtimeCalls: 0, requireCalls: [] as string[], engineCalls: 0 };
    const fixture = await createFixture({ moduleExports: createEngine(counters, mismatch) });
    t.after(fixture.cleanup);
    await expectLoaderCode(() => fixture.loader.load(), "engine-mismatch");
    assert.equal(fixture.counters.requireCalls.length, 1, label);
  }

  let getterCalls = 0;
  const accessorEngine: Record<string, unknown> = {
    buildId: BUILD_ID,
    pluginVersion: PLUGIN_VERSION,
    renderDocx: async () => undefined,
  };
  Object.defineProperty(accessorEngine, "abiVersion", {
    enumerable: true,
    get: () => {
      getterCalls += 1;
      return 1;
    },
  });
  const getterFixture = await createFixture({ moduleExports: accessorEngine });
  t.after(getterFixture.cleanup);
  await expectLoaderCode(() => getterFixture.loader.load(), "engine-mismatch");
  assert.equal(getterCalls, 0);
});

void test("concurrent loads share one in-flight promise and reject recursive require", async (t) => {
  const fixture = await createFixture();
  t.after(fixture.cleanup);
  const [first, second] = await Promise.all([
    fixture.loader.load(),
    fixture.loader.load(),
  ]);
  assert.equal(first.buildId, second.buildId);
  assert.equal(fixture.counters.requireCalls.length, 1);

  let reentrantLoader: Loader | null = null;
  let reentrantResult: Promise<unknown> | null = null;
  const reentrant = await createFixture({
    moduleRequire: () => {
      reentrantResult = reentrantLoader?.load() ?? null;
      if (reentrantResult) {
        void reentrantResult.catch(() => {});
      }
      return createEngine(reentrant.counters);
    },
  });
  t.after(reentrant.cleanup);
  reentrantLoader = reentrant.loader;
  await reentrant.loader.load();
  assert.ok(reentrantResult);
  await expectLoaderCode(() => reentrantResult!, "recursive-load");
});

void test("cached identity is revalidated and unload cancels late work", async (t) => {
  const fixture = await createFixture();
  t.after(fixture.cleanup);
  await fixture.loader.load();
  await fs.writeFile(path.join(fixture.directory, ENGINE_PATH), "changed while enabled");
  await expectLoaderCode(() => fixture.loader.load(), "asset-mismatch");
  assert.equal(fixture.counters.requireCalls.length, 1);

  let releaseRead: (() => void) | undefined;
  const readGate = new Promise<void>((resolve) => {
    releaseRead = resolve;
  });
  const late = await createFixture({ beforeRead: () => readGate });
  t.after(late.cleanup);
  const target = {} as HTMLElement;
  const renderPromise = late.loader.renderDocx(new ArrayBuffer(0), target, {
    ownerDocument: {} as Document,
    isCancelled: () => false,
    yieldControl: async () => {},
  });
  late.loader.dispose();
  releaseRead?.();
  await expectLoaderCode(() => renderPromise, "disposed");
  assert.equal(late.counters.requireCalls.length, 0);
  assert.equal(late.counters.engineCalls, 0);
});

void test("engine-unavailable diagnostics preserve each admitted equation occurrence", async (t) => {
  const fixture = await createFixture({ isDesktopApp: false });
  t.after(fixture.cleanup);
  const buffer = await createEquationFixture("docx", {
    progId: "Equation.3",
    previewRelationship: { id: "image", target: "media/preview.png" },
    previewBytes: new Uint8Array([137]),
    opaqueEmbeddingBytes: new TextEncoder().encode("opaque payload"),
  });

  await assert.rejects(() => fixture.loader.load(buffer), (error: unknown) => {
    const issues = (error as { issues?: Array<Record<string, unknown>> }).issues;
    assert.equal((error as { code?: string }).code, "desktop-only");
    assert.equal(issues?.length, 1);
    assert.equal(issues?.[0].partPath, "word/document.xml");
    assert.equal(typeof issues?.[0].occurrenceId, "string");
    assert.equal(issues?.[0].code, "docx-engine-unavailable");
    return true;
  });
  assert.equal(fixture.counters.requireCalls.length, 0);
});
