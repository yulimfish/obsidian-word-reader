import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import { build } from "esbuild";
import { createRequire } from "node:module";
import Module from "node:module";

interface TestModule extends Module {
  _compile(source: string, filename: string): void;
}

const require = createRequire(__filename);

void test("desktop runtime uses Obsidian's injected CommonJS require", async () => {
  const result = await build({
    entryPoints: [path.resolve("src/desktopRuntime.ts")],
    bundle: true,
    platform: "node",
    format: "cjs",
    external: ["obsidian", "electron", "node:*"],
    write: false,
    logLevel: "silent",
  });
  const source = result.outputFiles[0].text;
  const obsidian = { Platform: { isDesktopApp: true } };
  const runtimeModule = new Module("obsidian-runtime-test") as TestModule;
  const load = (request: string): unknown => request === "obsidian" ? obsidian : require(request) as unknown;
  runtimeModule.require = load as typeof runtimeModule.require;
  runtimeModule._compile(source, "desktopRuntime.test.cjs");
  const runtime = runtimeModule.exports as {
    getDocxEngineNodeRuntime: () => { fs: { lstat: unknown }; path: { sep: string } };
  };
  const nodeRuntime = runtime.getDocxEngineNodeRuntime();
  assert.equal(nodeRuntime.fs.lstat, fs.promises.lstat);
  assert.equal(nodeRuntime.path.sep, path.sep);
});
