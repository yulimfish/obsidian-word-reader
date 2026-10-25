/* eslint-disable obsidianmd/no-global-this -- Worker descriptor snapshots require Node's test-thread global target. */
import assert from "node:assert/strict";
import { createJiti } from "jiti";
import { Worker } from "node:worker_threads";
import { clearTimeout as clearNodeTimeout, setTimeout as setNodeTimeout } from "node:timers";
import test from "node:test";
import fs from "node:fs";
import path from "node:path";
import { build, type Plugin } from "esbuild";
import JSZip from "jszip";
import { DOMParser } from "@xmldom/xmldom";

import { createEquationFixture } from "./fixtureBuilders";

const jiti = createJiti(__filename);
const WORKER_SOURCE = path.resolve(process.cwd(), "tests/fidelity/docxProduction.worker.ts");

interface WorkerMessage {
  type: string;
  [key: string]: unknown;
}

function normalizePath(value: string): string {
  return value.split(path.sep).join("/");
}

function installationPoints(): Array<{ owner: object; property: string }> {
  const sample = new DOMParser().parseFromString("<!--before--><root/>", "application/xml");
  return [
    { owner: globalThis, property: "DOMParser" },
    { owner: globalThis, property: "XMLSerializer" },
    { owner: globalThis, property: "Node" },
    { owner: Object.getPrototypeOf(sample) as object, property: "firstElementChild" },
    { owner: Object.getPrototypeOf(sample.documentElement) as object, property: "firstElementChild" },
  ];
}

function descriptorSnapshot(): unknown[] {
  return installationPoints().map(({ owner, property }) => {
    const chain: unknown[] = [];
    let current: object | null = owner;
    while (current) {
      chain.push({
        own: Object.getOwnPropertyDescriptor(current, property),
      });
      current = Object.getPrototypeOf(current) as object | null;
    }
    return chain;
  });
}

function startWorker(bundlePath: string, workerData: unknown): {
  worker: Worker;
  message: Promise<WorkerMessage>;
  exit: Promise<number>;
} {
  const worker = new Worker(bundlePath, { workerData });
  const message = new Promise<WorkerMessage>((resolve, reject) => {
    worker.once("message", resolve);
    worker.once("error", reject);
  });
  const exit = new Promise<number>((resolve) => {
    worker.once("exit", resolve);
  });
  return { worker, message, exit };
}

async function receiveMessage(run: ReturnType<typeof startWorker>): Promise<WorkerMessage> {
  let timeout: NodeJS.Timeout | undefined;
  try {
    return await Promise.race([
      run.message,
      new Promise<never>((_, reject) => {
        timeout = setNodeTimeout(() => reject(new Error("Parser worker timed out.")), 15_000);
      }),
    ]);
  } catch (error) {
    await run.worker.terminate();
    await run.exit;
    throw error;
  } finally {
    if (timeout) clearNodeTimeout(timeout);
  }
}

async function runSuccessfulWorker(
  bundlePath: string,
  workerData: unknown,
): Promise<WorkerMessage> {
  const run = startWorker(bundlePath, workerData);
  const message = await receiveMessage(run);
  const exitCode = await run.exit;
  assert.equal(message.type, "success", JSON.stringify(message));
  assert.equal(exitCode, 0);
  return message;
}

async function stopParkedWorker(
  run: ReturnType<typeof startWorker>,
): Promise<number> {
  const termination = run.worker.terminate();
  await termination;
  return await run.exit;
}

function assertReadOnlyBundle(metafile: Awaited<ReturnType<typeof build>>["metafile"]): void {
  const inputs = Object.keys(metafile?.inputs ?? {}).map(normalizePath);
  assert.ok(inputs.some((input) => input.endsWith("src/ooxml/jszipReader.ts")), "metafile must contain the actual read-only ZIP host");
  assert.ok(inputs.some((input) => input.includes("node_modules/pako/lib/inflate.js")), "metafile must contain pako inflate");
  assert.equal(inputs.some((input) => /node_modules\/jszip\/lib\/(?:generate|compressions)\//.test(input)), false);
  assert.equal(inputs.some((input) => input.endsWith("node_modules/pako/lib/deflate.js")), false);
}

async function loadProductionZipPlugin(root: string): Promise<(root: string) => Plugin> {
  let module: Record<string, unknown>;
  try {
    module = await jiti.import(path.join(root, "scripts/read-only-zip-plugin.mjs"));
  } catch (error) {
    assert.fail(
      `shared production read-only ZIP plugin is missing: ${error instanceof Error ? error.message : String(error)}`,
    );
  }
  assert.equal(typeof module.createReadOnlyZipPlugin, "function");
  return module.createReadOnlyZipPlugin as (root: string) => Plugin;
}

async function equationFixtureWithSurroundingText(): Promise<ArrayBuffer> {
  const buffer = await createEquationFixture("docx", {
    progId: "Equation.3",
    previewRelationship: { id: "image", target: "media/preview.png" },
    previewBytes: Uint8Array.from([137, 80, 78, 71, 13, 10, 26, 10, 0]),
    opaqueEmbeddingBytes: new TextEncoder().encode("opaque OLE bytes must not be read"),
  });
  const zip = await JSZip.loadAsync(buffer);
  const document = zip.file("word/document.xml")!;
  const xml = await document.async("string");
  zip.file(
    "word/document.xml",
    xml
      .replace("<w:body><w:p><w:r>", "<w:body><w:p><w:r><w:t>before</w:t></w:r><w:r>")
      .replace("</w:r></w:p>", "</w:r><w:r><w:t>after</w:t></w:r></w:p>"),
  );
  return zip.generateAsync({ type: "arraybuffer", compression: "STORE" });
}

void test("real docx-preview parses the sanitized package through the production read-only ZIP alias in isolated workers", async () => {
  const root = process.cwd();
  const pluginFactory = await loadProductionZipPlugin(root);
  const tempDir = fs.mkdtempSync(path.join(__dirname, "docx-production-"));
  const workerBundle = path.join(tempDir, "docxProduction.worker.cjs");
  let fixture: ArrayBuffer | null = null;
  let inputBefore: Uint8Array | null = null;
  try {
    assert.equal(
      path.dirname(tempDir),
      __dirname,
      "worker bundle temp directory must stay beneath this test's existing output directory",
    );
    const result = await build({
      absWorkingDir: root,
      entryPoints: [WORKER_SOURCE],
      outfile: workerBundle,
      bundle: true,
      platform: "node",
      format: "cjs",
      target: "es2018",
      metafile: true,
      logLevel: "silent",
      plugins: [pluginFactory(root)],
    });
    assertReadOnlyBundle(result.metafile);

    fixture = await equationFixtureWithSurroundingText();
    inputBefore = new Uint8Array(fixture).slice();
    const parentDescriptors = descriptorSnapshot();
    const [first, second] = await Promise.all([
      runSuccessfulWorker(workerBundle, { mode: "parse", buffer: fixture.slice(0) }),
      runSuccessfulWorker(workerBundle, { mode: "parse", buffer: fixture.slice(0) }),
    ]);
    assert.notEqual(first.threadId, second.threadId);
    assert.equal(first.globalTargetIsWorkerGlobal, true);
    assert.equal(second.globalTargetIsWorkerGlobal, true);
    assert.deepEqual(descriptorSnapshot(), parentDescriptors);
    assert.deepEqual(new Uint8Array(fixture), inputBefore);

    for (const message of [first, second]) {
      assert.equal(message.distinctSanitizedBuffer, true);
      assert.equal(message.originalUnchanged, true);
      assert.equal(message.inventoryOccurrences, 1);
      assert.deepEqual(message.issues, []);
      assert.equal(message.imageCount, 1);
      assert.ok((message.text as string[]).includes("before"));
      assert.ok((message.text as string[]).includes("after"));
      assert.ok((message.readableImageParts as string[]).length >= 1);
      assert.deepEqual(message.ordinaryPayloadReads, []);
      assert.equal(message.oleReferencesRemain, false);
      assert.equal(message.externalResourcesRemain, false);
      assert.equal(message.altChunksRemain, false);
      assert.equal(message.shimRestored, true);
    }

    const probe = await runSuccessfulWorker(workerBundle, { mode: "shim-probe" });
    assert.equal(probe.actualXmldomSemantics, true);
    assert.equal(probe.descriptorsRestored, true);

    const partial = await runSuccessfulWorker(workerBundle, { mode: "partial-install" });
    assert.equal(partial.parserCallbackEntered, false);
    assert.equal(partial.rollbackExact, true);

    const sharedMarker = new SharedArrayBuffer(Int32Array.BYTES_PER_ELEMENT);
    const pending = startWorker(workerBundle, {
      mode: "pending-sibling",
      marker: sharedMarker,
    });
    const pendingMessage = await receiveMessage(pending);
    assert.equal(pendingMessage.type, "aggregate-rejected");
    assert.equal(pendingMessage.shimStillInstalled, true);
    assert.equal(Atomics.load(new Int32Array(sharedMarker), 0), 0);
    await stopParkedWorker(pending);
    assert.equal(Atomics.load(new Int32Array(sharedMarker), 0), 0);

    const malformed = startWorker(workerBundle, { mode: "parser-failure" });
    const failure = await receiveMessage(malformed);
    assert.equal(failure.type, "failure");
    assert.equal(failure.stage, "parseAsync");
    await stopParkedWorker(malformed);
  } finally {
    if (fixture && inputBefore) assert.deepEqual(new Uint8Array(fixture), inputBefore);
    fs.rmSync(tempDir, { recursive: true, force: true });
  }
});
