import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";
import test from "node:test";
import { isDeepStrictEqual } from "node:util";

import {
  assertCapturedEngineBinding,
  assertControlSnapshotEqual,
  bindControlInputs,
  bindMetafileInputs,
  captureDocxEvidenceAttempt3,
  createEvidencePersistenceForTesting,
  inspectControlFile,
  mergeBoundInputs,
  pinnedPhysicalLoad,
  prepareEvidenceInputs,
  preparePinnedToolchainMetadata,
  recheckPinnedSnapshots,
  verifyAttempt2Unchanged,
  type BoundControlInput,
  type BoundMetafileInput,
  type ControlCategory,
  type ControlInputManifest,
  type EvidencePersistenceIo,
  type EvidencePersistencePayload,
  type EvidencePreGuardFiles,
  type PhysicalInputInspection,
  type PreservedAttemptReadIo,
  type RawProjectFileSnapshot,
  type VirtualLoadRecord,
} from "./docxEvidenceCapture";

const VIRTUAL_NAMESPACE = "read-only-zip-disabled-readable-stream";
const VIRTUAL_PATH = "jszip-readable-stream-disabled";
const VIRTUAL_KEY = `${VIRTUAL_NAMESPACE}:${VIRTUAL_PATH}`;
const VIRTUAL_TEXT = "module.exports = Object.freeze({});";
const EVIDENCE_ROOT = "/private/var/folders/60/0btch0dj3111k_1grpv_9xmw0000gn/T/opencode";
const EVIDENCE_CHILD = path.join(EVIDENCE_ROOT, "obsidian-word-reader-W69-S5.3-evidence-attempt-3");
const EVIDENCE_PARENT = path.dirname(EVIDENCE_ROOT);
const PROJECT_ROOT = "/workspace/project";
const sha256 = (bytes: Uint8Array): string => createHash("sha256").update(bytes).digest("hex");

function stableValue(value: unknown): unknown {
  if (Array.isArray(value)) return value.map(stableValue);
  if (value && typeof value === "object") {
    return Object.fromEntries(Object.keys(value).sort().map((key) => [key, stableValue((value as Record<string, unknown>)[key])]));
  }
  return value;
}

function stableJson(value: unknown): string {
  return JSON.stringify(stableValue(value));
}

interface MatrixVariant {
  readonly label: string;
  readonly expected: unknown;
  readonly run: () => Promise<unknown>;
  readonly matches?: (actual: unknown) => boolean;
}

type Observation<T> =
  | { readonly outcome: "returned"; readonly value: T }
  | { readonly outcome: "threw"; readonly errorName: string; readonly errorMessage: string };

async function assertMatrixOutcomes(name: string, variants: readonly MatrixVariant[]): Promise<void> {
  const settled = await Promise.allSettled(variants.map((variant) => variant.run()));
  const records = variants.map((variant, index) => {
    const result = settled[index];
    const actual = result.status === "fulfilled"
      ? result.value
      : {
        outcome: "setup-threw",
        errorName: result.reason instanceof Error ? result.reason.name : typeof result.reason,
        errorMessage: result.reason instanceof Error ? result.reason.message : String(result.reason),
      };
    const matched = variant.matches
      ? result.status === "fulfilled" && variant.matches(actual)
      : result.status === "fulfilled" && isDeepStrictEqual(variant.expected, actual);
    return { label: variant.label, expected: variant.expected, actual, matched };
  });
  const mismatches = records.filter((record) => !record.matched);
  console.log(`${name} matrix outcomes:\n${JSON.stringify(records, null, 2)}`);
  assert.equal(
    mismatches.length,
    0,
    `${name} matrix mismatches; all expected/actual outcomes follow:\n${JSON.stringify(records, null, 2)}`,
  );
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function summarizeObservation<T>(
  observation: Observation<T>,
  summarizeValue: (value: T) => unknown = (value) => value,
): Record<string, unknown> {
  if (observation.outcome === "threw") {
    return {
      outcome: observation.outcome,
      errorName: observation.errorName,
      errorMessage: observation.errorMessage,
    };
  }
  const value = summarizeValue(observation.value);
  return value === undefined
    ? { outcome: observation.outcome }
    : { outcome: observation.outcome, value };
}

function fieldOutcome(record: unknown, field: string): string | undefined {
  if (!isRecord(record)) return undefined;
  const value = record[field];
  if (!isRecord(value)) return undefined;
  return typeof value.outcome === "string" ? value.outcome : undefined;
}

async function observe<T>(operation: () => T | Promise<T>): Promise<Observation<T>> {
  try {
    return { outcome: "returned", value: await operation() };
  } catch (error) {
    return {
      outcome: "threw",
      errorName: error instanceof Error ? error.name : typeof error,
      errorMessage: error instanceof Error ? error.message : String(error),
    };
  }
}

function rawSnapshot(
  relativePath: string,
  overrides: Partial<RawProjectFileSnapshot> = {},
): RawProjectFileSnapshot {
  const contents = Buffer.from(overrides.contents ?? `fixture:${relativePath}`);
  const canonicalPath = overrides.canonicalPath ?? path.resolve(PROJECT_ROOT, relativePath);
  return {
    projectRoot: overrides.projectRoot ?? PROJECT_ROOT,
    canonicalPath,
    contents,
    byteLength: overrides.byteLength ?? contents.byteLength,
    sha256: overrides.sha256 ?? sha256(contents),
    device: overrides.device ?? "memory-device",
    inode: overrides.inode ?? `memory-inode:${relativePath}`,
  };
}

function compiledInput(
  relativePath: string,
  overrides: Partial<RawProjectFileSnapshot> = {},
): BoundMetafileInput {
  const snapshot = rawSnapshot(relativePath, overrides);
  return {
    namespace: "file",
    key: relativePath,
    path: snapshot.canonicalPath,
    canonicalPath: snapshot.canonicalPath,
    loader: "ts",
    resolveDir: path.dirname(snapshot.canonicalPath),
    byteLength: snapshot.byteLength,
    sha256: snapshot.sha256,
    device: snapshot.device,
    inode: snapshot.inode,
    entries: ["engine"],
    imports: [],
    bytes: Buffer.from(snapshot.contents),
  };
}

function frozenManifest(
  rows: readonly { readonly path: string; readonly categories?: readonly ControlCategory[] }[],
): ControlInputManifest {
  return Object.freeze(rows.map((row) => Object.freeze({
    path: row.path,
    categories: Object.freeze([...(row.categories ?? ["calculateBuildId"])]),
  })));
}

function summarizeControl(input: BoundControlInput): Record<string, unknown> {
  return {
    path: input.path,
    canonicalPath: input.canonicalPath,
    role: input.role,
    categories: input.categories,
    compiledKeys: input.compiledKeys,
    device: input.device,
    inode: input.inode,
    byteLength: input.byteLength,
    sha256: input.sha256,
  };
}

function rawSnapshotFromCompiled(input: BoundMetafileInput): RawProjectFileSnapshot {
  return {
    projectRoot: PROJECT_ROOT,
    canonicalPath: input.canonicalPath,
    contents: input.bytes,
    byteLength: input.byteLength,
    sha256: input.sha256,
    device: String(input.device),
    inode: String(input.inode),
  };
}

function openBasenames(events: readonly string[]): string[] {
  return events
    .filter((event) => event.startsWith("open:"))
    .map((event) => path.basename(event.split(":")[1]));
}

function virtualLoad(overrides: Partial<VirtualLoadRecord> = {}): VirtualLoadRecord {
  return {
    entry: "engine",
    namespace: VIRTUAL_NAMESPACE,
    path: VIRTUAL_PATH,
    suffix: "",
    importAttributes: {},
    loader: "js",
    resolveDir: "",
    contents: VIRTUAL_TEXT,
    ...overrides,
  };
}

function metafile(inputs: Record<string, { imports: readonly unknown[] }>): { inputs: Record<string, { imports: readonly unknown[] }> } {
  return { inputs };
}

function physicalInspection(
  key: string,
  overrides: Partial<PhysicalInputInspection> = {},
): PhysicalInputInspection {
  const root = "/workspace/project";
  const canonicalPath = path.resolve(root, key);
  const contents = Buffer.from("export const captured = true;\n");
  return {
    projectRoot: root,
    canonicalPath,
    contents,
    loader: "ts",
    resolveDir: path.dirname(canonicalPath),
    device: 1,
    inode: key === "src/shared.ts" ? 9 : 10,
    sha256: sha256(contents),
    ...overrides,
  };
}

function inspectorFor(
  rows: Record<string, PhysicalInputInspection>,
  calls: string[],
): (key: string) => Promise<PhysicalInputInspection> {
  return async (key) => {
    calls.push(key);
    const result = rows[key];
    assert.ok(result, `unexpected physical inspection: ${key}`);
    return result;
  };
}

void test("capture binding: exact virtual input binds with no physical inspection", async () => {
  const captureSource = await fs.readFile(
    path.resolve(process.cwd(), "tests/fidelity/docxEvidenceCapture.ts"),
    "utf8",
  );
  const hashesObject = captureSource.match(
    /const\s+ATTEMPT_1_HASHES\s*=\s*Object\.freeze\(\{([\s\S]*?)\}\);/,
  );
  assert.ok(hashesObject, "ATTEMPT_1_HASHES source object should exist");
  const expectedHashes = [
    ["guard-error.json", "4c619ea4b9f9be801f0f0e40be80a157ba258f94ab743b52c6bcf523b36172b3"],
    ["receipt.json", "fcb4d587066de47d0a2f5f6116a7797ca2cf3cf26785523458beb280984c4d56"],
  ] as const;
  for (const [name, expectedHash] of expectedHashes) {
    const escapedName = name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const keyPattern = new RegExp(`[\"']${escapedName}[\"']\\s*:`, "g");
    assert.equal([...hashesObject[1].matchAll(keyPattern)].length, 1, `${name} should occur exactly once`);
    const literalPattern = new RegExp(`[\"']${escapedName}[\"']\\s*:\\s*[\"']([0-9a-f]+)[\"']`, "g");
    const literals = [...hashesObject[1].matchAll(literalPattern)];
    assert.equal(literals.length, 1, `${name} should have exactly one hash literal`);
    assert.equal(literals[0][1], expectedHash);
    assert.match(literals[0][1], /^[0-9a-f]{64}$/);
  }

  const calls: string[] = [];
  const result = await bindMetafileInputs(
    metafile({ [VIRTUAL_KEY]: { imports: [] } }),
    [virtualLoad()],
    inspectorFor({}, calls),
  );

  assert.equal(calls.length, 0);
  assert.equal(result.length, 1);
  assert.equal(result[0].namespace, VIRTUAL_NAMESPACE);
  assert.equal(result[0].key, VIRTUAL_KEY);
  assert.equal(result[0].path, VIRTUAL_PATH);
  assert.equal(result[0].loader, "js");
  assert.equal(result[0].resolveDir, "");
  assert.equal(result[0].byteLength, Buffer.byteLength(VIRTUAL_TEXT, "utf8"));
  assert.equal(result[0].sha256, sha256(Buffer.from(VIRTUAL_TEXT, "utf8")));
  assert.deepEqual(result[0].entries, ["engine"]);
  const exposed = result[0].bytes;
  exposed[0] ^= 0xff;
  assert.deepEqual(Buffer.from(result[0].bytes), Buffer.from(VIRTUAL_TEXT, "utf8"));
});

void test("capture binding: styles control binds raw bytes without loader", async () => {
  const projectRoot = "/workspace/project";
  const relativePath = "styles.css";
  const contents = Buffer.from(".preview { color: #123456; }\n", "utf8");
  const snapshot: RawProjectFileSnapshot = {
    projectRoot,
    canonicalPath: path.resolve(projectRoot, relativePath),
    contents,
    byteLength: contents.byteLength,
    sha256: sha256(contents),
    device: "memory-device",
    inode: "memory-inode",
  };

  const result = await inspectControlFile(
    projectRoot,
    relativePath,
    async () => snapshot,
  );

  assert.equal(result.projectRoot, snapshot.projectRoot);
  assert.equal(result.canonicalPath, snapshot.canonicalPath);
  assert.deepEqual(Buffer.from(result.contents), contents);
  assert.equal(result.byteLength, snapshot.byteLength);
  assert.equal(result.sha256, snapshot.sha256);
  assert.equal(result.device, snapshot.device);
  assert.equal(result.inode, snapshot.inode);
  assert.equal("loader" in result, false);
  assert.equal("resolveDir" in result, false);
  assert.equal("entries" in result, false);
});

void test("capture binding: engine byte define is a strict safe numeric length", () => {
  type ExpectedResult =
    | { readonly outcome: "returned" }
    | {
      readonly outcome: "threw";
      readonly errorName: string;
      readonly errorMessage?: string;
    };
  type MatrixCase = {
    readonly label: string;
    readonly rawDefines: Record<string, string>;
    readonly expectedEngineLength: number;
    readonly expectedResult: ExpectedResult;
  };

  const expectedDigest = "a".repeat(64);
  const wrongDigest = "b".repeat(64);
  const genericMismatch = "main.js does not embed the captured engine digest/byte length.";
  const invalidExpectedLength = "Invalid expected DOCX engine byte length.";
  const invalidEmbeddedLength = "Invalid embedded DOCX engine byte length.";
  const defines = (rawByteLength: string, digest = expectedDigest): Record<string, string> => ({
    __DOCX_ENGINE_SHA256__: JSON.stringify(digest),
    __DOCX_ENGINE_BYTES__: rawByteLength,
  });
  const throws = (errorMessage: string, errorName = "Error"): ExpectedResult => ({
    outcome: "threw",
    errorName,
    errorMessage,
  });

  const matrix: readonly MatrixCase[] = [
    {
      label: "valid JSON number seven returns",
      rawDefines: defines("7"),
      expectedEngineLength: 7,
      expectedResult: { outcome: "returned" },
    },
    {
      label: "valid JSON number zero returns",
      rawDefines: defines("0"),
      expectedEngineLength: 0,
      expectedResult: { outcome: "returned" },
    },
    {
      label: "JSON digit string rejects",
      rawDefines: defines(JSON.stringify("7")),
      expectedEngineLength: 7,
      expectedResult: throws(invalidEmbeddedLength),
    },
    {
      label: "wrong numeric length uses original mismatch diagnostic",
      rawDefines: defines("8"),
      expectedEngineLength: 7,
      expectedResult: throws(genericMismatch),
    },
    {
      label: "negative embedded length is invalid",
      rawDefines: defines("-1"),
      expectedEngineLength: 7,
      expectedResult: throws(invalidEmbeddedLength),
    },
    {
      label: "fractional embedded length is invalid",
      rawDefines: defines("0.5"),
      expectedEngineLength: 7,
      expectedResult: throws(invalidEmbeddedLength),
    },
    {
      label: "unsafe embedded length is invalid",
      rawDefines: defines(String(Number.MAX_SAFE_INTEGER + 1)),
      expectedEngineLength: 7,
      expectedResult: throws(invalidEmbeddedLength),
    },
    {
      label: "null embedded length is invalid",
      rawDefines: defines("null"),
      expectedEngineLength: 7,
      expectedResult: throws(invalidEmbeddedLength),
    },
    {
      label: "boolean embedded length is invalid",
      rawDefines: defines("true"),
      expectedEngineLength: 7,
      expectedResult: throws(invalidEmbeddedLength),
    },
    {
      label: "malformed JSON preserves SyntaxError",
      rawDefines: defines("{"),
      expectedEngineLength: 7,
      expectedResult: { outcome: "threw", errorName: "SyntaxError" },
    },
    {
      label: "wrong digest uses original mismatch diagnostic",
      rawDefines: defines("7", wrongDigest),
      expectedEngineLength: 7,
      expectedResult: throws(genericMismatch),
    },
    {
      label: "negative expected length is invalid before byte comparison",
      rawDefines: defines(JSON.stringify("-1")),
      expectedEngineLength: -1,
      expectedResult: throws(invalidExpectedLength),
    },
    {
      label: "fractional expected length is invalid before byte comparison",
      rawDefines: defines(JSON.stringify("0.5")),
      expectedEngineLength: 0.5,
      expectedResult: throws(invalidExpectedLength),
    },
    {
      label: "unsafe expected length is invalid before byte comparison",
      rawDefines: defines(JSON.stringify(String(Number.MAX_SAFE_INTEGER + 1))),
      expectedEngineLength: Number.MAX_SAFE_INTEGER + 1,
      expectedResult: throws(invalidExpectedLength),
    },
    {
      label: "NaN expected length is invalid before byte comparison",
      rawDefines: defines(JSON.stringify("NaN")),
      expectedEngineLength: Number.NaN,
      expectedResult: throws(invalidExpectedLength),
    },
    {
      label: "infinite expected length is invalid before byte comparison",
      rawDefines: defines(JSON.stringify("Infinity")),
      expectedEngineLength: Number.POSITIVE_INFINITY,
      expectedResult: throws(invalidExpectedLength),
    },
    {
      label: "runtime string expected length is invalid before byte comparison",
      rawDefines: defines(JSON.stringify("7")),
      expectedEngineLength: "7" as unknown as number,
      expectedResult: throws(invalidExpectedLength),
    },
  ];

  const outcomes = matrix.map((scenario) => {
    let actual: { readonly outcome: "returned" } | {
      readonly outcome: "threw";
      readonly errorName: string;
      readonly errorMessage: string;
    };
    try {
      assertCapturedEngineBinding(
        scenario.rawDefines,
        expectedDigest,
        scenario.expectedEngineLength,
      );
      actual = { outcome: "returned" };
    } catch (error) {
      actual = {
        outcome: "threw",
        errorName: error instanceof Error ? error.name : typeof error,
        errorMessage: error instanceof Error ? error.message : String(error),
      };
    }
    return {
      label: scenario.label,
      rawDefines: scenario.rawDefines,
      expectedEngineLength: scenario.expectedEngineLength,
      expectedEngineLengthDisplay: String(scenario.expectedEngineLength),
      actual,
      expectedResult: scenario.expectedResult,
    };
  });

  const mismatches = outcomes.filter(({ actual, expectedResult }) => {
    if (expectedResult.outcome === "returned") return actual.outcome !== "returned";
    return actual.outcome !== "threw" ||
      actual.errorName !== expectedResult.errorName ||
      (expectedResult.errorMessage !== undefined && actual.errorMessage !== expectedResult.errorMessage);
  });
  assert.equal(
    mismatches.length,
    0,
    `Engine binding matrix mismatches and all outcomes:\n${JSON.stringify({ mismatches, outcomes }, null, 2)}`,
  );
});

void test("capture binding: physical inputs retain canonical loader identity and content validation", async () => {
  const calls: string[] = [];
  const inspection = physicalInspection("src/shared.ts");
  const result = await bindMetafileInputs(
    metafile({ "src/shared.ts": { imports: [] } }),
    [],
    inspectorFor({ "src/shared.ts": inspection }, calls),
  );

  assert.deepEqual(calls, ["src/shared.ts"]);
  assert.equal(result.length, 1);
  assert.equal(result[0].namespace, "file");
  assert.equal(result[0].canonicalPath, inspection.canonicalPath);
  assert.equal(result[0].loader, "ts");
  assert.equal(result[0].resolveDir, path.dirname(inspection.canonicalPath));
  assert.equal(result[0].sha256, sha256(inspection.contents));
  assert.equal(result[0].device, inspection.device);
  assert.equal(result[0].inode, inspection.inode);
  assert.deepEqual(Buffer.from(result[0].bytes), Buffer.from(inspection.contents));
});

void test("capture binding: rejects unknown namespace before physical inspection", async () => {
  const calls: string[] = [];
  await assert.rejects(
    () => bindMetafileInputs(
      metafile({ "src/ordinary.ts": { imports: [] }, "another-namespace:module": { imports: [] } }),
      [],
      inspectorFor({ "src/ordinary.ts": physicalInspection("src/ordinary.ts") }, calls),
    ),
    /unknown namespace/i,
  );
  assert.deepEqual(calls, []);
});

void test("capture binding: rejects unknown virtual path before physical inspection", async () => {
  const calls: string[] = [];
  await assert.rejects(
    () => bindMetafileInputs(
      metafile({ "src/ordinary.ts": { imports: [] }, "read-only-zip-disabled-readable-stream:other": { imports: [] } }),
      [],
      inspectorFor({ "src/ordinary.ts": physicalInspection("src/ordinary.ts") }, calls),
    ),
    /unknown namespace|virtual/i,
  );
  assert.deepEqual(calls, []);
});

void test("capture binding: rejects wrong virtual namespace before physical inspection", async () => {
  const calls: string[] = [];
  await assert.rejects(
    () => bindMetafileInputs(
      metafile({ "src/ordinary.ts": { imports: [] }, "wrong-namespace:jszip-readable-stream-disabled": { imports: [] } }),
      [],
      inspectorFor({ "src/ordinary.ts": physicalInspection("src/ordinary.ts") }, calls),
    ),
    /unknown namespace/i,
  );
  assert.deepEqual(calls, []);
});

void test("capture binding: rejects missing virtual callback before physical inspection", async () => {
  const calls: string[] = [];
  await assert.rejects(
    () => bindMetafileInputs(
      metafile({ "src/ordinary.ts": { imports: [] }, [VIRTUAL_KEY]: { imports: [] } }),
      [],
      inspectorFor({ "src/ordinary.ts": physicalInspection("src/ordinary.ts") }, calls),
    ),
    /exactly one.*callback|callback.*exactly one/i,
  );
  assert.deepEqual(calls, []);
});

for (const [label, overrides] of [
  ["namespace", { namespace: "other-namespace" }],
  ["path", { path: "other-virtual-path" }],
  ["suffix", { suffix: "?raw" }],
  ["attributes", { importAttributes: { type: "json" } }],
  ["malformed attributes", { importAttributes: null }],
  ["contents", { contents: "module.exports = {};" }],
  ["loader", { loader: "ts" }],
  ["resolveDir", { resolveDir: "/workspace" }],
] as const) {
  void test(`capture binding: rejects virtual ${label} drift before physical inspection`, async () => {
    const calls: string[] = [];
    await assert.rejects(
      () => bindMetafileInputs(
        metafile({ "src/ordinary.ts": { imports: [] }, [VIRTUAL_KEY]: { imports: [] } }),
        [virtualLoad(overrides)],
        inspectorFor({ "src/ordinary.ts": physicalInspection("src/ordinary.ts") }, calls),
      ),
      /virtual|namespace|path|suffix|attribute|content|loader|resolveDir/i,
    );
    assert.deepEqual(calls, []);
  });
}

void test("capture binding: rejects duplicate virtual callbacks before physical inspection", async () => {
  const calls: string[] = [];
  await assert.rejects(
    () => bindMetafileInputs(
      metafile({ "src/ordinary.ts": { imports: [] }, [VIRTUAL_KEY]: { imports: [] } }),
      [virtualLoad(), virtualLoad()],
      inspectorFor({ "src/ordinary.ts": physicalInspection("src/ordinary.ts") }, calls),
    ),
    /duplicate|exactly one/i,
  );
  assert.deepEqual(calls, []);
});

void test("capture binding: rejects virtual imports before physical inspection", async () => {
  const calls: string[] = [];
  await assert.rejects(
    () => bindMetafileInputs(
      metafile({ "src/ordinary.ts": { imports: [] }, [VIRTUAL_KEY]: { imports: [{ path: "x", kind: "import-statement" }] } }),
      [virtualLoad()],
      inspectorFor({ "src/ordinary.ts": physicalInspection("src/ordinary.ts") }, calls),
    ),
    /virtual.*imports|imports.*virtual/i,
  );
  assert.deepEqual(calls, []);
});

void test("capture binding: rejects duplicate raw serialized input keys before physical inspection", async () => {
  const calls: string[] = [];
  const raw = `{"inputs":{"${VIRTUAL_KEY}":{"imports":[]},"${VIRTUAL_KEY}":{"imports":[]}}}`;
  await assert.rejects(
    () => bindMetafileInputs(raw, [virtualLoad()], inspectorFor({}, calls)),
    /duplicate.*key/i,
  );
  assert.deepEqual(calls, []);
});

void test("capture binding: rejects physical escape and hash drift", async () => {
  const escapeCalls: string[] = [];
  await assert.rejects(
    () => bindMetafileInputs(
      metafile({ "../outside.ts": { imports: [] } }),
      [],
      inspectorFor({}, escapeCalls),
    ),
    /physical.*path|outside|escape|canonical/i,
  );
  assert.deepEqual(escapeCalls, []);

  const hashCalls: string[] = [];
  const correct = physicalInspection("src/shared.ts");
  await assert.rejects(
    () => bindMetafileInputs(
      metafile({ "src/shared.ts": { imports: [] } }),
      [],
      inspectorFor({ "src/shared.ts": { ...correct, sha256: "0".repeat(64) } }, hashCalls),
    ),
    /hash|digest/i,
  );
  assert.deepEqual(hashCalls, ["src/shared.ts"]);
});

void test("capture binding: rejects ambiguous physical path or inode aliases", async () => {
  const calls: string[] = [];
  const shared = physicalInspection("src/first.ts", { canonicalPath: "/workspace/project/src/first.ts", inode: 77 });
  const alias = physicalInspection("src/second.ts", { canonicalPath: "/workspace/project/src/first.ts", inode: 77 });
  await assert.rejects(
    () => bindMetafileInputs(
      metafile({ "src/first.ts": { imports: [] }, "src/second.ts": { imports: [] } }),
      [],
      inspectorFor({ "src/first.ts": shared, "src/second.ts": alias }, calls),
    ),
    /ambiguous|alias|one-to-one|identity/i,
  );
  assert.deepEqual(calls, ["src/first.ts", "src/second.ts"]);
});

void test("capture binding: merges matching engine/main virtual tuples with explicit membership", async () => {
  const calls: string[] = [];
  const engine = await bindMetafileInputs(
    metafile({ [VIRTUAL_KEY]: { imports: [] } }),
    [virtualLoad({ entry: "engine" })],
    inspectorFor({}, calls),
  );
  const main = await bindMetafileInputs(
    metafile({ [VIRTUAL_KEY]: { imports: [] } }),
    [virtualLoad({ entry: "main" })],
    inspectorFor({}, calls),
  );
  const merged = mergeBoundInputs([
    { entry: "engine", inputs: engine },
    { entry: "main", inputs: main },
  ]);
  const virtual = merged.find((input) => input.key === VIRTUAL_KEY);
  assert.ok(virtual);
  assert.deepEqual(virtual.entries, ["engine", "main"]);
  assert.equal(calls.length, 0);
});

void test("capture binding: physical CSS rejection", async () => {
  await assertMatrixOutcomes("physical CSS rejection", [{
    label: "physical styles.css key rejects during discovery binding",
    expected: {
      outcome: "throws Unsupported physical input loader extension .css",
      inspectionCalls: ["styles.css"],
    },
    async run() {
      const inspectionCalls: string[] = [];
      const operation = await observe(() => bindMetafileInputs(
        metafile({ "styles.css": { imports: [] } }),
        [],
        async (key) => {
          inspectionCalls.push(key);
          return physicalInspection(key, { loader: "css" });
        },
      ));
      return {
        operation: summarizeObservation(operation, (inputs) => inputs.map((input) => ({
          key: input.key,
          canonicalPath: input.canonicalPath,
          loader: input.loader,
        }))),
        inspectionCalls,
      };
    },
    matches(actual) {
      if (!isRecord(actual) || !isRecord(actual.operation)) return false;
      return actual.operation.outcome === "threw" &&
        actual.operation.errorName === "Error" &&
        actual.operation.errorMessage === `Unsupported physical input loader extension .css: ${path.resolve(PROJECT_ROOT, "styles.css")}` &&
        isDeepStrictEqual(actual.inspectionCalls, ["styles.css"]);
    },
  }]);
});

void test("capture binding: equal dual-role snapshot", async () => {
  await assertMatrixOutcomes("equal dual-role snapshot", [{
    label: "same exact canonical path remains allowed across compiled and control roles",
      expected: {
        binding: "returns one control row linked to the compiled key",
        equality: "returns",
        recheck: "returns after one deduplicated raw read; enumeration count is recorded",
        bindingCalls: ["src/shared.ts"],
        recheckCalls: ["src/shared.ts"],
    },
    async run() {
      const compiled = compiledInput("src/shared.ts");
      const manifest = frozenManifest([{ path: "src/shared.ts" }]);
      const bindingCalls: string[] = [];
      const binding = await observe(() => bindControlInputs(
        manifest,
        [compiled],
        async (relativePath) => {
          bindingCalls.push(relativePath);
          return rawSnapshot(relativePath);
        },
      ));
      const controls = binding.outcome === "returned" ? binding.value : [];
      const equality = await observe(() => {
        const control = controls[0];
        if (!control) throw new Error("Dual-role fixture did not produce a bound control.");
        return assertControlSnapshotEqual(rawSnapshotFromCompiled(compiled), control);
      });
      let enumerationCalls = 0;
      const recheckCalls: string[] = [];
      const recheck = await observe(() => recheckPinnedSnapshots(
        PROJECT_ROOT,
        [compiled],
        controls,
        manifest,
        async () => {
          enumerationCalls += 1;
          return frozenManifest([{ path: "src/shared.ts" }]);
        },
        async (relativePath) => {
          recheckCalls.push(relativePath);
          return rawSnapshot(relativePath);
        },
      ));
      return {
        binding: summarizeObservation(binding, (rows) => rows.map(summarizeControl)),
        equality: summarizeObservation(equality),
        recheck: summarizeObservation(recheck),
        boundRows: controls.map(summarizeControl),
        bindingCalls,
        enumerationCalls,
        recheckCalls,
      };
    },
    matches(actual) {
      if (!isRecord(actual)) return false;
      const rows = actual.boundRows;
      return fieldOutcome(actual, "binding") === "returned" &&
        fieldOutcome(actual, "equality") === "returned" &&
        fieldOutcome(actual, "recheck") === "returned" &&
        Array.isArray(rows) && rows.length === 1 && isRecord(rows[0]) &&
        isDeepStrictEqual(rows[0].compiledKeys, ["src/shared.ts"]) &&
        isDeepStrictEqual(actual.bindingCalls, ["src/shared.ts"]) &&
        isDeepStrictEqual(actual.recheckCalls, ["src/shared.ts"]);
    },
  }]);
});

void test("capture binding: identity drift", async () => {
  const variants: MatrixVariant[] = [
    {
      label: "same-path dual-role identity drift must reject before raw reinspection",
      expected: { recheck: "throws on differing device/inode", recheckCalls: [] },
      async run() {
        const compiled = compiledInput("src/shared.ts");
        const controlSnapshot = rawSnapshot("src/shared.ts", {
          device: "different-device",
          inode: "different-inode",
        });
        const manifest = frozenManifest([{ path: "src/shared.ts" }]);
        const binding = await observe(() => bindControlInputs(
          manifest,
          [compiled],
          async () => controlSnapshot,
        ));
        const controls = binding.outcome === "returned" ? binding.value : [];
        const equality = await observe(() => {
          const control = controls[0];
          if (!control) throw new Error("Identity-drift fixture did not produce a bound control.");
          return assertControlSnapshotEqual(rawSnapshotFromCompiled(compiled), control);
        });
        let enumerationCalls = 0;
        const recheckCalls: string[] = [];
        const recheck = await observe(() => recheckPinnedSnapshots(
          PROJECT_ROOT,
          [compiled],
          controls,
          manifest,
          async () => {
            enumerationCalls += 1;
            return frozenManifest([{ path: "src/shared.ts" }]);
          },
          async (relativePath) => {
            recheckCalls.push(relativePath);
            return controlSnapshot;
          },
        ));
        return {
          compiledIdentity: { device: compiled.device, inode: compiled.inode },
          controlIdentity: { device: controlSnapshot.device, inode: controlSnapshot.inode },
          binding: summarizeObservation(binding, (rows) => rows.map(summarizeControl)),
          equality: summarizeObservation(equality),
          legacyOutcome: binding.outcome === "returned" && recheck.outcome === "returned"
            ? "unexpected same-path dual-role identity acceptance"
            : "identity drift rejected",
          recheck: summarizeObservation(recheck),
          enumerationCalls,
          recheckCalls,
        };
      },
      matches(actual) {
        if (!isRecord(actual)) return false;
        return fieldOutcome(actual, "binding") === "threw" ||
          (fieldOutcome(actual, "binding") === "returned" &&
            fieldOutcome(actual, "recheck") === "threw" &&
            actual.enumerationCalls === 1 &&
            isDeepStrictEqual(actual.recheckCalls, []));
      },
    },
    {
      label: "distinct cross-role canonical paths sharing device/inode must reject before raw reads",
      expected: { recheck: "throws on cross-role identity alias", recheckCalls: [] },
      async run() {
        const compiled = compiledInput("src/compiled.ts", {
          device: "shared-device",
          inode: "shared-inode",
        });
        const controlSnapshot = rawSnapshot("src/control.ts", {
          contents: compiled.bytes,
          device: "shared-device",
          inode: "shared-inode",
        });
        const manifest = frozenManifest([{ path: "src/control.ts" }]);
        const bindingCalls: string[] = [];
        const binding = await observe(() => bindControlInputs(
          manifest,
          [compiled],
          async (relativePath) => {
            bindingCalls.push(relativePath);
            return controlSnapshot;
          },
        ));
        const controls = binding.outcome === "returned" ? binding.value : [];
        let enumerationCalls = 0;
        const recheckCalls: string[] = [];
        const snapshots: Record<string, RawProjectFileSnapshot> = {
          "src/compiled.ts": rawSnapshotFromCompiled(compiled),
          "src/control.ts": controlSnapshot,
        };
        const recheck = await observe(() => recheckPinnedSnapshots(
          PROJECT_ROOT,
          [compiled],
          controls,
          manifest,
          async () => {
            enumerationCalls += 1;
            return frozenManifest([{ path: "src/control.ts" }]);
          },
          async (relativePath) => {
            recheckCalls.push(relativePath);
            return snapshots[relativePath] ?? rawSnapshot(relativePath);
          },
        ));
        return {
          binding: summarizeObservation(binding, (rows) => rows.map(summarizeControl)),
          bindingCalls,
          compiledIdentity: { device: compiled.device, inode: compiled.inode },
          controlIdentity: { device: controlSnapshot.device, inode: controlSnapshot.inode },
          legacyOutcome: binding.outcome === "returned" && recheck.outcome === "returned"
            ? "unexpected cross-role inode alias acceptance"
            : "cross-role identity alias rejected",
          recheck: summarizeObservation(recheck),
          enumerationCalls,
          recheckCalls,
        };
      },
      matches(actual) {
        if (!isRecord(actual)) return false;
        return fieldOutcome(actual, "binding") === "threw" ||
          (fieldOutcome(actual, "binding") === "returned" &&
            fieldOutcome(actual, "recheck") === "threw" &&
            actual.enumerationCalls === 1 &&
            isDeepStrictEqual(actual.bindingCalls, ["src/control.ts"]) &&
            isDeepStrictEqual(actual.recheckCalls, []));
      },
    },
    {
      label: "later raw inspector device/inode mutation rejects after one inspection",
      expected: { recheck: "throws after one raw inspection", recheckCalls: ["src/compiled.ts"] },
      async run() {
        const compiled = compiledInput("src/compiled.ts");
        const expectedManifest = frozenManifest([]);
        const mutatedSnapshot = rawSnapshot("src/compiled.ts", {
          contents: compiled.bytes,
          device: "mutated-device",
          inode: "mutated-inode",
        });
        let enumerationCalls = 0;
        const recheckCalls: string[] = [];
        const recheck = await observe(() => recheckPinnedSnapshots(
          PROJECT_ROOT,
          [compiled],
          [],
          expectedManifest,
          async () => {
            enumerationCalls += 1;
            return frozenManifest([]);
          },
          async (relativePath) => {
            recheckCalls.push(relativePath);
            return mutatedSnapshot;
          },
        ));
        return {
          recheck: summarizeObservation(recheck),
          enumerationCalls,
          recheckCalls,
          expectedIdentity: { device: compiled.device, inode: compiled.inode },
          inspectedIdentity: { device: mutatedSnapshot.device, inode: mutatedSnapshot.inode },
        };
      },
      matches(actual) {
        return fieldOutcome(actual, "recheck") === "threw" &&
          isRecord(actual) &&
          isDeepStrictEqual(actual.recheckCalls, ["src/compiled.ts"]);
      },
    },
  ];

  await assertMatrixOutcomes("identity drift", variants);
});

void test("capture binding: byte drift", async () => {
  const variants: MatrixVariant[] = [
    {
      label: "overlapping compiled/control byte and hash drift rejects before raw reinspection",
      expected: {
        equality: "throws on byte/hash mismatch",
        recheck: "throws before raw inspection",
        recheckCalls: [],
      },
      async run() {
        const compiled = compiledInput("src/shared.ts");
        const changedBytes = Buffer.from(compiled.bytes);
        changedBytes[0] ^= 0xff;
        const controlSnapshot = rawSnapshot("src/shared.ts", { contents: changedBytes });
        const manifest = frozenManifest([{ path: "src/shared.ts" }]);
        const binding = await observe(() => bindControlInputs(
          manifest,
          [compiled],
          async () => controlSnapshot,
        ));
        const controls = binding.outcome === "returned" ? binding.value : [];
        const equality = await observe(() => {
          const control = controls[0];
          if (!control) throw new Error("Byte-drift fixture did not produce a bound control.");
          return assertControlSnapshotEqual(rawSnapshotFromCompiled(compiled), control);
        });
        let enumerationCalls = 0;
        const recheckCalls: string[] = [];
        const recheck = await observe(() => recheckPinnedSnapshots(
          PROJECT_ROOT,
          [compiled],
          controls,
          manifest,
          async () => {
            enumerationCalls += 1;
            return frozenManifest([{ path: "src/shared.ts" }]);
          },
          async (relativePath) => {
            recheckCalls.push(relativePath);
            return controlSnapshot;
          },
        ));
        return {
          binding: summarizeObservation(binding, (rows) => rows.map(summarizeControl)),
          equality: summarizeObservation(equality),
          recheck: summarizeObservation(recheck),
          enumerationCalls,
          recheckCalls,
          compiledDigest: compiled.sha256,
          controlDigest: controlSnapshot.sha256,
          compiledByteLength: compiled.byteLength,
          controlByteLength: controlSnapshot.byteLength,
        };
      },
      matches(actual) {
        return fieldOutcome(actual, "equality") === "threw" &&
          fieldOutcome(actual, "recheck") === "threw" &&
          isRecord(actual) &&
          isDeepStrictEqual(actual.recheckCalls, []);
      },
    },
    {
      label: "later raw inspector byte mutation rejects after one inspection",
      expected: { recheck: "throws after one raw inspection", recheckCalls: ["src/compiled.ts"] },
      async run() {
        const compiled = compiledInput("src/compiled.ts");
        const changedBytes = Buffer.from(compiled.bytes);
        changedBytes[0] ^= 0xff;
        const mutatedSnapshot = rawSnapshot("src/compiled.ts", { contents: changedBytes });
        const expectedManifest = frozenManifest([]);
        let enumerationCalls = 0;
        const recheckCalls: string[] = [];
        const recheck = await observe(() => recheckPinnedSnapshots(
          PROJECT_ROOT,
          [compiled],
          [],
          expectedManifest,
          async () => {
            enumerationCalls += 1;
            return frozenManifest([]);
          },
          async (relativePath) => {
            recheckCalls.push(relativePath);
            return mutatedSnapshot;
          },
        ));
        return {
          recheck: summarizeObservation(recheck),
          enumerationCalls,
          recheckCalls,
          compiledDigest: compiled.sha256,
          inspectedDigest: mutatedSnapshot.sha256,
          compiledByteLength: compiled.byteLength,
          inspectedByteLength: mutatedSnapshot.byteLength,
        };
      },
      matches(actual) {
        return fieldOutcome(actual, "recheck") === "threw" &&
          isRecord(actual) &&
          isDeepStrictEqual(actual.recheckCalls, ["src/compiled.ts"]);
      },
    },
  ];

  await assertMatrixOutcomes("byte drift", variants);
});

void test("capture binding: unknown control", async () => {
  const variants: MatrixVariant[] = [
    {
      label: "unknown per-row control role field rejects",
      expected: { binding: "throws on unknown role" },
      async run() {
        const manifest = [{
          path: "src/unknown-role.ts",
          categories: ["calculateBuildId"],
          role: "unknown-control-role",
        }] as unknown as ControlInputManifest;
        const inspectCalls: string[] = [];
        const binding = await observe(() => bindControlInputs(
          manifest,
          [],
          async (relativePath) => {
            inspectCalls.push(relativePath);
            return rawSnapshot(relativePath);
          },
        ));
        return {
          candidateRole: "unknown-control-role",
          binding: summarizeObservation(binding, (rows) => rows.map(summarizeControl)),
          inspectCalls,
        };
      },
      matches(actual) {
        return fieldOutcome(actual, "binding") === "threw";
      },
    },
    {
      label: "unknown per-row control category rejects",
      expected: { binding: "throws on unknown category" },
      async run() {
        const manifest = [{
          path: "src/unknown-category.ts",
          categories: ["unknown-control-category"],
        }] as unknown as ControlInputManifest;
        const inspectCalls: string[] = [];
        const binding = await observe(() => bindControlInputs(
          manifest,
          [],
          async (relativePath) => {
            inspectCalls.push(relativePath);
            return rawSnapshot(relativePath);
          },
        ));
        return {
          candidateCategories: ["unknown-control-category"],
          binding: summarizeObservation(binding, (rows) => rows.map(summarizeControl)),
          inspectCalls,
        };
      },
      matches(actual) {
        return fieldOutcome(actual, "binding") === "threw";
      },
    },
    {
      label: "noncanonical manifest spelling rejects despite a canonical injected snapshot",
      expected: { binding: "throws on noncanonical spec path" },
      async run() {
        const manifest = frozenManifest([{ path: "src/../control.ts" }]);
        const inspectCalls: string[] = [];
        const binding = await observe(() => bindControlInputs(
          manifest,
          [],
          async (relativePath) => {
            inspectCalls.push(relativePath);
            return rawSnapshot(relativePath);
          },
        ));
        return {
          candidatePath: manifest[0].path,
          binding: summarizeObservation(binding, (rows) => rows.map(summarizeControl)),
          inspectCalls,
        };
      },
      matches(actual) {
        return fieldOutcome(actual, "binding") === "threw";
      },
    },
    {
      label: "control-control inode alias rejects before raw recheck reads",
      expected: {
        binding: "C0 binds both distinct specs",
        recheck: "rejects alias before inspection",
        recheckCalls: [],
      },
      async run() {
        const manifest = frozenManifest([
          { path: "src/control-a.ts" },
          { path: "src/control-b.ts" },
        ]);
        const sharedBytes = Buffer.from("same aliased control bytes");
        const snapshots: Record<string, RawProjectFileSnapshot> = {
          "src/control-a.ts": rawSnapshot("src/control-a.ts", {
            contents: sharedBytes,
            device: "shared-device",
            inode: "shared-inode",
          }),
          "src/control-b.ts": rawSnapshot("src/control-b.ts", {
            contents: sharedBytes,
            device: "shared-device",
            inode: "shared-inode",
          }),
        };
        const bindingCalls: string[] = [];
        const binding = await observe(() => bindControlInputs(
          manifest,
          [],
          async (relativePath) => {
            bindingCalls.push(relativePath);
            return snapshots[relativePath] ?? rawSnapshot(relativePath);
          },
        ));
        const controls = binding.outcome === "returned" ? binding.value : [];
        let enumerationCalls = 0;
        const recheckCalls: string[] = [];
        const recheck = await observe(() => recheckPinnedSnapshots(
          PROJECT_ROOT,
          [],
          controls,
          manifest,
          async () => {
            enumerationCalls += 1;
            return frozenManifest([
              { path: "src/control-a.ts" },
              { path: "src/control-b.ts" },
            ]);
          },
          async (relativePath) => {
            recheckCalls.push(relativePath);
            return snapshots[relativePath] ?? rawSnapshot(relativePath);
          },
        ));
        return {
          binding: summarizeObservation(binding, (rows) => rows.map(summarizeControl)),
          bindingCalls,
          controlBindingOutcome: binding.outcome === "returned"
            ? "unexpected control inode alias acceptance"
            : "control inode alias rejected",
          recheck: summarizeObservation(recheck),
          enumerationCalls,
          recheckCalls,
          identities: Object.fromEntries(Object.entries(snapshots).map(([relativePath, snapshot]) => [
            relativePath,
            { canonicalPath: snapshot.canonicalPath, device: snapshot.device, inode: snapshot.inode },
          ])),
        };
      },
      matches(actual) {
        if (!isRecord(actual)) return false;
        const binding = isRecord(actual.binding) ? actual.binding : undefined;
        const rows = binding?.value;
        return fieldOutcome(actual, "binding") === "threw" ||
          (fieldOutcome(actual, "binding") === "returned" &&
            fieldOutcome(actual, "recheck") === "threw" &&
            actual.enumerationCalls === 1 &&
            Array.isArray(actual.bindingCalls) && actual.bindingCalls.length === 2 &&
            isDeepStrictEqual(actual.recheckCalls, []) &&
            Array.isArray(rows) && rows.length === 2);
      },
    },
  ];

  await assertMatrixOutcomes("unknown control", variants);
});

void test("capture binding: recheck", async () => {
  const variants: Array<{
    readonly label: string;
    readonly candidateRows: () => readonly { readonly path: string; readonly categories?: readonly ControlCategory[] }[];
    readonly expectedRejection: boolean;
  }> = [
    {
      label: "unchanged membership enumerates once and inspects one deduplicated path",
      candidateRows: () => [{ path: "src/control.ts" }],
      expectedRejection: false,
    },
    {
      label: "membership addition rejects before raw inspection",
      candidateRows: () => [{ path: "src/control.ts" }, { path: "src/added.ts" }],
      expectedRejection: true,
    },
    {
      label: "membership removal rejects before raw inspection",
      candidateRows: () => [],
      expectedRejection: true,
    },
    {
      label: "duplicate manifest member rejects before raw inspection",
      candidateRows: () => [{ path: "src/control.ts" }, { path: "src/control.ts" }],
      expectedRejection: true,
    },
    {
      label: "category drift rejects before raw inspection",
      candidateRows: () => [{ path: "src/control.ts", categories: ["project-package-lock-config-tsconfig"] }],
      expectedRejection: true,
    },
  ];
  const matrix: MatrixVariant[] = variants.map((variant) => ({
    label: variant.label,
    expected: variant.expectedRejection
      ? { recheck: "throws before raw inspection", enumerationCalls: 1, inspectCalls: [] }
      : { recheck: "returns after one deduplicated inspection", enumerationCalls: 1, inspectCalls: ["src/control.ts"] },
    async run() {
      const expectedManifest = frozenManifest([{ path: "src/control.ts" }]);
      const compiled = compiledInput("src/control.ts");
      const bindingCalls: string[] = [];
      const binding = await observe(() => bindControlInputs(
        expectedManifest,
        [compiled],
        async (relativePath) => {
          bindingCalls.push(relativePath);
          return rawSnapshot(relativePath);
        },
      ));
      const controls = binding.outcome === "returned" ? binding.value : [];
      let enumerationCalls = 0;
      const inspectCalls: string[] = [];
      const recheck = await observe(() => recheckPinnedSnapshots(
        PROJECT_ROOT,
        [compiled],
        controls,
        expectedManifest,
        async () => {
          enumerationCalls += 1;
          return frozenManifest(variant.candidateRows());
        },
        async (relativePath) => {
          inspectCalls.push(relativePath);
          return rawSnapshot(relativePath);
        },
      ));
      return {
        binding: summarizeObservation(binding, (rows) => rows.map(summarizeControl)),
        bindingCalls,
        recheck: summarizeObservation(recheck),
        enumerationCalls,
        inspectCalls,
      };
    },
    matches(actual) {
      if (!isRecord(actual) || actual.enumerationCalls !== 1) return false;
      if (variant.expectedRejection) {
        return fieldOutcome(actual, "recheck") === "threw" && isDeepStrictEqual(actual.inspectCalls, []);
      }
      return fieldOutcome(actual, "recheck") === "returned" &&
        isDeepStrictEqual(actual.inspectCalls, ["src/control.ts"]);
    },
  }));

  await assertMatrixOutcomes("recheck", matrix);
});

void test("capture binding: pinned loader", async () => {
  const variants: MatrixVariant[] = [
    {
      label: "pinned physical engine input returns its exact bytes, loader and resolve directory",
      expected: {
        operation: "returns the pinned TS binding",
        contents: "fixture:src/pinned.ts",
        loader: "ts",
        resolveDir: "/workspace/project/src",
      },
      async run() {
        const input = compiledInput("src/pinned.ts");
        const operation = await observe(() => pinnedPhysicalLoad(
          [input],
          { namespace: "file", path: input.canonicalPath, suffix: "", with: {} },
          "engine",
        ));
        const summary = summarizeObservation(operation, (loaded) => ({
          contents: Buffer.from(loaded.contents).toString("utf8"),
          loader: loaded.loader,
          resolveDir: loaded.resolveDir,
        }));
        return { operation: summary };
      },
      matches(actual) {
        if (!isRecord(actual) || !isRecord(actual.operation)) return false;
        const value = actual.operation.value;
        return actual.operation.outcome === "returned" && isRecord(value) &&
          value.contents === "fixture:src/pinned.ts" && value.loader === "ts" &&
          value.resolveDir === path.resolve(PROJECT_ROOT, "src");
      },
    },
    {
      label: "controls-only styles.css is unavailable to the physical loader",
      expected: { operation: "throws because controls are not in the physical binding map" },
      async run() {
        const operation = await observe(() => pinnedPhysicalLoad(
          [],
          { namespace: "file", path: path.resolve(PROJECT_ROOT, "styles.css"), suffix: "", with: {} },
          "main",
        ));
        return { operation: summarizeObservation(operation) };
      },
      matches(actual) { return fieldOutcome(actual, "operation") === "threw"; },
    },
    {
      label: "unknown load namespace rejects",
      expected: { operation: "throws on unknown namespace" },
      async run() {
        const input = compiledInput("src/pinned.ts");
        const operation = await observe(() => pinnedPhysicalLoad(
          [input],
          { namespace: "other", path: input.canonicalPath, suffix: "", with: {} },
          "engine",
        ));
        return { operation: summarizeObservation(operation) };
      },
      matches(actual) { return fieldOutcome(actual, "operation") === "threw"; },
    },
    {
      label: "nonempty file suffix rejects",
      expected: { operation: "throws on suffix" },
      async run() {
        const input = compiledInput("src/pinned.ts");
        const operation = await observe(() => pinnedPhysicalLoad(
          [input],
          { namespace: "file", path: input.canonicalPath, suffix: "?raw", with: {} },
          "engine",
        ));
        return { operation: summarizeObservation(operation) };
      },
      matches(actual) { return fieldOutcome(actual, "operation") === "threw"; },
    },
    {
      label: "nonempty import attributes reject",
      expected: { operation: "throws on import attributes" },
      async run() {
        const input = compiledInput("src/pinned.ts");
        const operation = await observe(() => pinnedPhysicalLoad(
          [input],
          { namespace: "file", path: input.canonicalPath, suffix: "", with: { type: "json" } },
          "engine",
        ));
        return { operation: summarizeObservation(operation) };
      },
      matches(actual) { return fieldOutcome(actual, "operation") === "threw"; },
    },
    {
      label: "unlisted canonical path rejects",
      expected: { operation: "throws on unpinned file path" },
      async run() {
        const input = compiledInput("src/pinned.ts");
        const operation = await observe(() => pinnedPhysicalLoad(
          [input],
          { namespace: "file", path: path.resolve(PROJECT_ROOT, "src/other.ts"), suffix: "", with: {} },
          "engine",
        ));
        return { operation: summarizeObservation(operation) };
      },
      matches(actual) { return fieldOutcome(actual, "operation") === "threw"; },
    },
    {
      label: "physical input absent from requested entry membership rejects",
      expected: { operation: "throws on wrong entry membership" },
      async run() {
        const input = compiledInput("src/pinned.ts");
        const operation = await observe(() => pinnedPhysicalLoad(
          [input],
          { namespace: "file", path: input.canonicalPath, suffix: "", with: {} },
          "main",
        ));
        return { operation: summarizeObservation(operation) };
      },
      matches(actual) { return fieldOutcome(actual, "operation") === "threw"; },
    },
  ];

  await assertMatrixOutcomes("pinned loader", variants);
});

function fakeEvidenceIo(options: { failOnPath?: string } = {}): EvidencePersistenceIo & {
  readonly events: string[];
  readonly files: Map<string, Buffer>;
} {
  const uid = process.getuid();
  const directories = new Map<string, { uid: number; mode: number }>([
    [EVIDENCE_PARENT, { uid, mode: 0o700 }],
    [EVIDENCE_ROOT, { uid, mode: 0o755 }],
  ]);
  const files = new Map<string, Buffer>();
  const events: string[] = [];
  const makeStats = (type: "directory" | "file", uidValue: number, mode: number) => ({
    uid: uidValue,
    mode: mode | (type === "directory" ? 0o040000 : 0o100000),
    isDirectory: () => type === "directory",
    isFile: () => type === "file",
    isSymbolicLink: () => false,
  });
  const io: EvidencePersistenceIo & { readonly events: string[]; readonly files: Map<string, Buffer> } = {
    events,
    files,
    async mkdir(target, optionsArg) {
      events.push(`mkdir:${target}:${optionsArg.mode}:${optionsArg.recursive}`);
      assert.equal(optionsArg.recursive, false);
      assert.equal(optionsArg.mode, 0o700);
      if (directories.has(target) || files.has(target)) throw Object.assign(new Error("exists"), { code: "EEXIST" });
      const parent = path.dirname(target);
      assert.ok(directories.has(parent), `missing fake parent ${parent}`);
      directories.set(target, { uid, mode: 0o700 });
    },
    async open(target, flags, mode) {
      events.push(`open:${target}:${flags}:${mode}`);
      assert.equal(flags, "wx");
      assert.equal(mode, 0o600);
      if (directories.has(target) || files.has(target)) throw Object.assign(new Error("exists"), { code: "EEXIST" });
      assert.ok(directories.has(path.dirname(target)), `missing fake file parent ${path.dirname(target)}`);
      let closed = false;
      return {
        async writeFile(data) {
          assert.equal(closed, false);
          events.push(`write:${target}`);
          if (options.failOnPath === target) throw new Error(`simulated write failure: ${target}`);
          files.set(target, Buffer.from(data));
        },
        async close() {
          closed = true;
          events.push(`close:${target}`);
        },
      };
    },
    async lstat(target) {
      events.push(`lstat:${target}`);
      if (directories.has(target)) {
        const item = directories.get(target)!;
        return makeStats("directory", item.uid, item.mode);
      }
      if (files.has(target)) return makeStats("file", uid, 0o600);
      throw Object.assign(new Error(`missing: ${target}`), { code: "ENOENT" });
    },
    async realpath(target) {
      events.push(`realpath:${target}`);
      if (!directories.has(target) && !files.has(target)) throw Object.assign(new Error(`missing: ${target}`), { code: "ENOENT" });
      return target;
    },
    async readdir(target) {
      events.push(`readdir:${target}`);
      assert.ok(directories.has(target));
      const children = new Set<string>();
      for (const candidate of [...directories.keys(), ...files.keys()]) {
        if (path.dirname(candidate) === target) children.add(path.basename(candidate));
      }
      return [...children].sort().map((name) => {
        const child = path.join(target, name);
        return {
          name,
          isDirectory: () => directories.has(child),
          isFile: () => files.has(child),
          isSymbolicLink: () => false,
        };
      });
    },
    async readFile(target) {
      events.push(`readFile:${target}`);
      const data = files.get(target);
      if (!data) throw Object.assign(new Error(`missing: ${target}`), { code: "ENOENT" });
      return Buffer.from(data);
    },
  };
  return io;
}

function preGuardFiles(): EvidencePersistencePayload {
  const fixture = persistedControlFixture();
  return persistedControlPayload(fixture, fixture.prepared.controlInputs);
}

function finalReceiptForPayload(payload: EvidencePersistencePayload, guardError: {
  readonly name: string;
  readonly message: string;
  readonly stack: string;
}): Buffer {
  const inputManifest = JSON.parse(Buffer.from(payload.files["inputs.json"]).toString("utf8")) as Record<string, unknown>;
  return Buffer.from(JSON.stringify({
    schemaVersion: 2,
    status: "guard failed / evidence only",
    workId: "W69",
    attempt: 3,
    root: EVIDENCE_ROOT,
    child: EVIDENCE_CHILD,
    buildId: inputManifest.buildId,
    actualInputManifestSha256: inputManifest.actualInputManifestSha256,
    guardError,
  }));
}

function evidenceIoEventSummary(io: ReturnType<typeof fakeEvidenceIo>): Record<string, unknown> {
  return {
    openBasenames: openBasenames(io.events),
    writeBasenames: io.events
      .filter((event) => event.startsWith("write:"))
      .map((event) => path.basename(event.slice("write:".length))),
    fileBasenames: [...io.files.keys()].map((file) => path.basename(file)).sort(),
  };
}

async function runFreshEvidenceScenario(
  operation: (
    persistence: ReturnType<typeof createEvidencePersistenceForTesting>,
    io: ReturnType<typeof fakeEvidenceIo>,
  ) => Promise<unknown>,
): Promise<unknown> {
  let io: ReturnType<typeof fakeEvidenceIo> | undefined;
  try {
    io = fakeEvidenceIo();
    const persistence = createEvidencePersistenceForTesting(io);
    return await operation(persistence, io);
  } catch (error) {
    return {
      outcome: "scenario-setup-threw",
      errorName: error instanceof Error ? error.name : typeof error,
      errorMessage: error instanceof Error ? error.message : String(error),
      ...(io ? evidenceIoEventSummary(io) : {}),
    };
  }
}

void test("capture binding: persistence failure retains partial files without cleanup", async () => {
  const failingPath = path.join(EVIDENCE_CHILD, "docx-engine.cjs");
  const io = fakeEvidenceIo({ failOnPath: failingPath });
  const persistence = createEvidencePersistenceForTesting(io);
  await persistence.createChild();
  const payload = preGuardFiles();

  const preGuard = await observe(() => persistence.persistBeforeGuard(payload));
  assert.equal(preGuard.outcome, "threw");
  if (preGuard.outcome === "threw") assert.match(preGuard.errorMessage, /simulated write failure/);

  const partialFileSnapshot = new Map<string, Buffer>(
    [...io.files].map(([file, bytes]) => [file, Buffer.from(bytes)] as const),
  );
  const guardError = {
    name: "Error",
    message: "A JSZip generation/deflate module is present in the runtime closure: node_modules/jszip/lib/crc32.js",
    stack: "guard stack",
  };
  const finalReceipt = finalReceiptForPayload(payload, guardError);
  const receiptAttempt = await observe(() => persistence.persistFinalReceipt(finalReceipt, guardError));
  const completeAttempt = await observe(() => persistence.writeComplete(guardError, async () => {}));
  const openOrder = openBasenames(io.events);
  const sortedFileEntries = (files: Map<string, Buffer>) =>
    [...files.entries()].sort(([left], [right]) => left.localeCompare(right));

  console.log("persistence failure finalization observations:\n" + JSON.stringify({
    preGuard: summarizeObservation(preGuard),
    receipt: summarizeObservation(receiptAttempt),
    complete: summarizeObservation(completeAttempt),
    receiptOpenAttempts: openOrder.filter((name) => name === "receipt.json").length,
    completeOpenAttempts: openOrder.filter((name) => name === "COMPLETE.json").length,
    partialFileBasenames: [...partialFileSnapshot.keys()].map((file) => path.basename(file)).sort(),
  }, null, 2));

  assert.equal(receiptAttempt.outcome, "threw");
  assert.equal(completeAttempt.outcome, "threw");
  assert.equal(openOrder.filter((name) => name === "receipt.json").length, 0);
  assert.equal(openOrder.filter((name) => name === "COMPLETE.json").length, 0);
  assert.ok(io.files.has(path.join(EVIDENCE_CHILD, "main.js")));
  assert.ok([...io.files.keys()].some((file) => file.startsWith(`${path.join(EVIDENCE_CHILD, "inputs")}${path.sep}`)));
  assert.deepEqual(sortedFileEntries(io.files), sortedFileEntries(partialFileSnapshot));
  assert.equal(io.files.has(path.join(EVIDENCE_CHILD, "receipt.json")), false);
  assert.equal(io.files.has(path.join(EVIDENCE_CHILD, "COMPLETE.json")), false);
  assert.equal(io.events.some((event) => /^(?:rm|unlink|rmdir):/.test(event)), false);
});

void test("capture binding: persistence-success aggregate", async () => {
  const guardError = {
    name: "Error",
    message: "A JSZip generation/deflate module is present in the runtime closure: node_modules/jszip/lib/crc32.js",
    stack: "guard stack",
  };
  const finalReceipt = finalReceiptForPayload(preGuardFiles(), guardError);
  const variants: MatrixVariant[] = [
    {
      label: "1 inventory immediately after persistBeforeGuard has no receipt",
      expected: { receiptPresent: false, receiptOpenAttempts: 0, persistBeforeGuard: "returned" },
      async run() {
        return runFreshEvidenceScenario(async (persistence, io) => {
          const createChild = await observe(() => persistence.createChild());
          const preGuard = await observe(() => persistence.persistBeforeGuard(preGuardFiles()));
          const events = evidenceIoEventSummary(io);
          const openOrder = events.openBasenames as string[];
          return {
            createChild: summarizeObservation(createChild),
            persistBeforeGuard: summarizeObservation(preGuard),
            receiptPresent: io.files.has(path.join(EVIDENCE_CHILD, "receipt.json")),
            receiptOpenAttempts: openOrder.filter((name) => name === "receipt.json").length,
            ...events,
          };
        });
      },
      matches(actual) {
        return isRecord(actual) && actual.receiptPresent === false &&
          actual.receiptOpenAttempts === 0 &&
          fieldOutcome(actual, "persistBeforeGuard") === "returned";
      },
    },
    {
      label: "2 final receipt before pre-guard persistence rejects without opening receipt",
      expected: { receipt: "throws on state", receiptOpenAttempts: 0, receiptPresent: false },
      async run() {
        return runFreshEvidenceScenario(async (persistence, io) => {
          const createChild = await observe(() => persistence.createChild());
          const receipt = await observe(() => persistence.persistFinalReceipt(finalReceipt, guardError));
          const events = evidenceIoEventSummary(io);
          const openOrder = events.openBasenames as string[];
          return {
            createChild: summarizeObservation(createChild),
            receipt: summarizeObservation(receipt),
            receiptOpenAttempts: openOrder.filter((name) => name === "receipt.json").length,
            receiptPresent: io.files.has(path.join(EVIDENCE_CHILD, "receipt.json")),
            ...events,
          };
        });
      },
      matches(actual) {
        return fieldOutcome(actual, "receipt") === "threw" && isRecord(actual) &&
          actual.receiptOpenAttempts === 0 && actual.receiptPresent === false;
      },
    },
    {
      label: "3 final receipt before guard-error rejects without a new receipt open",
      expected: { receipt: "throws on lifecycle state", newReceiptOpenAttempts: 0 },
      async run() {
        return runFreshEvidenceScenario(async (persistence, io) => {
          const createChild = await observe(() => persistence.createChild());
          const preGuard = await observe(() => persistence.persistBeforeGuard(preGuardFiles()));
          const before = evidenceIoEventSummary(io).openBasenames as string[];
          const receipt = await observe(() => persistence.persistFinalReceipt(finalReceipt, guardError));
          const after = evidenceIoEventSummary(io).openBasenames as string[];
          const events = evidenceIoEventSummary(io);
          return {
            createChild: summarizeObservation(createChild),
            persistBeforeGuard: summarizeObservation(preGuard),
            receipt: summarizeObservation(receipt),
            receiptOpenAttemptsBeforeFinalCall: before.filter((name) => name === "receipt.json").length,
            newReceiptOpenAttempts: after.filter((name) => name === "receipt.json").length -
              before.filter((name) => name === "receipt.json").length,
            ...events,
          };
        });
      },
      matches(actual) {
        return fieldOutcome(actual, "receipt") === "threw" && isRecord(actual) &&
          actual.newReceiptOpenAttempts === 0;
      },
    },
    {
      label: "4 COMPLETE before receipt/guard-error rejects without opening COMPLETE",
      expected: { complete: "throws before guard error", completeOpenAttempts: 0 },
      async run() {
        return runFreshEvidenceScenario(async (persistence, io) => {
          const createChild = await observe(() => persistence.createChild());
          const preGuard = await observe(() => persistence.persistBeforeGuard(preGuardFiles()));
          const complete = await observe(() => persistence.writeComplete(guardError, async () => {}));
          const events = evidenceIoEventSummary(io);
          const openOrder = events.openBasenames as string[];
          return {
            createChild: summarizeObservation(createChild),
            persistBeforeGuard: summarizeObservation(preGuard),
            complete: summarizeObservation(complete),
            completeOpenAttempts: openOrder.filter((name) => name === "COMPLETE.json").length,
            ...events,
          };
        });
      },
      matches(actual) {
        return fieldOutcome(actual, "complete") === "threw" && isRecord(actual) &&
          actual.completeOpenAttempts === 0;
      },
    },
    {
      label: "5 pre-guard then guard-error then final receipt then COMPLETE succeeds in exclusive order",
      expected: {
        preGuard: "returned",
        guardError: "returned",
        receipt: "returned",
        complete: "returned",
        openOrder: "guard-error.json < receipt.json < COMPLETE.json, with one receipt open",
      },
      async run() {
        return runFreshEvidenceScenario(async (persistence, io) => {
          const createChild = await observe(() => persistence.createChild());
          const preGuard = await observe(() => persistence.persistBeforeGuard(preGuardFiles()));
          const persistedGuardError = await observe(() => persistence.persistGuardError(guardError));
          const receipt = await observe(() => persistence.persistFinalReceipt(finalReceipt, guardError));
          const complete = await observe(() => persistence.writeComplete(guardError, async () => {}));
          const postCompleteGuardError = await observe(() => persistence.persistGuardError(guardError));
          const events = evidenceIoEventSummary(io);
          const openOrder = events.openBasenames as string[];
          const guardIndex = openOrder.indexOf("guard-error.json");
          const receiptIndices = openOrder.flatMap((name, index) => name === "receipt.json" ? [index] : []);
          const completeIndex = openOrder.indexOf("COMPLETE.json");
          const completeFileHashes = complete.outcome === "returned"
            ? complete.value.files.map((file) => {
              const bytes = io.files.get(path.join(EVIDENCE_CHILD, file.path));
              return {
                path: file.path,
                expectedSha256: file.sha256,
                actualSha256: bytes ? sha256(bytes) : null,
                matches: Boolean(bytes) && sha256(bytes!) === file.sha256,
              };
            })
            : [];
          const completeValue = complete.outcome === "returned"
            ? {
              status: complete.value.status,
              completeSha256: complete.value.completeSha256,
              priorFileCount: complete.value.files.length,
              excludesComplete: !complete.value.files.some((file) => file.path === "COMPLETE.json"),
            }
            : undefined;
          return {
            createChild: summarizeObservation(createChild),
            persistBeforeGuard: summarizeObservation(preGuard),
            guardError: summarizeObservation(persistedGuardError),
            receipt: summarizeObservation(receipt),
            complete: summarizeObservation(complete, () => completeValue),
            postCompleteGuardError: summarizeObservation(postCompleteGuardError),
            openOrder,
            receiptOpenAttempts: receiptIndices.length,
            completeFileHashes,
            completeOpenWasLast: completeIndex === openOrder.length - 1,
            writeOrder: events.writeBasenames,
            fileBasenames: events.fileBasenames,
            cleanupEvents: io.events.filter((event) => /^(?:rm|unlink|rmdir):/.test(event)),
          };
        });
      },
      matches(actual) {
        if (!isRecord(actual) || !Array.isArray(actual.openOrder)) return false;
        const openOrder = actual.openOrder;
        const guardIndex = openOrder.indexOf("guard-error.json");
        const receiptIndex = openOrder.indexOf("receipt.json");
        const completeIndex = openOrder.indexOf("COMPLETE.json");
        return fieldOutcome(actual, "persistBeforeGuard") === "returned" &&
          fieldOutcome(actual, "guardError") === "returned" &&
          fieldOutcome(actual, "receipt") === "returned" &&
          fieldOutcome(actual, "complete") === "returned" &&
          fieldOutcome(actual, "postCompleteGuardError") === "threw" &&
          actual.receiptOpenAttempts === 1 && guardIndex >= 0 &&
          receiptIndex > guardIndex && completeIndex > receiptIndex &&
          actual.completeOpenWasLast === true && Array.isArray(actual.cleanupEvents) &&
          actual.cleanupEvents.length === 0 &&
          isRecord(actual.complete) && isRecord(actual.complete.value) &&
          actual.complete.value.status === "guard failed / evidence only" &&
          actual.complete.value.excludesComplete === true &&
          Array.isArray(actual.completeFileHashes) &&
          actual.completeFileHashes.every((item) => isRecord(item) && item.matches === true);
      },
    },
    {
      label: "6 duplicate final receipt records first result and rejects second without opening",
      expected: {
        guardError: "returned",
        firstReceipt: "returned",
        secondReceipt: "throws on duplicate state",
        secondReceiptOpenAttempts: 0,
      },
      async run() {
        return runFreshEvidenceScenario(async (persistence, io) => {
          const createChild = await observe(() => persistence.createChild());
          const preGuard = await observe(() => persistence.persistBeforeGuard(preGuardFiles()));
          const persistedGuardError = await observe(() => persistence.persistGuardError(guardError));
          const beforeFirst = evidenceIoEventSummary(io).openBasenames as string[];
          const firstReceipt = await observe(() => persistence.persistFinalReceipt(finalReceipt, guardError));
          const afterFirst = evidenceIoEventSummary(io).openBasenames as string[];
          const secondReceipt = await observe(() => persistence.persistFinalReceipt(finalReceipt, guardError));
          const afterSecond = evidenceIoEventSummary(io).openBasenames as string[];
          const events = evidenceIoEventSummary(io);
          return {
            createChild: summarizeObservation(createChild),
            persistBeforeGuard: summarizeObservation(preGuard),
            guardError: summarizeObservation(persistedGuardError),
            firstReceipt: summarizeObservation(firstReceipt),
            secondReceipt: summarizeObservation(secondReceipt),
            firstReceiptOpenAttempts: afterFirst.filter((name) => name === "receipt.json").length -
              beforeFirst.filter((name) => name === "receipt.json").length,
            secondReceiptOpenAttempts: afterSecond.filter((name) => name === "receipt.json").length -
              afterFirst.filter((name) => name === "receipt.json").length,
            ...events,
          };
        });
      },
      matches(actual) {
        return fieldOutcome(actual, "guardError") === "returned" &&
          fieldOutcome(actual, "firstReceipt") === "returned" &&
          fieldOutcome(actual, "secondReceipt") === "threw" &&
          isRecord(actual) && actual.secondReceiptOpenAttempts === 0;
      },
    },
    {
      label: "7 COMPLETE after caught guard but before final receipt rejects without opening COMPLETE",
      expected: { complete: "throws until final receipt", completeOpenAttempts: 0 },
      async run() {
        return runFreshEvidenceScenario(async (persistence, io) => {
          const createChild = await observe(() => persistence.createChild());
          const preGuard = await observe(() => persistence.persistBeforeGuard(preGuardFiles()));
          const persistedGuardError = await observe(() => persistence.persistGuardError(guardError));
          const complete = await observe(() => persistence.writeComplete(guardError, async () => {}));
          const events = evidenceIoEventSummary(io);
          const openOrder = events.openBasenames as string[];
          return {
            createChild: summarizeObservation(createChild),
            persistBeforeGuard: summarizeObservation(preGuard),
            guardError: summarizeObservation(persistedGuardError),
            complete: summarizeObservation(complete, (value) => ({
              status: value.status,
              completeSha256: value.completeSha256,
              priorFileCount: value.files.length,
            })),
            completeOpenAttempts: openOrder.filter((name) => name === "COMPLETE.json").length,
            ...events,
          };
        });
      },
      matches(actual) {
        return fieldOutcome(actual, "complete") === "threw" && isRecord(actual) &&
          actual.completeOpenAttempts === 0;
      },
    },
    {
      label: "8 mismatched receipt stack rejects before open, then exact stack receipt succeeds and reads back",
      expected: {
        guardError: "returned",
        mismatchedReceipt: "throws on stack mismatch",
        mismatchReceiptOpenAttempts: 0,
        correctReceipt: "returned",
        readBackStack: guardError.stack,
        receiptOpenAttempts: 1,
      },
      async run() {
        return runFreshEvidenceScenario(async (persistence, io) => {
          const createChild = await observe(() => persistence.createChild());
          const preGuard = await observe(() => persistence.persistBeforeGuard(preGuardFiles()));
          const persistedGuardError = await observe(() => persistence.persistGuardError(guardError));
          const mismatchedReceiptValue = JSON.parse(finalReceipt.toString("utf8")) as Record<string, unknown>;
          const mismatchedStackReceipt = Buffer.from(JSON.stringify({
            ...mismatchedReceiptValue,
            guardError: { ...guardError, stack: "different guard stack" },
          }));
          const receiptOpensBeforeMismatch = openBasenames(io.events)
            .filter((name) => name === "receipt.json").length;
          const mismatchedReceipt = await observe(() =>
            persistence.persistFinalReceipt(mismatchedStackReceipt, guardError));
          const receiptOpensAfterMismatch = openBasenames(io.events)
            .filter((name) => name === "receipt.json").length;
          const correctReceipt = await observe(() => persistence.persistFinalReceipt(finalReceipt, guardError));
          const readBack = await observe(async () =>
            JSON.parse(Buffer.from(await io.readFile(path.join(EVIDENCE_CHILD, "receipt.json"))).toString("utf8")) as Record<string, unknown>);
          const readBackStack = readBack.outcome === "returned" && isRecord(readBack.value.guardError)
            ? readBack.value.guardError.stack
            : null;
          const events = evidenceIoEventSummary(io);
          const openOrder = events.openBasenames as string[];
          return {
            createChild: summarizeObservation(createChild),
            persistBeforeGuard: summarizeObservation(preGuard),
            guardError: summarizeObservation(persistedGuardError),
            mismatchedReceipt: summarizeObservation(mismatchedReceipt),
            mismatchReceiptOpenAttempts: receiptOpensAfterMismatch - receiptOpensBeforeMismatch,
            correctReceipt: summarizeObservation(correctReceipt),
            readBack: summarizeObservation(readBack),
            readBackStack,
            receiptOpenAttempts: openOrder.filter((name) => name === "receipt.json").length,
            receiptPresent: io.files.has(path.join(EVIDENCE_CHILD, "receipt.json")),
            ...events,
          };
        });
      },
      matches(actual) {
        return fieldOutcome(actual, "createChild") === "returned" &&
          fieldOutcome(actual, "persistBeforeGuard") === "returned" &&
          fieldOutcome(actual, "guardError") === "returned" &&
          fieldOutcome(actual, "mismatchedReceipt") === "threw" &&
          isRecord(actual) && actual.mismatchReceiptOpenAttempts === 0 &&
          fieldOutcome(actual, "correctReceipt") === "returned" &&
          fieldOutcome(actual, "readBack") === "returned" &&
          actual.readBackStack === guardError.stack &&
          actual.receiptOpenAttempts === 1 && actual.receiptPresent === true;
      },
    },
  ];

  await assertMatrixOutcomes("persistence-success aggregate", variants);
});

void test("capture attempt-3: retains complete evidence with expected ownership rejection", {
  skip: process.env.W69_CAPTURE_ATTEMPT_3 !== "1",
}, async () => {
  assert.equal(process.env.W69_CAPTURE_ATTEMPT_3, "1");
  const result = await captureDocxEvidenceAttempt3();
  assert.equal(result.status, "guard failed / evidence only");
  assert.equal(result.guardError.message, "A JSZip generation/deflate module is present in the runtime closure: node_modules/jszip/lib/crc32.js");
  assert.equal(result.complete, true);
});

interface PersistedControlFixture {
  readonly compiledInputs: readonly BoundMetafileInput[];
  readonly controls: readonly BoundControlInput[];
  readonly expectedManifest: ControlInputManifest;
  readonly bindingContext: NonNullable<EvidencePersistencePayload["bindingContext"]>;
  readonly manifestBytes: Buffer;
  readonly stylesBytes: Buffer;
  readonly prepared: ReturnType<typeof prepareEvidenceInputs>;
}

type PersistedControlRowFixture = Record<string, unknown>;

function persistedControlFixture(options: {
  readonly controlPath?: string;
  readonly compiledPath?: string;
  readonly categories?: readonly ControlCategory[];
} = {}): PersistedControlFixture {
  const controlPath = options.controlPath ?? "src/shared.ts";
  const compiledPath = options.compiledPath ?? controlPath;
  const categories: readonly ControlCategory[] = options.categories ?? (controlPath === "src/shared.ts"
    ? ["calculateBuildId"]
    : ["project-package-lock-config-tsconfig"]);
  const compiled = compiledInput(compiledPath);
  const snapshot = rawSnapshot(controlPath);
  const control: BoundControlInput = Object.freeze({
    ...snapshot,
    path: controlPath,
    categories: Object.freeze([...categories]),
    role: "control",
    compiledKeys: Object.freeze(controlPath === compiledPath ? [compiled.key] : []),
  });
  const expectedManifest = frozenManifest([{ path: controlPath, categories }]);
  const compiledInputs = Object.freeze([compiled]);
  const controls = Object.freeze([control]);
  const manifestBytes = Buffer.from("independent manifest fixture bytes", "utf8");
  const stylesBytes = Buffer.from("independent styles fixture bytes", "utf8");
  const prepared = prepareEvidenceInputs(compiledInputs, controls, manifestBytes, stylesBytes);
  const bindingContext = Object.freeze({ expectedManifest, compiledInputs, controls });
  return {
    compiledInputs,
    controls,
    expectedManifest,
    bindingContext,
    manifestBytes,
    stylesBytes,
    prepared,
  };
}

function serializedControlFixture(control: BoundControlInput): PersistedControlRowFixture {
  return {
    role: "control",
    path: control.path,
    canonicalPath: control.canonicalPath,
    device: control.device,
    inode: control.inode,
    byteLength: control.byteLength,
    sha256: control.sha256,
    categories: [...control.categories],
    compiledKeys: [...control.compiledKeys],
    blob: `inputs/${control.sha256}.bin`,
  };
}

function persistedControlPayload(
  fixture: PersistedControlFixture,
  candidateControlRows: readonly unknown[],
  candidateCompiledRows: readonly unknown[] = fixture.prepared.inputs,
): EvidencePersistencePayload {
  const blobInventory = fixture.prepared.blobs.map((blob) => ({
    path: `inputs/${blob.sha256}.bin`,
    sha256: blob.sha256,
    byteLength: blob.bytes.byteLength,
  }));
  const inputsJson = Buffer.from(JSON.stringify({
    schemaVersion: 2,
    workId: "W69",
    attempt: 3,
    root: PROJECT_ROOT,
    child: EVIDENCE_CHILD,
    buildId: "memory-build",
    inputs: candidateCompiledRows,
    controlInputs: candidateControlRows,
    actualInputManifestSha256: sha256(Buffer.from(stableJson({
      inputs: candidateCompiledRows,
      controlInputs: candidateControlRows,
    }))),
    blobInventory,
  }), "utf8");
  const files: EvidencePreGuardFiles = {
    "main.js": Buffer.from("main fixture", "utf8"),
    "docx-engine.cjs": Buffer.from("engine fixture", "utf8"),
    "main.metafile.json": Buffer.from("{}\n", "utf8"),
    "engine.metafile.json": Buffer.from("{}\n", "utf8"),
    "inputs.json": inputsJson,
  };
  return {
    files,
    blobs: fixture.prepared.blobs,
    bindingContext: fixture.bindingContext,
  };
}

async function executePersistedControlPayload(
  payload: EvidencePersistencePayload,
): Promise<{
  readonly outcome: Record<string, unknown>;
  readonly files: Map<string, Buffer>;
  readonly events: string[];
}> {
  const io = fakeEvidenceIo();
  const persistence = createEvidencePersistenceForTesting(io);
  const createChild = await observe(() => persistence.createChild());
  const persistBeforeGuard = await observe(() => persistence.persistBeforeGuard(payload));
  return {
    outcome: {
      createChild: summarizeObservation(createChild),
      persistBeforeGuard: summarizeObservation(persistBeforeGuard),
      ...evidenceIoEventSummary(io),
    },
    files: io.files,
    events: io.events,
  };
}

interface PersistedControlNegativeVariant {
  readonly label: string;
  readonly categories?: readonly ControlCategory[];
  readonly rows: (
    fixture: PersistedControlFixture,
    validRow: PersistedControlRowFixture,
  ) => readonly unknown[];
}

async function assertPersistedControlRejectionMatrix(
  name: string,
  variants: readonly PersistedControlNegativeVariant[],
): Promise<void> {
  const settled = await Promise.allSettled(variants.map(async (variant) => {
    const fixture = persistedControlFixture({ categories: variant.categories });
    const validRow = serializedControlFixture(fixture.controls[0]);
    const payload = persistedControlPayload(fixture, variant.rows(fixture, validRow));
    const execution = await executePersistedControlPayload(payload);
    return execution.outcome;
  }));
  const records = variants.map((variant, index) => {
    const result = settled[index];
    const actual = result.status === "fulfilled"
      ? result.value
      : {
        outcome: "scenario-threw",
        errorName: result.reason instanceof Error ? result.reason.name : typeof result.reason,
        errorMessage: result.reason instanceof Error ? result.reason.message : String(result.reason),
      };
    const matched = fieldOutcome(actual, "createChild") === "returned" &&
      fieldOutcome(actual, "persistBeforeGuard") === "threw";
    return {
      label: variant.label,
      expected: "persisted-control rejection",
      actual,
      matched,
    };
  });
  const mismatches = records.filter((record) => !record.matched);
  const accepted = records
    .filter((record) => fieldOutcome(record.actual, "persistBeforeGuard") === "returned")
    .map((record) => record.label);
  console.log(`${name} persisted-control outcomes:\n${JSON.stringify(records, null, 2)}`);
  const acceptedDiagnostic = accepted.map((label) =>
    `expected persisted-control rejection, observed acceptance: ${label}`,
  );
  const otherDiagnostics = mismatches
    .filter((record) => fieldOutcome(record.actual, "persistBeforeGuard") !== "returned")
    .map((record) => `expected persisted-control rejection after fresh child creation: ${record.label}`);
  assert.equal(
    mismatches.length,
    0,
    `${name}: ${[...acceptedDiagnostic, ...otherDiagnostics].join("; ")}\nAll outcomes:\n${JSON.stringify(records, null, 2)}`,
  );
}

function exactRecordMismatches(
  label: string,
  actual: unknown,
  expected: Record<string, unknown>,
): string[] {
  if (!isRecord(actual)) return [`${label}: expected a record, received ${JSON.stringify(actual)}`];
  const mismatches: string[] = [];
  for (const key of Object.keys(expected).sort()) {
    if (!(key in actual)) mismatches.push(`${label}.${key}: missing`);
    else if (!isDeepStrictEqual(actual[key], expected[key])) {
      mismatches.push(`${label}.${key}: expected ${JSON.stringify(expected[key])}, received ${JSON.stringify(actual[key])}`);
    }
  }
  for (const key of Object.keys(actual).sort()) {
    if (!(key in expected)) mismatches.push(`${label}.${key}: unexpected field`);
  }
  return mismatches;
}

void test("capture binding: persisted extra field", async () => {
  await assertPersistedControlRejectionMatrix("persisted extra field", [{
    label: "unknown control field",
    rows: (_fixture, row) => [{ ...row, unexpectedField: "not in schema" }],
  }]);
});

void test("capture binding: persisted field types", async () => {
  await assertPersistedControlRejectionMatrix("persisted field types", [
    {
      label: "missing device field",
      rows: (_fixture, row) => {
        const candidate = { ...row };
        delete candidate.device;
        return [candidate];
      },
    },
    {
      label: "numeric canonicalPath field",
      rows: (_fixture, row) => [{ ...row, canonicalPath: 7 }],
    },
    {
      label: "string category field",
      rows: (_fixture, row) => [{ ...row, categories: "calculateBuildId" }],
    },
    {
      label: "string compiledKeys field",
      rows: (_fixture, row) => [{ ...row, compiledKeys: "src/shared.ts" }],
    },
  ]);
});

void test("capture binding: persisted role", async () => {
  await assertPersistedControlRejectionMatrix("persisted role", [{
    label: "compiled role in control row",
    rows: (_fixture, row) => [{ ...row, role: "compiled" }],
  }]);
});

void test("capture binding: persisted path", async () => {
  await assertPersistedControlRejectionMatrix("persisted path", [
    {
      label: "noncanonical control path",
      rows: (_fixture, row) => [{ ...row, path: "src/../src/shared.ts" }],
    },
    {
      label: "outside control path",
      rows: (_fixture, row) => [{ ...row, path: "../outside.ts" }],
    },
    {
      label: "wrong expected control path",
      rows: (_fixture, row) => [{ ...row, path: "src/other.ts" }],
    },
  ]);
});

void test("capture binding: persisted categories", async () => {
  const expectedCategories = ["calculateBuildId", "project-package-lock-config-tsconfig"] as const;
  await assertPersistedControlRejectionMatrix("persisted categories", [
    {
      label: "unknown control category",
      rows: (_fixture, row) => [{ ...row, categories: ["unknown-category"] }],
    },
    {
      label: "duplicate control category",
      rows: (_fixture, row) => [{ ...row, categories: ["calculateBuildId", "calculateBuildId"] }],
    },
    {
      label: "unsorted known categories",
      categories: expectedCategories,
      rows: (_fixture, row) => [{
        ...row,
        categories: [...expectedCategories].reverse(),
      }],
    },
  ]);
});

void test("capture binding: persisted size", async () => {
  await assertPersistedControlRejectionMatrix("persisted size", [
    {
      label: "negative control size",
      rows: (_fixture, row) => [{ ...row, byteLength: -1 }],
    },
    {
      label: "fractional control size",
      rows: (_fixture, row) => [{ ...row, byteLength: 0.5 }],
    },
    {
      label: "unsafe control size",
      rows: (_fixture, row) => [{ ...row, byteLength: Number.MAX_SAFE_INTEGER + 1 }],
    },
  ]);
});

void test("capture binding: persisted digest", async () => {
  await assertPersistedControlRejectionMatrix("persisted digest", [
    {
      label: "malformed control digest",
      rows: (_fixture, row) => [{ ...row, sha256: "not-a-sha256" }],
    },
    {
      label: "wrong well-formed control digest",
      rows: (_fixture, row) => [{ ...row, sha256: "0".repeat(64) }],
    },
  ]);
});

void test("capture binding: persisted blob reference", async () => {
  await assertPersistedControlRejectionMatrix("persisted blob reference", [
    {
      label: "wrong reference to a different present blob",
      rows: (fixture, row) => [{ ...row, blob: `inputs/${sha256(fixture.manifestBytes)}.bin` }],
    },
    {
      label: "unreferenced valid-shaped blob path",
      rows: (_fixture, row) => [{ ...row, blob: `inputs/${sha256(Buffer.from("not in the payload", "utf8"))}.bin` }],
    },
  ]);
});

void test("capture binding: persisted compiled keys", async () => {
  await assertPersistedControlRejectionMatrix("persisted compiled keys", [
    {
      label: "nonexistent compiled key",
      rows: (_fixture, row) => [{ ...row, compiledKeys: ["src/missing.ts"] }],
    },
    {
      label: "duplicate compiled key",
      rows: (_fixture, row) => [{ ...row, compiledKeys: ["src/shared.ts", "src/shared.ts"] }],
    },
    {
      label: "wrong-path compiled key",
      rows: (_fixture, row) => [{ ...row, compiledKeys: ["/workspace/project/src/shared.ts"] }],
    },
  ]);
});

void test("capture binding: persisted cross-role identity", async () => {
  await assertPersistedControlRejectionMatrix("persisted cross-role identity", [
    {
      label: "device differs from pinned compiled identity",
      rows: (_fixture, row) => [{ ...row, device: "different-device" }],
    },
    {
      label: "inode differs from pinned compiled identity",
      rows: (_fixture, row) => [{ ...row, inode: "different-inode" }],
    },
  ]);
});

void test("capture binding: persisted cross-role bytes", async () => {
  await assertPersistedControlRejectionMatrix("persisted cross-role bytes", [
    {
      label: "safe control size differs from pinned compiled size",
      rows: (_fixture, row) => [{ ...row, byteLength: Number(row.byteLength) + 1 }],
    },
    {
      label: "control digest differs from pinned compiled digest",
      rows: (_fixture, row) => [{ ...row, sha256: "e".repeat(64) }],
    },
    {
      label: "control blob resolves to bytes different from pinned compiled bytes",
      rows: (fixture, row) => [{ ...row, blob: `inputs/${sha256(fixture.manifestBytes)}.bin` }],
    },
  ]);
});

void test("capture binding: persisted membership", async () => {
  await assertPersistedControlRejectionMatrix("persisted membership", [
    {
      label: "missing control member",
      rows: () => [],
    },
    {
      label: "extra control member",
      rows: (fixture, row) => {
        const extraSnapshot = rawSnapshot("src/extra.ts", {
          contents: fixture.controls[0].contents,
          device: "extra-device",
          inode: "extra-inode",
        });
        const extraControl: BoundControlInput = {
          ...extraSnapshot,
          path: "src/extra.ts",
          categories: Object.freeze(["calculateBuildId"]),
          role: "control",
          compiledKeys: Object.freeze([]),
        };
        return [row, serializedControlFixture(extraControl)];
      },
    },
    {
      label: "duplicate control member",
      rows: (_fixture, row) => [row, { ...row }],
    },
  ]);
});

void test("capture binding: persisted all-control blob retention", async () => {
  const fixture = persistedControlFixture({
    controlPath: "package.json",
    compiledPath: "src/compiled.ts",
    categories: ["project-package-lock-config-tsconfig"],
  });
  const control = fixture.controls[0];
  const digest = control.sha256;
  const payload = persistedControlPayload(fixture, [serializedControlFixture(control)]);
  const execution = await executePersistedControlPayload(payload);
  const preparedBlob = fixture.prepared.blobs.find((blob) => blob.sha256 === digest);
  const persistedBlob = execution.files.get(path.join(EVIDENCE_CHILD, "inputs", `${digest}.bin`));
  const preparedHasBytes = Boolean(preparedBlob) && Buffer.from(preparedBlob!.bytes).equals(Buffer.from(control.contents));
  const persistedHasBytes = Boolean(persistedBlob) && persistedBlob!.equals(Buffer.from(control.contents));
  const mismatches = [
    ...(!preparedHasBytes ? ["control blob missing from prepared blobs"] : []),
    ...(!persistedHasBytes ? ["control blob missing from persisted files"] : []),
    ...(fieldOutcome(execution.outcome, "createChild") !== "returned" ? ["fake child creation failed"] : []),
    ...(fieldOutcome(execution.outcome, "persistBeforeGuard") !== "returned" ? ["persistBeforeGuard did not return"] : []),
  ];
  const outcomes = {
    digest,
    preparedBlobPresent: Boolean(preparedBlob),
    preparedHasBytes,
    persistedBlobPresent: Boolean(persistedBlob),
    persistedHasBytes,
    persistence: execution.outcome,
  };
  console.log(`persisted all-control blob retention outcomes:\n${JSON.stringify(outcomes, null, 2)}`);
  assert.equal(mismatches.length, 0, `control blob missing: ${JSON.stringify({ mismatches, outcomes }, null, 2)}`);
});

void test("capture binding: persisted serialized roles and links", async () => {
  const fixture = persistedControlFixture();
  const compiled = fixture.compiledInputs[0];
  const control = fixture.controls[0];
  const expectedCompiledRow: Record<string, unknown> = {
    role: "compiled",
    namespace: compiled.namespace,
    canonicalPath: compiled.canonicalPath,
    key: compiled.key,
    loader: compiled.loader,
    resolveDir: compiled.resolveDir,
    byteLength: compiled.byteLength,
    sha256: compiled.sha256,
    device: compiled.device,
    inode: compiled.inode,
    entries: [...compiled.entries],
    blob: `inputs/${compiled.sha256}.bin`,
  };
  const expectedControlRow = serializedControlFixture(control);
  const preparedCompiledRow = fixture.prepared.inputs[0];
  const preparedControlRow = fixture.prepared.controlInputs[0];
  const projection = fixture.prepared.controlInputs.map((row) => isRecord(row)
    ? { path: row.path, categories: row.categories }
    : row);
  const payload = persistedControlPayload(fixture, fixture.prepared.controlInputs);
  const execution = await executePersistedControlPayload(payload);
  const mismatches = [
    ...exactRecordMismatches("compiled row", preparedCompiledRow, expectedCompiledRow),
    ...exactRecordMismatches("control row", preparedControlRow, expectedControlRow),
    ...(!isDeepStrictEqual(projection, fixture.expectedManifest)
      ? [`control membership projection differs: expected ${JSON.stringify(fixture.expectedManifest)}, received ${JSON.stringify(projection)}`]
      : []),
    ...(fieldOutcome(execution.outcome, "createChild") !== "returned" ? ["fake child creation failed"] : []),
    ...(fieldOutcome(execution.outcome, "persistBeforeGuard") !== "returned" ? ["persistBeforeGuard did not return"] : []),
  ];
  const outcomes = {
    expectedCompiledRow,
    preparedCompiledRow,
    expectedControlRow,
    preparedControlRow,
    expectedMembership: fixture.expectedManifest,
    projectedMembership: projection,
    persistence: execution.outcome,
    mismatches,
  };
  console.log(`persisted serialized roles and links outcomes:\n${JSON.stringify(outcomes, null, 2)}`);
  assert.equal(mismatches.length, 0, `serialized role/link fields differ: ${JSON.stringify(outcomes, null, 2)}`);
});

void test("capture binding: persisted single blob for shared digest", async () => {
  const fixture = persistedControlFixture();
  const compiled = fixture.compiledInputs[0];
  const control = fixture.controls[0];
  const digest = compiled.sha256;
  const payload = persistedControlPayload(fixture, fixture.prepared.controlInputs);
  const execution = await executePersistedControlPayload(payload);
  const inputsJsonBytes = execution.files.get(path.join(EVIDENCE_CHILD, "inputs.json"));
  const inputsJson = inputsJsonBytes
    ? JSON.parse(inputsJsonBytes.toString("utf8")) as Record<string, unknown>
    : null;
  const blobInventory = isRecord(inputsJson) && Array.isArray(inputsJson.blobInventory)
    ? inputsJson.blobInventory
    : [];
  const digestInventoryCount = blobInventory.filter((item) => isRecord(item) && item.sha256 === digest).length;
  const blobPath = path.join(EVIDENCE_CHILD, "inputs", `${digest}.bin`);
  const exclusiveOpenCount = execution.events.filter((event) => event.startsWith(`open:${blobPath}:wx:`)).length;
  const persistedBlob = execution.files.get(blobPath);
  const byteEquality = Boolean(persistedBlob) &&
    persistedBlob!.equals(Buffer.from(compiled.bytes)) &&
    persistedBlob!.equals(Buffer.from(control.contents));
  const preparedDigestCount = fixture.prepared.blobs.filter((blob) => blob.sha256 === digest).length;
  const mismatches = [
    ...(preparedDigestCount !== 1 ? [`prepared digest member count is ${preparedDigestCount}, expected 1`] : []),
    ...(digestInventoryCount !== 1 ? [`blob inventory digest count is ${digestInventoryCount}, expected 1`] : []),
    ...(exclusiveOpenCount !== 1 ? [`exclusive blob open count is ${exclusiveOpenCount}, expected 1`] : []),
    ...(!byteEquality ? ["persisted shared blob bytes differ from compiled/control bytes"] : []),
    ...(fieldOutcome(execution.outcome, "createChild") !== "returned" ? ["fake child creation failed"] : []),
    ...(fieldOutcome(execution.outcome, "persistBeforeGuard") !== "returned" ? ["persistBeforeGuard did not return"] : []),
  ];
  const outcomes = {
    digest,
    preparedDigestCount,
    digestInventoryCount,
    exclusiveOpenCount,
    byteEquality,
    persistence: execution.outcome,
    mismatches,
  };
  console.log(`persisted single blob for shared digest outcomes:\n${JSON.stringify(outcomes, null, 2)}`);
  assert.equal(mismatches.length, 0, `single blob for shared digest mismatches: ${JSON.stringify(outcomes, null, 2)}`);
});

void test("capture binding: preserved attempt-2 lifecycle is exact", async () => {
  const attempt2Directory = path.join(
    EVIDENCE_ROOT,
    "obsidian-word-reader-W69-S5.3-evidence-attempt-2",
  );
  const expectedAttempt2Files = [
    {
      path: "guard-error.json",
      byteLength: 1469,
      sha256: "85a8b2555598836ba10fd328ec9ade79c6023e7e1b1c4d1ae366735431224b4d",
    },
    {
      path: "receipt.json",
      byteLength: 1504,
      sha256: "a8d87b311ee9f98b5d469d7a471c0fe1074510578bbcd04d4d954d1283c681ea",
    },
  ] as const;
  type EntryKind = "directory" | "file" | "symlink";
  interface ObservedStat {
    size: number;
    uid: number;
    gid: number;
    mode: number;
    kind: EntryKind;
  }
  interface ObservedFile {
    stat: ObservedStat;
    canonicalPath: string;
    fingerprint: { byteLength: number; sha256: string };
    readError?: Error;
  }
  interface Attempt2Observations {
    directory: ObservedStat;
    directoryCanonicalPath: string;
    entries: Array<{ name: string; kind: EntryKind }>;
    files: Map<string, ObservedFile>;
  }
  const makeObservations = (): Attempt2Observations => ({
    directory: { size: 128, uid: 501, gid: 20, mode: 0o40700, kind: "directory" },
    directoryCanonicalPath: attempt2Directory,
    entries: expectedAttempt2Files.map(({ path: name }) => ({ name, kind: "file" })),
    files: new Map(expectedAttempt2Files.map((expected) => [expected.path, {
      stat: {
        size: expected.byteLength,
        uid: 501,
        gid: 20,
        mode: 0o100600,
        kind: "file" as const,
      },
      canonicalPath: path.join(attempt2Directory, expected.path),
      fingerprint: { byteLength: expected.byteLength, sha256: expected.sha256 },
    }] as const)),
  });
  const statView = (observed: ObservedStat) => ({
    size: observed.size,
    uid: observed.uid,
    gid: observed.gid,
    mode: observed.mode,
    isDirectory: () => observed.kind === "directory",
    isFile: () => observed.kind === "file",
    isSymbolicLink: () => observed.kind === "symlink",
  });
  const makeIo = (observations: Attempt2Observations): PreservedAttemptReadIo => ({
    async lstat(target) {
      if (target === attempt2Directory) return statView(observations.directory);
      const name = path.basename(target);
      const file = observations.files.get(name);
      if (!file || path.dirname(target) !== attempt2Directory) throw new Error(`unexpected fake lstat: ${target}`);
      return statView(file.stat);
    },
    async realpath(target) {
      if (target === attempt2Directory) return observations.directoryCanonicalPath;
      const file = observations.files.get(path.basename(target));
      if (!file || path.dirname(target) !== attempt2Directory) throw new Error(`unexpected fake realpath: ${target}`);
      return file.canonicalPath;
    },
    async readdir(target) {
      if (target !== attempt2Directory) throw new Error(`unexpected fake readdir: ${target}`);
      return observations.entries.map((entry) => ({
        name: entry.name,
        isDirectory: () => entry.kind === "directory",
        isFile: () => entry.kind === "file",
        isSymbolicLink: () => entry.kind === "symlink",
      }));
    },
    async readFileFingerprint(target) {
      const file = observations.files.get(path.basename(target));
      if (!file || path.dirname(target) !== attempt2Directory) throw new Error(`unexpected fake fingerprint: ${target}`);
      if (file.readError) throw file.readError;
      return { ...file.fingerprint };
    },
  });
  const expectedFingerprintResult = expectedAttempt2Files;
  const rejected = { outcome: "threw" };
  const rejects = (actual: unknown): boolean => isRecord(actual) && actual.outcome === "threw";
  const variants: MatrixVariant[] = [
    {
      label: "baseline: exact sorted inventory, metadata, sizes and independent hashes",
      expected: { outcome: "returned", value: expectedFingerprintResult },
      async run() {
        return summarizeObservation(await observe(() => verifyAttempt2Unchanged(makeIo(makeObservations()))));
      },
    },
    {
      label: "missing guard-error.json",
      expected: rejected,
      async run() {
        const observations = makeObservations();
        observations.entries = observations.entries.filter((entry) => entry.name !== "guard-error.json");
        return summarizeObservation(await observe(() => verifyAttempt2Unchanged(makeIo(observations))));
      },
      matches: rejects,
    },
    {
      label: "extra COMPLETE.json",
      expected: rejected,
      async run() {
        const observations = makeObservations();
        observations.entries.push({ name: "COMPLETE.json", kind: "file" });
        return summarizeObservation(await observe(() => verifyAttempt2Unchanged(makeIo(observations))));
      },
      matches: rejects,
    },
    {
      label: "extra inputs directory",
      expected: rejected,
      async run() {
        const observations = makeObservations();
        observations.entries.push({ name: "inputs", kind: "directory" });
        return summarizeObservation(await observe(() => verifyAttempt2Unchanged(makeIo(observations))));
      },
      matches: rejects,
    },
    {
      label: "attempt-2 directory wrong type",
      expected: rejected,
      async run() {
        const observations = makeObservations();
        observations.directory.kind = "file";
        return summarizeObservation(await observe(() => verifyAttempt2Unchanged(makeIo(observations))));
      },
      matches: rejects,
    },
    {
      label: "attempt-2 directory wrong UID",
      expected: rejected,
      async run() {
        const observations = makeObservations();
        observations.directory.uid = 502;
        return summarizeObservation(await observe(() => verifyAttempt2Unchanged(makeIo(observations))));
      },
      matches: rejects,
    },
    {
      label: "attempt-2 directory wrong GID",
      expected: rejected,
      async run() {
        const observations = makeObservations();
        observations.directory.gid = 21;
        return summarizeObservation(await observe(() => verifyAttempt2Unchanged(makeIo(observations))));
      },
      matches: rejects,
    },
    {
      label: "attempt-2 directory wrong mode",
      expected: rejected,
      async run() {
        const observations = makeObservations();
        observations.directory.mode = 0o40710;
        return summarizeObservation(await observe(() => verifyAttempt2Unchanged(makeIo(observations))));
      },
      matches: rejects,
    },
    {
      label: "attempt-2 directory symlink",
      expected: rejected,
      async run() {
        const observations = makeObservations();
        observations.directory.kind = "symlink";
        return summarizeObservation(await observe(() => verifyAttempt2Unchanged(makeIo(observations))));
      },
      matches: rejects,
    },
    {
      label: "attempt-2 directory noncanonical path",
      expected: rejected,
      async run() {
        const observations = makeObservations();
        observations.directoryCanonicalPath = `${attempt2Directory}/different`;
        return summarizeObservation(await observe(() => verifyAttempt2Unchanged(makeIo(observations))));
      },
      matches: rejects,
    },
    {
      label: "guard-error.json dirent has wrong type",
      expected: rejected,
      async run() {
        const observations = makeObservations();
        observations.entries[0].kind = "directory";
        return summarizeObservation(await observe(() => verifyAttempt2Unchanged(makeIo(observations))));
      },
      matches: rejects,
    },
    {
      label: "guard-error.json stat has wrong type",
      expected: rejected,
      async run() {
        const observations = makeObservations();
        observations.files.get("guard-error.json")!.stat.kind = "directory";
        return summarizeObservation(await observe(() => verifyAttempt2Unchanged(makeIo(observations))));
      },
      matches: rejects,
    },
    {
      label: "guard-error.json wrong UID",
      expected: rejected,
      async run() {
        const observations = makeObservations();
        observations.files.get("guard-error.json")!.stat.uid = 502;
        return summarizeObservation(await observe(() => verifyAttempt2Unchanged(makeIo(observations))));
      },
      matches: rejects,
    },
    {
      label: "guard-error.json wrong GID",
      expected: rejected,
      async run() {
        const observations = makeObservations();
        observations.files.get("guard-error.json")!.stat.gid = 21;
        return summarizeObservation(await observe(() => verifyAttempt2Unchanged(makeIo(observations))));
      },
      matches: rejects,
    },
    {
      label: "guard-error.json wrong mode",
      expected: rejected,
      async run() {
        const observations = makeObservations();
        observations.files.get("guard-error.json")!.stat.mode = 0o100640;
        return summarizeObservation(await observe(() => verifyAttempt2Unchanged(makeIo(observations))));
      },
      matches: rejects,
    },
    {
      label: "guard-error.json symlink",
      expected: rejected,
      async run() {
        const observations = makeObservations();
        observations.files.get("guard-error.json")!.stat.kind = "symlink";
        return summarizeObservation(await observe(() => verifyAttempt2Unchanged(makeIo(observations))));
      },
      matches: rejects,
    },
    {
      label: "receipt.json noncanonical path",
      expected: rejected,
      async run() {
        const observations = makeObservations();
        observations.files.get("receipt.json")!.canonicalPath = `${attempt2Directory}/../receipt.json`;
        return summarizeObservation(await observe(() => verifyAttempt2Unchanged(makeIo(observations))));
      },
      matches: rejects,
    },
    {
      label: "stat.size differs from observed fingerprint byteLength",
      expected: rejected,
      async run() {
        const observations = makeObservations();
        observations.files.get("receipt.json")!.stat.size += 1;
        return summarizeObservation(await observe(() => verifyAttempt2Unchanged(makeIo(observations))));
      },
      matches: rejects,
    },
    {
      label: "fingerprint length differs from fixed preserved byte length",
      expected: rejected,
      async run() {
        const observations = makeObservations();
        const file = observations.files.get("receipt.json")!;
        file.fingerprint.byteLength += 1;
        file.stat.size += 1;
        return summarizeObservation(await observe(() => verifyAttempt2Unchanged(makeIo(observations))));
      },
      matches: rejects,
    },
    {
      label: "guard-error.json wrong SHA-256",
      expected: rejected,
      async run() {
        const observations = makeObservations();
        observations.files.get("guard-error.json")!.fingerprint.sha256 = "0".repeat(64);
        return summarizeObservation(await observe(() => verifyAttempt2Unchanged(makeIo(observations))));
      },
      matches: rejects,
    },
    {
      label: "non-numeric stat.size",
      expected: rejected,
      async run() {
        const observations = makeObservations();
        observations.files.get("guard-error.json")!.stat.size = Number.NaN;
        return summarizeObservation(await observe(() => verifyAttempt2Unchanged(makeIo(observations))));
      },
      matches: rejects,
    },
    {
      label: "fingerprint IO error propagates",
      expected: rejected,
      async run() {
        const observations = makeObservations();
        observations.files.get("guard-error.json")!.readError = new Error("injected fingerprint IO failure");
        return summarizeObservation(await observe(() => verifyAttempt2Unchanged(makeIo(observations))));
      },
      matches: rejects,
    },
    {
      label: "valid baseline followed by late attempt-2 drift rejects",
      expected: { initial: "returned", final: "threw" },
      async run() {
        const observations = makeObservations();
        const io = makeIo(observations);
        const initial = await observe(() => verifyAttempt2Unchanged(io));
        observations.files.get("receipt.json")!.fingerprint.sha256 = "0".repeat(64);
        const final = await observe(() => verifyAttempt2Unchanged(io));
        return { initial: summarizeObservation(initial), final: summarizeObservation(final) };
      },
      matches(actual) {
        return isRecord(actual) && fieldOutcome(actual, "initial") === "returned" &&
          fieldOutcome(actual, "final") === "threw";
      },
    },
    {
      label: "real writeComplete orders final inventory and receipt readback before callback then exclusive COMPLETE open",
      expected: {
        completion: "returned",
        everyInventoryHashReadBeforeCallback: true,
        receiptReadbackBeforeCallback: true,
        completeOpenImmediatelyAfterCallback: true,
      },
      async run() {
        const io = fakeEvidenceIo();
        const persistence = createEvidencePersistenceForTesting(io);
        const guardError = {
          name: "Error",
          message: "A JSZip generation/deflate module is present in the runtime closure: node_modules/jszip/lib/crc32.js",
          stack: "attempt-2 lifecycle guard stack",
        };
        const payload = preGuardFiles();
        await persistence.createChild();
        await persistence.persistBeforeGuard(payload);
        await persistence.persistGuardError(guardError);
        await persistence.persistFinalReceipt(finalReceiptForPayload(payload, guardError), guardError);
        const priorFilePaths = [...io.files.keys()];
        const completionStart = io.events.length;
        const completion = await observe(() => persistence.writeComplete(guardError, async () => {
          io.events.push("old-attempts-check:start");
          await verifyAttempt2Unchanged(makeIo(makeObservations()));
          io.events.push("old-attempts-check:end");
        }));
        const events = io.events.slice(completionStart);
        const callbackStart = events.indexOf("old-attempts-check:start");
        const callbackEnd = events.indexOf("old-attempts-check:end");
        const completeOpen = events.indexOf(`open:${path.join(EVIDENCE_CHILD, "COMPLETE.json")}:wx:384`);
        const readsBeforeCallback = events.slice(0, callbackStart).filter((event) => event.startsWith("readFile:"));
        const receiptPath = path.join(EVIDENCE_CHILD, "receipt.json");
        const result = {
          completion: summarizeObservation(completion),
          everyInventoryHashReadBeforeCallback: callbackStart > 0 &&
            priorFilePaths.every((file) => readsBeforeCallback.includes(`readFile:${file}`)),
          receiptReadbackBeforeCallback: readsBeforeCallback.filter((event) => event === `readFile:${receiptPath}`).length >= 2,
          completeOpenImmediatelyAfterCallback: callbackEnd >= 0 && completeOpen === callbackEnd + 1,
          eventOrder: events,
        };
        return result;
      },
      matches(actual) {
        return isRecord(actual) && fieldOutcome(actual, "completion") === "returned" &&
          actual.everyInventoryHashReadBeforeCallback === true &&
          actual.receiptReadbackBeforeCallback === true &&
          actual.completeOpenImmediatelyAfterCallback === true;
      },
    },
    {
      label: "late old-attempt drift rejects the actual writeComplete callback and suppresses COMPLETE open",
      expected: { baseline: "returned", completion: "threw", completeOpenAttempts: 0 },
      async run() {
        const io = fakeEvidenceIo();
        const persistence = createEvidencePersistenceForTesting(io);
        const guardError = {
          name: "Error",
          message: "A JSZip generation/deflate module is present in the runtime closure: node_modules/jszip/lib/crc32.js",
          stack: "attempt-2 callback rejection stack",
        };
        const payload = preGuardFiles();
        await persistence.createChild();
        await persistence.persistBeforeGuard(payload);
        await persistence.persistGuardError(guardError);
        await persistence.persistFinalReceipt(finalReceiptForPayload(payload, guardError), guardError);
        const observations = makeObservations();
        const attempt2Io = makeIo(observations);
        const baseline = await observe(() => verifyAttempt2Unchanged(attempt2Io));
        observations.files.get("receipt.json")!.fingerprint.sha256 = "0".repeat(64);
        const completion = await observe(() => persistence.writeComplete(guardError, async () => {
          io.events.push("old-attempts-check:start");
          await verifyAttempt2Unchanged(attempt2Io);
          io.events.push("old-attempts-check:end");
        }));
        return {
          baseline: summarizeObservation(baseline),
          completion: summarizeObservation(completion),
          completeOpenAttempts: openBasenames(io.events).filter((name) => name === "COMPLETE.json").length,
          completeWriteAttempts: io.events.filter((event) => event === `write:${path.join(EVIDENCE_CHILD, "COMPLETE.json")}`).length,
          callbackEvents: io.events.filter((event) => event.startsWith("old-attempts-check:")),
        };
      },
      matches(actual) {
        return isRecord(actual) && fieldOutcome(actual, "baseline") === "returned" &&
          fieldOutcome(actual, "completion") === "threw" && actual.completeOpenAttempts === 0 &&
          actual.completeWriteAttempts === 0;
      },
    },
    {
      label: "missing beforeComplete callback rejects at runtime without opening COMPLETE",
      expected: { completion: "threw", completeOpenAttempts: 0 },
      async run() {
        const io = fakeEvidenceIo();
        const persistence = createEvidencePersistenceForTesting(io);
        const guardError = {
          name: "Error",
          message: "A JSZip generation/deflate module is present in the runtime closure: node_modules/jszip/lib/crc32.js",
          stack: "missing callback guard stack",
        };
        const payload = preGuardFiles();
        await persistence.createChild();
        await persistence.persistBeforeGuard(payload);
        await persistence.persistGuardError(guardError);
        await persistence.persistFinalReceipt(finalReceiptForPayload(payload, guardError), guardError);
        const completion = await observe(() => Reflect.apply(
          persistence.writeComplete,
          persistence,
          [guardError],
        ));
        return {
          completion: summarizeObservation(completion),
          completeOpenAttempts: openBasenames(io.events).filter((name) => name === "COMPLETE.json").length,
        };
      },
      matches(actual) {
        return isRecord(actual) && fieldOutcome(actual, "completion") === "threw" &&
          actual.completeOpenAttempts === 0;
      },
    },
    {
      label: "production source wires both pre-create and final immutable checks and preserves absence-to-mkdir adjacency",
      expected: {
        bothBaselinesBeforeCreate: true,
        closureRechecksAndComparesBoth: true,
        closurePassedToWriteComplete: true,
        finalAwaitImmediatelyBeforeCompleteWrite: true,
        attempt3AbsenceImmediatelyBeforeMkdir: true,
        productionFingerprintUsesRawBytesAndNumericStatSize: true,
      },
      async run() {
        const source = await fs.readFile(
          path.resolve(process.cwd(), "tests/fidelity/docxEvidenceCapture.ts"),
          "utf8",
        );
        const writeStart = source.indexOf("async writeComplete(");
        const writeEnd = source.indexOf("async writeIncomplete(", writeStart);
        const writeBody = source.slice(writeStart, writeEnd);
        return {
          bothBaselinesBeforeCreate: /const beforeAttempt1 = await verifyAttempt1Unchanged\(\);\s*const beforeAttempt2 = await verifyAttempt2Unchanged\(defaultPreservedAttemptReadIo\);[\s\S]*?await persistence\.createChild\(\);/.test(source),
          closureRechecksAndComparesBoth: /const beforeComplete = async \(\): Promise<void> => \{\s*const afterAttempt1 = await verifyAttempt1Unchanged\(\);\s*if \(stableJson\(beforeAttempt1\) !== stableJson\(afterAttempt1\)\)[\s\S]*?const afterAttempt2 = await verifyAttempt2Unchanged\(defaultPreservedAttemptReadIo\);\s*if \(stableJson\(beforeAttempt2\) !== stableJson\(afterAttempt2\)\)/.test(source),
          closurePassedToWriteComplete: /persistence\.writeComplete\(caughtGuardError, beforeComplete\)/.test(source),
          finalAwaitImmediatelyBeforeCompleteWrite: /await beforeComplete\(\);\s*await writeExclusive\("COMPLETE\.json", completeBytes\);/.test(writeBody),
          attempt3AbsenceImmediatelyBeforeMkdir: /if \(await lstatOrNull\(io, ATTEMPT_3_DIRECTORY\)\) throw new Error\("Attempt-3 evidence child already exists in any form; refusing reuse\."\);\s*await io\.mkdir\(ATTEMPT_3_DIRECTORY, \{ recursive: false, mode: 0o700 \}\);/.test(source),
          productionFingerprintUsesRawBytesAndNumericStatSize: /readFileFingerprint: async \(file\) => \{\s*const bytes = await fs\.readFile\(file\);\s*return \{\s*byteLength: bytes\.byteLength,\s*sha256: sha256\(bytes\),\s*\};\s*\}/.test(source) &&
            /stat\.size !== fingerprint\.byteLength/.test(source),
        };
      },
      matches(actual) {
        return isRecord(actual) && Object.values(actual).every((value) => value === true);
      },
    },
  ];

  await assertMatrixOutcomes("preserved attempt-2 lifecycle", variants);
});

void test("capture binding: toolchain receipt metadata uses pinned package bytes", async () => {
  const rejected = { outcome: "threw" };
  const rejects = (actual: unknown): boolean => isRecord(actual) && actual.outcome === "threw";
  const fixedPackagePaths = [
    "node_modules/esbuild/package.json",
    "node_modules/typescript/package.json",
  ] as const;
  const expectedManifest: ControlInputManifest = Object.freeze(fixedPackagePaths.map((relativePath) =>
    Object.freeze({
      path: relativePath,
      categories: Object.freeze(["node_modules-package-metadata"] as const),
    }),
  ));
  const projectRoot = await fs.realpath(process.cwd());
  const [esbuildDiskBytes, typescriptDiskBytes] = await Promise.all(fixedPackagePaths.map((relativePath) =>
    fs.readFile(path.resolve(projectRoot, relativePath)),
  ));
  const realPinnedBytes = new Map<string, Buffer>([
    [fixedPackagePaths[0], Buffer.from(esbuildDiskBytes)],
    [fixedPackagePaths[1], Buffer.from(typescriptDiskBytes)],
  ]);
  const independentlyParsedVersion = (bytes: Uint8Array): string => {
    const parsed: unknown = JSON.parse(Buffer.from(bytes).toString("utf8"));
    assert.ok(isRecord(parsed));
    assert.equal(typeof parsed.version, "string");
    return parsed.version as string;
  };
  const expectedRealEsbuildVersion = independentlyParsedVersion(esbuildDiskBytes);
  const expectedRealTypeScriptVersion = independentlyParsedVersion(typescriptDiskBytes);
  const bindSnapshots = async (
    root: string,
    snapshots: ReadonlyMap<string, Uint8Array>,
  ): Promise<readonly BoundControlInput[]> => bindControlInputs(
    expectedManifest,
    [],
    async (relativePath) => {
      const contents = snapshots.get(relativePath);
      assert.ok(contents, `unexpected pinned toolchain path: ${relativePath}`);
      return {
        projectRoot: root,
        canonicalPath: path.resolve(root, relativePath),
        contents: Buffer.from(contents),
        byteLength: contents.byteLength,
        sha256: sha256(contents),
        device: "independent-memory-device",
        inode: `independent-memory-inode:${relativePath}`,
      };
    },
  );
  const realControls = await bindSnapshots(projectRoot, realPinnedBytes);
  const fixedFixtureBytes = new Map<string, Buffer>([
    [fixedPackagePaths[0], Buffer.from('{"name":"esbuild","version":"0.0.0-w130-pinned-esbuild"}\n')],
    [fixedPackagePaths[1], Buffer.from('{"name":"typescript","version":"0.0.0-w130-pinned-typescript"}\n')],
  ]);
  const fixtureControls = await bindSnapshots(PROJECT_ROOT, fixedFixtureBytes);
  const nodeVersion = "node-version-passed-explicitly";
  const pluginVersion = "plugin-version-passed-explicitly";
  const expectedRealMetadata = {
    toolchain: {
      node: nodeVersion,
      esbuild: expectedRealEsbuildVersion,
      typescript: expectedRealTypeScriptVersion,
    },
    versions: {
      node: nodeVersion,
      esbuild: expectedRealEsbuildVersion,
      typescript: expectedRealTypeScriptVersion,
      plugin: pluginVersion,
    },
  };
  const expectedFixtureMetadata = {
    toolchain: {
      node: nodeVersion,
      esbuild: "0.0.0-w130-pinned-esbuild",
      typescript: "0.0.0-w130-pinned-typescript",
    },
    versions: {
      node: nodeVersion,
      esbuild: "0.0.0-w130-pinned-esbuild",
      typescript: "0.0.0-w130-pinned-typescript",
      plugin: pluginVersion,
    },
  };
  assert.notEqual(expectedRealEsbuildVersion, expectedFixtureMetadata.toolchain.esbuild);
  assert.notEqual(expectedRealTypeScriptVersion, expectedFixtureMetadata.toolchain.typescript);
  const cloneControl = (
    control: BoundControlInput,
    changes: Record<string, unknown>,
  ): BoundControlInput => Object.freeze({ ...control, ...changes }) as unknown as BoundControlInput;
  const replaceControl = (
    controls: readonly BoundControlInput[],
    target: BoundControlInput,
    replacement: BoundControlInput,
  ): readonly BoundControlInput[] => Object.freeze(controls.map((control) =>
    control === target ? replacement : control,
  ));
  const esbuildControl = fixtureControls.find((control) => control.path === fixedPackagePaths[0]);
  const typescriptControl = fixtureControls.find((control) => control.path === fixedPackagePaths[1]);
  assert.ok(esbuildControl);
  assert.ok(typescriptControl);
  const coherentContents = (control: BoundControlInput, source: string): BoundControlInput => {
    const contents = Buffer.from(source, "utf8");
    return cloneControl(control, {
      contents,
      byteLength: contents.byteLength,
      sha256: sha256(contents),
    });
  };
  const variants: MatrixVariant[] = [
    {
      label: "installed package versions derive from the once-read pinned package bytes and serialize exactly",
      expected: expectedRealMetadata,
      async run() {
        const metadata = preparePinnedToolchainMetadata(realControls, nodeVersion, pluginVersion);
        return JSON.parse(JSON.stringify({ toolchain: metadata.toolchain, versions: metadata.versions })) as unknown;
      },
    },
    {
      label: "independent fixture bytes override disk versions without filesystem rereads",
      expected: expectedFixtureMetadata,
      async run() {
        const metadata = preparePinnedToolchainMetadata(fixtureControls, nodeVersion, pluginVersion);
        return JSON.parse(JSON.stringify({ toolchain: metadata.toolchain, versions: metadata.versions })) as unknown;
      },
    },
    {
      label: "wrong control role",
      expected: rejected,
      run() {
        const candidate = cloneControl(esbuildControl, { role: "compiled" });
        const controls = replaceControl(fixtureControls, esbuildControl, candidate);
        return observe(() => preparePinnedToolchainMetadata(controls, nodeVersion, pluginVersion)).then(summarizeObservation);
      },
      matches: rejects,
    },
    {
      label: "wrong fixed project-relative path",
      expected: rejected,
      run() {
        const candidate = cloneControl(esbuildControl, { path: "node_modules/other/package.json" });
        const controls = replaceControl(fixtureControls, esbuildControl, candidate);
        return observe(() => preparePinnedToolchainMetadata(controls, nodeVersion, pluginVersion)).then(summarizeObservation);
      },
      matches: rejects,
    },
    {
      label: "noncanonical canonical-path identity",
      expected: rejected,
      run() {
        const candidate = cloneControl(esbuildControl, {
          canonicalPath: `${PROJECT_ROOT}/node_modules/esbuild/../esbuild/package.json`,
        });
        const controls = replaceControl(fixtureControls, esbuildControl, candidate);
        return observe(() => preparePinnedToolchainMetadata(controls, nodeVersion, pluginVersion)).then(summarizeObservation);
      },
      matches: rejects,
    },
    {
      label: "missing node_modules-package-metadata category",
      expected: rejected,
      run() {
        const candidate = cloneControl(esbuildControl, { categories: Object.freeze([]) });
        const controls = replaceControl(fixtureControls, esbuildControl, candidate);
        return observe(() => preparePinnedToolchainMetadata(controls, nodeVersion, pluginVersion)).then(summarizeObservation);
      },
      matches: rejects,
    },
    {
      label: "wrong control category without node_modules-package-metadata",
      expected: rejected,
      run() {
        const candidate = cloneControl(esbuildControl, { categories: Object.freeze(["calculateBuildId"]) });
        const controls = replaceControl(fixtureControls, esbuildControl, candidate);
        return observe(() => preparePinnedToolchainMetadata(controls, nodeVersion, pluginVersion)).then(summarizeObservation);
      },
      matches: rejects,
    },
    {
      label: "contents byteLength mismatch against pinned snapshot",
      expected: rejected,
      run() {
        const candidate = cloneControl(esbuildControl, { byteLength: esbuildControl.byteLength + 1 });
        const controls = replaceControl(fixtureControls, esbuildControl, candidate);
        return observe(() => preparePinnedToolchainMetadata(controls, nodeVersion, pluginVersion)).then(summarizeObservation);
      },
      matches: rejects,
    },
    {
      label: "pinned SHA-256 mismatch against package contents",
      expected: rejected,
      run() {
        const candidate = cloneControl(esbuildControl, { sha256: "0".repeat(64) });
        const controls = replaceControl(fixtureControls, esbuildControl, candidate);
        return observe(() => preparePinnedToolchainMetadata(controls, nodeVersion, pluginVersion)).then(summarizeObservation);
      },
      matches: rejects,
    },
    {
      label: "missing fixed toolchain control",
      expected: rejected,
      run() {
        const controls = Object.freeze(fixtureControls.filter((control) => control !== esbuildControl));
        return observe(() => preparePinnedToolchainMetadata(controls, nodeVersion, pluginVersion)).then(summarizeObservation);
      },
      matches: rejects,
    },
    {
      label: "duplicate fixed toolchain control",
      expected: rejected,
      run() {
        const controls = Object.freeze([...fixtureControls, esbuildControl]);
        return observe(() => preparePinnedToolchainMetadata(controls, nodeVersion, pluginVersion)).then(summarizeObservation);
      },
      matches: rejects,
    },
    {
      label: "duplicate version JSON key",
      expected: rejected,
      run() {
        const candidate = coherentContents(typescriptControl, '{"name":"typescript","version":"first","version":"second"}\n');
        const controls = replaceControl(fixtureControls, typescriptControl, candidate);
        return observe(() => preparePinnedToolchainMetadata(controls, nodeVersion, pluginVersion)).then(summarizeObservation);
      },
      matches: rejects,
    },
    {
      label: "malformed package JSON",
      expected: rejected,
      run() {
        const candidate = coherentContents(typescriptControl, "{");
        const controls = replaceControl(fixtureControls, typescriptControl, candidate);
        return observe(() => preparePinnedToolchainMetadata(controls, nodeVersion, pluginVersion)).then(summarizeObservation);
      },
      matches: rejects,
    },
    {
      label: "nonplain package JSON root",
      expected: rejected,
      run() {
        const candidate = coherentContents(typescriptControl, "[]\n");
        const controls = replaceControl(fixtureControls, typescriptControl, candidate);
        return observe(() => preparePinnedToolchainMetadata(controls, nodeVersion, pluginVersion)).then(summarizeObservation);
      },
      matches: rejects,
    },
    {
      label: "absent version field",
      expected: rejected,
      run() {
        const candidate = coherentContents(typescriptControl, '{"name":"typescript"}\n');
        const controls = replaceControl(fixtureControls, typescriptControl, candidate);
        return observe(() => preparePinnedToolchainMetadata(controls, nodeVersion, pluginVersion)).then(summarizeObservation);
      },
      matches: rejects,
    },
    {
      label: "non-string version field",
      expected: rejected,
      run() {
        const candidate = coherentContents(typescriptControl, '{"version":7}\n');
        const controls = replaceControl(fixtureControls, typescriptControl, candidate);
        return observe(() => preparePinnedToolchainMetadata(controls, nodeVersion, pluginVersion)).then(summarizeObservation);
      },
      matches: rejects,
    },
    {
      label: "empty version field",
      expected: rejected,
      run() {
        const candidate = coherentContents(typescriptControl, '{"version":""}\n');
        const controls = replaceControl(fixtureControls, typescriptControl, candidate);
        return observe(() => preparePinnedToolchainMetadata(controls, nodeVersion, pluginVersion)).then(summarizeObservation);
      },
      matches: rejects,
    },
    {
      label: "padded version field",
      expected: rejected,
      run() {
        const candidate = coherentContents(typescriptControl, '{"version":" 1.2.3 "}\n');
        const controls = replaceControl(fixtureControls, typescriptControl, candidate);
        return observe(() => preparePinnedToolchainMetadata(controls, nodeVersion, pluginVersion)).then(summarizeObservation);
      },
      matches: rejects,
    },
    {
      label: "real capture uses pinned toolchain for inputs and receipt without package rereads",
      expected: {
        helperCallUsesCapturedControlsAndVersions: true,
        manifestUsesPinnedToolchain: true,
        receiptClosesOverPinnedVersions: true,
        captureHasNoPackageMetadataReread: true,
        pureHelperHasNoIOFallback: true,
      },
      async run() {
        const source = await fs.readFile(
          path.resolve(process.cwd(), "tests/fidelity/docxEvidenceCapture.ts"),
          "utf8",
        );
        const helperStart = source.indexOf("export function preparePinnedToolchainMetadata(");
        const helperEnd = source.indexOf("\nfunction parseMetafile(", helperStart);
        const helperSource = source.slice(helperStart, helperEnd);
        return {
          helperCallUsesCapturedControlsAndVersions: /preparePinnedToolchainMetadata\(controls, process\.version, version as string\)/.test(source) &&
            /bindingContext: Object\.freeze\(\{\s*expectedManifest: controlManifest,\s*compiledInputs: mergedInputs,\s*controls,/.test(source) &&
            /prepareEvidenceInputs\(mergedInputs, controls, manifestBytes, stylesBytes\)/.test(source),
          manifestUsesPinnedToolchain: /toolchain:\s*pinnedToolchainMetadata\.toolchain/.test(source),
          receiptClosesOverPinnedVersions: /const versions = pinnedToolchainMetadata\.versions/.test(source) &&
            /versions,\s*serializedProductionOptions:/.test(source),
          captureHasNoPackageMetadataReread: !/fs\.readFile\(path\.join\(projectRoot,\s*"node_modules\/(?:esbuild|typescript)\/package\.json"/.test(source),
          pureHelperHasNoIOFallback: helperStart >= 0 && helperEnd > helperStart && !/\bfs\b|readFile/.test(helperSource),
        };
      },
      matches(actual) {
        return isRecord(actual) && Object.values(actual).every((value) => value === true);
      },
    },
  ];

  await assertMatrixOutcomes("pinned toolchain metadata", variants);
});
