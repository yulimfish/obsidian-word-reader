import { createHash } from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";

import { writeStoredZip } from "../../src/ooxml/storedZipWriter";

const VIRTUAL_NAMESPACE = "read-only-zip-disabled-readable-stream";
const VIRTUAL_PATH = "jszip-readable-stream-disabled";
const VIRTUAL_KEY = `${VIRTUAL_NAMESPACE}:${VIRTUAL_PATH}`;
const VIRTUAL_TEXT = "module.exports = Object.freeze({});";
const APPROVED_EVIDENCE_ROOT = "/private/var/folders/60/0btch0dj3111k_1grpv_9xmw0000gn/T/opencode";
const ATTEMPT_1_DIRECTORY = path.join(APPROVED_EVIDENCE_ROOT, "obsidian-word-reader-W69-S5.3-evidence");
const ATTEMPT_2_DIRECTORY = path.join(APPROVED_EVIDENCE_ROOT, "obsidian-word-reader-W69-S5.3-evidence-attempt-2");
const ATTEMPT_3_DIRECTORY = path.join(APPROVED_EVIDENCE_ROOT, "obsidian-word-reader-W69-S5.3-evidence-attempt-3");
const ATTEMPT_1_HASHES = Object.freeze({
  "guard-error.json": "4c619ea4b9f9be801f0f0e40be80a157ba258f94ab743b52c6bcf523b36172b3",
  "receipt.json": "fcb4d587066de47d0a2f5f6116a7797ca2cf3cf26785523458beb280984c4d56",
});
const ATTEMPT_2_HASHES = Object.freeze({
  "guard-error.json": "85a8b2555598836ba10fd328ec9ade79c6023e7e1b1c4d1ae366735431224b4d",
  "receipt.json": "a8d87b311ee9f98b5d469d7a471c0fe1074510578bbcd04d4d954d1283c681ea",
});
const ATTEMPT_2_LENGTHS = Object.freeze({
  "guard-error.json": 1469,
  "receipt.json": 1504,
});
const PRE_GUARD_FILE_NAMES = Object.freeze([
  "main.js",
  "docx-engine.cjs",
  "main.metafile.json",
  "engine.metafile.json",
  "inputs.json",
]);
const ALL_TOP_LEVEL_FILE_NAMES = Object.freeze([
  ...PRE_GUARD_FILE_NAMES,
  "receipt.json",
  "guard-error.json",
  "COMPLETE.json",
]);
const MAIN_LIMIT_BYTES = 512_000;
const ARCHIVE_LIMIT_BYTES = 8_388_608;
const ARCHIVE_PAYLOAD_LIMIT_BYTES = 8_388_192;
const ARCHIVE_OVERHEAD_BYTES = 416;
const EXPECTED_GUARD_MESSAGE = "A JSZip generation/deflate module is present in the runtime closure: node_modules/jszip/lib/crc32.js";
const SHA256_PATTERN = /^[0-9a-f]{64}$/;

export interface VirtualLoadRecord {
  readonly entry: "engine" | "main";
  readonly namespace: string;
  readonly path: string;
  readonly suffix: string;
  readonly importAttributes: unknown;
  readonly loader: string;
  readonly resolveDir: string;
  readonly contents: string;
}

export interface PhysicalInputInspection {
  readonly projectRoot: string;
  readonly canonicalPath: string;
  readonly contents: Uint8Array;
  readonly loader: string;
  readonly resolveDir: string;
  readonly device: string | number | bigint;
  readonly inode: string | number | bigint;
  readonly sha256: string;
}

export interface RawProjectFileSnapshot {
  readonly projectRoot: string;
  readonly canonicalPath: string;
  readonly contents: Uint8Array;
  readonly byteLength: number;
  readonly sha256: string;
  readonly device: string;
  readonly inode: string;
}

export type ControlCategory =
  | "calculateBuildId"
  | "project-package-lock-config-tsconfig"
  | "node_modules-package-metadata";

export interface ControlInputSpec {
  readonly path: string;
  readonly categories: readonly ControlCategory[];
}

export type ControlInputManifest = readonly ControlInputSpec[];

export interface SerializedControlInput {
  readonly role: "control";
  readonly path: string;
  readonly canonicalPath: string;
  readonly device: string;
  readonly inode: string;
  readonly byteLength: number;
  readonly sha256: string;
  readonly categories: readonly ControlCategory[];
  readonly compiledKeys: readonly string[];
  readonly blob: string;
}

export type BoundControlInput = RawProjectFileSnapshot & ControlInputSpec & {
  readonly role: "control";
  readonly compiledKeys: readonly string[];
};

export interface BoundMetafileInput {
  readonly namespace: string;
  readonly key: string;
  readonly path: string;
  readonly canonicalPath: string;
  readonly loader: string;
  readonly resolveDir: string;
  readonly byteLength: number;
  readonly sha256: string;
  readonly device: string | number | bigint | null;
  readonly inode: string | number | bigint | null;
  readonly entries: readonly ("engine" | "main")[];
  readonly imports: readonly unknown[];
  readonly bytes: Uint8Array;
}

export type PhysicalInputInspector = (
  rawKey: string,
) => PhysicalInputInspection | Promise<PhysicalInputInspection>;

interface MetafileInputLike {
  readonly imports: readonly unknown[];
}

interface ParsedMetafile {
  readonly inputs: Readonly<Record<string, MetafileInputLike>>;
}

interface EvidenceFileHash {
  readonly path: string;
  readonly byteLength: number;
  readonly sha256: string;
}

export interface EvidencePreGuardFiles {
  readonly "main.js": Uint8Array;
  readonly "docx-engine.cjs": Uint8Array;
  readonly "main.metafile.json": Uint8Array;
  readonly "engine.metafile.json": Uint8Array;
  readonly "inputs.json": Uint8Array;
}

export interface EvidenceFileHandle {
  writeFile(bytes: Uint8Array): Promise<void>;
  close(): Promise<void>;
}

export interface EvidenceStatLike {
  readonly uid: number;
  readonly mode: number;
  isDirectory(): boolean;
  isFile(): boolean;
  isSymbolicLink(): boolean;
}

export interface EvidenceDirentLike {
  readonly name: string;
  isDirectory(): boolean;
  isFile(): boolean;
  isSymbolicLink(): boolean;
}

export interface EvidencePersistenceIo {
  mkdir(directory: string, options: { readonly recursive: false; readonly mode: number }): Promise<void>;
  open(file: string, flags: "wx", mode: number): Promise<EvidenceFileHandle>;
  lstat(target: string): Promise<EvidenceStatLike>;
  realpath(target: string): Promise<string>;
  readdir(directory: string, options: { readonly withFileTypes: true }): Promise<readonly EvidenceDirentLike[]>;
  readFile(file: string): Promise<Uint8Array>;
}

export interface EvidenceBlob {
  readonly sha256: string;
  readonly bytes: Uint8Array;
}

export interface EvidencePersistencePayload {
  readonly files: EvidencePreGuardFiles;
  readonly blobs: readonly EvidenceBlob[];
  readonly bindingContext?: {
    readonly expectedManifest: ControlInputManifest;
    readonly compiledInputs: readonly BoundMetafileInput[];
    readonly controls: readonly BoundControlInput[];
  };
}

export interface ExpectedGuardError {
  readonly name: string;
  readonly message: string;
  readonly stack: string;
}

export interface CompleteReceipt {
  readonly status: "guard failed / evidence only";
  readonly files: readonly EvidenceFileHash[];
  readonly completeSha256: string;
}

interface EvidencePersistence {
  createChild(): Promise<void>;
  persistBeforeGuard(payload: EvidencePersistencePayload): Promise<void>;
  persistFinalReceipt(bytes: Uint8Array, error: ExpectedGuardError): Promise<void>;
  persistGuardError(error: ExpectedGuardError): Promise<void>;
  writeComplete(error: ExpectedGuardError, beforeComplete: () => Promise<void>): Promise<CompleteReceipt>;
  writeIncomplete(error: unknown): Promise<void>;
}

export interface PreservedAttemptStat {
  readonly size: number;
  readonly uid: number;
  readonly gid: number;
  readonly mode: number;
  isDirectory(): boolean;
  isFile(): boolean;
  isSymbolicLink(): boolean;
}

export interface PreservedAttemptDirent {
  readonly name: string;
  isDirectory(): boolean;
  isFile(): boolean;
  isSymbolicLink(): boolean;
}

export interface PreservedAttemptReadIo {
  lstat(target: string): Promise<PreservedAttemptStat>;
  realpath(target: string): Promise<string>;
  readdir(directory: string): Promise<readonly PreservedAttemptDirent[]>;
  readFileFingerprint(file: string): Promise<{
    readonly byteLength: number;
    readonly sha256: string;
  }>;
}

const defaultEvidenceIo: EvidencePersistenceIo = {
  mkdir: (directory, options) => fs.mkdir(directory, options),
  open: async (file, flags, mode) => fs.open(file, flags, mode),
  lstat: async (target) => fs.lstat(target),
  realpath: async (target) => fs.realpath(target),
  readdir: async (directory, options) => fs.readdir(directory, options),
  readFile: async (file) => fs.readFile(file),
};

const defaultPreservedAttemptReadIo: PreservedAttemptReadIo = {
  lstat: async (target) => fs.lstat(target),
  realpath: async (target) => fs.realpath(target),
  readdir: async (directory) => fs.readdir(directory, { withFileTypes: true }),
  readFileFingerprint: async (file) => {
    const bytes = await fs.readFile(file);
    return {
      byteLength: bytes.byteLength,
      sha256: sha256(bytes),
    };
  },
};

function sha256(bytes: Uint8Array): string {
  return createHash("sha256").update(bytes).digest("hex");
}

function compareText(left: string, right: string): number {
  return left < right ? -1 : left > right ? 1 : 0;
}

function stableValue(value: unknown): unknown {
  if (Array.isArray(value)) return value.map(stableValue);
  if (value && typeof value === "object") {
    const record = value as Record<string, unknown>;
    return Object.fromEntries(Object.keys(record).sort(compareText).map((key) => [key, stableValue(record[key])]));
  }
  return value;
}

function stableJson(value: unknown): string {
  return JSON.stringify(stableValue(value));
}

function deepFreezeJson(value: unknown): unknown {
  if (Array.isArray(value)) {
    for (const item of value) deepFreezeJson(item);
    return Object.freeze(value);
  }
  if (value && typeof value === "object") {
    for (const item of Object.values(value as Record<string, unknown>)) deepFreezeJson(item);
    return Object.freeze(value);
  }
  return value;
}

function assertNoDuplicateJsonKeys(source: string): void {
  let index = 0;
  const fail = (): never => { throw new SyntaxError(`Invalid or duplicate-key JSON at byte ${index}.`); };
  const skipWhitespace = (): void => {
    while (index < source.length && /\s/.test(source[index])) index += 1;
  };
  const parseString = (): string => {
    if (source[index] !== '"') return fail();
    const start = index;
    index += 1;
    while (index < source.length) {
      const character = source[index];
      if (character === "\\") {
        index += 2;
        continue;
      }
      index += 1;
      if (character === '"') {
        return JSON.parse(source.slice(start, index)) as string;
      }
    }
    return fail();
  };
  const parseValue = (depth: number): void => {
    if (depth > 256) return fail();
    skipWhitespace();
    const character = source[index];
    if (character === "{") {
      index += 1;
      skipWhitespace();
      const keys = new Set<string>();
      if (source[index] === "}") { index += 1; return; }
      while (index < source.length) {
        skipWhitespace();
        const key = parseString();
        if (keys.has(key)) throw new SyntaxError(`Duplicate JSON object key ${JSON.stringify(key)}.`);
        keys.add(key);
        skipWhitespace();
        if (source[index] !== ":") return fail();
        index += 1;
        parseValue(depth + 1);
        skipWhitespace();
        if (source[index] === "}") { index += 1; return; }
        if (source[index] !== ",") return fail();
        index += 1;
      }
      return fail();
    }
    if (character === "[") {
      index += 1;
      skipWhitespace();
      if (source[index] === "]") { index += 1; return; }
      while (index < source.length) {
        parseValue(depth + 1);
        skipWhitespace();
        if (source[index] === "]") { index += 1; return; }
        if (source[index] !== ",") return fail();
        index += 1;
      }
      return fail();
    }
    if (character === '"') { parseString(); return; }
    const primitive = /^(?:true|false|null)/.exec(source.slice(index));
    if (primitive) { index += primitive[0].length; return; }
    const number = /^-?(?:0|[1-9][0-9]*)(?:\.[0-9]+)?(?:[eE][+-]?[0-9]+)?/.exec(source.slice(index));
    if (number) { index += number[0].length; return; }
    return fail();
  };

  parseValue(0);
  skipWhitespace();
  if (index !== source.length) fail();
}

function isPlainRecord(value: unknown): value is Record<string, unknown> {
  if (!value || typeof value !== "object" || Array.isArray(value)) return false;
  const prototype = Object.getPrototypeOf(value);
  return prototype === Object.prototype || prototype === null;
}

const PINNED_TOOLCHAIN_PACKAGE_PATHS = Object.freeze({
  esbuild: "node_modules/esbuild/package.json",
  typescript: "node_modules/typescript/package.json",
});

export function preparePinnedToolchainMetadata(
  controls: readonly BoundControlInput[],
  nodeVersion: string,
  pluginVersion: string,
): {
  readonly toolchain: {
    readonly node: string;
    readonly esbuild: string;
    readonly typescript: string;
  };
  readonly versions: {
    readonly node: string;
    readonly esbuild: string;
    readonly typescript: string;
    readonly plugin: string;
  };
} {
  if (!Array.isArray(controls)) throw new TypeError("Pinned toolchain metadata requires bound control rows.");
  if (typeof nodeVersion !== "string" || typeof pluginVersion !== "string") {
    throw new TypeError("Pinned toolchain node and plugin versions must be strings.");
  }

  const readPinnedVersion = (relativePath: string): string => {
    const matches = controls.filter((control) =>
      isPlainRecord(control) && control.path === relativePath,
    );
    if (matches.length !== 1) {
      throw new Error(`Pinned toolchain control must occur exactly once: ${relativePath}`);
    }
    const control = matches[0] as BoundControlInput;
    if (control.role !== "control") throw new Error(`Pinned toolchain row has an incorrect role: ${relativePath}`);
    if (control.path !== relativePath) throw new Error(`Pinned toolchain path is not canonical: ${relativePath}`);
    if (
      !Array.isArray(control.categories) ||
      !control.categories.includes("node_modules-package-metadata")
    ) throw new Error(`Pinned toolchain control lacks its package-metadata category: ${relativePath}`);
    if (
      typeof control.projectRoot !== "string" || !path.isAbsolute(control.projectRoot) ||
      path.resolve(control.projectRoot) !== control.projectRoot ||
      typeof control.canonicalPath !== "string" ||
      path.resolve(control.canonicalPath) !== control.canonicalPath ||
      control.canonicalPath !== path.resolve(control.projectRoot, relativePath) ||
      path.relative(control.projectRoot, control.canonicalPath).replaceAll("\\", "/") !== relativePath
    ) throw new Error(`Pinned toolchain control has a noncanonical project-relative identity: ${relativePath}`);
    if (!(control.contents instanceof Uint8Array)) {
      throw new TypeError(`Pinned toolchain control has no raw package bytes: ${relativePath}`);
    }
    if (
      !Number.isSafeInteger(control.byteLength) || control.byteLength < 0 ||
      control.byteLength !== control.contents.byteLength
    ) throw new Error(`Pinned toolchain control byte length differs from its package bytes: ${relativePath}`);
    if (
      typeof control.sha256 !== "string" || !SHA256_PATTERN.test(control.sha256) ||
      sha256(control.contents) !== control.sha256
    ) throw new Error(`Pinned toolchain control hash differs from its package bytes: ${relativePath}`);

    const source = Buffer.from(control.contents).toString("utf8");
    assertNoDuplicateJsonKeys(source);
    const metadata: unknown = JSON.parse(source);
    if (!isPlainRecord(metadata)) {
      throw new TypeError(`Pinned toolchain package metadata must be a plain JSON object: ${relativePath}`);
    }
    if (!Object.prototype.hasOwnProperty.call(metadata, "version")) {
      throw new Error(`Pinned toolchain package metadata has no own version field: ${relativePath}`);
    }
    const version = metadata.version;
    if (typeof version !== "string" || version.length === 0 || version.trim() !== version) {
      throw new Error(`Pinned toolchain package version must be a nonempty unpadded string: ${relativePath}`);
    }
    return version;
  };

  const esbuild = readPinnedVersion(PINNED_TOOLCHAIN_PACKAGE_PATHS.esbuild);
  const typescript = readPinnedVersion(PINNED_TOOLCHAIN_PACKAGE_PATHS.typescript);
  return {
    toolchain: { node: nodeVersion, esbuild, typescript },
    versions: { node: nodeVersion, esbuild, typescript, plugin: pluginVersion },
  };
}

function parseMetafile(value: string | unknown): ParsedMetafile {
  let parsed: unknown = value;
  if (typeof value === "string") {
    assertNoDuplicateJsonKeys(value);
    parsed = JSON.parse(value) as unknown;
  }
  if (!isPlainRecord(parsed) || !isPlainRecord(parsed.inputs)) {
    throw new TypeError("Metafile must contain a plain string-keyed inputs map.");
  }
  if (Reflect.ownKeys(parsed.inputs).some((key) => typeof key !== "string")) {
    throw new TypeError("Metafile inputs cannot contain symbol keys.");
  }
  const inputs: Record<string, MetafileInputLike> = Object.create(null) as Record<string, MetafileInputLike>;
  for (const [key, input] of Object.entries(parsed.inputs)) {
    if (!isPlainRecord(input) || !Array.isArray(input.imports)) {
      throw new TypeError(`Metafile input ${JSON.stringify(key)} is malformed.`);
    }
    inputs[key] = input as unknown as MetafileInputLike;
  }
  return { inputs };
}

type ClassifiedKey =
  | { readonly kind: "virtual"; readonly key: string; readonly input: MetafileInputLike }
  | { readonly kind: "physical"; readonly key: string; readonly input: MetafileInputLike };

function classifyRawInputKey(key: string, input: MetafileInputLike): ClassifiedKey {
  if (key === VIRTUAL_KEY) return { kind: "virtual", key, input };
  if (!key || key.includes("\0") || key.includes("?") || key.includes("#")) {
    throw new Error(`Invalid physical input key or suffix/query: ${JSON.stringify(key)}.`);
  }
  const colon = key.indexOf(":");
  const slash = key.indexOf("/");
  if (colon >= 0 && (slash < 0 || colon < slash)) {
    throw new Error(`Unknown namespace or virtual input key: ${key}`);
  }
  if (path.isAbsolute(key) || path.win32.isAbsolute(key) || key.includes("\\") || key.includes(":")) {
    throw new Error(`Invalid physical input path: ${key}`);
  }
  const parts = key.split("/");
  if (parts.some((part) => !part || part === "." || part === "..")) {
    throw new Error(`Physical input path is not canonical or escapes its project: ${key}`);
  }
  if (path.posix.normalize(key) !== key) {
    throw new Error(`Physical input path is not canonical: ${key}`);
  }
  return { kind: "physical", key, input };
}

function hasExactEmptyAttributes(value: unknown): boolean {
  return isPlainRecord(value) && Object.getPrototypeOf(value) === Object.prototype && Reflect.ownKeys(value).length === 0;
}

function assertVirtualTuple(event: VirtualLoadRecord): void {
  if (event.entry !== "engine" && event.entry !== "main") throw new Error("Virtual input callback has an unknown entry label.");
  if (event.namespace !== VIRTUAL_NAMESPACE || event.path !== VIRTUAL_PATH) {
    throw new Error("Virtual input callback namespace/path does not exactly match the known read-only module.");
  }
  if (event.suffix !== "") throw new Error("Virtual input callback has a nonempty suffix/query.");
  if (!hasExactEmptyAttributes(event.importAttributes)) throw new Error("Virtual input callback has nonempty or malformed import attributes.");
  if (event.loader !== "js") throw new Error("Virtual input callback loader drifted from js.");
  if (event.resolveDir !== "") throw new Error("Virtual input callback resolveDir drifted from empty.");
  if (typeof event.contents !== "string" || event.contents !== VIRTUAL_TEXT) {
    throw new Error("Virtual input callback contents drifted from the exact read-only module.");
  }
}

function loaderForPath(file: string): string {
  const extension = path.extname(file).toLowerCase();
  const loaders: Readonly<Record<string, string>> = {
    ".js": "js",
    ".mjs": "js",
    ".cjs": "js",
    ".ts": "ts",
    ".tsx": "tsx",
    ".jsx": "jsx",
    ".json": "json",
  };
  const loader = loaders[extension];
  if (!loader) throw new Error(`Unsupported physical input loader extension ${extension}: ${file}`);
  return loader;
}

export function pinnedPhysicalLoad(
  inputs: readonly BoundMetafileInput[],
  args: { readonly namespace: string; readonly path: string; readonly suffix: string; readonly with: unknown },
  entry: "engine" | "main",
): { contents: Uint8Array; loader: string; resolveDir: string } {
  if (args.namespace !== "file") throw new Error(`Unknown authoritative load namespace: ${args.namespace}`);
  if (args.suffix !== "" || !hasExactEmptyAttributes(args.with)) {
    throw new Error(`Authoritative file load has suffix or import attributes: ${args.path}`);
  }
  const binding = inputs.find((input) => input.namespace === "file" && input.canonicalPath === args.path);
  if (!binding || !binding.entries.includes(entry)) {
    throw new Error(`Unknown or unpinned authoritative file load: ${args.path}`);
  }
  return { contents: binding.bytes, loader: binding.loader, resolveDir: binding.resolveDir };
}

export function assertCapturedEngineBinding(
  mainDefines: Readonly<Record<string, string>>,
  engineSha256: string,
  engineByteLength: number,
): void {
  if (typeof engineByteLength !== "number" || !Number.isSafeInteger(engineByteLength) || engineByteLength < 0) {
    throw new Error("Invalid expected DOCX engine byte length.");
  }
  const embeddedEngineByteLength: unknown = JSON.parse(mainDefines.__DOCX_ENGINE_BYTES__);
  if (
    typeof embeddedEngineByteLength !== "number" ||
    !Number.isSafeInteger(embeddedEngineByteLength) ||
    embeddedEngineByteLength < 0
  ) {
    throw new Error("Invalid embedded DOCX engine byte length.");
  }
  if (
    embeddedEngineByteLength !== engineByteLength ||
    JSON.parse(mainDefines.__DOCX_ENGINE_SHA256__) !== engineSha256
  ) {
    throw new Error("main.js does not embed the captured engine digest/byte length.");
  }
}

function freezeBinding(
  value: Omit<BoundMetafileInput, "bytes">,
  sourceBytes: Uint8Array,
): BoundMetafileInput {
  const privateBytes = Uint8Array.from(sourceBytes);
  const imports = deepFreezeJson(Array.from(value.imports)) as readonly unknown[];
  const entries = Object.freeze([...value.entries]);
  const binding = {
    ...value,
    imports,
    entries,
    get bytes(): Uint8Array { return Uint8Array.from(privateBytes); },
  };
  return Object.freeze(binding);
}

/** Classify every raw key before calling the physical inspector for any input. */
export async function bindMetafileInputs(
  metafileValue: string | unknown,
  virtualLoads: readonly VirtualLoadRecord[],
  physicalInspector: PhysicalInputInspector,
): Promise<readonly BoundMetafileInput[]> {
  if (!Array.isArray(virtualLoads) || typeof physicalInspector !== "function") {
    throw new TypeError("Metafile input binding requires callback records and a physical inspector.");
  }
  const metafile = parseMetafile(metafileValue);
  const classified = Object.keys(metafile.inputs)
    .sort(compareText)
    .map((key) => classifyRawInputKey(key, metafile.inputs[key]));
  const virtualKeys = classified.filter((input) => input.kind === "virtual");
  const hasVirtualKey = virtualKeys.length === 1;
  if (virtualKeys.length > 1) throw new Error("Duplicate or ambiguous exact virtual metafile key.");
  if (hasVirtualKey) {
    if (virtualLoads.length !== 1) throw new Error("The virtual metafile key requires exactly one onVirtualLoad callback for its entry.");
    const event = virtualLoads[0];
    assertVirtualTuple(event);
    if (metafile.inputs[VIRTUAL_KEY].imports.length !== 0) {
      throw new Error("The virtual metafile input must have imports: [].");
    }
    const bytes = Buffer.from(event.contents, "utf8");
    const virtualBinding = freezeBinding({
      namespace: VIRTUAL_NAMESPACE,
      key: VIRTUAL_KEY,
      path: VIRTUAL_PATH,
      canonicalPath: `${VIRTUAL_NAMESPACE}:${VIRTUAL_PATH}`,
      loader: "js",
      resolveDir: "",
      byteLength: bytes.byteLength,
      sha256: sha256(bytes),
      device: null,
      inode: null,
      entries: [event.entry],
      imports: metafile.inputs[VIRTUAL_KEY].imports,
    }, bytes);
    const physical = await bindPhysicalInputs(classified, physicalInspector);
    return Object.freeze([...physical, virtualBinding].sort((left, right) => compareText(left.key, right.key)));
  }
  if (virtualLoads.length !== 0) throw new Error("An onVirtualLoad callback was supplied without the exact virtual metafile key.");
  return Object.freeze(await bindPhysicalInputs(classified, physicalInspector));
}

async function bindPhysicalInputs(
  classified: readonly ClassifiedKey[],
  physicalInspector: PhysicalInputInspector,
): Promise<BoundMetafileInput[]> {
  const result: BoundMetafileInput[] = [];
  const keyToCanonical = new Map<string, string>();
  const canonicalToKey = new Map<string, string>();
  const identityToKey = new Map<string, string>();
  for (const item of classified) {
    if (item.kind !== "physical") continue;
    const inspection = await physicalInspector(item.key);
    const root = inspection.projectRoot;
    const canonicalRoot = path.resolve(root);
    const expectedPath = path.resolve(canonicalRoot, item.key);
    const canonicalPath = path.resolve(inspection.canonicalPath);
    if (!path.isAbsolute(root) || root !== canonicalRoot || canonicalPath !== inspection.canonicalPath) {
      throw new Error(`Physical input inspector returned a noncanonical project/path for ${item.key}.`);
    }
    const relative = path.relative(canonicalRoot, canonicalPath);
    if (!relative || relative === ".." || relative.startsWith(`..${path.sep}`) || path.isAbsolute(relative)) {
      throw new Error(`Physical input escapes its canonical project root: ${item.key}`);
    }
    if (canonicalPath !== expectedPath || path.posix.normalize(item.key) !== item.key) {
      throw new Error(`Physical input key/path mapping is ambiguous or noncanonical: ${item.key}`);
    }
    if (inspection.resolveDir !== path.dirname(canonicalPath)) {
      throw new Error(`Physical input resolveDir drifted from its canonical parent: ${item.key}`);
    }
    const expectedLoader = loaderForPath(canonicalPath);
    if (inspection.loader !== expectedLoader) throw new Error(`Physical input loader drifted for ${item.key}.`);
    if (!(inspection.contents instanceof Uint8Array)) throw new Error(`Physical input inspector returned no bytes for ${item.key}.`);
    const bytes = Uint8Array.from(inspection.contents);
    const digest = sha256(bytes);
    if (!SHA256_PATTERN.test(inspection.sha256) || digest !== inspection.sha256) {
      throw new Error(`Physical input hash/digest mismatch for ${item.key}.`);
    }
    const identity = `${String(inspection.device)}:${String(inspection.inode)}`;
    const priorCanonical = keyToCanonical.get(item.key);
    if (priorCanonical && priorCanonical !== canonicalPath) throw new Error(`Metafile key maps to multiple physical files: ${item.key}`);
    const priorKey = canonicalToKey.get(canonicalPath);
    if (priorKey && priorKey !== item.key) throw new Error(`Multiple metafile keys map to one physical path: ${canonicalPath}`);
    const priorIdentity = identityToKey.get(identity);
    if (priorIdentity && priorIdentity !== item.key) throw new Error(`Physical input inode identity alias: ${priorIdentity} and ${item.key}`);
    keyToCanonical.set(item.key, canonicalPath);
    canonicalToKey.set(canonicalPath, item.key);
    identityToKey.set(identity, item.key);
    result.push(freezeBinding({
      namespace: "file",
      key: item.key,
      path: canonicalPath,
      canonicalPath,
      loader: inspection.loader,
      resolveDir: inspection.resolveDir,
      byteLength: bytes.byteLength,
      sha256: digest,
      device: inspection.device,
      inode: inspection.inode,
      entries: [],
      imports: item.input.imports,
    }, bytes));
  }
  return result;
}

export function mergeBoundInputs(
  entryBindings: readonly { readonly entry: "engine" | "main"; readonly inputs: readonly BoundMetafileInput[] }[],
): readonly BoundMetafileInput[] {
  const seenPerEntry = new Set<string>();
  const merged = new Map<string, { binding: BoundMetafileInput; entries: Set<"engine" | "main"> }>();
  const canonicalToKey = new Map<string, string>();
  const identityToKey = new Map<string, string>();
  for (const entryGroup of entryBindings) {
    if (entryGroup.entry !== "engine" && entryGroup.entry !== "main") throw new Error("Unknown metafile entry membership.");
    for (const input of entryGroup.inputs) {
      const membershipKey = `${entryGroup.entry}\0${input.key}`;
      if (seenPerEntry.has(membershipKey)) throw new Error(`Duplicate metafile binding for ${entryGroup.entry}:${input.key}`);
      seenPerEntry.add(membershipKey);
      if (input.namespace === VIRTUAL_NAMESPACE) {
        if (input.entries.length !== 1 || input.entries[0] !== entryGroup.entry) {
          throw new Error(`Virtual input entry membership mismatch for ${entryGroup.entry}.`);
        }
      } else if (input.entries.length !== 0) {
        throw new Error(`Physical input binding unexpectedly carries virtual entry metadata: ${input.key}`);
      }
      const previous = merged.get(input.key);
      if (previous) {
        if (
          previous.binding.namespace !== input.namespace ||
          previous.binding.path !== input.path ||
          previous.binding.canonicalPath !== input.canonicalPath ||
          previous.binding.loader !== input.loader ||
          previous.binding.resolveDir !== input.resolveDir ||
          previous.binding.byteLength !== input.byteLength ||
          previous.binding.sha256 !== input.sha256 ||
          previous.binding.device !== input.device ||
          previous.binding.inode !== input.inode ||
          !Buffer.from(previous.binding.bytes).equals(Buffer.from(input.bytes))
        ) {
          throw new Error(`Shared input bytes/identity/tuple drift between build entries: ${input.key}`);
        }
        if (stableJson(previous.binding.imports) !== stableJson(input.imports)) {
          throw new Error(`Shared input import-edge drift between build entries: ${input.key}`);
        }
        previous.entries.add(entryGroup.entry);
      } else {
        merged.set(input.key, { binding: input, entries: new Set([entryGroup.entry]) });
      }
      if (input.namespace === "file") {
        const oldKey = canonicalToKey.get(input.canonicalPath);
        if (oldKey && oldKey !== input.key) throw new Error(`Physical canonical-path alias across entries: ${oldKey} and ${input.key}`);
        canonicalToKey.set(input.canonicalPath, input.key);
        const identity = `${String(input.device)}:${String(input.inode)}`;
        const identityKey = identityToKey.get(identity);
        if (identityKey && identityKey !== input.key) throw new Error(`Physical inode alias across entries: ${identityKey} and ${input.key}`);
        identityToKey.set(identity, input.key);
      }
    }
  }
  const virtual = merged.get(VIRTUAL_KEY);
  if (!virtual || virtual.entries.size !== 2 || !virtual.entries.has("engine") || !virtual.entries.has("main")) {
    throw new Error("Exact virtual input must be present with identical bytes and explicit engine/main membership.");
  }
  return Object.freeze([...merged.entries()].map(([key, value]) => freezeBinding({
    ...value.binding,
    entries: [...value.entries].sort(compareText),
  }, value.binding.bytes)).sort((left, right) => compareText(left.key, right.key)));
}

function isContained(base: string, candidate: string): boolean {
  const relative = path.relative(base, candidate);
  return relative !== "" && relative !== ".." && !relative.startsWith(`..${path.sep}`) && !path.isAbsolute(relative);
}

async function lstatOrNull(io: EvidencePersistenceIo, target: string): Promise<EvidenceStatLike | null> {
  try { return await io.lstat(target); }
  catch (error) {
    if (typeof error === "object" && error !== null && (error as NodeJS.ErrnoException).code === "ENOENT") return null;
    throw error;
  }
}

function modeOf(stat: EvidenceStatLike): number {
  return stat.mode & 0o777;
}

function errno(error: unknown): string | undefined {
  return typeof error === "object" && error !== null ? (error as NodeJS.ErrnoException).code : undefined;
}

function makeEvidencePersistence(io: EvidencePersistenceIo): EvidencePersistence {
  let childCreated = false;
  let inputsCreated = false;
  let preGuardWritten = false;
  let guardErrorWritten = false;
  let receiptWritten = false;
  let completed = false;
  let expectedBuildId: string | null = null;
  let expectedInputManifestSha256: string | null = null;
  let persistedGuardError: ExpectedGuardError | null = null;
  let persistedReceiptBytes: Buffer | null = null;
  let expectedBlobs: readonly string[] = [];

  const assertWritable = (): void => {
    if (completed) throw new Error("Evidence persistence is already complete.");
  };

  const verifyDirectory = async (directory: string, mode: number): Promise<EvidenceStatLike> => {
    const stat = await io.lstat(directory);
    if (!stat.isDirectory() || stat.isSymbolicLink()) throw new Error(`Evidence path is not a real directory: ${directory}`);
    if (stat.uid !== process.getuid() || modeOf(stat) !== mode) throw new Error(`Evidence directory owner/mode mismatch: ${directory}`);
    if (await io.realpath(directory) !== directory) throw new Error(`Evidence directory canonical path mismatch: ${directory}`);
    return stat;
  };

  const writeExclusive = async (relative: string, bytes: Uint8Array): Promise<void> => {
    assertWritable();
    const isInputBlob = /^inputs\/[0-9a-f]{64}\.bin$/.test(relative);
    if (!ALL_TOP_LEVEL_FILE_NAMES.includes(relative) && !PRE_GUARD_FILE_NAMES.includes(relative) && !isInputBlob) {
      throw new Error(`Evidence file is outside the exact allowlist: ${relative}`);
    }
    if (isInputBlob && !inputsCreated) throw new Error("Evidence input directory has not been exclusively created.");
    const absolute = path.join(ATTEMPT_3_DIRECTORY, relative);
    if (!isContained(ATTEMPT_3_DIRECTORY, absolute)) throw new Error(`Evidence path escapes the attempt-3 child: ${relative}`);
    const handle = await io.open(absolute, "wx", 0o600);
    try { await handle.writeFile(Uint8Array.from(bytes)); }
    finally { await handle.close(); }
    const stat = await io.lstat(absolute);
    if (!stat.isFile() || stat.isSymbolicLink() || stat.uid !== process.getuid() || modeOf(stat) !== 0o600) {
      throw new Error(`Evidence file owner/mode/type mismatch: ${relative}`);
    }
  };

  const readInventory = async (includeComplete: boolean): Promise<EvidenceFileHash[]> => {
    await verifyDirectory(ATTEMPT_3_DIRECTORY, 0o700);
    const topEntries = await io.readdir(ATTEMPT_3_DIRECTORY, { withFileTypes: true });
    const expectedNames = [
      ...PRE_GUARD_FILE_NAMES,
      ...(receiptWritten ? ["receipt.json"] : []),
      ...(guardErrorWritten ? ["guard-error.json"] : []),
      "inputs",
      ...(includeComplete ? ["COMPLETE.json"] : []),
    ].sort(compareText);
    const actualNames = topEntries.map((entry) => entry.name).sort(compareText);
    if (stableJson(actualNames) !== stableJson(expectedNames)) throw new Error("Evidence child has an unexpected top-level inventory.");
    const inputDir = path.join(ATTEMPT_3_DIRECTORY, "inputs");
    if (inputsCreated) await verifyDirectory(inputDir, 0o700);
    const inputEntries = await io.readdir(inputDir, { withFileTypes: true });
    const actualBlobs = inputEntries.map((entry) => entry.name).sort(compareText);
    const expectedBlobNames = [...expectedBlobs].map((item) => path.basename(item)).sort(compareText);
    if (stableJson(actualBlobs) !== stableJson(expectedBlobNames)) throw new Error("Evidence input blob inventory differs from inputs.json.");
    for (const entry of [...topEntries, ...inputEntries]) {
      if (entry.isSymbolicLink()) throw new Error(`Symlink in evidence tree: ${entry.name}`);
      const isInputsDirectory = entry.name === "inputs" && topEntries.includes(entry);
      if (isInputsDirectory ? !entry.isDirectory() : !entry.isFile()) throw new Error(`Nonregular evidence entry: ${entry.name}`);
    }
    const relativePaths = [
      ...PRE_GUARD_FILE_NAMES,
      ...(receiptWritten ? ["receipt.json"] : []),
      ...(guardErrorWritten ? ["guard-error.json"] : []),
      ...expectedBlobs,
      ...(includeComplete ? ["COMPLETE.json"] : []),
    ].sort(compareText);
    const hashes: EvidenceFileHash[] = [];
    for (const relative of relativePaths) {
      const absolute = path.join(ATTEMPT_3_DIRECTORY, relative);
      const stat = await io.lstat(absolute);
      if (!stat.isFile() || stat.isSymbolicLink() || stat.uid !== process.getuid() || modeOf(stat) !== 0o600) {
        throw new Error(`Evidence inventory contains an invalid file: ${relative}`);
      }
      const bytes = await io.readFile(absolute);
      hashes.push({ path: relative, byteLength: bytes.byteLength, sha256: sha256(bytes) });
    }
    return hashes;
  };

  return {
    async createChild(): Promise<void> {
      if (childCreated) throw new Error("Attempt-3 evidence child creation was already attempted.");
      const uid = process.getuid();
      const parent = path.dirname(APPROVED_EVIDENCE_ROOT);
      const parentStat = await io.lstat(parent);
      if (!parentStat.isDirectory() || parentStat.isSymbolicLink() || parentStat.uid !== uid || modeOf(parentStat) !== 0o700) {
        throw new Error("Approved evidence root parent must be an owner-only 0700 nonsymlink directory owned by the current uid.");
      }
      if (await io.realpath(parent) !== parent) throw new Error("Approved evidence root parent canonical path mismatch.");
      const rootStat = await io.lstat(APPROVED_EVIDENCE_ROOT);
      if (!rootStat.isDirectory() || rootStat.isSymbolicLink() || rootStat.uid !== uid || (modeOf(rootStat) & 0o022) !== 0) {
        throw new Error("Approved evidence root must be an owned, nonsymlink directory that is not group/world writable.");
      }
      if (await io.realpath(APPROVED_EVIDENCE_ROOT) !== APPROVED_EVIDENCE_ROOT) throw new Error("Approved evidence root canonical path mismatch.");
      if (!isContained(parent, APPROVED_EVIDENCE_ROOT)) throw new Error("Approved evidence root escaped its verified parent.");
      if (await lstatOrNull(io, ATTEMPT_3_DIRECTORY)) throw new Error("Attempt-3 evidence child already exists in any form; refusing reuse.");
      await io.mkdir(ATTEMPT_3_DIRECTORY, { recursive: false, mode: 0o700 });
      childCreated = true;
      await verifyDirectory(ATTEMPT_3_DIRECTORY, 0o700);
      if (!isContained(APPROVED_EVIDENCE_ROOT, ATTEMPT_3_DIRECTORY)) throw new Error("Attempt-3 evidence child escaped approved root B.");
    },

    async persistBeforeGuard(payload: EvidencePersistencePayload): Promise<void> {
      assertWritable();
      if (!childCreated || preGuardWritten) throw new Error("Evidence child is missing or pre-guard persistence already ran.");
      const fileNames = Object.keys(payload.files).sort(compareText);
      if (stableJson(fileNames) !== stableJson([...PRE_GUARD_FILE_NAMES].sort(compareText))) {
        throw new Error("Pre-guard evidence file set differs from the exact allowlist.");
      }
      if (!(payload.files["inputs.json"] instanceof Uint8Array)) throw new TypeError("inputs.json must contain serialized bytes.");
      const parsedInputManifest: unknown = JSON.parse(Buffer.from(payload.files["inputs.json"]).toString("utf8"));
      if (!isPlainRecord(parsedInputManifest)) throw new TypeError("inputs.json must be a plain JSON object.");
      const inputManifest = parsedInputManifest;
      const blobs = new Map<string, Buffer>();
      for (const blob of payload.blobs) {
        if (!blob || !(blob.bytes instanceof Uint8Array)) throw new TypeError("Input blob must contain raw bytes.");
        if (!SHA256_PATTERN.test(blob.sha256) || sha256(blob.bytes) !== blob.sha256) throw new Error("Input blob digest mismatch.");
        const bytes = Buffer.from(blob.bytes);
        const previous = blobs.get(blob.sha256);
        if (previous && !previous.equals(bytes)) throw new Error("Duplicate input blob digest has different bytes.");
        blobs.set(blob.sha256, bytes);
      }
      const binding = validatePersistedEvidenceBindings(payload, inputManifest, blobs);
      const inventory = inputManifest.blobInventory as Array<{ path?: unknown; sha256?: unknown; byteLength?: unknown }>;
      if (!Array.isArray(inventory)) throw new Error("inputs.json has no blob inventory.");
      const references = inventory.map((item) => {
        if (
          !item || typeof item !== "object" ||
          typeof item.sha256 !== "string" || !SHA256_PATTERN.test(item.sha256) ||
          item.path !== `inputs/${item.sha256}.bin` ||
          !Number.isSafeInteger(item.byteLength) || (item.byteLength as number) < 0
        ) throw new Error("inputs.json contains an invalid input-blob reference.");
        const bytes = blobs.get(item.sha256);
        if (!bytes || bytes.byteLength !== item.byteLength || sha256(bytes) !== item.sha256) {
          throw new Error(`inputs.json references an absent or mismatched input blob: ${item.sha256}`);
        }
        return item.path;
      }).sort(compareText);
      if (new Set(references).size !== references.length || references.length !== blobs.size) {
        throw new Error("Input blobs are not uniquely referenced by inputs.json.");
      }
      expectedBuildId = binding.buildId;
      expectedInputManifestSha256 = binding.actualInputManifestSha256;
      await io.mkdir(path.join(ATTEMPT_3_DIRECTORY, "inputs"), { recursive: false, mode: 0o700 });
      inputsCreated = true;
      await verifyDirectory(path.join(ATTEMPT_3_DIRECTORY, "inputs"), 0o700);
      expectedBlobs = Object.freeze(references);
      for (const relative of references) {
        const digest = path.basename(relative, ".bin");
        await writeExclusive(relative, blobs.get(digest)!);
      }
      for (const name of PRE_GUARD_FILE_NAMES) await writeExclusive(name, payload.files[name]);
      preGuardWritten = true;
    },

    async persistFinalReceipt(bytes: Uint8Array, error: ExpectedGuardError): Promise<void> {
      assertWritable();
      if (!preGuardWritten || !guardErrorWritten || receiptWritten || !persistedGuardError) {
        throw new Error("Final receipt requires completed pre-guard persistence and the caught expected guard error, and may be written only once.");
      }
      if (
        error.name !== persistedGuardError.name ||
        error.message !== persistedGuardError.message ||
        error.stack !== persistedGuardError.stack
      ) {
        throw new Error("Final receipt guard error differs from the persisted caught guard error.");
      }
      const receiptValue: unknown = JSON.parse(Buffer.from(bytes).toString("utf8"));
      if (!isPlainRecord(receiptValue)) throw new Error("Final receipt must be a plain JSON object.");
      const receiptStatus = receiptValue.status ??
        (isPlainRecord(receiptValue.guard) ? receiptValue.guard.status : undefined);
      if (receiptStatus !== "guard failed / evidence only") {
        throw new Error("Final receipt must explicitly remain evidence-only with the expected guard failure.");
      }
      if (
        receiptValue.schemaVersion !== 2 || receiptValue.workId !== "W69" || receiptValue.attempt !== 3 ||
        receiptValue.root !== APPROVED_EVIDENCE_ROOT || receiptValue.child !== ATTEMPT_3_DIRECTORY ||
        typeof receiptValue.buildId !== "string" || receiptValue.buildId.length === 0 ||
        (expectedBuildId !== null && receiptValue.buildId !== expectedBuildId) ||
        receiptValue.actualInputManifestSha256 !== expectedInputManifestSha256
      ) {
        throw new Error("Final receipt work, attempt or buildId binding is invalid.");
      }
      const receiptGuard = isPlainRecord(receiptValue.guardError)
        ? receiptValue.guardError
        : isPlainRecord(receiptValue.guard)
          ? receiptValue.guard
          : null;
      if (
        !receiptGuard || receiptGuard.name !== persistedGuardError.name ||
        receiptGuard.message !== persistedGuardError.message ||
        receiptGuard.name !== error.name || receiptGuard.message !== error.message ||
        receiptGuard.stack !== error.stack || receiptGuard.stack !== persistedGuardError.stack
      ) {
        throw new Error("Final receipt does not bind the persisted expected guard error.");
      }
      await writeExclusive("receipt.json", bytes);
      receiptWritten = true;
      persistedReceiptBytes = Buffer.from(bytes);
    },

    async persistGuardError(error: ExpectedGuardError): Promise<void> {
      assertWritable();
      if (!preGuardWritten || guardErrorWritten || receiptWritten) throw new Error("Expected guard error cannot be persisted before outputs, after a receipt, or more than once.");
      if (
        !error || typeof error.name !== "string" || typeof error.message !== "string" ||
        typeof error.stack !== "string" || error.stack.length === 0
      ) {
        throw new Error("Expected guard error requires exact name, message and stack.");
      }
      if (error.message !== EXPECTED_GUARD_MESSAGE) throw new Error("Unexpected ownership guard message; evidence remains incomplete.");
      const payload = {
        schemaVersion: 2,
        status: "expected ownership rejection / evidence only",
        workId: "W69",
        attempt: 3,
        root: APPROVED_EVIDENCE_ROOT,
        child: ATTEMPT_3_DIRECTORY,
        name: error.name,
        message: error.message,
        stack: error.stack,
        s5_3Pass: false,
      };
      await writeExclusive("guard-error.json", Buffer.from(`${JSON.stringify(payload, null, 2)}\n`));
      guardErrorWritten = true;
      persistedGuardError = Object.freeze({
        name: error.name,
        message: error.message,
        stack: error.stack,
      });
    },

    async writeComplete(
      error: ExpectedGuardError,
      beforeComplete: () => Promise<void>,
    ): Promise<CompleteReceipt> {
      if (typeof beforeComplete !== "function") {
        throw new TypeError("COMPLETE requires a final old-attempt verification callback.");
      }
      assertWritable();
      if (!preGuardWritten || !guardErrorWritten || !receiptWritten || !persistedGuardError) {
        throw new Error("COMPLETE requires persisted outputs, the caught expected guard error and the final receipt.");
      }
      if (
        error.name !== persistedGuardError.name || error.message !== persistedGuardError.message ||
        error.stack !== persistedGuardError.stack || error.message !== EXPECTED_GUARD_MESSAGE
      ) throw new Error("COMPLETE cannot bind an unexpected ownership result.");
      const files = await readInventory(false);
      const receiptBytes = Buffer.from(await io.readFile(path.join(ATTEMPT_3_DIRECTORY, "receipt.json")));
      if (!persistedReceiptBytes || !receiptBytes.equals(persistedReceiptBytes)) {
        throw new Error("Final receipt readback differs from the exclusively persisted receipt bytes.");
      }
      const receiptValue: unknown = JSON.parse(receiptBytes.toString("utf8"));
      const receiptStatus = isPlainRecord(receiptValue)
        ? receiptValue.status ?? (isPlainRecord(receiptValue.guard) ? receiptValue.guard.status : undefined)
        : undefined;
      const receiptGuard = isPlainRecord(receiptValue) && isPlainRecord(receiptValue.guardError)
        ? receiptValue.guardError
        : isPlainRecord(receiptValue) && isPlainRecord(receiptValue.guard)
          ? receiptValue.guard
          : null;
      if (
        !isPlainRecord(receiptValue) || receiptStatus !== "guard failed / evidence only" ||
        receiptValue.schemaVersion !== 2 || receiptValue.workId !== "W69" || receiptValue.attempt !== 3 ||
        receiptValue.root !== APPROVED_EVIDENCE_ROOT || receiptValue.child !== ATTEMPT_3_DIRECTORY ||
        typeof receiptValue.buildId !== "string" || receiptValue.buildId.length === 0 ||
        (expectedBuildId !== null && receiptValue.buildId !== expectedBuildId) ||
        receiptValue.actualInputManifestSha256 !== expectedInputManifestSha256 ||
        !receiptGuard || receiptGuard.name !== error.name || receiptGuard.message !== error.message ||
        receiptGuard.stack !== error.stack
      ) throw new Error("Final receipt readback does not verify against the persisted expected guard error and build binding.");
      const completePayload = {
        schemaVersion: 2,
        status: "guard failed / evidence only" as const,
        workId: "W69",
        attempt: 3,
        root: APPROVED_EVIDENCE_ROOT,
        child: ATTEMPT_3_DIRECTORY,
        buildId: receiptValue.buildId,
        expectedGuardError: { name: error.name, message: error.message },
        files,
      };
      const completeBytes = Buffer.from(`${JSON.stringify(completePayload, null, 2)}\n`);
      await beforeComplete();
      await writeExclusive("COMPLETE.json", completeBytes);
      completed = true;
      const finalFiles = await readInventory(true);
      const receiptReadback = JSON.parse(Buffer.from(await io.readFile(path.join(ATTEMPT_3_DIRECTORY, "COMPLETE.json"))).toString("utf8")) as typeof completePayload;
      if (stableJson(receiptReadback.files) !== stableJson(files)) throw new Error("COMPLETE file hash list changed after exclusive write.");
      const priorHashes = finalFiles.filter((file) => file.path !== "COMPLETE.json");
      if (stableJson(priorHashes) !== stableJson(files)) throw new Error("Evidence file hashes changed after COMPLETE write.");
      const completeFile = finalFiles.find((file) => file.path === "COMPLETE.json");
      if (!completeFile) throw new Error("COMPLETE.json is absent after its exclusive write.");
      return { status: "guard failed / evidence only", files: Object.freeze(files), completeSha256: completeFile.sha256 };
    },

    async writeIncomplete(error: unknown): Promise<void> {
      if (!childCreated || completed) return;
      const failure = {
        name: error instanceof Error ? error.name : "Error",
        message: error instanceof Error ? error.message : String(error),
        stack: error instanceof Error ? error.stack ?? "" : "",
      };
      const guardDiagnostic = Buffer.from(`${JSON.stringify({ schemaVersion: 2, workId: "W69", attempt: 3, root: APPROVED_EVIDENCE_ROOT, child: ATTEMPT_3_DIRECTORY, status: "incomplete / evidence only", failure, complete: false }, null, 2)}\n`);
      const receiptDiagnostic = Buffer.from(`${JSON.stringify({ schemaVersion: 2, workId: "W69", attempt: 3, root: APPROVED_EVIDENCE_ROOT, child: ATTEMPT_3_DIRECTORY, status: "incomplete / evidence only", failure, complete: false }, null, 2)}\n`);
      for (const [name, bytes] of [["guard-error.json", guardDiagnostic], ["receipt.json", receiptDiagnostic]] as const) {
        const target = path.join(ATTEMPT_3_DIRECTORY, name);
        if (await lstatOrNull(io, target)) continue;
        try { await writeExclusive(name, bytes); }
        catch (writeError) { if (errno(writeError) !== "EEXIST") throw writeError; }
      }
    },
  };
}

/** Test-only persistence seam; it accepts an IO adapter but never an evidence path. */
export function createEvidencePersistenceForTesting(io: EvidencePersistenceIo): EvidencePersistence {
  if (!io || io === defaultEvidenceIo) throw new TypeError("Persistence tests require a non-production in-memory IO adapter.");
  return makeEvidencePersistence(io);
}

async function verifyAttempt1Unchanged(): Promise<Readonly<Record<string, string>>> {
  const uid = process.getuid();
  const stat = await fs.lstat(ATTEMPT_1_DIRECTORY);
  if (!stat.isDirectory() || stat.isSymbolicLink() || stat.uid !== uid || (stat.mode & 0o777) !== 0o700) {
    throw new Error("Attempt-1 evidence directory is missing or has changed ownership/mode/type.");
  }
  if (await fs.realpath(ATTEMPT_1_DIRECTORY) !== ATTEMPT_1_DIRECTORY) throw new Error("Attempt-1 evidence path is not canonical.");
  const names = (await fs.readdir(ATTEMPT_1_DIRECTORY)).sort(compareText);
  if (stableJson(names) !== stableJson(["guard-error.json", "receipt.json"])) {
    throw new Error("Attempt-1 evidence inventory changed; refusing to continue.");
  }
  const hashes: Record<string, string> = {};
  for (const [name, expectedHash] of Object.entries(ATTEMPT_1_HASHES)) {
    const file = path.join(ATTEMPT_1_DIRECTORY, name);
    const fileStat = await fs.lstat(file);
    if (!fileStat.isFile() || fileStat.isSymbolicLink() || fileStat.uid !== uid || (fileStat.mode & 0o777) !== 0o600) {
      throw new Error(`Attempt-1 evidence file changed ownership/mode/type: ${name}`);
    }
    const digest = sha256(await fs.readFile(file));
    if (digest !== expectedHash) throw new Error(`Attempt-1 evidence file changed hash: ${name}`);
    hashes[name] = digest;
  }
  return Object.freeze(hashes);
}

export async function verifyAttempt2Unchanged(
  io: PreservedAttemptReadIo,
): Promise<readonly EvidenceFileHash[]> {
  if (
    !io || typeof io.lstat !== "function" || typeof io.realpath !== "function" ||
    typeof io.readdir !== "function" || typeof io.readFileFingerprint !== "function"
  ) throw new TypeError("Attempt-2 verification requires read-only filesystem observations.");
  if (
    !path.isAbsolute(ATTEMPT_2_DIRECTORY) || path.resolve(ATTEMPT_2_DIRECTORY) !== ATTEMPT_2_DIRECTORY ||
    path.dirname(ATTEMPT_2_DIRECTORY) !== APPROVED_EVIDENCE_ROOT
  ) throw new Error("Attempt-2 evidence directory is not the exact canonical child of approved root B.");

  const assertStat = (
    stat: PreservedAttemptStat,
    label: string,
    expectedType: "directory" | "file",
    expectedMode: number,
  ): void => {
    if (
      !stat || !Number.isSafeInteger(stat.size) || stat.size < 0 ||
      !Number.isSafeInteger(stat.uid) || !Number.isSafeInteger(stat.gid) ||
      !Number.isSafeInteger(stat.mode)
    ) throw new Error(`Attempt-2 stat metadata is malformed: ${label}`);
    const validType = expectedType === "directory"
      ? stat.isDirectory() && !stat.isFile()
      : stat.isFile() && !stat.isDirectory();
    if (!validType || stat.isSymbolicLink()) throw new Error(`Attempt-2 path has the wrong type or is a symlink: ${label}`);
    if (stat.uid !== 501 || stat.gid !== 20 || (stat.mode & 0o777) !== expectedMode) {
      throw new Error(`Attempt-2 path owner, group or mode differs from its fixed policy: ${label}`);
    }
  };

  const directoryStat = await io.lstat(ATTEMPT_2_DIRECTORY);
  assertStat(directoryStat, ATTEMPT_2_DIRECTORY, "directory", 0o700);
  if (await io.realpath(ATTEMPT_2_DIRECTORY) !== ATTEMPT_2_DIRECTORY) {
    throw new Error("Attempt-2 evidence directory is not at its exact canonical path.");
  }

  const entries = await io.readdir(ATTEMPT_2_DIRECTORY);
  if (!Array.isArray(entries) || entries.some((entry) => !entry || typeof entry.name !== "string")) {
    throw new Error("Attempt-2 evidence inventory is malformed.");
  }
  const expectedNames = Object.keys(ATTEMPT_2_HASHES).sort(compareText);
  const actualNames = entries.map((entry) => entry.name).sort(compareText);
  if (stableJson(actualNames) !== stableJson(expectedNames)) {
    throw new Error("Attempt-2 evidence inventory differs from the exact preserved two-file set.");
  }

  const fingerprints: EvidenceFileHash[] = [];
  for (const name of expectedNames) {
    const entry = entries.find((candidate) => candidate.name === name)!;
    if (!entry.isFile() || entry.isDirectory() || entry.isSymbolicLink()) {
      throw new Error(`Attempt-2 inventory entry is not a regular nonsymlink file: ${name}`);
    }
    const file = path.join(ATTEMPT_2_DIRECTORY, name);
    if (path.resolve(file) !== file || path.dirname(file) !== ATTEMPT_2_DIRECTORY) {
      throw new Error(`Attempt-2 evidence file path is not canonical: ${name}`);
    }
    const stat = await io.lstat(file);
    assertStat(stat, name, "file", 0o600);
    if (await io.realpath(file) !== file) throw new Error(`Attempt-2 evidence file is not at its exact canonical path: ${name}`);

    const fingerprint = await io.readFileFingerprint(file);
    const expectedLength = ATTEMPT_2_LENGTHS[name as keyof typeof ATTEMPT_2_LENGTHS];
    const expectedHash = ATTEMPT_2_HASHES[name as keyof typeof ATTEMPT_2_HASHES];
    if (
      !fingerprint || !Number.isSafeInteger(fingerprint.byteLength) || fingerprint.byteLength < 0 ||
      typeof fingerprint.sha256 !== "string" || !SHA256_PATTERN.test(fingerprint.sha256)
    ) throw new Error(`Attempt-2 raw-byte fingerprint is malformed: ${name}`);
    if (
      stat.size !== fingerprint.byteLength || stat.size !== expectedLength ||
      fingerprint.byteLength !== expectedLength
    ) throw new Error(`Attempt-2 evidence size differs from its fixed byte length: ${name}`);
    if (fingerprint.sha256 !== expectedHash) throw new Error(`Attempt-2 evidence hash differs from its fixed fingerprint: ${name}`);
    fingerprints.push({ path: name, byteLength: fingerprint.byteLength, sha256: fingerprint.sha256 });
  }
  return Object.freeze(fingerprints
    .sort((left, right) => compareText(left.path, right.path))
    .map((fingerprint) => Object.freeze(fingerprint)));
}

export async function inspectProjectBytes(
  projectRoot: string,
  relativePath: string,
): Promise<RawProjectFileSnapshot> {
  const canonicalRoot = await fs.realpath(projectRoot);
  const absolutePath = path.resolve(canonicalRoot, relativePath);
  if (!isContained(canonicalRoot, absolutePath)) throw new Error(`Physical input escapes project root: ${relativePath}`);
  let cursor = canonicalRoot;
  const parts = relativePath.split(/[\\/]+/);
  for (let index = 0; index < parts.length; index += 1) {
    if (!parts[index] || parts[index] === "." || parts[index] === "..") throw new Error(`Physical path is not canonical: ${relativePath}`);
    cursor = path.join(cursor, parts[index]);
    const component = await fs.lstat(cursor);
    if (component.isSymbolicLink()) throw new Error(`Symlink in physical input path: ${cursor}`);
    if (index < parts.length - 1 && !component.isDirectory()) throw new Error(`Non-directory physical input ancestor: ${cursor}`);
    if (index === parts.length - 1 && !component.isFile()) throw new Error(`Physical input is not a regular file: ${cursor}`);
  }
  const canonicalPath = await fs.realpath(absolutePath);
  if (canonicalPath !== absolutePath) throw new Error(`Physical input path is not canonical: ${relativePath}`);
  const before = await fs.lstat(absolutePath);
  const contents = await fs.readFile(absolutePath);
  const after = await fs.lstat(absolutePath);
  if (
    after.isSymbolicLink() || !after.isFile() ||
    String(before.dev) !== String(after.dev) || String(before.ino) !== String(after.ino) ||
    contents.byteLength !== after.size
  ) throw new Error(`Physical input changed during pinning: ${relativePath}`);
  return {
    projectRoot: canonicalRoot,
    canonicalPath,
    contents: Buffer.from(contents),
    byteLength: contents.byteLength,
    sha256: sha256(contents),
    device: String(after.dev),
    inode: String(after.ino),
  };
}

export async function inspectControlFile(
  projectRoot: string,
  relativePath: string,
  inspectRaw: (
    projectRoot: string,
    relativePath: string,
  ) => Promise<RawProjectFileSnapshot> = inspectProjectBytes,
): Promise<RawProjectFileSnapshot> {
  return inspectRaw(projectRoot, relativePath);
}

async function inspectProjectFile(projectRoot: string, relativePath: string): Promise<PhysicalInputInspection> {
  const snapshot = await inspectProjectBytes(projectRoot, relativePath);
  return {
    ...snapshot,
    loader: loaderForPath(snapshot.canonicalPath),
    resolveDir: path.dirname(snapshot.canonicalPath),
  };
}

async function collectProjectMetadataPaths(projectRoot: string): Promise<Set<string>> {
  const paths = new Set<string>([
    "package.json",
    "package-lock.json",
    "manifest.json",
    "styles.css",
    "esbuild.config.mjs",
    "scripts/docx-build-config.mjs",
    "scripts/read-only-zip-plugin.mjs",
    "tsconfig.json",
  ]);
  const skippedDirectories = new Set([".git", "node_modules", "dist", "release", ".test-dist", ".benchmark-dist", ".lifecycle-dist", ".zip-smoke-dist"]);
  const walk = async (absolute: string, relative: string): Promise<void> => {
    const entries = await fs.readdir(absolute, { withFileTypes: true });
    for (const entry of entries.sort((left, right) => compareText(left.name, right.name))) {
      const childRelative = relative ? `${relative}/${entry.name}` : entry.name;
      const childAbsolute = path.join(absolute, entry.name);
      if (entry.isSymbolicLink()) continue;
      if (entry.isDirectory()) {
        if (!skippedDirectories.has(entry.name)) await walk(childAbsolute, childRelative);
        continue;
      }
      if (!entry.isFile()) continue;
      if (
        /(?:^|\/)(?:package(?:-lock)?\.json|npm-shrinkwrap\.json|pnpm-lock\.yaml|yarn\.lock|tsconfig[^/]*\.json|[^/]*\.config\.[cm]?[jt]s|[^/]*-config\.[cm]?[jt]s)$/.test(childRelative) ||
        childRelative === "manifest.json" || childRelative === "styles.css"
      ) paths.add(childRelative);
    }
  };
  await walk(projectRoot, "");
  for (const required of paths) {
    await inspectProjectBytes(projectRoot, required);
  }
  return paths;
}

export async function collectControlInputManifest(
  projectRoot: string,
  buildIdInputs: readonly { readonly path: string; readonly sha256: string }[],
  physicalInputs: readonly BoundMetafileInput[],
): Promise<ControlInputManifest> {
  const categories = new Map<string, Set<ControlCategory>>();
  const add = (relativePath: string, category: ControlCategory): void => {
    const values = categories.get(relativePath) ?? new Set<ControlCategory>();
    values.add(category);
    categories.set(relativePath, values);
  };
  for (const input of buildIdInputs) {
    classifyRawInputKey(input.path, { imports: [] });
    add(input.path, "calculateBuildId");
  }
  for (const file of await collectProjectMetadataPaths(projectRoot)) add(file, "project-package-lock-config-tsconfig");

  const nodeModulesRoot = path.join(projectRoot, "node_modules");
  const nodePackagePaths = new Set<string>([
    "node_modules/esbuild/package.json",
    "node_modules/typescript/package.json",
  ]);
  for (const input of physicalInputs) {
    const relative = path.relative(nodeModulesRoot, input.canonicalPath);
    if (relative === ".." || relative.startsWith(`..${path.sep}`) || path.isAbsolute(relative)) continue;
    let cursor = path.dirname(input.canonicalPath);
    let found = false;
    while (cursor === nodeModulesRoot || isContained(nodeModulesRoot, cursor)) {
      const packageFile = path.join(cursor, "package.json");
      try {
        const stat = await fs.lstat(packageFile);
        if (stat.isSymbolicLink() || !stat.isFile()) throw new Error(`Invalid dependency package metadata: ${packageFile}`);
        nodePackagePaths.add(path.relative(projectRoot, packageFile).replaceAll("\\", "/"));
        found = true;
      } catch (error) {
        if (typeof error === "object" && error !== null && (error as NodeJS.ErrnoException).code === "ENOENT") {
          // Keep walking to the package root; scoped-directory levels need not have a manifest.
        } else throw error;
      }
      if (cursor === nodeModulesRoot) break;
      cursor = path.dirname(cursor);
    }
    if (!found) throw new Error(`Resolved dependency input has no package.json ancestor: ${input.key}`);
  }
  for (const file of nodePackagePaths) add(file, "node_modules-package-metadata");

  return Object.freeze([...categories.entries()]
    .sort(([left], [right]) => compareText(left, right))
    .map(([relativePath, labels]) => Object.freeze({
      path: relativePath,
      categories: Object.freeze([...labels].sort(compareText)),
    })));
}

const CONTROL_CATEGORIES: ReadonlySet<string> = new Set([
  "calculateBuildId",
  "project-package-lock-config-tsconfig",
  "node_modules-package-metadata",
]);

function assertCanonicalControlPath(value: unknown): asserts value is string {
  if (typeof value !== "string") throw new TypeError("Control input path must be a string.");
  const classified = classifyRawInputKey(value, { imports: [] });
  if (classified.kind !== "physical") throw new Error(`Control input path is not a physical canonical path: ${value}`);
}

function parseControlInputManifest(value: unknown, label: string): ControlInputManifest {
  if (!Array.isArray(value)) throw new TypeError(`${label} control manifest must be an array.`);
  const rows: ControlInputSpec[] = [];
  let previousPath: string | null = null;
  for (const row of value) {
    if (!isPlainRecord(row)) throw new TypeError(`${label} control manifest row must be a plain object.`);
    const keys = Reflect.ownKeys(row);
    if (
      keys.length !== 2 || !keys.includes("path") || !keys.includes("categories") ||
      keys.some((key) => typeof key !== "string")
    ) throw new Error(`${label} control manifest row has unknown or missing fields.`);
    assertCanonicalControlPath(row.path);
    if (previousPath !== null && compareText(previousPath, row.path) >= 0) {
      throw new Error(`${label} control manifest has duplicate or noncanonical row ordering: ${row.path}`);
    }
    if (!Array.isArray(row.categories) || row.categories.length === 0) {
      throw new Error(`${label} control manifest categories must be a nonempty canonical array: ${row.path}`);
    }
    const categories: ControlCategory[] = [];
    let previousCategory: string | null = null;
    for (const category of row.categories) {
      if (typeof category !== "string" || !CONTROL_CATEGORIES.has(category)) {
        throw new Error(`${label} control manifest has an unknown category for ${row.path}.`);
      }
      if (previousCategory !== null && compareText(previousCategory, category) >= 0) {
        throw new Error(`${label} control manifest has duplicate or noncanonical categories for ${row.path}.`);
      }
      categories.push(category as ControlCategory);
      previousCategory = category;
    }
    rows.push(Object.freeze({ path: row.path, categories: Object.freeze(categories) }));
    previousPath = row.path;
  }
  return Object.freeze(rows);
}

export function validateControlInputManifest(
  candidate: unknown,
  expected: ControlInputManifest,
): ControlInputManifest {
  const expectedRows = parseControlInputManifest(expected, "Expected");
  const candidateRows = parseControlInputManifest(candidate, "Candidate");
  if (stableJson(candidateRows) !== stableJson(expectedRows)) {
    throw new Error("Control input manifest membership, category or ordering differs from the frozen expected manifest.");
  }
  return candidateRows;
}

const SERIALIZED_CONTROL_FIELDS = Object.freeze([
  "role",
  "path",
  "canonicalPath",
  "device",
  "inode",
  "byteLength",
  "sha256",
  "categories",
  "compiledKeys",
  "blob",
]);

export function projectControlInputSpecs(
  rows: unknown,
  expectedManifest: ControlInputManifest,
): ControlInputManifest {
  if (!Array.isArray(rows)) throw new TypeError("Serialized control inputs must be an array.");
  const projected: ControlInputSpec[] = [];
  for (const row of rows) {
    if (!isPlainRecord(row)) throw new TypeError("Serialized control input row must be a plain object.");
    const keys = Reflect.ownKeys(row);
    if (
      keys.length !== SERIALIZED_CONTROL_FIELDS.length ||
      keys.some((key) => typeof key !== "string" || !SERIALIZED_CONTROL_FIELDS.includes(key)) ||
      SERIALIZED_CONTROL_FIELDS.some((key) => !Object.prototype.hasOwnProperty.call(row, key))
    ) throw new Error("Serialized control input row has unknown or missing fields.");
    if (row.role !== "control") throw new Error("Serialized control input has an unknown role.");
    assertCanonicalControlPath(row.path);
    if (
      typeof row.canonicalPath !== "string" || !path.isAbsolute(row.canonicalPath) ||
      path.resolve(row.canonicalPath) !== row.canonicalPath || row.canonicalPath.includes("\0") ||
      !row.canonicalPath.endsWith(`${path.sep}${row.path.split("/").join(path.sep)}`)
    ) throw new Error(`Serialized control canonical path does not match its relative path: ${String(row.path)}`);
    if (typeof row.device !== "string" || row.device.length === 0) {
      throw new Error(`Serialized control device identity is invalid: ${row.path}`);
    }
    if (typeof row.inode !== "string" || row.inode.length === 0) {
      throw new Error(`Serialized control inode identity is invalid: ${row.path}`);
    }
    if (!Number.isSafeInteger(row.byteLength) || (row.byteLength as number) < 0) {
      throw new Error(`Serialized control byte length is invalid: ${row.path}`);
    }
    if (typeof row.sha256 !== "string" || !SHA256_PATTERN.test(row.sha256)) {
      throw new Error(`Serialized control digest is invalid: ${row.path}`);
    }
    if (row.blob !== `inputs/${row.sha256}.bin`) {
      throw new Error(`Serialized control blob reference does not match its digest: ${row.path}`);
    }
    if (!Array.isArray(row.compiledKeys)) throw new Error(`Serialized control compiledKeys must be an array: ${row.path}`);
    let previousKey: string | null = null;
    for (const key of row.compiledKeys) {
      assertCanonicalControlPath(key);
      if (previousKey !== null && compareText(previousKey, key) >= 0) {
        throw new Error(`Serialized control compiledKeys are duplicate or noncanonical: ${row.path}`);
      }
      previousKey = key;
    }
    projected.push({ path: row.path, categories: row.categories as readonly ControlCategory[] });
  }
  return validateControlInputManifest(projected, expectedManifest);
}

function validateRawControlSnapshot(
  snapshot: RawProjectFileSnapshot,
  projectRoot: string,
  relativePath: string,
): RawProjectFileSnapshot {
  if (!isPlainRecord(snapshot)) throw new TypeError(`Raw control inspector returned a malformed snapshot: ${relativePath}`);
  if (
    typeof snapshot.projectRoot !== "string" || !path.isAbsolute(snapshot.projectRoot) ||
    path.resolve(snapshot.projectRoot) !== snapshot.projectRoot || snapshot.projectRoot !== projectRoot
  ) throw new Error(`Control snapshot has a noncanonical or mismatched project root: ${relativePath}`);
  const expectedPath = path.resolve(projectRoot, relativePath);
  if (
    typeof snapshot.canonicalPath !== "string" ||
    path.resolve(snapshot.canonicalPath) !== snapshot.canonicalPath ||
    snapshot.canonicalPath !== expectedPath || !isContained(projectRoot, snapshot.canonicalPath)
  ) throw new Error(`Control spec path does not exactly match its canonical snapshot path: ${relativePath}`);
  if (!(snapshot.contents instanceof Uint8Array)) throw new Error(`Control snapshot has no raw bytes: ${relativePath}`);
  if (!Number.isSafeInteger(snapshot.byteLength) || snapshot.byteLength < 0 || snapshot.byteLength !== snapshot.contents.byteLength) {
    throw new Error(`Control snapshot byte length is invalid: ${relativePath}`);
  }
  if (
    typeof snapshot.sha256 !== "string" || !SHA256_PATTERN.test(snapshot.sha256) ||
    sha256(snapshot.contents) !== snapshot.sha256
  ) throw new Error(`Control snapshot digest is invalid: ${relativePath}`);
  if (
    typeof snapshot.device !== "string" || snapshot.device.length === 0 ||
    typeof snapshot.inode !== "string" || snapshot.inode.length === 0
  ) throw new Error(`Control snapshot identity is invalid: ${relativePath}`);
  return {
    projectRoot,
    canonicalPath: snapshot.canonicalPath,
    contents: Buffer.from(snapshot.contents),
    byteLength: snapshot.byteLength,
    sha256: snapshot.sha256,
    device: snapshot.device,
    inode: snapshot.inode,
  };
}

function samePhysicalSnapshot(left: RawProjectFileSnapshot, right: RawProjectFileSnapshot): boolean {
  return left.device === right.device && left.inode === right.inode &&
    left.byteLength === right.byteLength && left.sha256 === right.sha256 &&
    Buffer.from(left.contents).equals(Buffer.from(right.contents));
}

function assertPhysicalIdentityIndex(
  projectRoot: string,
  compiledInputs: readonly BoundMetafileInput[],
  controls: readonly BoundControlInput[],
): Map<string, RawProjectFileSnapshot> {
  if (!path.isAbsolute(projectRoot) || path.resolve(projectRoot) !== projectRoot) {
    throw new Error("Control identity index requires a canonical absolute project root.");
  }
  const byIdentity = new Map<string, { canonicalPath: string; rawKey: string; snapshot: RawProjectFileSnapshot; roles: Set<string> }>();
  const byCanonicalPath = new Map<string, { rawKey: string; identity: string; snapshot: RawProjectFileSnapshot }>();
  const byRawKey = new Map<string, { canonicalPath: string; identity: string; snapshot: RawProjectFileSnapshot }>();
  const union = new Map<string, RawProjectFileSnapshot>();

  const add = (
    snapshot: RawProjectFileSnapshot,
    canonicalPath: string,
    rawKey: string,
    role: "compiled" | "control",
  ): void => {
    assertCanonicalControlPath(rawKey);
    if (path.resolve(projectRoot, rawKey) !== canonicalPath || snapshot.canonicalPath !== canonicalPath) {
      throw new Error(`Physical input path/key mapping differs in the control identity index: ${rawKey}`);
    }
    const checked = validateRawControlSnapshot(snapshot, projectRoot, rawKey);
    const identity = JSON.stringify([checked.device, checked.inode]);
    const identityPrior = byIdentity.get(identity);
    if (identityPrior) {
      if (identityPrior.canonicalPath !== canonicalPath || identityPrior.rawKey !== rawKey) {
        throw new Error(`Physical input device/inode identity alias across paths or raw keys: ${identityPrior.rawKey} and ${rawKey}`);
      }
      if (!samePhysicalSnapshot(identityPrior.snapshot, checked)) {
        throw new Error(`Physical input identity, size, digest or bytes drifted across roles: ${rawKey}`);
      }
      identityPrior.roles.add(role);
    } else {
      byIdentity.set(identity, { canonicalPath, rawKey, snapshot: checked, roles: new Set([role]) });
    }
    const canonicalPrior = byCanonicalPath.get(canonicalPath);
    if (canonicalPrior && (canonicalPrior.rawKey !== rawKey || canonicalPrior.identity !== identity)) {
      throw new Error(`Physical canonical path has an alternate raw key or identity: ${canonicalPath}`);
    }
    if (canonicalPrior && !samePhysicalSnapshot(canonicalPrior.snapshot, checked)) {
      throw new Error(`Physical canonical path snapshot drifted across roles: ${canonicalPath}`);
    }
    if (!canonicalPrior) byCanonicalPath.set(canonicalPath, { rawKey, identity, snapshot: checked });
    const rawKeyPrior = byRawKey.get(rawKey);
    if (rawKeyPrior && (rawKeyPrior.canonicalPath !== canonicalPath || rawKeyPrior.identity !== identity)) {
      throw new Error(`Physical raw key maps to an alternate canonical path or identity: ${rawKey}`);
    }
    if (rawKeyPrior && !samePhysicalSnapshot(rawKeyPrior.snapshot, checked)) {
      throw new Error(`Physical raw key snapshot drifted across roles: ${rawKey}`);
    }
    if (!rawKeyPrior) byRawKey.set(rawKey, { canonicalPath, identity, snapshot: checked });
    if (!union.has(canonicalPath)) union.set(canonicalPath, checked);
  };

  for (const input of compiledInputs) {
    if (input.namespace !== "file") continue;
    if (input.device === null || input.inode === null) throw new Error(`Physical input lacks device/inode identity: ${input.key}`);
    const snapshot: RawProjectFileSnapshot = {
      projectRoot,
      canonicalPath: input.canonicalPath,
      contents: input.bytes,
      byteLength: input.byteLength,
      sha256: input.sha256,
      device: String(input.device),
      inode: String(input.inode),
    };
    add(snapshot, input.canonicalPath, input.key, "compiled");
  }
  for (const control of controls) {
    if (control.role !== "control") throw new Error(`Bound control has an unknown role: ${String(control.role)}`);
    add(control, control.canonicalPath, control.path, "control");
  }
  return union;
}

function serializeCompiledInput(input: BoundMetafileInput): Record<string, unknown> {
  return {
    role: "compiled",
    namespace: input.namespace,
    canonicalPath: input.canonicalPath,
    key: input.key,
    loader: input.loader,
    resolveDir: input.resolveDir,
    byteLength: input.byteLength,
    sha256: input.sha256,
    device: input.device === null ? null : String(input.device),
    inode: input.inode === null ? null : String(input.inode),
    entries: input.entries,
    blob: `inputs/${input.sha256}.bin`,
  };
}

function serializeControlInput(input: BoundControlInput): SerializedControlInput {
  return {
    role: "control",
    path: input.path,
    canonicalPath: input.canonicalPath,
    device: input.device,
    inode: input.inode,
    byteLength: input.byteLength,
    sha256: input.sha256,
    categories: input.categories,
    compiledKeys: input.compiledKeys,
    blob: `inputs/${input.sha256}.bin`,
  };
}

function validatePersistedEvidenceBindings(
  payload: EvidencePersistencePayload,
  inputManifest: Record<string, unknown>,
  blobs: ReadonlyMap<string, Buffer>,
): { readonly buildId: string; readonly actualInputManifestSha256: string } {
  if (inputManifest.schemaVersion === 2 && !payload.bindingContext) {
    throw new Error("Schema-2 evidence persistence requires an independent binding context.");
  }
  if (inputManifest.schemaVersion !== 2 || inputManifest.workId !== "W69" || inputManifest.attempt !== 3) {
    throw new Error("Evidence inputs must bind W69 schema version 2, attempt 3.");
  }
  const bindingContext = payload.bindingContext;
  if (
    !bindingContext || !Array.isArray(bindingContext.expectedManifest) ||
    !Array.isArray(bindingContext.compiledInputs) || !Array.isArray(bindingContext.controls)
  ) throw new Error("Schema-2 evidence persistence requires independent expected manifest and pinned rows.");
  const { expectedManifest, compiledInputs, controls } = bindingContext;
  if (controls.length === 0) throw new Error("Schema-2 evidence persistence requires pinned control snapshots.");
  const projectRoot = controls[0].projectRoot;
  if (
    typeof projectRoot !== "string" || !path.isAbsolute(projectRoot) ||
    path.resolve(projectRoot) !== projectRoot || inputManifest.root !== projectRoot ||
    inputManifest.child !== ATTEMPT_3_DIRECTORY
  ) throw new Error("Evidence inputs root or fixed attempt-3 child binding is invalid.");
  if (
    typeof inputManifest.buildId !== "string" || inputManifest.buildId.length === 0 ||
    typeof inputManifest.actualInputManifestSha256 !== "string" ||
    !SHA256_PATTERN.test(inputManifest.actualInputManifestSha256)
  ) throw new Error("Evidence inputs buildId or role-manifest digest is invalid.");

  validateControlInputManifest(expectedManifest, expectedManifest);
  const projectedSpecs = projectControlInputSpecs(inputManifest.controlInputs, expectedManifest);
  const pinnedSpecs = validateControlInputManifest(
    controls.map((control) => ({ path: control.path, categories: control.categories })),
    expectedManifest,
  );
  if (stableJson(projectedSpecs) !== stableJson(pinnedSpecs)) {
    throw new Error("Persisted controls do not match the independently pinned control manifest.");
  }

  const pinnedControlRows = controls.map(serializeControlInput).sort((left, right) => compareText(left.path, right.path));
  if (stableJson(inputManifest.controlInputs) !== stableJson(pinnedControlRows)) {
    throw new Error("Persisted control rows differ from their independently pinned snapshots.");
  }
  if (!Array.isArray(inputManifest.inputs)) throw new Error("inputs.json has no compiled input rows.");
  const pinnedCompiledRows = compiledInputs
    .map(serializeCompiledInput)
    .sort((left, right) => compareText(left.key as string, right.key as string));
  if (stableJson(inputManifest.inputs) !== stableJson(pinnedCompiledRows)) {
    throw new Error("Persisted compiled rows differ from their independently pinned inputs.");
  }

  const expectedInputManifestSha256 = sha256(Buffer.from(stableJson({
    inputs: inputManifest.inputs,
    controlInputs: inputManifest.controlInputs,
  })));
  if (inputManifest.actualInputManifestSha256 !== expectedInputManifestSha256) {
    throw new Error("Evidence inputs role-manifest digest does not bind the serialized compiled and control rows.");
  }

  assertPhysicalIdentityIndex(projectRoot, compiledInputs, controls);
  for (const control of controls) {
    const compiledKeys = compiledInputs
      .filter((input) => input.namespace === "file" && input.canonicalPath === control.canonicalPath)
      .map((input) => input.key)
      .sort(compareText);
    if (stableJson(control.compiledKeys) !== stableJson(compiledKeys)) {
      throw new Error(`Pinned control compiled-key links differ from the physical input map: ${control.path}`);
    }
  }

  const assertPinnedBlob = (
    label: string,
    digest: string,
    byteLength: number,
    bytes: Uint8Array,
  ): void => {
    if (
      !(bytes instanceof Uint8Array) || !Number.isSafeInteger(byteLength) || byteLength < 0 ||
      bytes.byteLength !== byteLength || sha256(bytes) !== digest
    ) throw new Error(`${label} pinned bytes, size or digest are inconsistent.`);
    const blob = blobs.get(digest);
    if (!blob || blob.byteLength !== byteLength || !blob.equals(Buffer.from(bytes))) {
      throw new Error(`${label} blob differs from its independently pinned bytes: ${digest}`);
    }
  };
  for (const input of compiledInputs) {
    assertPinnedBlob(`Compiled input ${input.key}`, input.sha256, input.byteLength, input.bytes);
  }
  for (const control of controls) {
    assertPinnedBlob(`Control input ${control.path}`, control.sha256, control.byteLength, control.contents);
  }

  return {
    buildId: inputManifest.buildId,
    actualInputManifestSha256: inputManifest.actualInputManifestSha256,
  };
}

export async function bindControlInputs(
  manifest: ControlInputManifest,
  compiledInputs: readonly BoundMetafileInput[],
  inspectRaw: (relativePath: string) => Promise<RawProjectFileSnapshot>,
): Promise<readonly BoundControlInput[]> {
  if (!Array.isArray(compiledInputs) || typeof inspectRaw !== "function") {
    throw new TypeError("Control input binding requires compiled rows and a raw snapshot inspector.");
  }
  const specs = validateControlInputManifest(manifest, manifest);
  const result: BoundControlInput[] = [];
  for (const spec of specs) {
    const inspected = await inspectRaw(spec.path);
    if (!inspected || typeof inspected.projectRoot !== "string") {
      throw new Error(`Raw control inspector returned no project root: ${spec.path}`);
    }
    const projectRoot = path.resolve(inspected.projectRoot);
    if (inspected.projectRoot !== projectRoot) throw new Error(`Raw control inspector returned a noncanonical project root: ${spec.path}`);
    const snapshot = validateRawControlSnapshot(inspected, projectRoot, spec.path);
    const compiledKeys = compiledInputs
      .filter((input) => input.namespace === "file" && input.canonicalPath === snapshot.canonicalPath)
      .map((input) => input.key)
      .sort(compareText);
    result.push({
      ...snapshot,
      path: spec.path,
      categories: spec.categories,
      role: "control",
      compiledKeys: Object.freeze(compiledKeys),
    });
  }
  const bound = Object.freeze(result);
  if (bound.length > 0) assertPhysicalIdentityIndex(bound[0].projectRoot, compiledInputs, bound);
  return bound;
}

async function collectControlInputs(
  projectRoot: string,
  buildIdInputs: readonly { readonly path: string; readonly sha256: string }[],
  physicalInputs: readonly BoundMetafileInput[],
  manifest: ControlInputManifest,
): Promise<readonly BoundControlInput[]> {
  const result = await bindControlInputs(
    manifest,
    physicalInputs,
    (relativePath) => inspectProjectBytes(projectRoot, relativePath),
  );
  const buildIdHashMap = new Map(buildIdInputs.map((item) => [item.path, item.sha256]));
  for (const item of result) {
    const expectedHash = buildIdHashMap.get(item.path);
    if (expectedHash && item.sha256 !== expectedHash) {
      throw new Error(`calculateBuildId input changed during discovery: ${item.path}`);
    }
  }
  return result;
}

export function assertControlSnapshotEqual(
  expected: RawProjectFileSnapshot,
  actual: RawProjectFileSnapshot,
): void {
  if (
    expected.sha256 !== actual.sha256 ||
    !Buffer.from(expected.contents).equals(Buffer.from(actual.contents))
  ) {
    throw new Error(`Physical input and control snapshot differ: ${actual.canonicalPath}`);
  }
}

export async function recheckPinnedSnapshots(
  projectRoot: string,
  compiledInputs: readonly BoundMetafileInput[],
  controls: readonly BoundControlInput[],
  expectedManifest: ControlInputManifest,
  enumerateControls: () => Promise<ControlInputManifest>,
  inspectRaw: (relativePath: string) => Promise<RawProjectFileSnapshot>,
): Promise<void> {
  if (!Array.isArray(compiledInputs) || !Array.isArray(controls) || typeof enumerateControls !== "function" || typeof inspectRaw !== "function") {
    throw new TypeError("Snapshot recheck requires pinned rows and independent enumeration/raw-inspection callbacks.");
  }
  if (!path.isAbsolute(projectRoot) || path.resolve(projectRoot) !== projectRoot) {
    throw new Error("Snapshot recheck requires a canonical absolute project root.");
  }

  const frozenExpectedManifest = validateControlInputManifest(expectedManifest, expectedManifest);
  const enumeratedManifest = await enumerateControls();
  validateControlInputManifest(enumeratedManifest, frozenExpectedManifest);

  const projectedControls = controls.map((control) => {
    if (!control || typeof control !== "object" || control.role !== "control") {
      throw new Error("Bound control snapshot has an unknown role.");
    }
    return { path: control.path, categories: control.categories };
  });
  validateControlInputManifest(projectedControls, frozenExpectedManifest);

  const expected = assertPhysicalIdentityIndex(projectRoot, compiledInputs, controls);
  for (const [canonicalPath, snapshot] of expected) {
    const relativePath = path.relative(projectRoot, canonicalPath).replaceAll("\\", "/");
    const current = await inspectRaw(relativePath);
    if (
      current.projectRoot !== projectRoot || current.canonicalPath !== canonicalPath ||
      !(current.contents instanceof Uint8Array) ||
      String(current.device) !== snapshot.device || String(current.inode) !== snapshot.inode ||
      current.byteLength !== snapshot.byteLength || current.contents.byteLength !== snapshot.byteLength ||
      current.sha256 !== snapshot.sha256 || sha256(current.contents) !== current.sha256 ||
      !Buffer.from(current.contents).equals(Buffer.from(snapshot.contents))
    ) throw new Error(`Pinned physical/control input changed between builds: ${canonicalPath}`);
  }
}

async function recheckSnapshots(
  projectRoot: string,
  compiledInputs: readonly BoundMetafileInput[],
  controls: readonly BoundControlInput[],
  expectedManifest: ControlInputManifest,
  enumerateControls: () => Promise<ControlInputManifest>,
): Promise<void> {
  await recheckPinnedSnapshots(
    projectRoot,
    compiledInputs,
    controls,
    expectedManifest,
    enumerateControls,
    (relativePath) => inspectProjectBytes(projectRoot, relativePath),
  );
}

function normalizeMetaPath(projectRoot: string, value: string): string {
  if (value.includes("\0") || value.includes("?")) throw new Error(`Metafile path contains NUL/query suffix: ${value}`);
  if (path.isAbsolute(value)) {
    const absolute = path.resolve(value);
    if (absolute !== projectRoot && !isContained(projectRoot, absolute)) throw new Error(`Metafile path escapes project root: ${value}`);
    return path.relative(projectRoot, absolute).replaceAll("\\", "/");
  }
  return value.replaceAll("\\", "/");
}

function normalizedMetafile(projectRoot: string, metafile: { readonly inputs: Record<string, unknown>; readonly outputs: Record<string, unknown> }): unknown {
  const inputs: Record<string, unknown> = Object.create(null) as Record<string, unknown>;
  for (const key of Object.keys(metafile.inputs).sort(compareText)) {
    if (key !== VIRTUAL_KEY) classifyRawInputKey(key, { imports: [] });
    const input = metafile.inputs[key] as Record<string, unknown>;
    const edges = (input.imports as Array<Record<string, unknown>>).map((edge) => ({
      ...edge,
      path: typeof edge.path === "string" ? normalizeMetaPath(projectRoot, edge.path) : edge.path,
    })).sort((left, right) => compareText(stableJson(left), stableJson(right)));
    inputs[key] = { ...input, imports: edges };
  }
  const outputs: Record<string, unknown> = Object.create(null) as Record<string, unknown>;
  for (const key of Object.keys(metafile.outputs).sort(compareText)) {
    const output = metafile.outputs[key] as Record<string, unknown>;
    const imports = (output.imports as Array<Record<string, unknown>>).map((edge) => ({
      ...edge,
      path: typeof edge.path === "string" ? normalizeMetaPath(projectRoot, edge.path) : edge.path,
    })).sort((left, right) => compareText(stableJson(left), stableJson(right)));
    outputs[normalizeMetaPath(projectRoot, key)] = {
      ...output,
      ...(typeof output.entryPoint === "string" ? { entryPoint: normalizeMetaPath(projectRoot, output.entryPoint) } : {}),
      imports,
    };
  }
  return stableValue({ inputs, outputs });
}

function jsonOptions(serialized: string): Record<string, unknown> {
  const value: unknown = JSON.parse(serialized);
  if (!isPlainRecord(value)) throw new Error("Serialized build options are malformed.");
  return value;
}

function assertExactVirtualEvents(events: readonly VirtualLoadRecord[]): Map<string, VirtualLoadRecord> {
  const result = new Map<string, VirtualLoadRecord>();
  for (const event of events) {
    assertVirtualTuple(event);
    if (result.has(event.entry)) throw new Error(`Duplicate onVirtualLoad callback for ${event.entry}.`);
    result.set(event.entry, event);
  }
  return result;
}

function makeInputBlobMap(
  inputs: readonly BoundMetafileInput[],
  controls: readonly BoundControlInput[],
  manifestBytes: Uint8Array,
  stylesBytes: Uint8Array,
): Map<string, Buffer> {
  const result = new Map<string, Buffer>();
  const add = (bytes: Uint8Array): void => {
    const digest = sha256(bytes);
    const copy = Buffer.from(bytes);
    const existing = result.get(digest);
    if (existing && !existing.equals(copy)) throw new Error(`Equal SHA-256 values produced unequal input bytes: ${digest}`);
    if (!existing) result.set(digest, copy);
  };
  for (const input of inputs) add(input.bytes);
  for (const control of controls) add(control.contents);
  add(manifestBytes);
  add(stylesBytes);
  return result;
}

export function prepareEvidenceInputs(
  inputs: readonly BoundMetafileInput[],
  controls: readonly BoundControlInput[],
  manifestBytes: Uint8Array,
  stylesBytes: Uint8Array,
): {
  readonly inputs: readonly unknown[];
  readonly controlInputs: readonly unknown[];
  readonly blobs: readonly EvidenceBlob[];
} {
  const inputRows = inputs
    .map(serializeCompiledInput)
    .sort((left, right) => compareText(left.key as string, right.key as string));
  const controlRows = controls
    .map(serializeControlInput)
    .sort((left, right) => compareText(left.path, right.path));
  const blobMap = makeInputBlobMap(inputs, controls, manifestBytes, stylesBytes);
  return {
    inputs: Object.freeze(inputRows),
    controlInputs: Object.freeze(controlRows),
    blobs: Object.freeze([...blobMap.entries()].map(([digest, bytes]) => Object.freeze({ sha256: digest, bytes }))),
  };
}

function bytesOf(value: Uint8Array): Buffer {
  return Buffer.from(value);
}

/** Runs one explicitly opted-in attempt-3 capture; its destination is fixed, never caller supplied. */
export async function captureDocxEvidenceAttempt3(): Promise<{
  readonly status: "guard failed / evidence only";
  readonly evidenceDirectory: string;
  readonly buildId: string;
  readonly actualInputManifestSha256: string;
  readonly mainSha256: string;
  readonly engineSha256: string;
  readonly completeSha256: string;
  readonly guardError: ExpectedGuardError;
  readonly complete: true;
}> {
  if (arguments.length !== 0) throw new TypeError("Evidence path arguments are not accepted.");
  if (process.env.W69_CAPTURE_ATTEMPT_3 !== "1") throw new Error("Attempt-3 capture requires W69_CAPTURE_ATTEMPT_3=1.");
  const projectRoot = await fs.realpath(process.cwd());
  const beforeAttempt1 = await verifyAttempt1Unchanged();
  const beforeAttempt2 = await verifyAttempt2Unchanged(defaultPreservedAttemptReadIo);
  const persistence = makeEvidencePersistence(defaultEvidenceIo);
  let childCreated = false;
  let preGuardPersisted = false;
  let buildFinalReceipt: ((error: ExpectedGuardError) => Buffer) | null = null;
  try {
    await persistence.createChild();
    childCreated = true;
    const { buildDocxDistribution } = await import(pathToFileURL(path.join(projectRoot, "scripts/docx-build-config.mjs")).href) as {
      buildDocxDistribution(options: Record<string, unknown>): Promise<unknown>;
    };
    let discovery: Record<string, unknown> | null = null;
    const discoveryVirtualEvents: VirtualLoadRecord[] = [];
    let discoveryGuardError: Error | null = null;
    try {
      await buildDocxDistribution({
        root: projectRoot,
        production: true,
        write: false,
        captureHooks: {
          onVirtualLoad: (event: VirtualLoadRecord) => discoveryVirtualEvents.push(event),
          beforeOwnershipGuard: (record: Record<string, unknown>) => { discovery = record; },
        },
      });
    } catch (error) {
      if (error instanceof Error) discoveryGuardError = error;
      else throw new Error(`Discovery build threw a non-Error: ${String(error)}`);
    }
    if (!discovery || !discoveryGuardError || discoveryGuardError.message !== EXPECTED_GUARD_MESSAGE) {
      throw new Error(`Discovery build did not reach the expected unchanged ownership rejection: ${discoveryGuardError?.message ?? "no guard error"}`);
    }
    const discoveryRecord = discovery as unknown as Record<string, unknown>;
    const initialVirtualByEntry = assertExactVirtualEvents(discoveryVirtualEvents);
    if (initialVirtualByEntry.size !== 2 || !initialVirtualByEntry.has("engine") || !initialVirtualByEntry.has("main")) {
      throw new Error("Discovery virtual callback inventory must contain exactly engine and main.");
    }
    const buildId = discoveryRecord.buildId;
    const buildIdInputs = discoveryRecord.buildIdInputs as Array<{ path: string; sha256: string }>;
    const buildIdConfiguration = discoveryRecord.buildIdConfiguration;
    const version = discoveryRecord.version;
    const discoveryEngine = discoveryRecord.engine as Record<string, unknown>;
    const discoveryMain = discoveryRecord.main as Record<string, unknown>;
    const discoveryMetafiles = {
      engine: discoveryEngine.metafile as { inputs: Record<string, unknown>; outputs: Record<string, unknown> },
      main: discoveryMain.metafile as { inputs: Record<string, unknown>; outputs: Record<string, unknown> },
    };
    const inspectionCache = new Map<string, PhysicalInputInspection>();
    const canonicalKeys = new Map<string, string>();
    const identities = new Map<string, string>();
    const physicalInspector: PhysicalInputInspector = async (key) => {
      const existing = [...inspectionCache.values()].find((item) => path.relative(projectRoot, item.canonicalPath).replaceAll("\\", "/") === key);
      if (existing) return existing;
      const inspected = await inspectProjectFile(projectRoot, key);
      const previousKey = canonicalKeys.get(inspected.canonicalPath);
      if (previousKey && previousKey !== key) throw new Error(`Multiple raw metafile keys resolve to one canonical physical file: ${previousKey}, ${key}`);
      const identity = `${String(inspected.device)}:${String(inspected.inode)}`;
      const identityKey = identities.get(identity);
      if (identityKey && identityKey !== key) throw new Error(`Physical input inode alias: ${identityKey}, ${key}`);
      canonicalKeys.set(inspected.canonicalPath, key);
      identities.set(identity, key);
      inspectionCache.set(inspected.canonicalPath, inspected);
      return inspected;
    };
    const entryBindings = [] as Array<{ entry: "engine" | "main"; inputs: readonly BoundMetafileInput[] }>;
    for (const entry of ["engine", "main"] as const) {
      const callbacks = discoveryVirtualEvents.filter((event) => event.entry === entry);
      const bindings = await bindMetafileInputs(discoveryMetafiles[entry], callbacks, physicalInspector);
      entryBindings.push({ entry, inputs: bindings });
    }
    const mergedInputs = mergeBoundInputs(entryBindings);
    const physicalInputs = mergedInputs.filter((input) => input.namespace === "file");
    const controlManifest = await collectControlInputManifest(projectRoot, buildIdInputs, physicalInputs);
    const controls = await collectControlInputs(projectRoot, buildIdInputs, physicalInputs, controlManifest);
    const enumerateControls = (): Promise<ControlInputManifest> =>
      collectControlInputManifest(projectRoot, buildIdInputs, physicalInputs);
    const manifestBytes = bytesOf(discoveryRecord.manifestBytes as Uint8Array);
    const stylesBytes = bytesOf(discoveryRecord.stylesBytes as Uint8Array);
    for (const [relative, asset] of [["manifest.json", manifestBytes], ["styles.css", stylesBytes]] as const) {
      const control = controls.find((item) => item.path === relative);
      if (!control || !Buffer.from(control.contents).equals(asset)) throw new Error(`Asset buffer does not match its pinned control file: ${relative}`);
    }
    await recheckSnapshots(projectRoot, mergedInputs, controls, controlManifest, enumerateControls);

    const bindingByCanonical = new Map(physicalInputs.map((input) => [input.canonicalPath, input]));
    const inputByKey = new Map(mergedInputs.map((input) => [input.key, input]));
    const loadEvents: Array<{
      entry: "engine" | "main";
      key: string;
      namespace: string;
      path: string;
      loader: string;
      resolveDir: string;
      byteLength: number;
      sha256: string;
    }> = [];
    const authoritativeVirtualEvents: VirtualLoadRecord[] = [];
    const authoritativeRecordRef: { value: Record<string, unknown> | null } = { value: null };
    const persistAtPreGuard = async (record: Record<string, unknown>): Promise<void> => {
      authoritativeRecordRef.value = record;
      await recheckSnapshots(projectRoot, mergedInputs, controls, controlManifest, enumerateControls);
      if (record.root !== projectRoot || record.production !== true || record.buildId !== buildId) {
        throw new Error("Authoritative build root, production mode or original buildId changed.");
      }
      if (record.postBuildId !== buildId || stableJson(record.postBuildIdInputs) !== stableJson(buildIdInputs)) {
        throw new Error("Original calculateBuildId inputs changed after authoritative builds.");
      }
      if (stableJson(record.buildIdInputs) !== stableJson(buildIdInputs) || stableJson(record.buildIdConfiguration) !== stableJson(buildIdConfiguration)) {
        throw new Error("Original calculateBuildId inputs/configuration changed between discovery and authoritative build.");
      }
      if (record.version !== version) throw new Error("Plugin version changed between discovery and authoritative build.");
      if (!bytesOf(record.manifestBytes as Uint8Array).equals(manifestBytes) || !bytesOf(record.stylesBytes as Uint8Array).equals(stylesBytes)) {
        throw new Error("Manifest/styles buffers changed between discovery and authoritative build.");
      }
      const authoritativeEngine = record.engine as Record<string, unknown>;
      const authoritativeMain = record.main as Record<string, unknown>;
      for (const entry of ["engine", "main"] as const) {
        const original = entry === "engine" ? discoveryEngine : discoveryMain;
        const actual = entry === "engine" ? authoritativeEngine : authoritativeMain;
        if (actual.serializedOptions !== original.serializedOptions) throw new Error(`${entry} serialized production options changed.`);
        if (stableJson(normalizedMetafile(projectRoot, actual.metafile as never)) !== stableJson(normalizedMetafile(projectRoot, original.metafile as never))) {
          throw new Error(`${entry} authoritative input set/import edges/metafile differ from discovery.`);
        }
        if (!bytesOf(actual.bytes as Uint8Array).equals(bytesOf(original.bytes as Uint8Array))) {
          throw new Error(`${entry} authoritative JS output differs byte-for-byte from discovery.`);
        }
      }
      const authEngineBytes = bytesOf(authoritativeEngine.bytes as Uint8Array);
      const authMainBytes = bytesOf(authoritativeMain.bytes as Uint8Array);
      const engineSha256 = sha256(authEngineBytes);
      const mainSha256 = sha256(authMainBytes);
      const engineOptions = jsonOptions(authoritativeEngine.serializedOptions as string);
      const mainOptions = jsonOptions(authoritativeMain.serializedOptions as string);
      const engineDefines = engineOptions.define as Record<string, string>;
      const mainDefines = mainOptions.define as Record<string, string>;
      if (JSON.parse(engineDefines.__DOCX_BUILD_ID__) !== buildId || JSON.parse(mainDefines.__DOCX_BUILD_ID__) !== buildId) {
        throw new Error("Embedded buildId differs from the original buildId.");
      }
      assertCapturedEngineBinding(mainDefines, engineSha256, authEngineBytes.byteLength);
      if (!authEngineBytes.includes(Buffer.from(`/* Office Reader DOCX engine ${version} ${buildId} */`)) || !authMainBytes.includes(Buffer.from(`WORD_READER_BUILD_ID:${buildId}`))) {
        throw new Error("Production output banners do not contain the original buildId.");
      }
      const virtualByEntry = assertExactVirtualEvents(authoritativeVirtualEvents);
      if (virtualByEntry.size !== 2) throw new Error("Authoritative virtual callback inventory is incomplete.");
      const loadedByEntry = new Map<string, Set<string>>();
      for (const event of loadEvents) {
        const keys = loadedByEntry.get(event.entry) ?? new Set<string>();
        keys.add(event.key);
        loadedByEntry.set(event.entry, keys);
        const bound = inputByKey.get(event.key);
        if (!bound || event.namespace !== bound.namespace || event.path !== bound.path || event.loader !== bound.loader || event.resolveDir !== bound.resolveDir || event.byteLength !== bound.byteLength || event.sha256 !== bound.sha256) {
          throw new Error(`Authoritative onLoad bytes/metadata were not covered by the frozen input map: ${event.key}`);
        }
      }
      for (const entry of ["engine", "main"] as const) {
        const entryRecord = entry === "engine" ? authoritativeEngine : authoritativeMain;
        const entryMetafile = entryRecord.metafile as { inputs: Record<string, unknown> };
        const expectedKeys = Object.keys(entryMetafile.inputs).sort(compareText);
        const actualKeys = [...(loadedByEntry.get(entry) ?? new Set<string>())].sort(compareText);
        if (stableJson(actualKeys) !== stableJson(expectedKeys)) throw new Error(`Authoritative loaded inputs do not exactly cover ${entry} metafile inputs.`);
        if (virtualByEntry.get(entry)?.contents !== initialVirtualByEntry.get(entry)?.contents) throw new Error(`Virtual module bytes changed in ${entry}.`);
      }
      const pinnedToolchainMetadata = preparePinnedToolchainMetadata(controls, process.version, version as string);
      const versions = pinnedToolchainMetadata.versions;
      const preparedEvidenceInputs = prepareEvidenceInputs(mergedInputs, controls, manifestBytes, stylesBytes);
      const inputManifestRows = preparedEvidenceInputs.inputs;
      const controlManifestRows = preparedEvidenceInputs.controlInputs;
      const actualInputManifestSha256 = sha256(Buffer.from(stableJson({ inputs: inputManifestRows, controlInputs: controlManifestRows })));
      const blobs = preparedEvidenceInputs.blobs;
      const blobInventory = blobs.map((blob) => ({
        path: `inputs/${blob.sha256}.bin`,
        sha256: blob.sha256,
        byteLength: blob.bytes.byteLength,
      })).sort((left, right) => compareText(left.path, right.path));
      const entryMetafiles = Object.fromEntries(["engine", "main"].map((entry) => {
        const metafile = (entry === "engine" ? authoritativeEngine : authoritativeMain).metafile as { inputs: Record<string, unknown> };
        return [entry, Object.fromEntries(Object.keys(metafile.inputs).sort(compareText).map((key) => [key, metafile.inputs[key]]))];
      }));
      const inputManifest = {
        schemaVersion: 2,
        workId: "W69",
        attempt: 3,
        root: projectRoot,
        child: ATTEMPT_3_DIRECTORY,
        buildId,
        actualInputManifestSha256,
        inputs: inputManifestRows,
        entryInputs: entryMetafiles,
        buildIdInputs,
        buildIdConfiguration,
        controlInputs: controlManifestRows,
        toolchain: pinnedToolchainMetadata.toolchain,
        assets: {
          manifest: {
            sha256: sha256(manifestBytes),
            byteLength: manifestBytes.byteLength,
            blob: `inputs/${sha256(manifestBytes)}.bin`,
          },
          styles: {
            sha256: sha256(stylesBytes),
            byteLength: stylesBytes.byteLength,
            blob: `inputs/${sha256(stylesBytes)}.bin`,
          },
        },
        blobInventory,
      };
      const inputsJson = Buffer.from(`${JSON.stringify(inputManifest, null, 2)}\n`);
      const mainMetafileBytes = Buffer.from(`${JSON.stringify(authoritativeMain.metafile, null, 2)}\n`);
      const engineMetafileBytes = Buffer.from(`${JSON.stringify(authoritativeEngine.metafile, null, 2)}\n`);
      const mainDiscoveryMetaBytes = Buffer.from(`${JSON.stringify(discoveryMain.metafile, null, 2)}\n`);
      const engineDiscoveryMetaBytes = Buffer.from(`${JSON.stringify(discoveryEngine.metafile, null, 2)}\n`);
      const engineAllowanceBytes = ARCHIVE_LIMIT_BYTES - authMainBytes.byteLength - manifestBytes.byteLength - stylesBytes.byteLength - ARCHIVE_OVERHEAD_BYTES;
      const payloadBytes = authMainBytes.byteLength + authEngineBytes.byteLength + manifestBytes.byteLength + stylesBytes.byteLength;
      const zipBuffer = await writeStoredZip([
        { path: "main.js", bytes: authMainBytes },
        { path: "docx-engine.cjs", bytes: authEngineBytes },
        { path: "manifest.json", bytes: manifestBytes },
        { path: "styles.css", bytes: stylesBytes },
      ], { isCancelled: () => false, yieldControl: async () => {} });
      const archiveBytes = zipBuffer.byteLength;
      if (
        authMainBytes.byteLength > MAIN_LIMIT_BYTES ||
        payloadBytes > ARCHIVE_PAYLOAD_LIMIT_BYTES ||
        archiveBytes > ARCHIVE_LIMIT_BYTES ||
        archiveBytes - payloadBytes !== ARCHIVE_OVERHEAD_BYTES ||
        engineAllowanceBytes <= 0 || authEngineBytes.byteLength > engineAllowanceBytes
      ) throw new Error("Captured four-member STORE ZIP violates a distribution size limit.");
      const comparisons = {
        buildIdEqual: record.buildId === discoveryRecord.buildId,
        postBuildIdEqual: record.postBuildId === buildId,
        buildIdInputsAndConfigurationEqual: stableJson(record.buildIdInputs) === stableJson(buildIdInputs) && stableJson(record.buildIdConfiguration) === stableJson(buildIdConfiguration),
        productionOptionsEqual: authoritativeEngine.serializedOptions === discoveryEngine.serializedOptions && authoritativeMain.serializedOptions === discoveryMain.serializedOptions,
        normalizedMetafilesAndImportEdgesEqual: stableJson(normalizedMetafile(projectRoot, authoritativeEngine.metafile as never)) === stableJson(normalizedMetafile(projectRoot, discoveryEngine.metafile as never)) && stableJson(normalizedMetafile(projectRoot, authoritativeMain.metafile as never)) === stableJson(normalizedMetafile(projectRoot, discoveryMain.metafile as never)),
        javascriptOutputsByteEqual: authEngineBytes.equals(bytesOf(discoveryEngine.bytes as Uint8Array)) && authMainBytes.equals(bytesOf(discoveryMain.bytes as Uint8Array)),
        exactOnLoadCoverage: true,
        exactVirtualTuple: true,
        pinnedPhysicalAndControlInputsUnchanged: true,
        embeddedBuildIdAndEngineDigestEqual: true,
        packageLimitsPass: true,
      };
      if (!Object.values(comparisons).every(Boolean)) throw new Error("An authoritative capture comparison did not pass.");
      buildFinalReceipt = (caughtGuardError) => Buffer.from(`${JSON.stringify({
        schemaVersion: 2,
        workId: "W69",
        attempt: 3,
        root: APPROVED_EVIDENCE_ROOT,
        child: ATTEMPT_3_DIRECTORY,
        buildId,
        originalBuildIdCoverage: "unchanged; actual resolved bytes are separately bound by actualInputManifestSha256",
        actualInputManifestSha256,
        versions,
        serializedProductionOptions: {
          engine: authoritativeEngine.serializedOptions,
          main: authoritativeMain.serializedOptions,
        },
        buildIdConfiguration,
        comparisons,
        discovery: {
          guardError: { name: discoveryGuardError.name, message: discoveryGuardError.message, stack: discoveryGuardError.stack ?? "" },
          outputs: {
            main: { byteLength: bytesOf(discoveryMain.bytes as Uint8Array).byteLength, sha256: sha256(bytesOf(discoveryMain.bytes as Uint8Array)) },
            engine: { byteLength: bytesOf(discoveryEngine.bytes as Uint8Array).byteLength, sha256: sha256(bytesOf(discoveryEngine.bytes as Uint8Array)) },
          },
          metafiles: {
            mainSha256: sha256(mainDiscoveryMetaBytes),
            engineSha256: sha256(engineDiscoveryMetaBytes),
          },
        },
        authoritative: {
          outputs: {
            main: { byteLength: authMainBytes.byteLength, sha256: mainSha256 },
            engine: { byteLength: authEngineBytes.byteLength, sha256: engineSha256 },
          },
          metafiles: {
            main: { byteLength: mainMetafileBytes.byteLength, sha256: sha256(mainMetafileBytes) },
            engine: { byteLength: engineMetafileBytes.byteLength, sha256: sha256(engineMetafileBytes) },
          },
          inputsJson: { byteLength: inputsJson.byteLength, sha256: sha256(inputsJson) },
        },
        sizes: {
          mainBytes: authMainBytes.byteLength,
          mainLimitBytes: MAIN_LIMIT_BYTES,
          engineBytes: authEngineBytes.byteLength,
          engineAllowanceBytes,
          manifestBytes: manifestBytes.byteLength,
          stylesBytes: stylesBytes.byteLength,
          payloadBytes,
          payloadLimitBytes: ARCHIVE_PAYLOAD_LIMIT_BYTES,
          actualStoreZipBytes: archiveBytes,
          measuredOverheadBytes: archiveBytes - payloadBytes,
          expectedOverheadBytes: ARCHIVE_OVERHEAD_BYTES,
          archiveLimitBytes: ARCHIVE_LIMIT_BYTES,
          packageArtifactWritten: false,
        },
        guard: {
          status: "guard failed / evidence only",
          name: caughtGuardError.name,
          message: caughtGuardError.message,
          stack: caughtGuardError.stack,
          completeIsNotS5_3Pass: true,
        },
      }, null, 2)}\n`);
      const files: EvidencePreGuardFiles = {
        "main.js": authMainBytes,
        "docx-engine.cjs": authEngineBytes,
        "main.metafile.json": mainMetafileBytes,
        "engine.metafile.json": engineMetafileBytes,
        "inputs.json": inputsJson,
      };
      await persistence.persistBeforeGuard({
        files,
        blobs,
        bindingContext: Object.freeze({
          expectedManifest: controlManifest,
          compiledInputs: mergedInputs,
          controls,
        }),
      });
      preGuardPersisted = true;
    };
    const captureHooks = {
      beforeBuild: async (): Promise<void> => recheckSnapshots(
        projectRoot,
        mergedInputs,
        controls,
        controlManifest,
        enumerateControls,
      ),
      onFileLoad: (args: { namespace: string; path: string; suffix: string; with: Record<string, string> }, entry: "engine" | "main") => {
        const loaded = pinnedPhysicalLoad(physicalInputs, args, entry);
        const binding = bindingByCanonical.get(args.path)!;
        const bytes = loaded.contents;
        loadEvents.push({
          entry,
          key: binding.key,
          namespace: args.namespace,
          path: args.path,
          loader: loaded.loader,
          resolveDir: loaded.resolveDir,
          byteLength: bytes.byteLength,
          sha256: sha256(bytes),
        });
        return loaded;
      },
      onVirtualLoad: (event: VirtualLoadRecord): void => {
        assertVirtualTuple(event);
        const binding = inputByKey.get(VIRTUAL_KEY);
        if (!binding || !binding.entries.includes(event.entry)) throw new Error(`Unknown virtual input load for ${event.entry}.`);
        const bytes = Buffer.from(event.contents, "utf8");
        if (sha256(bytes) !== binding.sha256 || !bytes.equals(Buffer.from(binding.bytes))) throw new Error("Authoritative virtual bytes differ from discovery pin.");
        authoritativeVirtualEvents.push(event);
        loadEvents.push({
          entry: event.entry,
          key: VIRTUAL_KEY,
          namespace: event.namespace,
          path: event.path,
          loader: event.loader,
          resolveDir: event.resolveDir,
          byteLength: bytes.byteLength,
          sha256: sha256(bytes),
        });
      },
      beforeOwnershipGuard: persistAtPreGuard,
    };
    let authoritativeReturned = false;
    let authoritativeError: Error | null = null;
    try {
      await buildDocxDistribution({ root: projectRoot, production: true, write: false, captureHooks });
      authoritativeReturned = true;
    } catch (error) {
      if (error instanceof Error) authoritativeError = error;
      else throw new Error(`Authoritative build threw a non-Error: ${String(error)}`);
    }
    if (authoritativeReturned || !preGuardPersisted || !authoritativeError || authoritativeError.message !== EXPECTED_GUARD_MESSAGE) {
      throw new Error(`Authoritative build did not produce the exact expected guard rejection after persistence: ${authoritativeError?.message ?? "no error"}`);
    }
    const authoritativeStack = authoritativeError.stack;
    if (typeof authoritativeStack !== "string" || authoritativeStack.length === 0) {
      throw new Error("Authoritative guard error did not provide its exact stack.");
    }
    const caughtGuardError: ExpectedGuardError = {
      name: authoritativeError.name,
      message: authoritativeError.message,
      stack: authoritativeStack,
    };
    await persistence.persistGuardError(caughtGuardError);
    if (!buildFinalReceipt) throw new Error("Final receipt was not prepared by pre-guard validation.");
    const finalReceiptBytes = buildFinalReceipt(caughtGuardError);
    await persistence.persistFinalReceipt(finalReceiptBytes, caughtGuardError);
    const beforeComplete = async (): Promise<void> => {
      const afterAttempt1 = await verifyAttempt1Unchanged();
      if (stableJson(beforeAttempt1) !== stableJson(afterAttempt1)) throw new Error("Attempt-1 evidence changed during attempt-3.");
      const afterAttempt2 = await verifyAttempt2Unchanged(defaultPreservedAttemptReadIo);
      if (stableJson(beforeAttempt2) !== stableJson(afterAttempt2)) throw new Error("Attempt-2 evidence changed during attempt-3.");
    };
    const completion = await persistence.writeComplete(caughtGuardError, beforeComplete);
    const receipt = JSON.parse(Buffer.from(await fs.readFile(path.join(ATTEMPT_3_DIRECTORY, "receipt.json"))).toString("utf8")) as Record<string, unknown>;
    return {
      status: "guard failed / evidence only",
      evidenceDirectory: ATTEMPT_3_DIRECTORY,
      buildId: receipt.buildId as string,
      actualInputManifestSha256: receipt.actualInputManifestSha256 as string,
      mainSha256: (receipt.authoritative as { outputs: { main: { sha256: string } } }).outputs.main.sha256,
      engineSha256: (receipt.authoritative as { outputs: { engine: { sha256: string } } }).outputs.engine.sha256,
      completeSha256: completion.completeSha256,
      guardError: caughtGuardError,
      complete: true,
    };
  } catch (error) {
    if (childCreated) await persistence.writeIncomplete(error);
    throw error;
  }
}
