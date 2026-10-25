/* eslint-disable obsidianmd/no-global-this -- This file is a Node worker harness for isolating parser globals. */
import assert from "node:assert/strict";
import { DOMParser, XMLSerializer } from "@xmldom/xmldom";
import { parseAsync } from "docx-preview";
import ReadOnlyZip from "jszip";
import { isMainThread, parentPort, threadId, workerData } from "node:worker_threads";

import { prepareDocxPreviewPackage } from "../../src/docx/docxPreviewPackage";

type DescriptorSnapshot = {
  owner: object;
  property: string;
  own: PropertyDescriptor | undefined;
  inherited: Array<{ owner: object; descriptor: PropertyDescriptor | undefined }>;
};

type InstallProperty = {
  owner: object;
  property: string;
  descriptor: PropertyDescriptor;
};

interface ZipEntryLike {
  name: string;
  dir: boolean;
  async(type: string): Promise<unknown>;
}

interface ZipLike {
  files: Record<string, ZipEntryLike>;
  file(name: string): ZipEntryLike | null;
}

interface ReadOnlyZipHost {
  loadAsync(buffer: ArrayBuffer): Promise<ZipLike>;
}

interface WorkerRequest {
  mode: string;
  buffer?: ArrayBuffer;
  marker?: SharedArrayBuffer;
}

function captureDescriptor(owner: object, property: string): DescriptorSnapshot {
  const inherited: DescriptorSnapshot["inherited"] = [];
  let current = Object.getPrototypeOf(owner) as object | null;
  while (current) {
    inherited.push({
      owner: current,
      descriptor: Object.getOwnPropertyDescriptor(current, property),
    });
    current = Object.getPrototypeOf(current) as object | null;
  }
  return {
    owner,
    property,
    own: Object.getOwnPropertyDescriptor(owner, property),
    inherited,
  };
}

function restoreDescriptor(snapshot: DescriptorSnapshot): void {
  if (snapshot.own) {
    Object.defineProperty(snapshot.owner, snapshot.property, snapshot.own);
  } else {
    Reflect.deleteProperty(snapshot.owner, snapshot.property);
  }
}

function descriptorsEqual(first: DescriptorSnapshot, second: DescriptorSnapshot): boolean {
  return first.owner === second.owner &&
    first.property === second.property &&
    descriptorEqual(first.own, second.own) &&
    first.inherited.length === second.inherited.length &&
    first.inherited.every((entry, index) => {
      const other = second.inherited[index];
      return other !== undefined && entry.owner === other.owner &&
        descriptorEqual(entry.descriptor, other.descriptor);
    });
}

function descriptorEqual(
  first: PropertyDescriptor | undefined,
  second: PropertyDescriptor | undefined,
): boolean {
  if (!first || !second) return first === second;
  return first.configurable === second.configurable &&
    first.enumerable === second.enumerable &&
    first.writable === second.writable &&
    first.value === second.value &&
    first.get === second.get &&
    first.set === second.set;
}

function installationSnapshots(
  globalTarget: object,
  documentPrototype: object,
  elementPrototype: object,
): DescriptorSnapshot[] {
  return [
    captureDescriptor(globalTarget, "DOMParser"),
    captureDescriptor(globalTarget, "XMLSerializer"),
    captureDescriptor(globalTarget, "Node"),
    captureDescriptor(documentPrototype, "firstElementChild"),
    captureDescriptor(elementPrototype, "firstElementChild"),
  ];
}

function installationProperties(
  globalTarget: object,
  documentPrototype: object,
  elementPrototype: object,
): InstallProperty[] {
  const firstElementChild: PropertyDescriptor = {
    configurable: true,
    enumerable: false,
    get(this: Node): Element | null {
      let child = this.firstChild;
      while (child && child.nodeType !== 1) child = child.nextSibling;
      return child as Element | null;
    },
  };
  return [
    {
      owner: globalTarget,
      property: "DOMParser",
      descriptor: { configurable: true, enumerable: true, writable: true, value: DOMParser },
    },
    {
      owner: globalTarget,
      property: "XMLSerializer",
      descriptor: { configurable: true, enumerable: true, writable: true, value: XMLSerializer },
    },
    {
      owner: globalTarget,
      property: "Node",
      descriptor: {
        configurable: true,
        enumerable: true,
        writable: true,
        value: Object.freeze({ ELEMENT_NODE: 1 }),
      },
    },
    { owner: documentPrototype, property: "firstElementChild", descriptor: firstElementChild },
    { owner: elementPrototype, property: "firstElementChild", descriptor: firstElementChild },
  ];
}

export function installParserDomShim(
  globalTarget: object,
  documentPrototype: object,
  elementPrototype: object,
): () => void {
  const snapshots = installationSnapshots(globalTarget, documentPrototype, elementPrototype);
  const properties = installationProperties(globalTarget, documentPrototype, elementPrototype);
  const attempted: DescriptorSnapshot[] = [];
  let restored = false;
  try {
    for (const property of properties) {
      const snapshot = snapshots.find((item) =>
        item.owner === property.owner && item.property === property.property
      );
      if (!snapshot) throw new Error("Missing parser shim descriptor snapshot.");
      attempted.push(snapshot);
      Object.defineProperty(property.owner, property.property, property.descriptor);
    }
  } catch (error) {
    for (const snapshot of attempted.reverse()) restoreDescriptor(snapshot);
    throw error;
  }

  return () => {
    if (restored) return;
    for (const snapshot of [...snapshots].reverse()) restoreDescriptor(snapshot);
    restored = true;
  };
}

function verifySnapshots(snapshots: readonly DescriptorSnapshot[]): boolean {
  return snapshots.every((snapshot) =>
    descriptorsEqual(snapshot, captureDescriptor(snapshot.owner, snapshot.property))
  );
}

function serializeModel(body: unknown): {
  text: string[];
  images: Array<{ src: string; cssStyle: Record<string, unknown> }>;
  drawings: Array<Record<string, unknown>>;
} {
  const result = {
    text: [] as string[],
    images: [] as Array<{ src: string; cssStyle: Record<string, unknown> }>,
    drawings: [] as Array<Record<string, unknown>>,
  };
  const visit = (value: unknown): void => {
    if (Array.isArray(value)) {
      for (const child of value) visit(child);
      return;
    }
    if (!value || typeof value !== "object") return;
    const node = value as Record<string, unknown>;
    if (node.type === "text" && typeof node.text === "string") result.text.push(node.text);
    if (node.type === "image") {
      result.images.push({
        src: typeof node.src === "string" ? node.src : "",
        cssStyle: (node.cssStyle ?? {}) as Record<string, unknown>,
      });
    }
    if (node.type === "drawing") result.drawings.push((node.cssStyle ?? {}) as Record<string, unknown>);
    if (Array.isArray(node.children)) visit(node.children);
  };
  visit(body);
  return result;
}

function assertNoOlePayloadReferences(names: readonly string[], documentXml: string, relationshipsXml: string, contentTypesXml: string): void {
  assert.equal(names.some((name) => /(?:^|\/)(?:embeddings|oleobjects)(?:\/|$)/i.test(name)), false);
  assert.equal(/oleObject|TargetMode="External"/i.test(relationshipsXml), false);
  assert.equal(/oleObject|ole-object/i.test(contentTypesXml), false);
  assert.equal(/oleObject|oleobjects|embeddings/i.test(documentXml), false);
  assert.equal(/<w:altChunk\b/i.test(documentXml), false);
  assert.equal(/TargetMode="External"/i.test(relationshipsXml), false);
}

function post(message: Record<string, unknown>): void {
  parentPort?.postMessage(message);
}

async function runParse(request: WorkerRequest): Promise<void> {
  if (!request.buffer) throw new Error("Worker input buffer is missing.");
  const source = request.buffer;
  const original = new Uint8Array(source).slice();
  const parsedSample = new DOMParser().parseFromString("<root/>", "application/xml");
  const documentPrototype = Object.getPrototypeOf(parsedSample) as object;
  const elementPrototype = Object.getPrototypeOf(parsedSample.documentElement) as object;
  const before = installationSnapshots(globalThis, documentPrototype, elementPrototype);
  const restore = installParserDomShim(globalThis, documentPrototype, elementPrototype);
  const reads: Array<{ path: string; mode: string }> = [];
  const zipApi = ReadOnlyZip as unknown as ReadOnlyZipHost;
  const loadOriginal = zipApi.loadAsync.bind(zipApi) as ReadOnlyZipHost["loadAsync"];
  zipApi.loadAsync = async (buffer: ArrayBuffer): Promise<ZipLike> => {
    const zip = await loadOriginal(buffer);
    const entries = Object.values(zip.files);
    for (const entry of entries) {
      const originalAsync = entry.async.bind(entry) as ZipEntryLike["async"];
      entry.async = async (mode: string): Promise<unknown> => {
        reads.push({ path: entry.name, mode });
        return await originalAsync(mode);
      };
    }
    return zip;
  };

  let successful = false;
  let result: Record<string, unknown> | null = null;
  let failure: unknown = null;
  try {
    const prepared = await prepareDocxPreviewPackage(source, {
      isCancelled: () => false,
      yieldControl: async () => {},
    });
    const document = await parseAsync(prepared.buffer, {
      trimXmlDeclaration: true,
      experimental: true,
    }) as { documentPart?: { body?: unknown } };
    const model = serializeModel(document.documentPart?.body);
    const preparedZip = await zipApi.loadAsync(prepared.buffer);
    const names = Object.keys(preparedZip.files);
    const documentXml = await preparedZip.file("word/document.xml")!.async("string") as string;
    const relationshipsXml = await preparedZip.file("word/_rels/document.xml.rels")!.async("string") as string;
    const contentTypesXml = await preparedZip.file("[Content_Types].xml")!.async("string") as string;
    assertNoOlePayloadReferences(names, documentXml, relationshipsXml, contentTypesXml);
    const imageParts: string[] = [];
    for (const name of names.filter((item) => item.startsWith("word/media/"))) {
      const image = await preparedZip.file(name)!.async("uint8array") as Uint8Array;
      if (image.byteLength > 0) imageParts.push(name);
    }
    const ordinaryPayloadReads = reads.filter((read) =>
      /(?:^|\/)(?:embeddings|oleobjects)(?:\/|$)/i.test(read.path)
    );
    result = {
      threadId,
      globalTargetIsWorkerGlobal:
        (globalThis as unknown as { DOMParser: typeof DOMParser; Node: { ELEMENT_NODE: number } }).DOMParser === DOMParser &&
        (globalThis as unknown as { Node: { ELEMENT_NODE: number } }).Node.ELEMENT_NODE === 1,
      distinctSanitizedBuffer: prepared.buffer !== source,
      originalUnchanged: original.every((byte, index) => new Uint8Array(source)[index] === byte),
      inventoryOccurrences: prepared.inventory.occurrences.length,
      issues: prepared.issues,
      text: model.text,
      imageCount: model.images.length,
      imageStyles: model.images.map((image) => image.cssStyle),
      drawingStyles: model.drawings,
      readableImageParts: imageParts,
      ordinaryPayloadReads,
      oleReferencesRemain: /oleObject|ole-object/i.test(relationshipsXml + contentTypesXml) || names.some((name) => /(?:^|\/)(?:embeddings|oleobjects)(?:\/|$)/i.test(name)),
      externalResourcesRemain: /TargetMode="External"/i.test(relationshipsXml),
      altChunksRemain: /<w:altChunk\b/i.test(documentXml),
    };
    successful = true;
  } catch (error) {
    failure = error;
  } finally {
    if (successful) {
      try {
        restore();
        (result as Record<string, unknown>).shimRestored = verifySnapshots(before);
        if (!(result as Record<string, unknown>).shimRestored) {
          failure = new Error("Parser shim descriptors did not restore exactly.");
          successful = false;
        }
      } catch (error) {
        failure = error;
        successful = false;
      }
    }
  }

  zipApi.loadAsync = loadOriginal;
  if (!successful || !result) {
    post({ type: "failure", stage: "parseAsync", error: failure instanceof Error ? failure.message : String(failure) });
    await new Promise<void>(() => {});
    return;
  }
  post({ type: "success", ...result });
}

function runShimProbe(): void {
  const sample = new DOMParser().parseFromString("<!--before--><?probe yes?><root>text<!--middle--><child/><empty/></root>", "application/xml");
  const documentPrototype = Object.getPrototypeOf(sample) as object;
  const elementPrototype = Object.getPrototypeOf(sample.documentElement) as object;
  const before = installationSnapshots(globalThis, documentPrototype, elementPrototype);
  const restore = installParserDomShim(globalThis, documentPrototype, elementPrototype);
  try {
    assert.equal((globalThis as unknown as { Node: { ELEMENT_NODE: number } }).Node.ELEMENT_NODE, 1);
    assert.equal(sample.firstElementChild?.nodeName, "root");
    assert.equal(sample.documentElement.firstElementChild?.nodeName, "child");
    const empty = sample.getElementsByTagName("empty")[0];
    assert.equal(empty.firstElementChild, null);
  } finally {
    restore();
  }
  post({ type: "success", mode: "shim-probe", threadId, actualXmldomSemantics: true, descriptorsRestored: verifySnapshots(before) });
}

function runPartialInstallProbe(): void {
  const sample = new DOMParser().parseFromString("<root/>", "application/xml");
  const documentPrototype = Object.getPrototypeOf(sample) as object;
  const elementPrototype = Object.getPrototypeOf(sample.documentElement) as object;
  const target = Object.create(null) as object;
  Object.defineProperty(target, "XMLSerializer", {
    configurable: false,
    enumerable: false,
    writable: false,
    value: "locked",
  });
  const before = installationSnapshots(target, documentPrototype, elementPrototype);
  let parserCallbackEntered = false;
  assert.throws(() => {
    const restore = installParserDomShim(target, documentPrototype, elementPrototype);
    try {
      parserCallbackEntered = true;
    } finally {
      restore();
    }
  });
  post({
    type: "success",
    mode: "partial-install",
    threadId,
    parserCallbackEntered,
    rollbackExact: verifySnapshots(before),
  });
}

async function runPendingSibling(request: WorkerRequest): Promise<void> {
  if (!request.marker) throw new Error("Shared test marker is missing.");
  const sample = new DOMParser().parseFromString("<root/>", "application/xml");
  const restore = installParserDomShim(
    globalThis,
    Object.getPrototypeOf(sample) as object,
    Object.getPrototypeOf(sample.documentElement) as object,
  );
  const marker = new Int32Array(request.marker);
  const gate = new Promise<void>((resolve) => {
    parentPort?.once("message", () => resolve());
  });
  const sibling = gate.then(() => {
    Atomics.store(marker, 0, 1);
    return (globalThis as unknown as { DOMParser: typeof DOMParser }).DOMParser;
  });
  try {
    await Promise.all([Promise.reject(new Error("aggregate rejection")), sibling]);
  } catch {
    post({
      type: "aggregate-rejected",
      threadId,
      siblingPending: true,
      marker: Atomics.load(marker, 0),
      shimStillInstalled: (globalThis as unknown as { DOMParser: typeof DOMParser }).DOMParser === DOMParser,
    });
    void restore;
    await new Promise<void>(() => {});
  }
}

async function runParserFailure(): Promise<void> {
  const sample = new DOMParser().parseFromString("<root/>", "application/xml");
  installParserDomShim(
    globalThis,
    Object.getPrototypeOf(sample) as object,
    Object.getPrototypeOf(sample.documentElement) as object,
  );
  try {
    await parseAsync(new ArrayBuffer(0));
  } catch (error) {
    post({ type: "failure", stage: "parseAsync", error: error instanceof Error ? error.message : String(error) });
    await new Promise<void>(() => {});
  }
  post({ type: "failure", stage: "parseAsync", error: "Expected malformed ZIP rejection." });
  await new Promise<void>(() => {});
}

async function dispatch(): Promise<void> {
  const request = workerData as WorkerRequest;
  if (request.mode === "parse") await runParse(request);
  else if (request.mode === "shim-probe") runShimProbe();
  else if (request.mode === "partial-install") runPartialInstallProbe();
  else if (request.mode === "pending-sibling") await runPendingSibling(request);
  else if (request.mode === "parser-failure") await runParserFailure();
  else throw new Error(`Unknown worker mode: ${request.mode}`);
  parentPort?.close();
}

if (!isMainThread && parentPort) {
  void dispatch().catch((error: unknown) => {
    post({ type: "failure", stage: "worker", error: error instanceof Error ? error.message : String(error) });
  });
}
