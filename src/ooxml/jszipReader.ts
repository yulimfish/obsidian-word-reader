import type JSZip from "jszip";
import defaults from "jszip/lib/defaults";
import loadAsync from "jszip/lib/load";
import ZipObject from "jszip/lib/zipObject";

type PreviewZipEntry = JSZip.JSZipObject;

interface PreviewReader {
  readBytes(type: "uint8array"): Promise<Uint8Array>;
  readOther(type: "string" | "uint8array", onUpdate?: JSZip.OnUpdateCallback): Promise<unknown>;
}

function previewReader(entry: PreviewZipEntry): PreviewReader {
  const originalAsync: PreviewZipEntry["async"] = entry.async.bind(entry) as PreviewZipEntry["async"];
  return {
    readBytes: (type) => originalAsync(type),
    readOther: (type, onUpdate) => originalAsync(type, onUpdate),
  };
}

const svgEntries = new WeakSet<JSZip.JSZipObject>();

interface InternalZipOptions {
  date?: Date;
  dir?: boolean;
  [key: string]: unknown;
}

/** Minimal host for the pinned JSZip loader, without archive generation. */
export default class ReadOnlyZip {
  readonly files: Record<string, JSZip.JSZipObject> = Object.create(null) as
    Record<string, JSZip.JSZipObject>;
  comment: string | null = null;

  file(name: string): JSZip.JSZipObject | null;
  file(name: string, data: unknown, options: InternalZipOptions): this;
  file(
    name: string,
    data?: unknown,
    options?: InternalZipOptions,
  ): JSZip.JSZipObject | null | this {
    if (arguments.length === 1) {
      const entry = this.files[name];
      if (!entry || entry.dir) return null;
      if (name.toLowerCase().endsWith(".svg") && !svgEntries.has(entry)) {
        svgEntries.add(entry);
        const svgEntry = entry;
        const reader = previewReader(svgEntry);
        svgEntry.async = ((type: "string" | "blob" | "uint8array", onUpdate?: JSZip.OnUpdateCallback) => {
          if (type !== "blob") return reader.readOther(type, onUpdate);
          return reader.readBytes("uint8array").then((bytes: Uint8Array) => {
            const buffer = new ArrayBuffer(bytes.byteLength);
            new Uint8Array(buffer).set(bytes);
            return new Blob([buffer], { type: "image/svg+xml" });
          });
        }) as typeof entry.async;
      }
      return entry;
    }
    const normalized = { ...defaults, ...options };
    normalized.date ??= new Date();
    this.files[name] = new ZipObject(
      name,
      normalized.dir || data === null || data === undefined ? "" : data,
      normalized,
    );
    return this;
  }

  static loadAsync(data: ArrayBuffer): Promise<ReadOnlyZip> {
    const pending = Reflect.apply(
      loadAsync as unknown as (...args: unknown[]) => unknown,
      new ReadOnlyZip(),
      [data, { createFolders: false, checkCRC32: false }],
    );
    return pending as Promise<ReadOnlyZip>;
  }
}

export async function loadOoxmlZip(buffer: ArrayBuffer): Promise<JSZip> {
  const zip = await ReadOnlyZip.loadAsync(buffer);
  return zip as unknown as JSZip;
}
