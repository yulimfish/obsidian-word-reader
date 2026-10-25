import type JSZip from "jszip";
import defaults from "jszip/lib/defaults";
import loadAsync from "jszip/lib/load";
import ZipObject from "jszip/lib/zipObject";

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
      return entry && !entry.dir ? entry : null;
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
