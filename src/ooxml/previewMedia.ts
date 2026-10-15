import { LruCache } from "../reader/lruCache";
import { decodeEmfScene } from "./emfScene";
import { MetafileSerializeError, serializeMetafileScene } from "./metafileSvg";
import { DEFAULT_METAFILE_LIMITS, parseMetafileRecords } from "./metafileRecords";
import { decodeWmfScene } from "./wmfScene";
import type { MetafileScene } from "./metafileScene";
import type { PreviewIssue } from "./previewTypes";

const PNG_SIGNATURE = Uint8Array.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
const MAX_DERIVED_ENTRIES = 24;
const MAX_DERIVED_BYTES = 32 * 1024 * 1024;
type PreviewMediaExtension = "png" | "jpg" | "jpeg" | "bmp" | "wmf" | "emf";

export type PreviewMediaErrorCode = "invalid-media" | "unsupported-media" | "cancelled" | "disposed";

export class PreviewMediaError extends Error {
  constructor(readonly code: PreviewMediaErrorCode, message: string) {
    super(message);
    this.name = "PreviewMediaError";
  }
}

export interface ResolvedPreviewMedia {
  readonly bytes: Uint8Array;
  readonly mimeType: "image/svg+xml" | "image/png" | "image/jpeg" | "image/bmp";
  readonly requiredFonts: readonly string[];
  readonly issues: readonly PreviewIssue[];
}

export interface PreviewMediaOptions {
  readonly isCancelled: () => boolean;
  readonly yieldControl: () => Promise<void>;
}

interface CachedMedia {
  readonly sourceBytes: Uint8Array;
  readonly result: ResolvedPreviewMedia;
  readonly cost: number;
}

export interface PreviewMediaResolver {
  resolve(path: string, bytes: Uint8Array, options: PreviewMediaOptions): Promise<ResolvedPreviewMedia>;
  getCacheDiagnostics(): { readonly entries: number; readonly bytes: number };
  dispose(): void;
}

function checkCancelled(options: PreviewMediaOptions): void {
  if (options.isCancelled()) throw new PreviewMediaError("cancelled", "Preview media resolution was cancelled.");
}

function extension(path: string): string {
  const separator = Math.max(path.lastIndexOf("/"), path.lastIndexOf("\\"));
  const dot = path.lastIndexOf(".");
  return dot > separator ? path.slice(dot + 1).toLowerCase() : "";
}

function beginsWith(bytes: Uint8Array, signature: Uint8Array): boolean {
  if (bytes.byteLength < signature.byteLength) return false;
  for (let index = 0; index < signature.byteLength; index += 1) {
    if (bytes[index] !== signature[index]) return false;
  }
  return true;
}

function isJpeg(bytes: Uint8Array): boolean {
  return bytes.byteLength >= 4 && bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff;
}

function isBmp(bytes: Uint8Array): boolean {
  return bytes.byteLength >= 14 && bytes[0] === 0x42 && bytes[1] === 0x4d;
}

function assertValidMedia(path: string, bytes: Uint8Array): PreviewMediaExtension {
  if (typeof path !== "string" || path.length === 0 || !(bytes instanceof Uint8Array)) {
    throw new PreviewMediaError("invalid-media", "Preview media path and bytes are required.");
  }
  const kind = extension(path);
  if (kind === "png" && beginsWith(bytes, PNG_SIGNATURE)) return kind;
  if ((kind === "jpg" || kind === "jpeg") && isJpeg(bytes)) return kind;
  if (kind === "bmp" && isBmp(bytes)) return kind;
  if (kind === "wmf" || kind === "emf") return kind;
  if (kind === "png" || kind === "jpg" || kind === "jpeg" || kind === "bmp") {
    throw new PreviewMediaError("invalid-media", "Preview media magic does not match its file extension.");
  }
  throw new PreviewMediaError("unsupported-media", "Only PNG, JPEG, BMP, WMF and EMF package media are supported.");
}

function warningIssues(path: string, scene: MetafileScene): readonly PreviewIssue[] {
  return scene.warnings.map((warning) => {
    const details: string[] = [];
    for (const [key, value] of Object.entries(warning)) {
      if (key === "kind" || key === "fontName") continue;
      if (typeof value === "string" || typeof value === "number") details.push(`${key}=${value}`);
      else if (Array.isArray(value)) details.push(`${key}=${value.join(",")}`);
      else if (value && typeof value === "object") details.push(`${key}=present`);
    }
    return {
      partPath: path,
      occurrenceId: path,
      code: warning.kind,
      detail: warning.kind === "font-glyph-gap" || warning.kind === "font-metrics-required"
        ? `${warning.kind}: ${warning.fontName ?? "unknown font"}`
        : [warning.kind, ...details].join("; "),
    };
  });
}

async function decodeVector(
  path: string,
  bytes: Uint8Array,
  format: "wmf" | "emf",
  options: PreviewMediaOptions,
): Promise<ResolvedPreviewMedia> {
  checkCancelled(options);
  await options.yieldControl();
  checkCancelled(options);
  const records = parseMetafileRecords(bytes, format, DEFAULT_METAFILE_LIMITS);
  checkCancelled(options);
  const scene = format === "wmf"
    ? await decodeWmfScene(records, options)
    : await decodeEmfScene(records, options);
  checkCancelled(options);
  const svg = serializeMetafileScene(scene);
  checkCancelled(options);
  return {
    bytes: svg,
    mimeType: "image/svg+xml",
    requiredFonts: [...scene.requiredFonts],
    issues: warningIssues(path, scene),
  };
}

export async function resolvePreviewMedia(
  path: string,
  bytes: Uint8Array,
  options: PreviewMediaOptions,
): Promise<ResolvedPreviewMedia> {
  if (!options || typeof options.isCancelled !== "function" || typeof options.yieldControl !== "function") {
    throw new PreviewMediaError("invalid-media", "Preview media cancellation and yield callbacks are required.");
  }
  checkCancelled(options);
  const kind = assertValidMedia(path, bytes);
  if (kind === "png" || kind === "jpg" || kind === "jpeg" || kind === "bmp") {
    await options.yieldControl();
    checkCancelled(options);
    return {
      bytes: bytes.slice(),
      mimeType: kind === "png" ? "image/png" : kind === "bmp" ? "image/bmp" : "image/jpeg",
      requiredFonts: [],
      issues: [],
    };
  }
  try {
    return await decodeVector(path, bytes, kind, options);
  } catch (error) {
    if (error instanceof MetafileSerializeError) throw error;
    if (error instanceof PreviewMediaError) throw error;
    throw error;
  }
}

function utf8ByteLength(value: string): number {
  let bytes = 0;
  for (let index = 0; index < value.length; index += 1) {
    const code = value.charCodeAt(index);
    if (code <= 0x7f) bytes += 1;
    else if (code <= 0x7ff) bytes += 2;
    else if (code >= 0xd800 && code <= 0xdbff) {
      bytes += 4;
      index += 1;
    } else bytes += 3;
  }
  return bytes;
}

function sameBytes(first: Uint8Array, second: Uint8Array): boolean {
  if (first.byteLength !== second.byteLength) return false;
  for (let index = 0; index < first.byteLength; index += 1) {
    if (first[index] !== second[index]) return false;
  }
  return true;
}

export function createPreviewMediaResolver(): PreviewMediaResolver {
  const cache = new LruCache<string, CachedMedia>(MAX_DERIVED_ENTRIES);
  const recency = new Map<string, number>();
  let cachedBytes = 0;
  let disposed = false;

  const remove = (key: string): void => {
    cache.delete(key);
    const cost = recency.get(key);
    if (cost !== undefined) cachedBytes -= cost;
    recency.delete(key);
  };

  const touch = (key: string): void => {
    const cost = recency.get(key);
    if (cost === undefined) return;
    recency.delete(key);
    recency.set(key, cost);
  };

  return {
    async resolve(path: string, bytes: Uint8Array, options: PreviewMediaOptions): Promise<ResolvedPreviewMedia> {
      if (disposed) throw new PreviewMediaError("disposed", "Preview media resolver has been disposed.");
      checkCancelled(options);
      const existing = cache.get(path);
      if (existing) {
        if (sameBytes(existing.sourceBytes, bytes)) {
          touch(path);
          return { ...existing.result, bytes: existing.result.bytes.slice() };
        }
      }
      const sourceBytes = bytes.slice();
      const result = await resolvePreviewMedia(path, sourceBytes, options);
      if (disposed) throw new PreviewMediaError("disposed", "Preview media resolver has been disposed.");
      checkCancelled(options);
      const concurrent = cache.get(path);
      if (concurrent) {
        touch(path);
        if (sameBytes(concurrent.sourceBytes, sourceBytes)) {
          return { ...concurrent.result, bytes: concurrent.result.bytes.slice() };
        }
        remove(path);
      }
      const cachedResult: ResolvedPreviewMedia = { ...result, bytes: result.bytes.slice() };
      const cost = utf8ByteLength(path) + sourceBytes.byteLength + cachedResult.bytes.byteLength;
      if (cost <= MAX_DERIVED_BYTES) {
        while (recency.size >= MAX_DERIVED_ENTRIES || cachedBytes > MAX_DERIVED_BYTES - cost) {
          const oldest = recency.keys().next().value as string | undefined;
          if (oldest === undefined) break;
          remove(oldest);
        }
        cache.set(path, { sourceBytes, result: cachedResult, cost });
        recency.set(path, cost);
        cachedBytes += cost;
      }
      return { ...result, bytes: result.bytes.slice() };
    },
    getCacheDiagnostics(): { readonly entries: number; readonly bytes: number } {
      return { entries: cache.size, bytes: cachedBytes };
    },
    dispose(): void {
      cache.clear();
      recency.clear();
      cachedBytes = 0;
      disposed = true;
    },
  };
}

function quoteCssFamily(name: string): string {
  let result = '"';
  for (const character of name) {
    const codePoint = character.codePointAt(0)!;
    if (character === "\\" || character === '"') result += `\\${codePoint.toString(16)} `;
    else if (codePoint === 0 || codePoint < 0x20 || codePoint === 0x7f) result += `\\${codePoint.toString(16)} `;
    else result += character;
  }
  return `${result}"`;
}

export async function checkPreviewFonts(
  ownerDocument: Document,
  names: readonly string[],
): Promise<readonly string[]> {
  const fontSet = (ownerDocument as Document & { fonts?: FontFaceSet }).fonts;
  const missing: string[] = [];
  const seen = new Set<string>();
  for (const name of names) {
    if (typeof name !== "string" || name.length === 0 || seen.has(name)) continue;
    seen.add(name);
    if (!fontSet || typeof fontSet.check !== "function") {
      missing.push(name);
      continue;
    }
    let loaded = false;
    try {
      loaded = fontSet.check(`12px ${quoteCssFamily(name)}`);
    } catch {
      loaded = false;
    }
    if (!loaded) missing.push(name);
  }
  return missing;
}
