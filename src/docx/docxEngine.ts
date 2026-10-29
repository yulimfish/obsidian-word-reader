// SECURITY: The production build packages this entire DOCX render closure in a
// separately verified local CJS sidecar. It parses only a prepared inert copy.
import {
  parseAsync,
  renderDocument,
  type Options,
} from "docx-preview";

import { annotateDocxHeadingLevels } from "./docxOutline";
import { prepareDocxPreviewPackage } from "./docxPreviewPackage";
import { resolvePreviewMedia } from "../ooxml/previewMedia";
import type {
  DocxEngine,
  DocxEngineHost,
  DocxRenderResult,
} from "./docxEngineContract";

export type { DocxEngine, DocxEngineHost, DocxRenderResult };

declare const __DOCX_PLUGIN_VERSION__: string;
declare const __DOCX_BUILD_ID__: string;

export const abiVersion = 1 as const;
export const pluginVersion = typeof __DOCX_PLUGIN_VERSION__ === "undefined"
  ? "3.5.0"
  : __DOCX_PLUGIN_VERSION__;
export const buildId = typeof __DOCX_BUILD_ID__ === "undefined"
  ? "development"
  : __DOCX_BUILD_ID__;

const DOCX_RENDER_OPTIONS: Partial<Options> = {
  className: "docx",
  inWrapper: true,
  ignoreWidth: false,
  ignoreHeight: false,
  ignoreFonts: false,
  breakPages: true,
  // Blob URLs avoid duplicating large embedded images as base64 strings.
  // DocxSession owns and revokes URLs found in the rendered detached target.
  useBase64URL: false,
  experimental: true,
  trimXmlDeclaration: true,
};

function throwIfCancelled(host: DocxEngineHost): void {
  if (host.isCancelled()) {
    const error = new Error("DOCX rendering was cancelled.");
    error.name = "AbortError";
    throw error;
  }
}

export async function resolvePreviewImage(
  path: string,
  bytes: Uint8Array,
  options: { isCancelled?: () => boolean; yieldControl?: () => Promise<void> },
) {
  const result = await resolvePreviewMedia(path, bytes, {
    isCancelled: options.isCancelled ?? (() => false),
    yieldControl: options.yieldControl ?? (async () => {}),
  });
  return { bytes: result.bytes, mimeType: result.mimeType };
}

export async function renderDocx(
  buffer: ArrayBuffer,
  targetEl: HTMLElement,
  host: DocxEngineHost,
): Promise<DocxRenderResult> {
  if (targetEl.ownerDocument !== host.ownerDocument) {
    throw new Error("DOCX render target does not belong to the owning document.");
  }
  throwIfCancelled(host);
  const prepared = await prepareDocxPreviewPackage(buffer, {
    isCancelled: host.isCancelled,
    yieldControl: host.yieldControl,
  });
  throwIfCancelled(host);
  const document: unknown = await parseAsync(prepared.buffer, DOCX_RENDER_OPTIONS);
  throwIfCancelled(host);
  annotateDocxHeadingLevels(document);
  await renderDocument(document, targetEl, undefined, DOCX_RENDER_OPTIONS);
  throwIfCancelled(host);
  return { inventory: prepared.inventory, issues: prepared.issues };
}
