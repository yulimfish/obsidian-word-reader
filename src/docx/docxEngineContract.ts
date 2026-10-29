import type { PreviewInventory, PreviewIssue } from "../ooxml/previewTypes";

export interface DocxRenderOptions {
  readonly isCancelled?: () => boolean;
  readonly yieldControl?: () => Promise<void>;
}

export interface DocxRenderResult {
  readonly inventory: PreviewInventory;
  readonly issues: readonly PreviewIssue[];
}

export interface DocxEngineHost {
  readonly ownerDocument: Document;
  readonly isCancelled: () => boolean;
  readonly yieldControl: () => Promise<void>;
}

export interface PreviewImage {
  readonly bytes: Uint8Array;
  readonly mimeType: "image/svg+xml" | "image/png" | "image/jpeg" | "image/bmp";
}

export interface DocxEngine {
  readonly abiVersion: 1;
  readonly pluginVersion: string;
  readonly buildId: string;
  readonly renderDocx: (
    buffer: ArrayBuffer,
    targetEl: HTMLElement,
    options: DocxEngineHost,
  ) => Promise<DocxRenderResult>;
  readonly resolvePreviewImage: (
    path: string,
    bytes: Uint8Array,
    options: DocxRenderOptions,
  ) => Promise<PreviewImage>;
}

export type BoundDocxRenderer = (
  buffer: ArrayBuffer,
  targetEl: HTMLElement,
  options?: DocxRenderOptions,
) => Promise<DocxRenderResult>;
