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

export interface DocxEngine {
  readonly abiVersion: 1;
  readonly pluginVersion: string;
  readonly buildId: string;
  readonly renderDocx: (
    buffer: ArrayBuffer,
    targetEl: HTMLElement,
    options: DocxEngineHost,
  ) => Promise<DocxRenderResult>;
}

export type BoundDocxRenderer = (
  buffer: ArrayBuffer,
  targetEl: HTMLElement,
  options?: DocxRenderOptions,
) => Promise<DocxRenderResult>;
