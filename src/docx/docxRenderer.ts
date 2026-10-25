import type {
  BoundDocxRenderer,
  DocxEngine,
  DocxRenderOptions,
  DocxRenderResult,
} from "./docxEngineContract";

export type {
  DocxEngine,
  DocxEngineHost,
  DocxRenderOptions,
  DocxRenderResult,
} from "./docxEngineContract";

export function createDocxRenderer(engine: DocxEngine): BoundDocxRenderer {
  return (buffer, targetEl, options = {}) => engine.renderDocx(
    buffer,
    targetEl,
    {
      ownerDocument: targetEl.ownerDocument,
      isCancelled: options.isCancelled ?? (() => false),
      yieldControl: options.yieldControl ?? (async () => {}),
    },
  );
}

/**
 * Kept as a fail-closed compatibility entry point. Runtime callers must bind a
 * verified engine with createDocxRenderer instead of importing the sidecar.
 */
export async function renderDocx(
  _buffer: ArrayBuffer,
  _targetEl: HTMLElement,
  _options: DocxRenderOptions = {},
): Promise<DocxRenderResult> {
  throw new Error("DOCX renderer requires a verified engine binding.");
}
