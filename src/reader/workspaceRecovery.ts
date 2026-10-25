import type { ReaderFormat } from "./readingState";

export type ReaderWorkspaceRecoveryReason =
  | "missing-file"
  | "format-changed";

export interface ReaderWorkspaceEntry {
  id: string;
  format: ReaderFormat;
  path: string | null;
  extension: string | null;
}

export interface ReaderWorkspaceRecoveryAction {
  id: string;
  reason: ReaderWorkspaceRecoveryReason | null;
  duplicate: boolean;
}

const FORMAT_EXTENSIONS: Record<ReaderFormat, readonly string[]> = {
  docx: ["docx", "doc"],
  pptx: ["pptx"],
  xlsx: ["xlsx", "xls"],
};

/** Creates a non-throwing recovery plan for all persisted reader leaves. */
export function planReaderWorkspaceRecovery(
  entries: readonly ReaderWorkspaceEntry[],
): ReaderWorkspaceRecoveryAction[] {
  const seen = new Set<string>();
  return entries.map((entry) => {
    const extension = entry.extension?.toLowerCase() ?? null;
    const reason = !entry.path || !extension
      ? "missing-file"
      : FORMAT_EXTENSIONS[entry.format].includes(extension)
        ? null
        : "format-changed";
    const key = entry.path ? `${entry.format}:${entry.path}` : null;
    const duplicate = key !== null && seen.has(key);
    if (key) {
      seen.add(key);
    }
    return { id: entry.id, reason, duplicate };
  });
}
