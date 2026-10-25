import type { WordReaderLanguage } from "../i18n";
import type { ReaderFileChangeKind } from "./fileRevision";
import type { ReaderWorkspaceRecoveryReason } from "./workspaceRecovery";

export type ReaderRecoveryReason =
  | ReaderFileChangeKind
  | ReaderWorkspaceRecoveryReason;

export interface ReaderRecoveryText {
  title: string;
  body: string;
  status: string;
  reload: string;
  details: string;
}

export function getReaderRecoveryText(
  language: WordReaderLanguage,
  reason: ReaderRecoveryReason,
): ReaderRecoveryText {
  const formatChanged = reason === "format-changed";
  const missing = reason === "deleted" || reason === "missing-file";
  if (language === "en") {
    return {
      title: formatChanged
        ? "Office file format changed"
        : missing
          ? "Office file is no longer available"
          : "Office file changed",
      body: formatChanged
        ? "The restored view no longer matches this file format. Its stale preview, search results, and caches were discarded."
        : missing
          ? "The restored source file could not be found. The reader released its pending work and cached resources."
          : "The source changed while it was being read. The stale session was invalidated before it could update the preview.",
      status: "Reader session invalidated safely",
      reload: "Reload current file",
      details: reason,
    };
  }
  return {
    title: formatChanged
      ? "Office 文件格式已更改"
      : missing
        ? "Office 文件已不可用"
        : "Office 文件已更改",
    body: formatChanged
      ? "恢复的视图与当前文件格式不再匹配；旧预览、搜索结果和缓存已丢弃。"
      : missing
        ? "找不到恢复视图对应的源文件；阅读器已释放待处理任务和缓存资源。"
        : "读取期间源文件发生变化；旧会话已在更新预览前安全失效。",
    status: "阅读会话已安全失效",
    reload: "重新加载当前文件",
    details: reason,
  };
}
