import {
  FileSystemAdapter,
  Platform,
  Plugin,
  TFile,
} from "obsidian";

import {
  PptxView,
  VIEW_TYPE_PPTX_READER,
} from "./PptxView";
import {
  WordView,
  VIEW_TYPE_WORD_READER,
} from "./WordView";
import {
  XlsxView,
  VIEW_TYPE_XLSX_READER,
} from "./XlsxView";
import { DOCX_ADAPTER } from "./docx/DocxAdapter";
import type { DocxEngine } from "./docx/docxEngineContract";
import {
  createDocxEngineLoader,
  DocxEngineLoadError,
  type DocxEngineBuildExpectation,
  type DocxEngineLoader,
} from "./docx/docxEngineLoader";
import { getDocxEngineNodeRuntime } from "./desktopRuntime";
import { getWordReaderText, type WordReaderText } from "./i18n";
import { PPTX_ADAPTER } from "./pptx/PptxAdapter";
import { getPptxReaderText } from "./pptx/pptxI18n";
import {
  XLSX_ADAPTER,
} from "./xlsx/XlsxAdapter";
import { getXlsxReaderText } from "./xlsx/xlsxI18n";
import {
  ReadingStateStore,
  type ReaderFileIdentity,
  type ReaderFormat,
  type ReaderViewState,
} from "./reader/readingState";
import {
  type ReaderFileChange,
  type ReaderFileDescriptor,
  type ReaderFileRevision,
  ReaderFileRevisionTracker,
} from "./reader/fileRevision";
import { recoverPluginData } from "./reader/pluginData";
import {
  isReaderSession,
  type ReaderSession,
} from "./reader/session";
import {
  planReaderWorkspaceRecovery,
  type ReaderWorkspaceEntry,
} from "./reader/workspaceRecovery";
import {
  DEFAULT_OFFICE_READER_SETTINGS,
  WordReaderSettingTab,
  type OfficeReaderSettings,
  normalizeOfficeReaderSettings,
} from "./settings";

const DATA_SAVE_DEBOUNCE_MS = 500;

declare const __DOCX_PLUGIN_VERSION__: string;
declare const __DOCX_BUILD_ID__: string;
declare const __DOCX_MANIFEST_MIN_APP_VERSION__: string;
declare const __DOCX_MANIFEST_SHA256__: string;
declare const __DOCX_MANIFEST_BYTES__: number;
declare const __DOCX_STYLES_SHA256__: string;
declare const __DOCX_STYLES_BYTES__: number;
declare const __DOCX_ENGINE_SHA256__: string;
declare const __DOCX_ENGINE_BYTES__: number;

const DOCX_BUILD_EXPECTATION: DocxEngineBuildExpectation = {
  pluginId: "word-reader",
  pluginVersion: typeof __DOCX_PLUGIN_VERSION__ === "undefined"
    ? "3.5.0"
    : __DOCX_PLUGIN_VERSION__,
  buildId: typeof __DOCX_BUILD_ID__ === "undefined"
    ? "development"
    : __DOCX_BUILD_ID__,
  abiVersion: 1,
  minAppVersion: typeof __DOCX_MANIFEST_MIN_APP_VERSION__ === "undefined"
    ? "1.12.7"
    : __DOCX_MANIFEST_MIN_APP_VERSION__,
  mainMaxBytes: 512_000,
  assets: {
    "manifest.json": {
      bytes: typeof __DOCX_MANIFEST_BYTES__ === "undefined"
        ? 0
        : __DOCX_MANIFEST_BYTES__,
      sha256: typeof __DOCX_MANIFEST_SHA256__ === "undefined"
        ? ""
        : __DOCX_MANIFEST_SHA256__,
    },
    "styles.css": {
      bytes: typeof __DOCX_STYLES_BYTES__ === "undefined"
        ? 0
        : __DOCX_STYLES_BYTES__,
      sha256: typeof __DOCX_STYLES_SHA256__ === "undefined"
        ? ""
        : __DOCX_STYLES_SHA256__,
    },
    "docx-engine.cjs": {
      bytes: typeof __DOCX_ENGINE_BYTES__ === "undefined"
        ? 0
        : __DOCX_ENGINE_BYTES__,
      sha256: typeof __DOCX_ENGINE_SHA256__ === "undefined"
        ? ""
        : __DOCX_ENGINE_SHA256__,
    },
  },
};

export default class WordReaderPlugin extends Plugin {
  settings: OfficeReaderSettings = normalizeOfficeReaderSettings(
    DEFAULT_OFFICE_READER_SETTINGS,
  );
  private readingStates = new ReadingStateStore();
  private readonly fileRevisions = new ReaderFileRevisionTracker();
  private dataSaveTimer: number | null = null;
  private dataSavePromise: Promise<void> = Promise.resolve();
  private docxEngineLoader: DocxEngineLoader | null = null;

  async onload(): Promise<void> {
    this.docxEngineLoader = createDocxEngineLoader(
      this,
      DOCX_BUILD_EXPECTATION,
      {
        isDesktopApp: () => Platform.isDesktopApp,
        FileSystemAdapter,
        getRuntime: getDocxEngineNodeRuntime,
      },
    );
    await this.loadSettings();
    this.addSettingTab(new WordReaderSettingTab(this.app, this));
    const text = this.text;
    const pptxText = getPptxReaderText(this.settings.common.language);
    const xlsxText = getXlsxReaderText(this.settings.common.language);

    this.registerView(
      VIEW_TYPE_WORD_READER,
      (leaf) => new WordView(leaf, this),
    );
    this.registerView(
      VIEW_TYPE_PPTX_READER,
      (leaf) => new PptxView(leaf, this),
    );
    this.registerView(
      VIEW_TYPE_XLSX_READER,
      (leaf) => new XlsxView(leaf, this),
    );
    this.registerExtensions(
      [...DOCX_ADAPTER.extensions],
      DOCX_ADAPTER.viewType,
    );
    this.registerExtensions(
      [...PPTX_ADAPTER.extensions],
      PPTX_ADAPTER.viewType,
    );
    this.registerExtensions(
      [...XLSX_ADAPTER.extensions],
      XLSX_ADAPTER.viewType,
    );
    this.registerEvent(this.app.vault.on("modify", (file) => {
      if (file instanceof TFile) {
        this.notifyReaderFileChange({ kind: "modified", path: file.path });
      }
    }));
    this.registerEvent(this.app.vault.on("delete", (file) => {
      if (file instanceof TFile) {
        this.notifyReaderFileChange({ kind: "deleted", path: file.path });
      }
    }));
    this.registerEvent(this.app.vault.on("rename", (file, oldPath) => {
      if (file instanceof TFile) {
        this.fileRevisions.markRenamed(oldPath, file.path);
        this.notifyReaderFileChange({
          kind: "renamed",
          path: oldPath,
          newPath: file.path,
          newExtension: file.extension,
        }, false);
      }
    }));
    this.app.workspace.onLayoutReady(() => {
      try {
        this.recoverWorkspaceViews();
      } catch (error) {
        console.error("Office Reader could not inspect restored views", error);
      }
    });

    this.registerReaderCommand(
      "reload",
      text.commands.reload,
      (session) => session.capabilities.reload,
      (session) => session.reload(),
    );
    this.registerReaderCommand(
      "copy-text",
      text.commands.copyText,
      (session) => session.capabilities.copyText && Boolean(session.copyText),
      (session) => session.copyText?.(),
    );
    this.registerReaderCommand(
      "copy-markdown",
      text.commands.copyMarkdown,
      (session) =>
        session.adapter.format === "docx" &&
        session.capabilities.copyMarkdown && Boolean(session.copyMarkdown),
      (session) => session.copyMarkdown?.(),
    );
    this.registerReaderCommand(
      "create-note",
      text.commands.createNote,
      (session) =>
        session.adapter.format === "docx" &&
        session.capabilities.summaryNote &&
        Boolean(session.createSummaryNote),
      (session) => session.createSummaryNote?.(),
    );
    this.registerReaderCommand(
      "open-external",
      text.commands.openExternal,
      (session) =>
        session.capabilities.openExternal && Boolean(session.openExternal),
      (session) => session.openExternal?.(),
    );

    this.registerReaderCommand(
      "previous-slide",
      pptxText.commands.previousSlide,
      (session) =>
        session.capabilities.paged && Boolean(session.previousPage),
      (session) => session.previousPage?.(),
    );
    this.registerReaderCommand(
      "next-slide",
      pptxText.commands.nextSlide,
      (session) =>
        session.capabilities.paged && Boolean(session.nextPage),
      (session) => session.nextPage?.(),
    );
    this.registerReaderCommand(
      "toggle-presentation-fullscreen",
      pptxText.commands.toggleFullscreen,
      (session) =>
        session.capabilities.fullscreen && Boolean(session.toggleFullscreen),
      (session) => session.toggleFullscreen?.(),
    );
    this.registerReaderCommand(
      "copy-presentation-text",
      pptxText.commands.copySlideText,
      (session) =>
        session.capabilities.paged &&
        session.capabilities.copyText &&
        Boolean(session.copyText),
      (session) => session.copyText?.(),
    );
    this.registerReaderCommand(
      "copy-presentation-render-diagnostics",
      pptxText.commands.copyRenderDiagnostics,
      (session) =>
        session.capabilities.diagnostics && Boolean(session.copyDiagnostics),
      (session) => session.copyDiagnostics?.(),
    );
    this.registerReaderCommand(
      "create-presentation-note",
      pptxText.commands.createSummaryNote,
      (session) =>
        session.capabilities.paged &&
        session.capabilities.summaryNote &&
        Boolean(session.createSummaryNote),
      (session) => session.createSummaryNote?.(),
    );
    this.registerReaderCommand(
      "toggle-presentation-notes",
      pptxText.commands.toggleNotes,
      (session) =>
        session.capabilities.notes && Boolean(session.toggleNotes),
      (session) => session.toggleNotes?.(),
    );
    this.registerReaderCommand(
      "search-presentation",
      pptxText.commands.focusSearch,
      (session) =>
        session.capabilities.paged &&
        session.capabilities.search &&
        Boolean(session.focusSearch),
      (session) => session.focusSearch?.(),
    );
    this.registerReaderCommand(
      "search-spreadsheet",
      xlsxText.commands.focusSearch,
      (session) =>
        session.adapter.format === "xlsx" &&
        session.capabilities.search &&
        Boolean(session.focusSearch),
      (session) => session.focusSearch?.(),
    );
    this.registerReaderCommand(
      "go-to-spreadsheet-cell",
      xlsxText.commands.goToCell,
      (session) =>
        session.adapter.format === "xlsx" &&
        session.capabilities.navigation &&
        Boolean(session.focusNameBox),
      (session) => session.focusNameBox?.(),
    );
    this.registerReaderCommand(
      "copy-spreadsheet-values",
      xlsxText.commands.copyValues,
      (session) =>
        session.adapter.format === "xlsx" &&
        session.capabilities.copyText &&
        Boolean(session.copyText),
      (session) => session.copyText?.(),
    );
    this.registerReaderCommand(
      "copy-spreadsheet-formulas",
      xlsxText.commands.copyFormulas,
      (session) =>
        session.adapter.format === "xlsx" &&
        session.capabilities.copyText &&
        Boolean(session.copyFormulas),
      (session) => session.copyFormulas?.(),
    );
    this.registerReaderCommand(
      "copy-spreadsheet-markdown",
      xlsxText.commands.copyMarkdown,
      (session) =>
        session.adapter.format === "xlsx" &&
        session.capabilities.copyMarkdown &&
        Boolean(session.copyMarkdown),
      (session) => session.copyMarkdown?.(),
    );
    this.registerReaderCommand(
      "create-spreadsheet-note",
      xlsxText.commands.createSummaryNote,
      (session) =>
        session.adapter.format === "xlsx" &&
        session.capabilities.summaryNote &&
        Boolean(session.createSummaryNote),
      (session) => session.createSummaryNote?.(),
    );
  }

  onunload(): void {
    this.docxEngineLoader?.dispose();
    this.docxEngineLoader = null;
    for (const leaf of this.app.workspace.getLeavesOfType(VIEW_TYPE_WORD_READER)) {
      const session = leaf.view;
      if (isReaderSession(session) && session.adapter.format === "docx") {
        void session.onClose().catch((error: unknown) => {
          console.error("Office Reader could not close a DOCX reader during plugin unload", error);
        });
      }
    }
    if (this.dataSaveTimer !== null) {
      window.clearTimeout(this.dataSaveTimer);
      this.dataSaveTimer = null;
      void this.persistData().catch((error: unknown) => {
        console.error("Office Reader could not save plugin data", error);
      });
    }
  }

  get text(): WordReaderText {
    return getWordReaderText(this.settings.common.language);
  }

  loadDocxEngine(buffer?: ArrayBuffer): Promise<DocxEngine> {
    if (!this.docxEngineLoader) {
      return Promise.reject(
        new DocxEngineLoadError("disposed", "The DOCX engine loader is unavailable."),
      );
    }
    return this.docxEngineLoader.load(buffer);
  }

  refreshWordReaderViews(): void {
    for (const viewType of [
      VIEW_TYPE_WORD_READER,
      VIEW_TYPE_PPTX_READER,
      VIEW_TYPE_XLSX_READER,
    ]) {
      for (const leaf of this.app.workspace.getLeavesOfType(viewType)) {
        if (isReaderSession(leaf.view)) {
          leaf.view.refreshInterfaceLanguage();
        }
      }
    }
  }

  async loadSettings(): Promise<void> {
    let loadedData: unknown;
    try {
      loadedData = await this.loadData();
    } catch (error) {
      console.error("Office Reader could not load plugin data", error);
      const recovered = recoverPluginData(undefined);
      this.settings = recovered.settings;
      this.readingStates = recovered.readingStates;
      return;
    }
    const recovered = recoverPluginData(loadedData);
    this.settings = recovered.settings;
    this.readingStates = recovered.readingStates;
    if (recovered.requiresRepair) {
      await this.persistData().catch((error: unknown) => {
        console.error("Office Reader could not repair plugin data", error);
      });
    }
  }

  async saveSettings(): Promise<void> {
    this.settings = normalizeOfficeReaderSettings(this.settings);
    await this.flushData();
  }

  getReadingState(
    file: TFile | ReaderFileIdentity | string,
  ): ReaderViewState | undefined {
    return typeof file === "string"
      ? this.readingStates.get(file)
      : this.readingStates.get(toReaderIdentity(file));
  }

  updateReadingState(
    file: TFile | ReaderFileIdentity | string,
    state: ReaderViewState,
  ): void {
    if (typeof file === "string") {
      this.readingStates.set(file, state);
    } else {
      this.readingStates.set(toReaderIdentity(file), state);
    }
    this.scheduleDataSave();
  }

  captureFileRevision(
    file: ReaderFileDescriptor,
    format: ReaderFormat,
  ): ReaderFileRevision {
    return this.fileRevisions.capture(file, format);
  }

  isFileRevisionCurrent(
    revision: ReaderFileRevision | null,
    file: ReaderFileDescriptor | null,
    format: ReaderFormat,
  ): boolean {
    return this.fileRevisions.isCurrent(revision, file, format);
  }

  async flushData(): Promise<void> {
    if (this.dataSaveTimer !== null) {
      window.clearTimeout(this.dataSaveTimer);
      this.dataSaveTimer = null;
    }
    await this.persistData().catch((error: unknown) => {
      console.error("Office Reader could not save plugin data", error);
    });
  }

  private registerReaderCommand(
    id: string,
    name: string,
    isAvailable: (session: ReaderSession) => boolean,
    run: (session: ReaderSession) => void | Promise<void>,
  ): void {
    this.addCommand({
      id,
      name,
      checkCallback: (checking) => {
        const session = this.getActiveReaderSession();
        if (!session?.file || !isAvailable(session)) {
          return false;
        }
        if (!checking) {
          void Promise.resolve(run(session)).catch((error: unknown) => {
            console.error("Office Reader command failed", error);
          });
        }
        return true;
      },
    });
  }

  private getActiveReaderSession(): ReaderSession | null {
    const view = this.app.workspace.getMostRecentLeaf()?.view;
    return isReaderSession(view) ? view : null;
  }

  private scheduleDataSave(): void {
    if (this.dataSaveTimer !== null) {
      window.clearTimeout(this.dataSaveTimer);
    }
    this.dataSaveTimer = window.setTimeout(() => {
      this.dataSaveTimer = null;
      void this.persistData().catch((error: unknown) => {
        console.error("Office Reader could not save plugin data", error);
      });
    }, DATA_SAVE_DEBOUNCE_MS);
  }

  private persistData(): Promise<void> {
    const data = {
      ...this.settings,
      readingStates: this.readingStates.serialize(),
    };
    this.dataSavePromise = this.dataSavePromise
      .catch(() => undefined)
      .then(async () => {
        await this.saveData(data);
      });
    return this.dataSavePromise;
  }

  private notifyReaderFileChange(
    change: ReaderFileChange,
    markChanged = true,
  ): void {
    if (markChanged) {
      this.fileRevisions.markChanged(change.path);
    }
    if (this.readingStates.invalidate(change.path)) {
      this.scheduleDataSave();
    }
    for (const session of this.getReaderSessions()) {
      try {
        session.handleFileChange?.(change);
      } catch (error) {
        console.error("Office Reader could not invalidate a changed file", error);
      }
    }
  }

  private recoverWorkspaceViews(): void {
    const sessions = this.getReaderSessions();
    const entries: ReaderWorkspaceEntry[] = sessions.map((session, index) => {
      const file = session.file;
      const currentFile = file
        ? this.app.vault.getAbstractFileByPath(file.path)
        : null;
      const availableFile = currentFile instanceof TFile ? currentFile : null;
      return {
        id: String(index),
        format: session.adapter.format,
        path: availableFile?.path ?? null,
        extension: availableFile?.extension ?? file?.extension ?? null,
      };
    });
    for (const action of planReaderWorkspaceRecovery(entries)) {
      if (!action.reason) {
        continue;
      }
      try {
        sessions[Number(action.id)]?.handleWorkspaceRecovery?.(action.reason);
      } catch (error) {
        console.error("Office Reader could not recover a workspace view", error);
      }
    }
  }

  private getReaderSessions(): ReaderSession[] {
    const sessions: ReaderSession[] = [];
    for (const viewType of [
      VIEW_TYPE_WORD_READER,
      VIEW_TYPE_PPTX_READER,
      VIEW_TYPE_XLSX_READER,
    ]) {
      for (const leaf of this.app.workspace.getLeavesOfType(viewType)) {
        if (isReaderSession(leaf.view)) {
          sessions.push(leaf.view);
        }
      }
    }
    return sessions;
  }
}

function toReaderIdentity(
  file: TFile | ReaderFileIdentity,
): ReaderFileIdentity {
  if ("format" in file) {
    return file;
  }
  return {
    path: file.path,
    mtime: file.stat.mtime,
    format: getReaderFormat(file.extension),
  };
}

function getReaderFormat(extension: string): ReaderFormat {
  switch (extension.toLowerCase()) {
    case "pptx":
      return "pptx";
    case "xlsx":
      return "xlsx";
    default:
      return "docx";
  }
}
