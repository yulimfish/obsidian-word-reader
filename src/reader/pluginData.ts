import {
  ReadingStateStore,
  type PersistedReaderViewState,
} from "./readingState";
import {
  type OfficeReaderSettings,
  migrateSettings,
} from "../settingsModel";

export interface RecoveredPluginData {
  settings: OfficeReaderSettings;
  readingStates: ReadingStateStore;
  normalizedData: OfficeReaderSettings & {
    readingStates: PersistedReaderViewState[];
  };
  requiresRepair: boolean;
}

const ROOT_KEYS = [
  "schemaVersion",
  "common",
  "docx",
  "pptx",
  "xlsx",
  "readingStates",
] as const;

/**
 * Recovers every independently valid field and drops partial/unknown data.
 * Callers should not persist the fallback if loading the source itself threw,
 * because doing so could overwrite a file that remains recoverable by hand.
 */
export function recoverPluginData(data: unknown): RecoveredPluginData {
  const settings = migrateSettings(data);
  const readingStates = new ReadingStateStore(
    undefined,
    getReadingStateData(data),
  );
  const normalizedData = {
    ...settings,
    readingStates: readingStates.serialize(),
  };
  return {
    settings,
    readingStates,
    normalizedData,
    requiresRepair: !isCanonicalPluginData(data, normalizedData),
  };
}

function getReadingStateData(data: unknown): unknown {
  return isRecord(data) ? data.readingStates : undefined;
}

function isCanonicalPluginData(
  data: unknown,
  normalized: RecoveredPluginData["normalizedData"],
): boolean {
  if (!isRecord(data)) {
    return false;
  }
  if (Object.keys(data).some((key) => !ROOT_KEYS.includes(
    key as (typeof ROOT_KEYS)[number],
  ))) {
    return false;
  }
  const comparable = {
    schemaVersion: data.schemaVersion,
    common: data.common,
    docx: data.docx,
    pptx: data.pptx,
    xlsx: data.xlsx,
    readingStates: data.readingStates ?? [],
  };
  return equalJsonValue(comparable, normalized);
}

function equalJsonValue(left: unknown, right: unknown): boolean {
  if (Object.is(left, right)) {
    return true;
  }
  if (Array.isArray(left) || Array.isArray(right)) {
    return Array.isArray(left) &&
      Array.isArray(right) &&
      left.length === right.length &&
      left.every((value, index) => equalJsonValue(value, right[index]));
  }
  if (!isRecord(left) || !isRecord(right)) {
    return false;
  }
  const leftKeys = Object.keys(left).sort();
  const rightKeys = Object.keys(right).sort();
  return leftKeys.length === rightKeys.length &&
    leftKeys.every((key, index) =>
      key === rightKeys[index] && equalJsonValue(left[key], right[key])
    );
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}
