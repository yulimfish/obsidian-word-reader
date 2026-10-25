import { LruCache } from "./lruCache";
import type { ReaderFormat } from "./readingState";

export interface ReaderFileDescriptor {
  path: string;
  extension: string;
  stat: {
    mtime: number;
    size: number;
  };
}

export interface ReaderFileRevision {
  path: string;
  format: ReaderFormat;
  mtime: number;
  size: number;
  generation: number;
}

export type ReaderFileChangeKind = "modified" | "deleted" | "renamed";

export interface ReaderFileChange {
  kind: ReaderFileChangeKind;
  path: string;
  newPath?: string;
  newExtension?: string;
}

const MAX_TRACKED_FILE_REVISIONS = 512;

/** Tracks vault events as well as mtime/size so same-stat replacements are invalidated. */
export class ReaderFileRevisionTracker {
  private readonly generations = new LruCache<string, number>(
    MAX_TRACKED_FILE_REVISIONS,
  );

  capture(
    file: ReaderFileDescriptor,
    format: ReaderFormat,
  ): ReaderFileRevision {
    return {
      path: file.path,
      format,
      mtime: file.stat.mtime,
      size: file.stat.size,
      generation: this.getGeneration(file.path),
    };
  }

  markChanged(path: string): number {
    return this.generations.set(path, this.getGeneration(path) + 1);
  }

  markRenamed(oldPath: string, newPath: string): void {
    this.markChanged(oldPath);
    this.markChanged(newPath);
  }

  isCurrent(
    revision: ReaderFileRevision | null,
    file: ReaderFileDescriptor | null,
    format: ReaderFormat,
  ): boolean {
    return Boolean(
      revision &&
      file &&
      revision.path === file.path &&
      revision.format === format &&
      revision.mtime === file.stat.mtime &&
      revision.size === file.stat.size &&
      revision.generation === this.getGeneration(file.path),
    );
  }

  private getGeneration(path: string): number {
    return this.generations.get(path) ?? 0;
  }
}
