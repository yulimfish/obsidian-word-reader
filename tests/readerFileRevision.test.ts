import assert from "node:assert/strict";
import test from "node:test";

import { ReaderFileRevisionTracker } from "../src/reader/fileRevision";

const file = {
  path: "reports/current.docx",
  extension: "docx",
  stat: { mtime: 100, size: 2048 },
};

void test("file revisions reject mtime, size, path, and format changes", () => {
  const tracker = new ReaderFileRevisionTracker();
  const revision = tracker.capture(file, "docx");
  assert.equal(tracker.isCurrent(revision, file, "docx"), true);
  assert.equal(tracker.isCurrent(revision, {
    ...file,
    stat: { ...file.stat, mtime: 101 },
  }, "docx"), false);
  assert.equal(tracker.isCurrent(revision, {
    ...file,
    stat: { ...file.stat, size: 4096 },
  }, "docx"), false);
  assert.equal(tracker.isCurrent(revision, file, "pptx"), false);
});

void test("vault events reject same-stat content replacement and rename", () => {
  const tracker = new ReaderFileRevisionTracker();
  const revision = tracker.capture(file, "docx");
  tracker.markChanged(file.path);
  assert.equal(tracker.isCurrent(revision, file, "docx"), false);

  const nextRevision = tracker.capture(file, "docx");
  tracker.markRenamed(file.path, "reports/renamed.docx");
  assert.equal(tracker.isCurrent(nextRevision, file, "docx"), false);
});

void test("a file replacement during an asynchronous read blocks stale commit", async () => {
  const tracker = new ReaderFileRevisionTracker();
  const revision = tracker.capture(file, "docx");
  let committed = false;
  const pendingRead = Promise.resolve().then(() => {
    if (tracker.isCurrent(revision, file, "docx")) {
      committed = true;
    }
  });

  tracker.markChanged(file.path);
  await pendingRead;
  assert.equal(committed, false);
});
