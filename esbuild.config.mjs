import fs from "node:fs";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

import { buildDocxDistribution } from "./scripts/docx-build-config.mjs";

const root = path.dirname(fileURLToPath(import.meta.url));
const production = process.argv[2] === "production";

function describeSnapshot(snapshot) {
  console.log(
    `DOCX distribution built: main=${snapshot.bytes.main} B, ` +
      `engine=${snapshot.bytes.engine} B, buildId=${snapshot.buildId}, ` +
      `mainSha256=${snapshot.hashes.main}, engineSha256=${snapshot.hashes.engine}`,
  );
}

async function buildAndPublish() {
  const snapshot = await buildDocxDistribution({
    root,
    production,
    write: true,
  });
  describeSnapshot(snapshot);
}

function watchInputs() {
  const watchers = new Set();
  const watchedDirectories = new Set();
  let rebuildTimer = null;
  let rebuilding = false;
  let rebuildAgain = false;

  const scheduleRebuild = () => {
    if (rebuildTimer !== null) {
      clearTimeout(rebuildTimer);
    }
    rebuildTimer = setTimeout(() => {
      rebuildTimer = null;
      void rebuild();
    }, 100);
  };

  const watchDirectory = async (directory) => {
    if (watchedDirectories.has(directory)) {
      return;
    }
    watchedDirectories.add(directory);
    const watcher = fs.watch(directory, (eventType, filename) => {
      if (filename) {
        const changedPath = path.join(directory, filename.toString());
        void fs.promises.lstat(changedPath).then((stats) => {
          if (stats.isDirectory() && !stats.isSymbolicLink()) {
            return watchDirectory(changedPath);
          }
          return undefined;
        }).catch((error) => {
          if (error?.code !== "ENOENT") {
            console.error(`Could not inspect changed build input ${changedPath}`, error);
          }
        });
      }
      scheduleRebuild();
    });
    watcher.on("error", (error) => {
      console.error(`Build input watcher failed for ${directory}`, error);
    });
    watchers.add(watcher);
    const entries = await fs.promises.readdir(directory, { withFileTypes: true });
    for (const entry of entries) {
      if (entry.isDirectory() && !entry.isSymbolicLink()) {
        await watchDirectory(path.join(directory, entry.name));
      }
    }
  };

  const watchFile = (file) => {
    const watcher = fs.watch(file, scheduleRebuild);
    watcher.on("error", (error) => {
      console.error(`Build input watcher failed for ${file}`, error);
    });
    watchers.add(watcher);
  };

  const rebuild = async () => {
    if (rebuilding) {
      rebuildAgain = true;
      return;
    }
    rebuilding = true;
    try {
      await buildAndPublish();
    } catch (error) {
      console.error("DOCX distribution rebuild failed; the previous complete snapshot was left untouched.", error);
    } finally {
      rebuilding = false;
      if (rebuildAgain) {
        rebuildAgain = false;
        scheduleRebuild();
      }
    }
  };

  return Promise.all([
    watchDirectory(path.join(root, "src")),
    watchDirectory(path.join(root, "scripts")),
  ]).then(() => {
    for (const file of [
      "package.json",
      "package-lock.json",
      "manifest.json",
      "styles.css",
      "esbuild.config.mjs",
    ]) {
      watchFile(path.join(root, file));
    }
    process.on("SIGINT", () => {
      for (const watcher of watchers) {
        watcher.close();
      }
      if (rebuildTimer !== null) {
        clearTimeout(rebuildTimer);
      }
      process.exit(0);
    });
    console.log("Watching source, build configuration, and package assets for paired DOCX builds.");
  });
}

await buildAndPublish();
if (!production) {
  await watchInputs();
}
