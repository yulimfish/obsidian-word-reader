import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import JSZip from "jszip";

import {
  createEquationFixture,
  type EquationFixtureOptions,
} from "./fixtureBuilders";

const rootDir = process.cwd();
const runnerPath = path.join(rootDir, "scripts/run-fidelity-tests.mjs");
const testRoot = path.join(rootDir, "tests");
const outputRoot = path.join(rootDir, ".test-dist");

type OwnedFixtureRoot = {
  directory: string;
  parentDirectory: string;
  entriesBefore: string[];
};

function isPathWithin(directory: string, candidate: string): boolean {
  const relativePath = path.relative(directory, candidate);
  return (
    relativePath === "" ||
    (relativePath !== ".." &&
      !relativePath.startsWith(`..${path.sep}`) &&
      !path.isAbsolute(relativePath))
  );
}

function createOwnedFixtureRoot(
  prefix: string,
  parentPath = outputRoot,
): OwnedFixtureRoot {
  const parentStats = fs.lstatSync(parentPath);
  if (parentStats.isSymbolicLink()) {
    throw new Error("Fixture root parent must not be a symbolic link.");
  }
  if (!parentStats.isDirectory()) {
    throw new Error("Fixture root parent must be a directory.");
  }

  const canonicalProjectRoot = fs.realpathSync(rootDir);
  const canonicalParent = fs.realpathSync(parentPath);
  if (!isPathWithin(canonicalProjectRoot, canonicalParent)) {
    throw new Error("Fixture root parent must be inside the repository.");
  }

  const entriesBefore = fs.readdirSync(canonicalParent).sort();
  const directory = fs.mkdtempSync(path.join(canonicalParent, prefix));
  try {
    const canonicalDirectory = fs.realpathSync(directory);
    if (
      canonicalDirectory === canonicalParent ||
      !isPathWithin(canonicalParent, canonicalDirectory)
    ) {
      throw new Error("Fixture root must be a new child of its parent.");
    }
  } catch (error) {
    fs.rmSync(directory, { recursive: true, force: true });
    throw error;
  }

  return {
    directory,
    parentDirectory: canonicalParent,
    entriesBefore,
  };
}

function removeOwnedFixtureRoot(ownedRoot: OwnedFixtureRoot): void {
  const rootStats = fs.lstatSync(ownedRoot.directory);
  if (
    rootStats.isSymbolicLink() ||
    !rootStats.isDirectory() ||
    !isPathWithin(
      ownedRoot.parentDirectory,
      fs.realpathSync(ownedRoot.directory),
    )
  ) {
    throw new Error("Refusing to remove a fixture root that is not owned.");
  }

  fs.rmSync(ownedRoot.directory, { recursive: true });
  assert.deepEqual(
    fs.readdirSync(ownedRoot.parentDirectory).sort(),
    ownedRoot.entriesBefore,
  );
}

function runRunner(args: string[]) {
  return spawnSync(process.execPath, [runnerPath, ...args], {
    cwd: rootDir,
    encoding: "utf8",
  });
}

void test("selected tests leave unrelated artifacts intact", () => {
  const fixtureDir = fs.mkdtempSync(
    path.join(testRoot, "fidelity-runner-fixture-"),
  );
  const artifactRoot = createOwnedFixtureRoot("unrelated-fidelity-artifact-");
  const artifactDir = artifactRoot.directory;
  const artifactPath = path.join(artifactDir, "sentinel.txt");

  try {
    const selectedPath = path.join(fixtureDir, "selected.test.ts");
    const unselectedPath = path.join(fixtureDir, "unselected.test.ts");
    const selectedMarker = path.join(fixtureDir, "selected.marker");
    const filteredMarker = path.join(fixtureDir, "filtered.marker");
    const unselectedMarker = path.join(fixtureDir, "unselected.marker");
    fs.writeFileSync(
      selectedPath,
      "import fs from \"node:fs\";\n" +
        "import test from \"node:test\";\n" +
        `test("selected pattern", () => fs.writeFileSync(${JSON.stringify(selectedMarker)}, "ran"));\n` +
        `test("filtered pattern", () => fs.writeFileSync(${JSON.stringify(filteredMarker)}, "ran"));\n`,
    );
    fs.writeFileSync(
      unselectedPath,
      "import fs from \"node:fs\";\n" +
        "import test from \"node:test\";\n" +
        `test("unselected file", () => fs.writeFileSync(${JSON.stringify(unselectedMarker)}, "ran"));\n`,
    );
    fs.writeFileSync(
      path.join(fixtureDir, "not-selected.test.ts"),
      "export const = ;\n",
    );
    fs.writeFileSync(artifactPath, "preserve-this-artifact");

    const result = runRunner([
      path.relative(rootDir, selectedPath),
      "--test-name-pattern",
      "selected pattern",
    ]);

    assert.equal(result.status, 0, `${result.stdout}\n${result.stderr}`);
    assert.equal(fs.readFileSync(selectedMarker, "utf8"), "ran");
    assert.equal(fs.existsSync(filteredMarker), false);
    assert.equal(fs.existsSync(unselectedMarker), false);
    assert.equal(fs.readFileSync(artifactPath, "utf8"), "preserve-this-artifact");
  } finally {
    fs.rmSync(fixtureDir, { recursive: true, force: true });
    removeOwnedFixtureRoot(artifactRoot);
  }
});

void test("runner rejects traversal, absolute, and non-test paths with exit 2", () => {
  const invalidPaths = [
    "../obsidian-word-reader/tests/fidelity/testRunner.test.ts",
    path.join(testRoot, "fidelity/testRunner.test.ts"),
    "tests/fidelity/fixtureBuilders.ts",
  ];

  for (const invalidPath of invalidPaths) {
    const result = runRunner([invalidPath]);
    assert.equal(
      result.status,
      2,
      `Expected exit 2 for ${invalidPath}, got ${result.status}: ${result.stderr}`,
    );
  }
});

void test("runner fails when a name pattern matches no test cases", () => {
  const result = runRunner([
    "tests/fidelity/testRunner.test.ts",
    "--test-name-pattern",
    "W7_NO_MATCHING_TEST_SHOULD_RUN",
  ]);

  assert.equal(result.status, 1);
});

void test("fixture roots reject symlink parents without touching their targets", () => {
  const ownedRoot = createOwnedFixtureRoot("fidelity-fixture-parent-check-");
  const targetDirectory = path.join(ownedRoot.directory, "target");
  const parentLink = path.join(ownedRoot.directory, "linked-parent");
  const sentinelPath = path.join(targetDirectory, "sentinel.txt");

  try {
    fs.mkdirSync(targetDirectory);
    fs.writeFileSync(sentinelPath, "preserve-target");
    const targetEntriesBefore = fs.readdirSync(targetDirectory).sort();
    fs.symlinkSync(targetDirectory, parentLink, "dir");

    assert.throws(
      () => createOwnedFixtureRoot("must-not-be-created-", parentLink),
      /Fixture root parent must not be a symbolic link\./,
    );
    assert.deepEqual(fs.readdirSync(targetDirectory).sort(), targetEntriesBefore);
    assert.equal(fs.readFileSync(sentinelPath, "utf8"), "preserve-target");
    assert.equal(fs.lstatSync(parentLink).isSymbolicLink(), true);
  } finally {
    removeOwnedFixtureRoot(ownedRoot);
  }
  assert.equal(fs.existsSync(ownedRoot.directory), false);
});

void test("runner rejects a symlinked output root before external writes", () => {
  const ownedRoot = createOwnedFixtureRoot("fidelity-output-root-fixture-");
  const projectDir = path.join(ownedRoot.directory, "project");
  const scriptsDir = path.join(projectDir, "scripts");
  const testsDir = path.join(projectDir, "tests");
  const outsideDir = path.join(ownedRoot.directory, "outside-output-target");
  const outputRootPath = path.join(projectDir, ".test-dist");
  const sentinelPath = path.join(outsideDir, "sentinel.txt");
  const outsideMarker = path.join(outsideDir, "runner-executed.test");
  let outputRootCreated = false;

  try {
    assert.ok(
      isPathWithin(
        fs.realpathSync(outputRoot),
        fs.realpathSync(ownedRoot.directory),
      ),
      "Fixture root must resolve under the repository .test-dist directory.",
    );
    fs.mkdirSync(scriptsDir, { recursive: true });
    fs.mkdirSync(testsDir);
    fs.mkdirSync(outsideDir);
    const canonicalProjectDir = fs.realpathSync(projectDir);
    const canonicalOutsideDir = fs.realpathSync(outsideDir);
    assert.ok(isPathWithin(fs.realpathSync(rootDir), canonicalOutsideDir));
    assert.equal(isPathWithin(canonicalProjectDir, canonicalOutsideDir), false);
    fs.copyFileSync(runnerPath, path.join(scriptsDir, "run-fidelity-tests.mjs"));
    fs.writeFileSync(
      path.join(testsDir, "outputRoot.test.ts"),
      "import fs from \"node:fs\";\n" +
        "import test from \"node:test\";\n" +
        `test("external execution marker", () => fs.writeFileSync(${JSON.stringify(outsideMarker)}, "unexpected"));\n`,
    );
    fs.writeFileSync(sentinelPath, "preserve-output-target-sentinel");
    const initialOutsideEntries = fs.readdirSync(outsideDir).sort();
    fs.symlinkSync(outsideDir, outputRootPath, "dir");
    outputRootCreated = true;

    const result = spawnSync(
      process.execPath,
      [path.join(scriptsDir, "run-fidelity-tests.mjs"), "tests/outputRoot.test.ts"],
      {
        cwd: projectDir,
        encoding: "utf8",
      },
    );

    assert.equal(result.status, 2, `${result.stdout}\n${result.stderr}`);
    assert.deepEqual(fs.readdirSync(outsideDir).sort(), initialOutsideEntries);
    assert.equal(
      fs.readFileSync(sentinelPath, "utf8"),
      "preserve-output-target-sentinel",
    );
    assert.equal(fs.existsSync(outsideMarker), false);
    assert.equal(fs.lstatSync(outputRootPath).isSymbolicLink(), true);
  } finally {
    if (outputRootCreated) {
      fs.unlinkSync(outputRootPath);
    }
    removeOwnedFixtureRoot(ownedRoot);
  }
});

void test("runner rejects a tests root symlink outside the project before writes", () => {
  const ownedRoot = createOwnedFixtureRoot(
    "fidelity-tests-root-symlink-fixture-",
  );
  const fixtureArea = ownedRoot.directory;
  const projectDir = path.join(fixtureArea, "project");
  const scriptsDir = path.join(projectDir, "scripts");
  const outputRootPath = path.join(projectDir, ".test-dist");
  const outsideTestsDir = path.join(fixtureArea, "outside-tests");
  const testsLink = path.join(projectDir, "tests");

  try {
    fs.mkdirSync(projectDir);
    fs.mkdirSync(scriptsDir);
    fs.mkdirSync(outputRootPath);
    fs.mkdirSync(outsideTestsDir);
    const canonicalRoot = fs.realpathSync(rootDir);
    const canonicalProject = fs.realpathSync(projectDir);
    const canonicalOutsideTests = fs.realpathSync(outsideTestsDir);
    assert.ok(isPathWithin(canonicalRoot, canonicalOutsideTests));
    assert.equal(isPathWithin(canonicalProject, canonicalOutsideTests), false);
    fs.copyFileSync(runnerPath, path.join(scriptsDir, "run-fidelity-tests.mjs"));
    fs.writeFileSync(
      path.join(outsideTestsDir, "probe.test.ts"),
      "import test from \"node:test\";\ntest(\"outside fixture probe\", () => {});\n",
    );
    fs.symlinkSync(outsideTestsDir, testsLink, "dir");

    const projectEntriesBefore = fs.readdirSync(projectDir).sort();
    const outsideEntriesBefore = fs.readdirSync(outsideTestsDir).sort();
    const result = spawnSync(
      process.execPath,
      [
        path.join(scriptsDir, "run-fidelity-tests.mjs"),
        "tests/probe.test.ts",
      ],
      {
        cwd: projectDir,
        encoding: "utf8",
      },
    );

    assert.equal(result.status, 2, `${result.stdout}\n${result.stderr}`);
    assert.deepEqual(fs.readdirSync(projectDir).sort(), projectEntriesBefore);
    assert.deepEqual(fs.readdirSync(outputRootPath), []);
    assert.deepEqual(fs.readdirSync(outsideTestsDir).sort(), outsideEntriesBefore);
    assert.equal(fs.lstatSync(testsLink).isSymbolicLink(), true);
  } finally {
    if (fs.existsSync(testsLink)) {
      fs.unlinkSync(testsLink);
    }
    removeOwnedFixtureRoot(ownedRoot);
  }
});

void test("runner maps an in-project tests symlink within its owned output directory", () => {
  const ownedRoot = createOwnedFixtureRoot(
    "fidelity-tests-root-inside-fixture-",
  );
  const fixtureArea = ownedRoot.directory;
  const projectDir = path.join(fixtureArea, "project");
  const scriptsDir = path.join(projectDir, "scripts");
  const outputRootPath = path.join(projectDir, ".test-dist");
  const sourceTestsDir = path.join(projectDir, "test-sources");
  const testsLink = path.join(projectDir, "tests");
  const markerPath = path.join(sourceTestsDir, "selected.marker");

  try {
    fs.mkdirSync(scriptsDir, { recursive: true });
    fs.mkdirSync(outputRootPath);
    fs.mkdirSync(sourceTestsDir);
    fs.copyFileSync(runnerPath, path.join(scriptsDir, "run-fidelity-tests.mjs"));
    fs.writeFileSync(
      path.join(sourceTestsDir, "probe.test.ts"),
      "import fs from \"node:fs\";\n" +
        "import test from \"node:test\";\n" +
        `test("canonical mapping probe", () => fs.writeFileSync(${JSON.stringify(markerPath)}, "ran"));\n`,
    );
    fs.symlinkSync(sourceTestsDir, testsLink, "dir");

    const result = spawnSync(
      process.execPath,
      [
        path.join(scriptsDir, "run-fidelity-tests.mjs"),
        "tests/probe.test.ts",
      ],
      {
        cwd: projectDir,
        encoding: "utf8",
      },
    );

    assert.equal(result.status, 0, `${result.stdout}\n${result.stderr}`);
    assert.equal(fs.readFileSync(markerPath, "utf8"), "ran");
    assert.deepEqual(fs.readdirSync(outputRootPath), []);
  } finally {
    if (fs.existsSync(testsLink)) {
      fs.unlinkSync(testsLink);
    }
    removeOwnedFixtureRoot(ownedRoot);
  }
});

void test("equation fixtures are deterministic synthetic DOCX and PPTX packages", async () => {
  const opaqueEmbeddingBytes = Uint8Array.from(
    Buffer.from("\0SYNTHETIC-OPAQUE-OLE-PAYLOAD\xff", "binary"),
  );
  const previewBytes = Uint8Array.from(
    Buffer.from(
      "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=",
      "base64",
    ),
  );

  for (const format of ["docx", "pptx"] as const) {
    const options: EquationFixtureOptions = {
      progId: "Equation.3",
      previewRelationship: {
        id: "rIdPreview",
        target: format === "docx" ? "media/preview.png" : "../media/preview.png",
      },
      previewBytes,
      opaqueEmbeddingBytes,
    };
    const first = await createEquationFixture(format, options);
    const second = await createEquationFixture(format, options);
    assert.deepEqual(new Uint8Array(first), new Uint8Array(second));

    const zip = await JSZip.loadAsync(first);
    const documentPath =
      format === "docx" ? "word/document.xml" : "ppt/slides/slide1.xml";
    const embeddingPath =
      format === "docx"
        ? "word/embeddings/oleObject1.bin"
        : "ppt/embeddings/oleObject1.bin";
    const previewPath =
      format === "docx" ? "word/media/preview.png" : "ppt/media/preview.png";
    const relationshipsPath =
      format === "docx"
        ? "word/_rels/document.xml.rels"
        : "ppt/slides/_rels/slide1.xml.rels";
    const xml = await zip.file(documentPath)?.async("string");
    const relationships = await zip.file(relationshipsPath)?.async("string");
    const embedding = await zip.file(embeddingPath)?.async("uint8array");
    const preview = await zip.file(previewPath)?.async("uint8array");

    assert.ok(zip.file("[Content_Types].xml"));
    assert.ok(xml);
    assert.ok(relationships);
    assert.ok(xml.includes("Equation.3"));
    assert.match(relationships, /Id="rIdPreview"/);
    assert.deepEqual(embedding, opaqueEmbeddingBytes);
    assert.deepEqual(preview, previewBytes);
    assert.notDeepEqual(embedding, preview);
    assert.doesNotMatch(xml, /SYNTHETIC-OPAQUE-OLE-PAYLOAD/);
  }
});

void test("equation fixtures expose optional alternative and safety cases", async () => {
  const options: EquationFixtureOptions = {
    progId: "Equation.DSMT4",
    previewRelationship: {
      id: "rIdPreview",
      target: "../media/preview.png",
    },
    previewBytes: Uint8Array.from([137, 80, 78, 71]),
    opaqueEmbeddingBytes: Uint8Array.from([0, 255, 1]),
    alternative: true,
    duplicate: true,
  };
  const duplicateFixture = await JSZip.loadAsync(
    await createEquationFixture("pptx", options),
  );
  const slideXml = await duplicateFixture
    .file("ppt/slides/slide1.xml")
    ?.async("string");
  assert.ok(slideXml);
  assert.equal((slideXml.match(/<mc:AlternateContent>/g) ?? []).length, 2);
  assert.equal((slideXml.match(/<p:oleObj\b/g) ?? []).length, 4);

  const missingPreview = await JSZip.loadAsync(
    await createEquationFixture("pptx", {
      ...options,
      duplicate: false,
      missingPreview: true,
    }),
  );
  const missingRelationships = await missingPreview
    .file("ppt/slides/_rels/slide1.xml.rels")
    ?.async("string");
  assert.ok(missingRelationships);
  assert.doesNotMatch(missingRelationships, /rIdPreview/);
  assert.equal(missingPreview.file("ppt/media/preview.png"), null);

  const unsafeTarget = await JSZip.loadAsync(
    await createEquationFixture("docx", {
      ...options,
      duplicate: false,
      alternative: false,
      previewRelationship: {
        id: "rIdPreview",
        target: "media/preview.png",
      },
      unsafeTarget: true,
    }),
  );
  const unsafeRelationships = await unsafeTarget
    .file("word/_rels/document.xml.rels")
    ?.async("string");
  assert.ok(unsafeRelationships);
  assert.match(unsafeRelationships, /\.\.\/\.\.\/unsafe-preview\.png/);
});
