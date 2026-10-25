import { spawnSync } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { build } from "esbuild";

const rootDir = fs.realpathSync(
  path.resolve(path.dirname(fileURLToPath(import.meta.url)), ".."),
);
const testsDir = path.join(rootDir, "tests");
const outputRoot = path.join(rootDir, ".test-dist");
const fidelityReporterSource = `import { Transform } from "node:stream";
export default class FidelityTestSummaryReporter extends Transform {
  constructor() {
    super({ writableObjectMode: true });
  }

  _transform(event, _encoding, callback) {
    if (event.type === "test:summary" && event.data.file) {
      const counts = event.data.counts;
      const executed = counts.passed + counts.failed + counts.todo;
      this.push(JSON.stringify({ file: event.data.file, executed }) + "\\n");
    }
    callback();
  }
}
`;

class InvalidTestPathError extends Error {}
class InvalidOutputRootError extends Error {}

function isPathInside(directory, candidate) {
  const relativePath = path.relative(directory, candidate);
  return (
    relativePath === "" ||
    (relativePath !== ".." &&
      !relativePath.startsWith(`..${path.sep}`) &&
      !path.isAbsolute(relativePath))
  );
}

function parseArguments(args) {
  const testPaths = [];
  let testNamePattern;
  let hasPattern = false;

  for (let index = 0; index < args.length; index += 1) {
    const argument = args[index];
    if (argument === "--test-name-pattern") {
      if (hasPattern || index + 1 >= args.length) {
        return { error: "--test-name-pattern requires one REGEX value." };
      }
      testNamePattern = args[index + 1];
      hasPattern = true;
      index += 1;
    } else {
      testPaths.push(argument);
    }
  }

  if (testPaths.length === 0) {
    return {
      error:
        "Usage: node scripts/run-fidelity-tests.mjs tests/path.test.ts [tests/other.test.ts ...] [--test-name-pattern REGEX]",
    };
  }

  return { testPaths, testNamePattern };
}

function resolveTestFiles(testPaths) {
  let realTestsDir;
  try {
    realTestsDir = fs.realpathSync(testsDir);
  } catch (error) {
    throw new InvalidTestPathError(`Unable to resolve tests/: ${error.message}`);
  }
  if (!isPathInside(rootDir, realTestsDir)) {
    throw new InvalidTestPathError("tests/ must resolve inside the project.");
  }

  const testFiles = testPaths.map((testPath) => {
    const segments = testPath.split(/[\\/]+/);
    if (
      path.isAbsolute(testPath) ||
      path.win32.isAbsolute(testPath) ||
      testPath.includes("\\") ||
      segments.includes("..") ||
      segments[0] !== "tests" ||
      !testPath.endsWith(".test.ts")
    ) {
      throw new InvalidTestPathError(
        `Not an allowed test file path: ${testPath}`,
      );
    }

    const absolutePath = path.resolve(rootDir, testPath);
    if (!fs.existsSync(absolutePath)) {
      throw new InvalidTestPathError(
        `Not an allowed test file path: ${testPath}`,
      );
    }
    const realPath = fs.realpathSync(absolutePath);
    if (
      !isPathInside(realTestsDir, realPath) ||
      !fs.statSync(realPath).isFile() ||
      !realPath.endsWith(".test.ts")
    ) {
      throw new InvalidTestPathError(
        `Test file must be inside tests/: ${testPath}`,
      );
    }

    return realPath;
  });
  return { testFiles, testsDir: realTestsDir };
}

function resolveOwnedOutputRoot() {
  let outputRootStats;
  try {
    outputRootStats = fs.lstatSync(outputRoot);
  } catch (error) {
    if (!(error instanceof Error) || error.code !== "ENOENT") {
      throw error;
    }
  }

  if (outputRootStats?.isSymbolicLink()) {
    throw new InvalidOutputRootError(".test-dist must not be a symbolic link.");
  }
  if (outputRootStats && !outputRootStats.isDirectory()) {
    throw new InvalidOutputRootError(".test-dist must be a directory.");
  }
  if (!outputRootStats) {
    fs.mkdirSync(outputRoot);
  }

  const currentStats = fs.lstatSync(outputRoot);
  if (currentStats.isSymbolicLink() || !currentStats.isDirectory()) {
    throw new InvalidOutputRootError(".test-dist must be a real directory.");
  }

  const realProjectRoot = fs.realpathSync(rootDir);
  const realOutputRoot = fs.realpathSync(outputRoot);
  const relativeOutputRoot = path.relative(realProjectRoot, realOutputRoot);
  if (
    relativeOutputRoot === "" ||
    relativeOutputRoot === ".." ||
    relativeOutputRoot.startsWith(`..${path.sep}`) ||
    path.isAbsolute(relativeOutputRoot)
  ) {
    throw new InvalidOutputRootError(
      ".test-dist must resolve to a directory inside the project.",
    );
  }
  return realOutputRoot;
}

async function main() {
  const parsed = parseArguments(process.argv.slice(2));
  if (parsed.error) {
    console.error(parsed.error);
    process.exitCode = 2;
    return;
  }

  let ownedOutputDir;
  try {
    const { testFiles, testsDir: canonicalTestsDir } = resolveTestFiles(
      parsed.testPaths,
    );
    const safeOutputRoot = resolveOwnedOutputRoot();
    ownedOutputDir = fs.mkdtempSync(path.join(safeOutputRoot, "fidelity-"));

    const result = await build({
      entryPoints: testFiles,
      outdir: ownedOutputDir,
      outbase: canonicalTestsDir,
      bundle: true,
      packages: "external",
      platform: "node",
      format: "cjs",
      target: "node18",
      sourcemap: "inline",
      logLevel: "warning",
    });

    if (result.errors.length > 0) {
      process.exitCode = 1;
      return;
    }

    const compiledTests = testFiles.map((file) => {
      const relativePath = path.relative(canonicalTestsDir, file);
      return path.resolve(
        ownedOutputDir,
        relativePath.replace(/\.ts$/, ".js"),
      );
    });
    const reporterPath = path.join(ownedOutputDir, "fidelity-test-reporter.mjs");
    const reporterOutputPath = path.join(
      ownedOutputDir,
      "fidelity-test-summaries.ndjson",
    );
    fs.writeFileSync(reporterPath, fidelityReporterSource, "utf8");

    const nodeArgs = [
      "--test",
      "--test-reporter=spec",
      "--test-reporter-destination=stdout",
      `--test-reporter=${reporterPath}`,
      `--test-reporter-destination=${reporterOutputPath}`,
    ];
    if (parsed.testNamePattern !== undefined) {
      nodeArgs.push("--test-name-pattern", parsed.testNamePattern);
    }
    nodeArgs.push(...compiledTests);

    const testRunEnv = { ...process.env };
    delete testRunEnv.NODE_TEST_CONTEXT;
    const testRun = spawnSync(process.execPath, nodeArgs, {
      cwd: rootDir,
      env: testRunEnv,
      stdio: "inherit",
    });
    if (testRun.error) {
      console.error(testRun.error);
      process.exitCode = 1;
    } else if (testRun.status !== 0) {
      process.exitCode = testRun.status ?? 1;
    } else {
      const summaries = fs
        .readFileSync(reporterOutputPath, "utf8")
        .split(/\r?\n/)
        .filter(Boolean)
        .map((line) => JSON.parse(line));
      if (
        summaries.length !== testFiles.length ||
        summaries.some(
          (summary) => !Number.isSafeInteger(summary.executed) || summary.executed < 0,
        )
      ) {
        console.error("Unable to verify executed test cases from Node's reporter.");
        process.exitCode = 1;
      } else if (
        summaries.reduce((total, summary) => total + summary.executed, 0) === 0
      ) {
        console.error("No test cases executed; the name pattern may match none.");
        process.exitCode = 1;
      } else {
        process.exitCode = 0;
      }
    }
  } catch (error) {
    if (
      error instanceof InvalidTestPathError ||
      error instanceof InvalidOutputRootError
    ) {
      console.error(error.message);
      process.exitCode = 2;
    } else {
      console.error(error);
      process.exitCode = 1;
    }
  } finally {
    if (ownedOutputDir) {
      fs.rmSync(ownedOutputDir, { recursive: true, force: true });
    }
  }
}

await main();
