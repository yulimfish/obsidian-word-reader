import { spawnSync } from "node:child_process";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { build } from "esbuild";

const rootDir = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const outputDir = path.join(rootDir, ".lifecycle-dist");
const outputPath = path.join(outputDir, "lifecycle-stress.cjs");
const resultDir = path.join(rootDir, "benchmark-results");
const resultPath = path.join(resultDir, "lifecycle-stress.latest.json");
const trendPath = path.join(resultDir, "lifecycle-stress-trend.json");
const summaryPath = path.join(resultDir, "lifecycle-stress-summary.md");
const budgetPath = path.join(rootDir, "benchmarks", "lifecycle-stress-budget.json");
const sequenceNames = ["docx", "pptx", "xlsx", "mixed"];
const resourceNames = [
  "blobUrls",
  "caches",
  "tasks",
  "timers",
  "observers",
  "listeners",
];
const staleTaskNames = [
  "parse",
  "search",
  "thumbnail",
  "worksheet",
  "drawing",
];

fs.rmSync(outputDir, { recursive: true, force: true });
fs.mkdirSync(outputDir, { recursive: true });

try {
  await build({
    entryPoints: [
      path.join(rootDir, "scripts", "lifecycle-stress-runner.ts"),
    ],
    outfile: outputPath,
    bundle: true,
    packages: "external",
    platform: "node",
    format: "cjs",
    target: "node18",
    logLevel: "warning",
  });

  process.stdout.write("Running 4 x 100 lifecycle stress cycles...\n");
  const run = spawnSync(process.execPath, ["--expose-gc", outputPath], {
    cwd: rootDir,
    encoding: "utf8",
    maxBuffer: 32 * 1024 * 1024,
  });
  if (run.status !== 0) {
    process.stderr.write(run.stderr);
    process.stdout.write(run.stdout);
    process.exit(run.status ?? 1);
  }
  const resultLine = run.stdout
    .split(/\r?\n/)
    .find((line) => line.startsWith("OFFICE_READER_LIFECYCLE_JSON="));
  if (!resultLine) {
    throw new Error("Lifecycle stress runner did not emit a result.");
  }

  const result = JSON.parse(
    resultLine.slice("OFFICE_READER_LIFECYCLE_JSON=".length),
  );
  result.ci = {
    commit: process.env.GITHUB_SHA ?? null,
    runId: process.env.GITHUB_RUN_ID ?? null,
    runAttempt: process.env.GITHUB_RUN_ATTEMPT ?? null,
  };
  result.runtime.cpuCount = os.cpus().length;
  const budget = JSON.parse(fs.readFileSync(budgetPath, "utf8"));
  const failures = validate(result, budget);
  result.passed = failures.length === 0;
  result.failures = failures;

  const trend = createTrendPoint(result);
  const summary = createSummary(result, budget);
  fs.mkdirSync(resultDir, { recursive: true });
  fs.writeFileSync(resultPath, `${JSON.stringify(result, null, 2)}\n`);
  fs.writeFileSync(trendPath, `${JSON.stringify(trend, null, 2)}\n`);
  fs.writeFileSync(summaryPath, summary);
  process.stdout.write(summary);

  if (failures.length > 0) {
    process.stderr.write("Lifecycle stress budgets failed:\n");
    failures.forEach((failure) => process.stderr.write(`- ${failure}\n`));
    process.exitCode = 1;
  }
} finally {
  fs.rmSync(outputDir, { recursive: true, force: true });
}

function validate(result, budget) {
  const failures = [];
  if (result.schemaVersion !== budget.schemaVersion) {
    failures.push(
      `schemaVersion=${result.schemaVersion} does not equal ${budget.schemaVersion}`,
    );
  }
  if (result.configuration.cyclesPerSequence < budget.minimumCyclesPerSequence) {
    failures.push("configured lifecycle cycles are below the required minimum");
  }
  if (
    result.configuration.errorIterationsPerCase <
      budget.minimumErrorIterationsPerCase
  ) {
    failures.push("configured error iterations are below the required minimum");
  }

  for (const name of sequenceNames) {
    const sequence = result.sequences.find((entry) => entry.name === name);
    if (!sequence) {
      failures.push(`sequence ${name} is missing`);
      continue;
    }
    if (sequence.cycles < budget.minimumCyclesPerSequence) {
      failures.push(`${name}.cycles=${sequence.cycles} is below the minimum`);
    }
    for (const operation of [
      "open",
      "reload",
      "quickNavigation",
      "fileSwitch",
      "close",
    ]) {
      if (sequence.operations[operation] < sequence.cycles) {
        failures.push(
          `${name}.operations.${operation}=${sequence.operations[operation]} ` +
            `is below ${sequence.cycles}`,
        );
      }
    }
    for (const resource of resourceNames) {
      if (sequence.resourcePeak[resource] < 1) {
        failures.push(`${name}.resourcePeak.${resource} was not exercised`);
      }
      if (sequence.resourceResidue[resource] !== 0) {
        failures.push(
          `${name}.resourceResidue.${resource}=` +
            `${sequence.resourceResidue[resource]} is not zero`,
        );
      }
    }
    if (sequence.formatCacheResidue !== 0) {
      failures.push(`${name}.formatCacheResidue is not zero`);
    }
    if (sequence.formatCachePeak < 1) {
      failures.push(`${name}.formatCachePeak was not exercised`);
    }
    for (const task of staleTaskNames) {
      if (sequence.staleUiCommits[task] !== 0) {
        failures.push(`${name}.staleUiCommits.${task} is not zero`);
      }
    }
    if (
      sequence.handledTaskErrors <
        sequence.cycles * budget.minimumHandledTaskErrorsPerCycle
    ) {
      failures.push(`${name}.handledTaskErrors did not cover every close cycle`);
    }
    if (
      sequence.revisionInvalidations <
        sequence.cycles * budget.minimumRevisionInvalidationsPerCycle
    ) {
      failures.push(`${name}.revisionInvalidations did not cover reload and switch`);
    }
    if (sequence.navigationFailures !== 0) {
      failures.push(`${name}.navigationFailures is not zero`);
    }
    if (name === "mixed") {
      if (
        JSON.stringify(sequence.formatOrder) !==
          JSON.stringify(["docx", "pptx", "xlsx"]) ||
        Object.values(sequence.formatCycles).some((cycles) => cycles < 1)
      ) {
        failures.push("mixed format order did not cycle DOCX, PPTX, then XLSX");
      }
    } else if (
      sequence.formatOrder.length !== 1 ||
      sequence.formatOrder[0] !== name ||
      sequence.formatCycles[name] !== sequence.cycles
    ) {
      failures.push(`${name} sequence did not exclusively exercise ${name}`);
    }
    if (!sequence.heap.gcAvailable) {
      failures.push(`${name}.heap.gcAvailable is false`);
    }
    if (
      sequence.heap.retainedGrowthMiB >
        budget.maximumRetainedHeapGrowthMiB
    ) {
      failures.push(
        `${name}.heap.retainedGrowthMiB=${sequence.heap.retainedGrowthMiB} ` +
          `exceeds ${budget.maximumRetainedHeapGrowthMiB}`,
      );
    }
    if (
      sequence.heap.maxConsecutiveIncreases >=
        budget.sustainedHeapGrowthWindow
    ) {
      failures.push(
        `${name}.heap.maxConsecutiveIncreases=` +
          `${sequence.heap.maxConsecutiveIncreases} reaches the sustained ` +
          `growth window ${budget.sustainedHeapGrowthWindow}`,
      );
    }
  }

  if (result.sequences.length !== sequenceNames.length) {
    failures.push(`expected ${sequenceNames.length} lifecycle sequences`);
  }
  const expectedErrorAttempts = 3 * 4 *
    result.errorMatrix.iterationsPerCase;
  if (result.errorMatrix.attempts !== expectedErrorAttempts) {
    failures.push(
      `errorMatrix.attempts=${result.errorMatrix.attempts} does not equal ` +
        expectedErrorAttempts,
    );
  }
  if (result.errorMatrix.recoverableErrors !== expectedErrorAttempts) {
    failures.push("not every damaged/encrypted/over-limit/unsupported file recovered");
  }
  if (result.errorMatrix.unexpectedSuccesses !== 0) {
    failures.push("the error matrix contains unexpected successes");
  }
  for (const resource of resourceNames) {
    if (result.errorMatrix.resourceResidue[resource] !== 0) {
      failures.push(`errorMatrix.resourceResidue.${resource} is not zero`);
    }
  }
  if (result.unhandledRejections !== 0) {
    failures.push(`unhandledRejections=${result.unhandledRejections} is not zero`);
  }
  return failures;
}

function createTrendPoint(result) {
  return {
    product: result.product,
    schemaVersion: result.schemaVersion,
    generatedAt: result.generatedAt,
    ci: result.ci,
    runtime: result.runtime,
    passed: result.passed,
    failures: result.failures,
    sequences: Object.fromEntries(result.sequences.map((sequence) => [
      sequence.name,
      {
        cycles: sequence.cycles,
        durationMs: sequence.durationMs,
        resourceResidue: sequence.resourceResidue,
        formatCachePeak: sequence.formatCachePeak,
        formatCacheResidue: sequence.formatCacheResidue,
        staleUiCommits: sequence.staleUiCommits,
        heap: sequence.heap,
      },
    ])),
    errorMatrix: result.errorMatrix,
    unhandledRejections: result.unhandledRejections,
  };
}

function createSummary(result, budget) {
  const lines = [
    "# Office Reader lifecycle stress",
    "",
    `Status: ${result.passed ? "passed" : "failed"}`,
    "",
    `Cycles per sequence: ${result.configuration.cyclesPerSequence}`,
    "",
    "| Sequence | Cycles | Duration | Heap growth | Monotonic run | Resource residue | Stale UI commits |",
    "| --- | ---: | ---: | ---: | ---: | ---: | ---: |",
  ];
  for (const sequence of result.sequences) {
    lines.push(
      `| ${sequence.name.toUpperCase()} | ${sequence.cycles} | ` +
        `${sequence.durationMs} ms | ${sequence.heap.retainedGrowthMiB} MiB | ` +
        `${sequence.heap.maxConsecutiveIncreases} / ` +
        `${budget.sustainedHeapGrowthWindow} | ` +
        `${sumValues(sequence.resourceResidue) + sequence.formatCacheResidue} | ` +
        `${sumValues(sequence.staleUiCommits)} |`,
    );
  }
  lines.push(
    "",
    `Recoverable error attempts: ${result.errorMatrix.recoverableErrors} / ` +
      `${result.errorMatrix.attempts}`,
    "",
    `Unhandled Promise rejections: ${result.unhandledRejections}`,
    "",
  );
  if (result.failures.length > 0) {
    lines.push(
      "## Failures",
      "",
      ...result.failures.map((failure) => `- ${failure}`),
      "",
    );
  }
  return `${lines.join("\n")}\n`;
}

function sumValues(record) {
  return Object.values(record).reduce((total, value) => total + value, 0);
}
