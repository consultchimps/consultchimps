import { existsSync, readFileSync, writeFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { format, resolveConfig } from "prettier";

import {
  collectProjectionProblems,
  columnOperations,
  GENERATED_BLOCK_END,
  GENERATED_BLOCK_START,
  renderPreservationMatrixBlock,
  undeclaredOperations,
} from "./preservation-matrix.ts";

// Contract <-> site conformance check for the Excel preservation matrix. The
// published table is generated from packages/xlsx/src/contract.ts rather than
// written by hand, so this check does two things: it fails when the projection
// in scripts/preservation-matrix.ts has no words for something the contract
// now declares, and it fails when the block committed to the page differs from
// what the contract renders today. Run it with --write to regenerate the page
// after changing the contract.

const workspaceRoot = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "..",
);
const matrixPagePath = path.join(
  workspaceRoot,
  "apps",
  "docs",
  "content",
  "docs",
  "tools",
  "excel-preservation.mdx",
);
const matrixPageLabel = "apps/docs/content/docs/tools/excel-preservation.mdx";
const matrixPageHref = "/docs/tools/excel-preservation";

// Every Excel operation guide states its own limitations, and each one must
// send the reader to the full matrix instead of growing a second, divergent
// list of what survives an operation.
const guidesLinkingToMatrix: readonly string[] = [
  "apps/docs/content/docs/tools/spreadsheet-split.mdx",
  "apps/docs/content/docs/tools/workbook-merge.mdx",
  "apps/docs/content/docs/tools/spreadsheet-consolidate.mdx",
];

const shouldWrite = process.argv.includes("--write");

function readTextFile(absolutePath: string): string {
  return readFileSync(absolutePath, "utf8").replace(/\r\n/g, "\n");
}

const problems: string[] = [];

// Check 1: the projection covers the contract. A structure, operation, or
// behavior with no human label, note, or status is reported here rather than
// rendered as a blank cell nobody notices.
problems.push(...collectProjectionProblems());

if (problems.length > 0) {
  throw new Error(
    `The Excel preservation matrix cannot describe the contract in packages/xlsx/src/contract.ts:\n${problems
      .map((problem) => `- ${problem}`)
      .join(
        "\n",
      )}\nAdd the missing label, note, or status to scripts/preservation-matrix.ts, then run \`node scripts/check-preservation-matrix.ts --write\`.`,
  );
}

if (!existsSync(matrixPagePath)) {
  throw new Error(
    `The Excel preservation page was not found at ${matrixPagePath}.`,
  );
}

// Prettier formats the generated block exactly as it formats the page it is
// spliced into, so `pnpm format:check` and this check cannot disagree about
// table padding or prose wrapping.
const prettierOptions = await resolveConfig(matrixPagePath);
const renderedBlock = await format(renderPreservationMatrixBlock(), {
  ...prettierOptions,
  filepath: matrixPagePath,
});

const pageText = readTextFile(matrixPagePath);
const blockStart = pageText.indexOf(GENERATED_BLOCK_START);
const blockEnd = pageText.indexOf(GENERATED_BLOCK_END);

if (blockStart === -1 || blockEnd === -1 || blockEnd < blockStart) {
  throw new Error(
    `${matrixPageLabel} must contain the generated matrix between ${GENERATED_BLOCK_START} and ${GENERATED_BLOCK_END}. Restore both markers, then run \`node scripts/check-preservation-matrix.ts --write\`.`,
  );
}

const currentBlock = pageText.slice(
  blockStart,
  blockEnd + GENERATED_BLOCK_END.length,
);
const expectedBlock = renderedBlock.trim();

// Check 2: the committed block is what the contract renders today.
if (currentBlock !== expectedBlock) {
  if (shouldWrite) {
    writeFileSync(
      matrixPagePath,
      pageText.slice(0, blockStart) +
        expectedBlock +
        pageText.slice(blockEnd + GENERATED_BLOCK_END.length),
      "utf8",
    );
    process.stdout.write(
      `Regenerated the preservation matrix in ${matrixPageLabel} from packages/xlsx/src/contract.ts.\n`,
    );
  } else {
    const currentLines = currentBlock.split("\n");
    const expectedLines = expectedBlock.split("\n");
    const firstDifference = expectedLines.findIndex(
      (line, index) => currentLines[index] !== line,
    );
    const difference =
      firstDifference === -1
        ? `the committed block has ${currentLines.length - expectedLines.length} extra line(s) at the end`
        : `first difference at line ${firstDifference + 1} of the block:\n  page:     ${
            currentLines[firstDifference] ?? "(missing)"
          }\n  contract: ${expectedLines[firstDifference]}`;
    throw new Error(
      `${matrixPageLabel} no longer matches the contract in packages/xlsx/src/contract.ts (${difference}).\nRun \`node scripts/check-preservation-matrix.ts --write\` to regenerate the page, and review the change: it means the promises the package makes about workbook structures have moved.`,
    );
  }
}

// Check 3: every Excel guide sends the reader to the matrix rather than
// keeping its own list of what an operation preserves.
const missingLinks = guidesLinkingToMatrix.filter((guideLabel) => {
  const guidePath = path.join(workspaceRoot, ...guideLabel.split("/"));
  if (!existsSync(guidePath)) {
    problems.push(`${guideLabel} does not exist`);
    return false;
  }
  return !readTextFile(guidePath).includes(matrixPageHref);
});

for (const guideLabel of missingLinks) {
  problems.push(
    `${guideLabel} does not link to ${matrixPageHref}; every Excel guide's limitations section must point at the full matrix`,
  );
}

if (problems.length > 0) {
  throw new Error(
    `The Excel preservation matrix is not reachable from every Excel guide:\n${problems
      .map((problem) => `- ${problem}`)
      .join("\n")}`,
  );
}

process.stdout.write(
  `Verified the preservation matrix in ${matrixPageLabel} against packages/xlsx/src/contract.ts: ${columnOperations().length} operation columns, ${undeclaredOperations().length} operations explained without one, and ${guidesLinkingToMatrix.length} guides linking to it.\n`,
);
