// Fresh-process arm of the determinism check: exports the committed fixture
// through the BUILT package and prints the two artifact digests. Run by
// export.test.ts as a subprocess, so nothing of this process's state can
// influence the bytes.
import { createHash } from "node:crypto";
import { readFileSync } from "node:fs";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";
import { exportPbiTables } from "../dist/index.js";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const bytes = new Uint8Array(
  readFileSync(path.join(root, "fixtures", "a-2018-fuzzy.pbix")),
);
const outcome = await exportPbiTables(bytes);
process.stdout.write(
  outcome.outputs
    .map((output) => createHash("sha256").update(output.bytes).digest("hex"))
    .join(" "),
);
