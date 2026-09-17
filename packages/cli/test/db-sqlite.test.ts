import { afterEach } from "vitest";

import { defineFormatTests } from "./db-format-suite.js";
import { cleanupDirectories } from "./db-support.js";

afterEach(cleanupDirectories);

defineFormatTests("sqlite");
