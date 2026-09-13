import { createElement } from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";

import { WorkspaceImportDiagnostics } from "./workspace-import-diagnostics";

describe("WorkspaceImportDiagnostics", () => {
  it("names an unsupported saved plan and explains how to replace it", () => {
    const legacyVersion = 2;
    const supportedVersion = legacyVersion + 1;
    const rendered = renderToStaticMarkup(
      createElement(WorkspaceImportDiagnostics, {
        ignoredPlans: [
          {
            name: ".consultchimps-import-old.sqlite",
            code: "DB_UNSUPPORTED_PREPARED_IMPORT_VERSION",
            message: `This import plan uses format version ${String(legacyVersion)}, but this build supports version ${String(supportedVersion)}. Regenerate the plan from its original sources with this build.`,
          },
        ],
      }),
    );

    expect(rendered).toContain(".consultchimps-import-old.sqlite");
    expect(rendered).toContain(`format version ${String(legacyVersion)}`);
    expect(rendered).toContain(`supports version ${String(supportedVersion)}`);
    expect(rendered).toContain("Regenerate the plan from its original sources");
    expect(rendered).toContain("DB_UNSUPPORTED_PREPARED_IMPORT_VERSION");
  });
});
