import { createElement } from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";

import { WorkspaceImportDiagnostics } from "./workspace-import-diagnostics";

describe("WorkspaceImportDiagnostics", () => {
  it("names an unsupported saved plan and explains how to replace it", () => {
    const rendered = renderToStaticMarkup(
      createElement(WorkspaceImportDiagnostics, {
        ignoredPlans: [
          {
            name: ".consultchimps-import-old.sqlite",
            code: "DB_UNSUPPORTED_PREPARED_IMPORT_VERSION",
            message:
              "This import plan uses format version 1, but this build supports version 2. Regenerate the plan from its original sources with this build.",
          },
        ],
      }),
    );

    expect(rendered).toContain(".consultchimps-import-old.sqlite");
    expect(rendered).toContain("format version 1");
    expect(rendered).toContain("supports version 2");
    expect(rendered).toContain("Regenerate the plan from its original sources");
    expect(rendered).toContain("DB_UNSUPPORTED_PREPARED_IMPORT_VERSION");
  });
});
