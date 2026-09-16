import { expect, test } from "@playwright/test";

const SCHEMA = {
  version: 1,
  tables: [
    {
      name: "datasets",
      recordId: { prefix: "DATASET", padding: 6 },
      columns: [{ name: "dataset_name", type: "text", nullable: false }],
    },
  ],
};

/**
 * Wrap the page's worker so one schemaApplied reply arrives with a summary
 * that could not be refreshed, the way the worker reports it when its own
 * inspection fails after a committed write.
 */
async function installStaleSummaryWorker(
  page: import("@playwright/test").Page,
): Promise<void> {
  await page.addInitScript(() => {
    const NativeWorker = window.Worker;
    class StaleSummaryWorker extends EventTarget {
      readonly worker: Worker;
      stale = true;

      constructor(url: string | URL, options?: WorkerOptions) {
        super();
        this.worker = new NativeWorker(url, options);
        this.worker.addEventListener("message", (event) => {
          let message: unknown = event.data;
          if (
            this.stale &&
            typeof message === "object" &&
            message !== null &&
            "type" in message &&
            message.type === "schemaApplied" &&
            "result" in message &&
            typeof message.result === "object" &&
            message.result !== null
          ) {
            this.stale = false;
            message = {
              ...message,
              result: {
                ...message.result,
                summary: {
                  state: "refresh-required",
                  code: "DB_BROWSER_SUMMARY_REFRESH_REQUIRED",
                  message:
                    "The database operation completed, but the browser could not refresh its summary. Keep this tab open and choose Refresh summary before continuing",
                },
              },
            };
          }
          this.dispatchEvent(new MessageEvent("message", { data: message }));
        });
        this.worker.addEventListener("error", () => {
          this.dispatchEvent(new Event("error"));
        });
      }

      postMessage(
        message: unknown,
        options?: StructuredSerializeOptions | Transferable[],
      ): void {
        if (Array.isArray(options)) {
          this.worker.postMessage(message, options);
        } else {
          this.worker.postMessage(message, options);
        }
      }

      terminate(): void {
        this.worker.terminate();
      }
    }
    Object.defineProperty(window, "Worker", {
      configurable: true,
      value: StaleSummaryWorker,
    });
  });
}

test.describe("database tool feedback", () => {
  test("loads batch history on its own and reports outcomes in place", async ({
    page,
  }) => {
    await page.goto("/tools/db");
    await page.getByTestId("workspace-new-name").fill("feedback.sqlite");
    await page.getByTestId("workspace-new").click();
    await expect(page.getByTestId("workspace-summary")).toBeVisible();

    // The creation notice sits in the start section, not at the page end.
    const notice = page.getByTestId("workspace-notice");
    await expect(notice).toHaveAttribute("data-section", "start");
    await expect(
      page.getByTestId("workspace-start").getByTestId("workspace-notice"),
    ).toContainText("Created a persistent");

    // History loaded without a click on Refresh.
    await expect(page.getByTestId("workspace-deliveries")).toContainText(
      "No import batches recorded",
    );

    // A schema error renders inside the schema section with the error style.
    await page.getByTestId("workspace-schema-input").fill("{ not json");
    await page.getByTestId("workspace-schema-plan").click();
    const error = page
      .getByTestId("workspace-schema")
      .getByTestId("workspace-error");
    await expect(error).toBeVisible();
    await expect(error).toHaveAttribute("data-section", "schema");
    await expect(error).toContainText("Something went wrong");
    // The error style is the heavier border; a notice has the thin one.
    await expect(error).toHaveClass(/border-2/);
    await expect(page.getByTestId("workspace-notice")).toHaveCount(0);
  });

  test("offers Refresh summary when a write completed but the summary is stale", async ({
    page,
  }) => {
    await installStaleSummaryWorker(page);
    await page.goto("/tools/db");
    await page.getByTestId("workspace-new-name").fill("stale.sqlite");
    await page.getByTestId("workspace-new").click();
    await expect(page.getByTestId("workspace-summary")).toBeVisible();
    await expect(page.getByTestId("workspace-table-count")).toHaveText("0");

    await page
      .getByTestId("workspace-schema-input")
      .fill(JSON.stringify(SCHEMA));
    await page.getByTestId("workspace-schema-plan").click();
    await expect(page.getByTestId("workspace-schema-apply")).toBeEnabled();
    await page.getByTestId("workspace-schema-apply").click();

    const error = page
      .getByTestId("workspace-schema")
      .getByTestId("workspace-error");
    await expect(error).toContainText("DB_BROWSER_SUMMARY_REFRESH_REQUIRED");
    // The stale summary is still on screen until the reader refreshes it.
    await expect(page.getByTestId("workspace-table-count")).toHaveText("0");
    await page.getByTestId("workspace-status-action").click();
    await expect(page.getByTestId("workspace-table-count")).toHaveText("1");
    await expect(page.getByTestId("workspace-notice")).toContainText(
      "Refreshed the database summary",
    );
    await expect(page.getByTestId("workspace-error")).toHaveCount(0);
  });
});
