import { expect, test, type Page } from "@playwright/test";

import { createWorkbookUpload } from "./fixtures";

async function failFirstSavedPlanListing(page: Page): Promise<void> {
  await page.addInitScript(() => {
    const NativeWorker = window.Worker;
    class ListingFailureWorker extends EventTarget {
      readonly worker: Worker;
      failedListing = false;

      constructor(url: string | URL, options?: WorkerOptions) {
        super();
        this.worker = new NativeWorker(url, options);
        this.worker.addEventListener("message", (event) => {
          const message: unknown = event.data;
          if (
            !this.failedListing &&
            typeof message === "object" &&
            message !== null &&
            Reflect.get(message, "type") === "importsListed"
          ) {
            this.failedListing = true;
            this.dispatchEvent(
              new MessageEvent("message", {
                data: {
                  id: Reflect.get(message, "id"),
                  type: "error",
                  code: "DB_BROWSER_IMPORT_CLEANUP_REQUIRED",
                  message:
                    "Saved import plans could not finish releasing their private resources. Choose Retry saved imports to finish cleanup before reopening those plans.",
                },
              }),
            );
            return;
          }
          this.dispatchEvent(new MessageEvent("message", { data: message }));
        });
        this.worker.addEventListener("error", () => {
          this.dispatchEvent(new Event("error"));
        });
        this.worker.addEventListener("messageerror", () => {
          this.dispatchEvent(new MessageEvent("messageerror"));
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
      value: ListingFailureWorker,
    });
  });
}

test("retries a failed saved import listing from the database page", async ({
  page,
}) => {
  const workbook = await createWorkbookUpload("saved-review.xlsx", [
    {
      name: "Inventory",
      rows: [["Region"], ["North"]],
    },
  ]);
  await page.goto("/tools/db");
  await page.getByTestId("workspace-new-format").selectOption("sqlite");
  await page.getByTestId("workspace-new-name").fill("saved-list-retry.sqlite");
  await page.getByTestId("workspace-new").click();
  await expect(page.getByTestId("workspace-summary")).toBeVisible();

  await page.getByTestId("workspace-import-input").setInputFiles(workbook);
  await page
    .getByTestId("workspace-import-source")
    .getByTestId("workspace-import-role")
    .fill("inventory");
  await page
    .getByTestId("workspace-import-source")
    .getByTestId("workspace-import-revision")
    .fill("Initial");
  await page.getByTestId("workspace-import-prepare").click();
  await expect(page.getByTestId("workspace-import-review")).toBeVisible();

  await failFirstSavedPlanListing(page);
  await page.reload();
  await page.getByTestId("workspace-reopen").first().click();
  await expect(page.getByTestId("workspace-import-retry-saved")).toBeVisible();
  await expect(page.getByTestId("workspace-error")).toContainText(
    "Choose Retry saved imports",
  );

  await page.getByTestId("workspace-import-retry-saved").click();
  await expect(page.getByTestId("workspace-import-retry-saved")).toHaveCount(0);
  await expect(page.getByTestId("workspace-error")).toHaveCount(0);
  await expect(page.getByTestId("workspace-import-resume")).toHaveCount(1);
  await page.getByTestId("workspace-import-resume").click();
  await expect(page.getByTestId("workspace-import-review")).toBeVisible();
});
