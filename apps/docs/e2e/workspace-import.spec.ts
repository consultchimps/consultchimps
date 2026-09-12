import { expect, test, type Page } from "@playwright/test";

import { createWorkbookUpload, type UploadFile } from "./fixtures";

function inventoryWorkbook(
  name = "inventory.xlsx",
  extraRows: ReadonlyArray<readonly (number | string)[]> = [],
): Promise<UploadFile> {
  return createWorkbookUpload(name, [
    {
      name: "Inventory",
      rows: [
        ["Dataset", "Attribute", "CDE"],
        ["Customers", "Customer ID", "true"],
        ["Orders", "Order ID", "false"],
        ...extraRows,
      ],
    },
  ]);
}

function mappingWorkbook(): Promise<UploadFile> {
  return createWorkbookUpload("mappings.xlsx", [
    {
      name: "Mappings",
      rows: [
        ["Attribute", "CDE name"],
        ["Customer ID", "Customer identifier"],
      ],
    },
  ]);
}

function generatedRows(count: number): ReadonlyArray<readonly string[]> {
  return Array.from({ length: count }, (_, index) => [
    `Dataset ${String(index)}`,
    `Attribute ${String(index)}`,
    index % 2 === 0 ? "true" : "false",
  ]);
}

async function create(page: Page, format: "duckdb" | "sqlite"): Promise<void> {
  await page.goto("/workspace");
  await page.getByTestId("workspace-new-format").selectOption(format);
  await page.getByTestId("workspace-new-name").fill(`imports.${format}`);
  await page.getByTestId("workspace-new").click();
  await expect(page.getByTestId("workspace-summary")).toBeVisible();
}

async function prepare(
  page: Page,
  files: readonly UploadFile[],
): Promise<void> {
  await page.getByTestId("workspace-import-input").setInputFiles(files);
  await expect(page.getByTestId("workspace-import-source")).toHaveCount(
    files.length,
  );
  await page
    .getByTestId("workspace-import-source")
    .first()
    .getByTestId("workspace-import-role")
    .fill("inventory");
  await page
    .getByTestId("workspace-import-source")
    .first()
    .getByTestId("workspace-import-revision")
    .fill("Iteration 1");
  await page.getByTestId("workspace-import-prepare").click();
  await expect(page.getByTestId("workspace-import-review")).toBeVisible();
}

async function resolveAndApply(page: Page): Promise<void> {
  await page.getByTestId("workspace-import-resolve").click();
  await expect(page.getByTestId("workspace-import-review")).toContainText(
    "ready",
  );
  await page.getByTestId("workspace-delivery-vendor").fill("Vendor A");
  await page.getByTestId("workspace-delivery-entity").fill("Entity North");
  await page.getByTestId("workspace-delivery-phase").fill("Iteration 1");
  await page.getByTestId("workspace-delivery-coverage").selectOption("partial");
  await page.getByTestId("workspace-import-apply").click();
  await expect(page.getByTestId("workspace-import-result")).toContainText(
    "Added",
  );
}

test.describe("reviewed workbook imports", () => {
  for (const format of ["sqlite", "duckdb"] as const) {
    test(`imports multiple workbooks into ${format}`, async ({ page }) => {
      await create(page, format);
      await prepare(page, [await inventoryWorkbook(), await mappingWorkbook()]);
      await expect(page.getByTestId("workspace-import-region")).toHaveCount(2);
      await page.getByTestId("workspace-import-preview").first().click();
      await expect(
        page.getByTestId("workspace-import-preview-page"),
      ).toContainText("Customers");
      await resolveAndApply(page);
      await expect(page.getByTestId("workspace-table")).toHaveCount(2);
      await expect(page.getByTestId("workspace-table").nth(0)).toContainText(
        "2 rows",
      );
      await expect(page.getByTestId("workspace-table").nth(1)).toContainText(
        "1 row",
      );
      await page.getByTestId("workspace-deliveries-refresh").click();
      await expect(page.getByTestId("workspace-delivery")).toHaveCount(1);
    });

    test(`resumes a saved ${format} review after reload without Excel`, async ({
      page,
    }) => {
      await create(page, format);
      await prepare(page, [await inventoryWorkbook()]);

      await page.reload();
      await page.getByTestId("workspace-reopen").first().click();
      await expect(page.getByTestId("workspace-import-resume")).toHaveCount(1);
      await page.getByTestId("workspace-import-resume").click();
      await expect(page.getByTestId("workspace-import-review")).toContainText(
        "inventory.xlsx",
      );
      await resolveAndApply(page);
      await expect(page.getByTestId("workspace-import-result")).toContainText(
        "Added 2 rows",
      );
      await expect(page.getByTestId("workspace-table")).toHaveCount(1);
      await expect(page.getByTestId("workspace-table")).toContainText("2 rows");

      await page.reload();
      await page.getByTestId("workspace-reopen").first().click();
      await expect(page.getByTestId("workspace-import-resume")).toHaveCount(0);
    });
  }

  test("recovers a committed import when its worker reply is lost", async ({
    page,
  }) => {
    await page.addInitScript(() => {
      const NativeWorker = window.Worker;
      class RecoverableWorker extends EventTarget {
        readonly worker: Worker;

        constructor(url: string | URL, options?: WorkerOptions) {
          super();
          this.worker = new NativeWorker(url, options);
          this.worker.addEventListener("message", (event) => {
            const message: unknown = event.data;
            if (
              typeof message === "object" &&
              message !== null &&
              "type" in message &&
              message.type === "importApplied" &&
              window.localStorage.getItem("drop-import-applied") === "yes"
            ) {
              window.localStorage.setItem("drop-import-applied", "done");
              this.worker.terminate();
              this.dispatchEvent(new Event("error"));
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
        value: RecoverableWorker,
      });
    });

    await create(page, "sqlite");
    await prepare(page, [await inventoryWorkbook()]);
    await page.getByTestId("workspace-import-resolve").click();
    await expect(page.getByTestId("workspace-import-review")).toContainText(
      "ready",
    );
    await page.getByTestId("workspace-delivery-vendor").fill("Vendor A");
    await page.evaluate(() => {
      window.localStorage.setItem("drop-import-applied", "yes");
    });
    await page.getByTestId("workspace-import-apply").click();
    await expect(page.getByTestId("workspace-error")).toContainText(
      "worker is no longer available",
    );

    await page.reload();
    await page.getByTestId("workspace-reopen").first().click();
    await expect(page.getByTestId("workspace-import-resume")).toContainText(
      "Finish recovery",
    );
    await page.getByTestId("workspace-import-resume").click();
    await expect(page.getByTestId("workspace-delivery-vendor")).toHaveValue(
      "Vendor A",
    );
    await page.getByTestId("workspace-import-apply").click();
    await expect(page.getByTestId("workspace-import-result")).toContainText(
      "already applied",
    );
    await expect(page.getByTestId("workspace-table")).toContainText("2 rows");
    await page.getByTestId("workspace-deliveries-refresh").click();
    await expect(page.getByTestId("workspace-delivery")).toHaveCount(1);
  });

  test("skips a repeat capture and can record another delivery", async ({
    page,
  }) => {
    await create(page, "sqlite");
    const workbook = await inventoryWorkbook();
    await prepare(page, [workbook]);
    await resolveAndApply(page);

    await prepare(page, [workbook]);
    await expect(page.getByTestId("workspace-import-duplicate")).toBeVisible();
    await page.getByTestId("workspace-delivery-vendor").fill("Vendor A");
    await page.getByTestId("workspace-delivery-phase").fill("Iteration 2");
    await page.getByTestId("workspace-delivery-record-reuse").click();
    await expect(page.getByTestId("workspace-import-result")).toContainText(
      "reused the captured rows",
    );
    await page.getByTestId("workspace-deliveries-refresh").click();
    await expect(page.getByTestId("workspace-delivery")).toHaveCount(2);
    await expect(page.getByTestId("workspace-delivery").last()).toContainText(
      "Reused captured data",
    );
  });

  test("appends a changed submission to an existing table", async ({
    page,
  }) => {
    await create(page, "sqlite");
    await prepare(page, [await inventoryWorkbook()]);
    await resolveAndApply(page);

    await prepare(page, [
      await inventoryWorkbook("inventory-v2.xlsx", [
        ["Payments", "Payment ID", "true"],
      ]),
    ]);
    await page.getByTestId("workspace-import-route").selectOption("append");
    await page.getByTestId("workspace-import-table").fill("Inventory");
    await resolveAndApply(page);
    await expect(page.getByTestId("workspace-table").first()).toContainText(
      "5 rows",
    );
  });

  test("cancels preparation without publishing rows", async ({ page }) => {
    await create(page, "sqlite");
    await page
      .getByTestId("workspace-import-input")
      .setInputFiles(
        await inventoryWorkbook("large.xlsx", generatedRows(5_000)),
      );
    await page.getByTestId("workspace-import-prepare").click();
    await page.getByTestId("workspace-cancel").click();
    await expect(page.getByTestId("workspace-notice")).toContainText(
      "Cancelled",
    );
    await expect(page.getByTestId("workspace-table")).toHaveCount(0);
  });

  test("cancels a SQLite apply and rolls back its table", async ({ page }) => {
    await create(page, "sqlite");
    await prepare(page, [
      await inventoryWorkbook("large-apply.xlsx", generatedRows(2_500)),
    ]);
    await page.getByTestId("workspace-import-resolve").click();
    await expect(page.getByTestId("workspace-import-review")).toContainText(
      "ready",
    );
    await page.getByTestId("workspace-import-apply").click();
    await page.getByTestId("workspace-cancel").click();
    await expect(page.getByTestId("workspace-notice")).toContainText(
      "Cancelled",
    );
    await expect(page.getByTestId("workspace-table")).toHaveCount(0);

    await page.reload();
    await page.getByTestId("workspace-reopen").first().click();
    await expect(page.getByTestId("workspace-import-resume")).toHaveCount(1);
    await page.getByTestId("workspace-import-resume").click();
    await page.getByTestId("workspace-import-apply").click();
    await expect(page.getByTestId("workspace-import-result")).toContainText(
      "Added 2,502 rows",
    );
    await expect(page.getByTestId("workspace-table")).toContainText(
      "2,502 rows",
    );
  });
});
