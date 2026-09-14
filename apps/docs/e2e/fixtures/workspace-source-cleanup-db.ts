export * from "@consultchimps/db";

import { createWorkbookImportSource as createSource } from "@consultchimps/db";

let sourceCount = 0;
let failureInjected = false;

export async function createWorkbookImportSource(
  options: Parameters<typeof createSource>[0],
): ReturnType<typeof createSource> {
  const opened = await createSource(options);
  sourceCount += 1;
  const sourceOrdinal = sourceCount;
  return {
    ...opened,
    async close() {
      if (sourceOrdinal === 2 && !failureInjected) {
        failureInjected = true;
        throw new Error("Injected workbook source close failure");
      }
      await opened.close();
    },
  };
}
