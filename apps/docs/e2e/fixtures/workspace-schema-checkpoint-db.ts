export * from "@consultchimps/db";

import { applySchema as apply } from "@consultchimps/db";

let failureInjected = false;

export async function applySchema(
  options: Parameters<typeof apply>[0],
): ReturnType<typeof apply> {
  const result = await apply(options);
  const checkpoint = options.database.checkpoint.bind(options.database);
  Object.defineProperty(options.database, "checkpoint", {
    configurable: true,
    value: async () => {
      if (!failureInjected) {
        failureInjected = true;
        throw new Error("Injected checkpoint failure");
      }
      await checkpoint();
    },
  });
  return result;
}
