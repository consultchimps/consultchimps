/**
 * Single-tab ownership of the database tool.
 *
 * The browser SQLite engine keeps exclusive synchronous access handles on its
 * origin-private pool files for as long as its worker lives. A second tab that
 * starts the engine cannot acquire those handles and fails with a raw browser
 * error. This module makes the decision explicit and early: the page claims an
 * exclusive Web Lock before it creates a worker, and a page that cannot claim
 * the lock reports the conflict instead of starting an engine.
 *
 * The lock is scoped to the origin by the browser, so it also covers the same
 * site served under a different base path. A browser without the Web Locks API
 * is treated as if the claim succeeded, matching the export-lease behavior in
 * workspace-files.ts; the engine start reports the conflict in that case.
 */

export const DATABASE_TOOL_LOCK_NAME = "consultchimps:database-tool";

export interface OwnershipLockManager {
  request<T>(
    name: string,
    options: {
      readonly mode: "exclusive";
      readonly ifAvailable?: boolean;
      readonly signal?: AbortSignal;
    },
    callback: (lock: object | null) => Promise<T>,
  ): Promise<T>;
}

export interface OwnedDatabaseTool {
  readonly state: "owned";
  /** Give the lock back; a no-op after the first call. */
  release(): void;
}

export type DatabaseToolOwnership =
  | OwnedDatabaseTool
  | { readonly state: "unavailable" }
  | { readonly state: "conflict" };

function databaseToolLockManager(): OwnershipLockManager | null {
  const candidate: unknown = (
    globalThis.navigator as Navigator & { readonly locks?: unknown }
  ).locks;
  return typeof candidate === "object" && candidate !== null
    ? (candidate as OwnershipLockManager)
    : null;
}

const UNAVAILABLE: DatabaseToolOwnership = { state: "unavailable" };

function requestOwnership(
  locks: OwnershipLockManager,
  options: { readonly ifAvailable: boolean; readonly signal?: AbortSignal },
): Promise<DatabaseToolOwnership> {
  let release = (): void => undefined;
  const held = new Promise<void>((resolve) => {
    release = resolve;
  });
  return new Promise<DatabaseToolOwnership>((resolve) => {
    // A refused request, whether it throws on the spot (a lock manager that
    // is present but disabled) or rejects later (an aborted wait), leaves
    // nothing to release. Report it as unavailable so the engine start
    // decides.
    try {
      locks
        .request(
          DATABASE_TOOL_LOCK_NAME,
          {
            mode: "exclusive",
            ...(options.ifAvailable ? { ifAvailable: true } : {}),
            ...(options.signal === undefined ? {} : { signal: options.signal }),
          },
          async (lock) => {
            if (lock === null) {
              resolve({ state: "conflict" });
              return;
            }
            resolve({ state: "owned", release });
            await held;
          },
        )
        .catch(() => resolve(UNAVAILABLE));
    } catch {
      resolve(UNAVAILABLE);
    }
  });
}

/**
 * Claim the database tool for this page without waiting. Resolves once the
 * browser has answered: `owned` holds the lock until `release` runs,
 * `conflict` means another page of this origin holds it, and `unavailable`
 * means the browser has no lock manager so nothing can be claimed.
 */
export function claimDatabaseTool(
  locks: OwnershipLockManager | null = databaseToolLockManager(),
): Promise<DatabaseToolOwnership> {
  if (locks === null) return Promise.resolve(UNAVAILABLE);
  return requestOwnership(locks, { ifAvailable: true });
}

/**
 * Wait until the database tool can be owned by this page. Resolves with the
 * held lock once the previous owner releases it. Aborting the signal resolves
 * with `unavailable` and takes nothing.
 */
export function waitForDatabaseTool(
  signal: AbortSignal,
  locks: OwnershipLockManager | null = databaseToolLockManager(),
): Promise<DatabaseToolOwnership> {
  if (locks === null || signal.aborted) return Promise.resolve(UNAVAILABLE);
  return requestOwnership(locks, { ifAvailable: false, signal });
}
