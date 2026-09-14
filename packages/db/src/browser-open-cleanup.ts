import { OwnedResources } from "@consultchimps/core";

export type BrowserOpenOwnerKind = "database" | "prepared" | "inspection";

interface RetainedBrowserOpenOwner {
  readonly kind: BrowserOpenOwnerKind;
  readonly close: () => Promise<void>;
  causes: unknown[];
}

export class BrowserOpenCleanupError extends Error {
  readonly storageName: string;
  readonly ownerKinds: readonly BrowserOpenOwnerKind[];
  readonly cleanupCauses: readonly unknown[];

  constructor(
    storageName: string,
    owners: readonly RetainedBrowserOpenOwner[],
  ) {
    const cleanupCauses = owners.flatMap((owner) => owner.causes);
    super("Browser open cleanup failed", {
      cause: new AggregateError(cleanupCauses, "Browser open cleanup failed"),
    });
    this.storageName = storageName;
    this.ownerKinds = [...new Set(owners.map((owner) => owner.kind))];
    this.cleanupCauses = cleanupCauses;
  }
}

export class BrowserOpenCleanupRegistry {
  readonly #owners = new Map<
    string,
    OwnedResources<RetainedBrowserOpenOwner>
  >();

  has(storageName: string): boolean {
    return (this.#owners.get(storageName)?.size ?? 0) > 0;
  }

  retain(
    storageName: string,
    kind: BrowserOpenOwnerKind,
    close: () => Promise<void>,
    causes: readonly unknown[],
  ): BrowserOpenCleanupError {
    const owner: RetainedBrowserOpenOwner = {
      kind,
      causes: [...causes],
      async close() {
        try {
          await close();
        } catch (error) {
          owner.causes.push(error);
          throw error;
        }
      },
    };
    const owners =
      this.#owners.get(storageName) ??
      new OwnedResources<RetainedBrowserOpenOwner>();
    owners.add(owner);
    this.#owners.set(storageName, owners);
    return new BrowserOpenCleanupError(storageName, [owner]);
  }

  async retry(storageName: string): Promise<void> {
    const retained = this.#owners.get(storageName);
    if (retained === undefined || retained.size === 0) return;
    const failures = await retained.close();
    if (retained.size === 0 && this.#owners.get(storageName) === retained) {
      this.#owners.delete(storageName);
    }
    if (failures.length > 0) {
      throw new BrowserOpenCleanupError(
        storageName,
        failures.map((failure) => failure.resource),
      );
    }
  }
}
