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
  readonly #owners = new Map<string, Set<RetainedBrowserOpenOwner>>();

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
      close,
      causes: [...causes],
    };
    const owners = this.#owners.get(storageName) ?? new Set();
    owners.add(owner);
    this.#owners.set(storageName, owners);
    return new BrowserOpenCleanupError(storageName, [owner]);
  }

  async retry(storageName: string): Promise<void> {
    const retained = this.#owners.get(storageName);
    if (retained === undefined || retained.size === 0) return;
    const owners = [...retained];
    const results = await Promise.allSettled(
      owners.map(async (owner) => owner.close()),
    );
    const stillOpen: RetainedBrowserOpenOwner[] = [];
    for (const [index, result] of results.entries()) {
      const owner = owners[index];
      if (owner === undefined) continue;
      if (result.status === "fulfilled") retained.delete(owner);
      else {
        owner.causes.push(result.reason);
        stillOpen.push(owner);
      }
    }
    if (retained.size === 0) this.#owners.delete(storageName);
    if (stillOpen.length > 0) {
      throw new BrowserOpenCleanupError(storageName, stillOpen);
    }
  }
}
