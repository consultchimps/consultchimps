export interface BrowserPublicationOperations<T> {
  backup(): Promise<void>;
  publishAndOpen(): Promise<T>;
  restore(): Promise<void>;
  cleanupBackups(): Promise<void>;
  cleanupCandidate(): Promise<void>;
}

export interface BrowserPublicationResult<T> {
  readonly value: T;
  readonly cleanupFailures: readonly unknown[];
}

export class BrowserPublicationRecoveryError extends Error {
  readonly publicationCause: unknown;
  readonly recoveryCause: unknown;

  constructor(publicationCause: unknown, recoveryCause: unknown) {
    super("Browser database replacement recovery failed", {
      cause: new AggregateError([publicationCause, recoveryCause]),
    });
    this.publicationCause = publicationCause;
    this.recoveryCause = recoveryCause;
  }
}

export interface BrowserExportPublicationOperations<T> {
  backup(): Promise<void>;
  publish(): Promise<T>;
  restore(): Promise<void>;
  cleanupBackup(): Promise<void>;
}

export interface BrowserExportPublicationResult<T> {
  readonly value: T;
  readonly cleanupFailures: readonly unknown[];
}

export class BrowserExportCleanupError extends Error {
  readonly exportCause: unknown;
  readonly cleanupCause: unknown;

  constructor(exportCause: unknown, cleanupCause: unknown) {
    super("Browser export backup cleanup failed", {
      cause: new AggregateError([exportCause, cleanupCause]),
    });
    this.exportCause = exportCause;
    this.cleanupCause = cleanupCause;
  }
}

export async function publishBrowserExport<T>(
  operations: BrowserExportPublicationOperations<T>,
): Promise<BrowserExportPublicationResult<T>> {
  try {
    await operations.backup();
  } catch (exportCause) {
    try {
      await operations.cleanupBackup();
    } catch (cleanupCause) {
      throw new BrowserExportCleanupError(exportCause, cleanupCause);
    }
    throw exportCause;
  }

  let value: T;
  try {
    value = await operations.publish();
  } catch (publicationCause) {
    try {
      await operations.restore();
    } catch (recoveryCause) {
      throw new BrowserPublicationRecoveryError(
        publicationCause,
        recoveryCause,
      );
    }
    try {
      await operations.cleanupBackup();
    } catch (cleanupCause) {
      throw new BrowserExportCleanupError(publicationCause, cleanupCause);
    }
    throw publicationCause;
  }

  try {
    await operations.cleanupBackup();
    return { value, cleanupFailures: [] };
  } catch (cleanupError) {
    return { value, cleanupFailures: [cleanupError] };
  }
}

async function cleanup(
  operations: Pick<
    BrowserPublicationOperations<unknown>,
    "cleanupBackups" | "cleanupCandidate"
  >,
): Promise<readonly unknown[]> {
  const results = await Promise.allSettled([
    operations.cleanupBackups(),
    operations.cleanupCandidate(),
  ]);
  return results.flatMap((result) =>
    result.status === "rejected" ? [result.reason] : [],
  );
}

export async function publishBrowserCandidate<T>(
  operations: BrowserPublicationOperations<T>,
): Promise<BrowserPublicationResult<T>> {
  try {
    await operations.backup();
  } catch (error) {
    await cleanup(operations);
    throw error;
  }

  let value: T;
  try {
    value = await operations.publishAndOpen();
  } catch (publicationCause) {
    try {
      await operations.restore();
    } catch (recoveryCause) {
      await operations.cleanupCandidate().catch(() => undefined);
      throw new BrowserPublicationRecoveryError(
        publicationCause,
        recoveryCause,
      );
    }
    await cleanup(operations);
    throw publicationCause;
  }

  return { value, cleanupFailures: await cleanup(operations) };
}
