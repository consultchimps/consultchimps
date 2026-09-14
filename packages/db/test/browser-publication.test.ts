import { describe, expect, test, vi } from "vitest";

import {
  BrowserExportCleanupError,
  BrowserPublicationCleanupError,
  BrowserPublicationRecoveryError,
  publishBrowserCandidate,
  publishBrowserExport,
  type BrowserPublicationOperations,
} from "../src/browser-publication.js";

interface StoredRows {
  current: string[];
  backup?: string[];
  candidate?: string[];
}

function operations(
  stored: StoredRows,
): BrowserPublicationOperations<readonly string[]> {
  return {
    async backup() {
      stored.backup = [...stored.current];
    },
    async publishAndOpen() {
      stored.current = [...(stored.candidate ?? [])];
      return stored.current;
    },
    async restore() {
      stored.current = [...(stored.backup ?? [])];
    },
    async cleanupBackups() {
      delete stored.backup;
    },
    async cleanupCandidate() {
      delete stored.candidate;
    },
  };
}

describe("browser database publication", () => {
  test("publishes the candidate and treats cleanup failure as post-commit", async () => {
    const stored: StoredRows = {
      current: ["old row"],
      candidate: ["replacement row"],
    };
    const replacement = operations(stored);
    replacement.cleanupBackups = vi
      .fn<() => Promise<void>>()
      .mockRejectedValue(new Error("cleanup failed"));

    const result = await publishBrowserCandidate(replacement);

    expect(result.value).toEqual(["replacement row"]);
    expect(result.cleanupFailures).toHaveLength(1);
    expect(stored.current).toEqual(["replacement row"]);
    expect(stored.candidate).toBeUndefined();
  });

  test("restores persisted rows when publication fails", async () => {
    const stored: StoredRows = {
      current: ["old row"],
      candidate: ["replacement row"],
    };
    const replacement = operations(stored);
    const failure = new Error("storage write failed");
    replacement.publishAndOpen = async () => {
      stored.current = ["partial replacement"];
      throw failure;
    };

    await expect(publishBrowserCandidate(replacement)).rejects.toBe(failure);
    expect(stored.current).toEqual(["old row"]);
    expect(stored.backup).toBeUndefined();
    expect(stored.candidate).toBeUndefined();
  });

  test("retains the recoverable backup when restoration fails", async () => {
    const stored: StoredRows = {
      current: ["old row"],
      candidate: ["replacement row"],
    };
    const replacement = operations(stored);
    replacement.publishAndOpen = async () => {
      stored.current = ["partial replacement"];
      throw new Error("storage write failed");
    };
    replacement.restore = async () => {
      throw new Error("storage restore failed");
    };

    await expect(publishBrowserCandidate(replacement)).rejects.toBeInstanceOf(
      BrowserPublicationRecoveryError,
    );
    expect(stored.backup).toEqual(["old row"]);
    expect(stored.candidate).toBeUndefined();
  });

  test("does not publish when the backup cannot be created", async () => {
    const stored: StoredRows = {
      current: ["old row"],
      candidate: ["replacement row"],
    };
    const replacement = operations(stored);
    replacement.backup = async () => {
      throw new Error("backup failed");
    };

    await expect(publishBrowserCandidate(replacement)).rejects.toThrow(
      "backup failed",
    );
    expect(stored.current).toEqual(["old row"]);
    expect(stored.candidate).toBeUndefined();
  });

  test("reports cleanup failure after partial backup creation", async () => {
    const stored: StoredRows = {
      current: ["old row"],
      candidate: ["replacement row"],
    };
    const replacement = operations(stored);
    const backupFailure = new Error("backup write failed");
    const cleanupFailure = new Error("backup cleanup failed");
    replacement.backup = async () => {
      stored.backup = ["partial old row"];
      throw backupFailure;
    };
    replacement.cleanupBackups = async () => {
      throw cleanupFailure;
    };

    let rejection: unknown;
    try {
      await publishBrowserCandidate(replacement);
    } catch (error) {
      rejection = error;
    }
    expect(rejection).toMatchObject({
      phase: "backup",
      operationCause: backupFailure,
      cleanupFailures: [cleanupFailure],
    });
    expect(rejection).toBeInstanceOf(BrowserPublicationCleanupError);
    expect((rejection as Error).cause).toBeInstanceOf(AggregateError);
    expect(stored.current).toEqual(["old row"]);
    expect(stored.backup).toEqual(["partial old row"]);
    expect(stored.candidate).toBeUndefined();
  });

  test("reports cleanup failure after restoring a failed publication", async () => {
    const stored: StoredRows = {
      current: ["old row"],
      candidate: ["replacement row"],
    };
    const replacement = operations(stored);
    const publicationFailure = new Error("publication failed");
    const cleanupFailure = new Error("candidate cleanup failed");
    replacement.publishAndOpen = async () => {
      stored.current = ["partial replacement"];
      throw publicationFailure;
    };
    replacement.cleanupCandidate = async () => {
      throw cleanupFailure;
    };

    await expect(publishBrowserCandidate(replacement)).rejects.toMatchObject({
      phase: "publication",
      operationCause: publicationFailure,
      cleanupFailures: [cleanupFailure],
    });
    expect(stored.current).toEqual(["old row"]);
    expect(stored.backup).toBeUndefined();
    expect(stored.candidate).toEqual(["replacement row"]);
  });

  test("retains candidate cleanup failure without deleting a recovery backup", async () => {
    const stored: StoredRows = {
      current: ["old row"],
      candidate: ["replacement row"],
    };
    const replacement = operations(stored);
    const cleanupFailure = new Error("candidate cleanup failed");
    replacement.publishAndOpen = async () => {
      stored.current = ["partial replacement"];
      throw new Error("publication failed");
    };
    replacement.restore = async () => {
      throw new Error("restoration failed");
    };
    replacement.cleanupCandidate = async () => {
      throw cleanupFailure;
    };

    await expect(publishBrowserCandidate(replacement)).rejects.toMatchObject({
      cleanupFailures: [cleanupFailure],
    });
    expect(stored.backup).toEqual(["old row"]);
    expect(stored.candidate).toEqual(["replacement row"]);
  });

  test("rethrows an abort after successful private cleanup", async () => {
    const stored: StoredRows = {
      current: ["old row"],
      candidate: ["replacement row"],
    };
    const replacement = operations(stored);
    const abort = new DOMException("The operation was aborted", "AbortError");
    replacement.backup = async () => {
      throw abort;
    };

    await expect(publishBrowserCandidate(replacement)).rejects.toBe(abort);
    expect(stored.current).toEqual(["old row"]);
    expect(stored.backup).toBeUndefined();
    expect(stored.candidate).toBeUndefined();
  });
});

describe("browser export publication", () => {
  test("restores the destination after a failed export", async () => {
    const stored: StoredRows = { current: ["old row"] };
    const failure = new Error("export failed");

    await expect(
      publishBrowserExport({
        async backup() {
          stored.backup = [...stored.current];
        },
        async publish() {
          stored.current = ["partial export"];
          throw failure;
        },
        async restore() {
          stored.current = [...(stored.backup ?? [])];
        },
        async cleanupBackup() {
          delete stored.backup;
        },
      }),
    ).rejects.toBe(failure);
    expect(stored.current).toEqual(["old row"]);
    expect(stored.backup).toBeUndefined();
  });

  test("retains the export backup when restoration fails", async () => {
    const stored: StoredRows = { current: ["old row"] };

    await expect(
      publishBrowserExport({
        async backup() {
          stored.backup = [...stored.current];
        },
        async publish() {
          stored.current = ["partial export"];
          throw new Error("export failed");
        },
        async restore() {
          throw new Error("restore failed");
        },
        async cleanupBackup() {
          delete stored.backup;
        },
      }),
    ).rejects.toBeInstanceOf(BrowserPublicationRecoveryError);
    expect(stored.backup).toEqual(["old row"]);
  });

  test("reports cleanup failure after restoring the destination", async () => {
    const stored: StoredRows = { current: ["old row"] };

    await expect(
      publishBrowserExport({
        async backup() {
          stored.backup = [...stored.current];
        },
        async publish() {
          stored.current = ["partial export"];
          throw new Error("export failed");
        },
        async restore() {
          stored.current = [...(stored.backup ?? [])];
        },
        async cleanupBackup() {
          throw new Error("cleanup failed");
        },
      }),
    ).rejects.toBeInstanceOf(BrowserExportCleanupError);
    expect(stored.current).toEqual(["old row"]);
    expect(stored.backup).toEqual(["old row"]);
  });
});
