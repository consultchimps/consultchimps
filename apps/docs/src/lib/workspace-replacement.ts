export class WorkspaceReplacementCleanupError extends Error {
  readonly replacementCause: unknown;
  readonly nextCleanupCause: unknown;

  constructor(replacementCause: unknown, nextCleanupCause: unknown) {
    super(
      "The workspace switch failed, and the newly opened working copy could not be released. Reload the workspace to restart its worker before trying again",
      {
        cause: new AggregateError([replacementCause, nextCleanupCause]),
      },
    );
    this.replacementCause = replacementCause;
    this.nextCleanupCause = nextCleanupCause;
  }
}

export async function closeTrackedResources<Key, Value>(options: {
  readonly resources: Map<Key, Value>;
  close(value: Value): Promise<void>;
}): Promise<void> {
  const entries = [...options.resources.entries()];
  const results = await Promise.allSettled(
    entries.map(([, value]) => options.close(value)),
  );
  const failures: unknown[] = [];
  for (const [index, result] of results.entries()) {
    const [key, value] = entries[index]!;
    if (result.status === "rejected") {
      failures.push(result.reason);
    } else if (options.resources.get(key) === value) {
      options.resources.delete(key);
    }
  }
  if (failures.length === 1) throw failures[0];
  if (failures.length > 1) {
    throw new AggregateError(failures, "Closing tracked resources failed");
  }
}

export async function replaceActiveWorkspace<Workspace, Prepared>(options: {
  readonly previous: Workspace | null;
  readonly next: Workspace;
  prepare(next: Workspace): Promise<Prepared>;
  closeDependencies(): Promise<void>;
  close(workspace: Workspace): Promise<void>;
  activate(next: Workspace): void;
}): Promise<Prepared> {
  let prepared: Prepared;
  try {
    prepared = await options.prepare(options.next);
    await options.closeDependencies();
    if (options.previous !== null) await options.close(options.previous);
  } catch (replacementCause) {
    try {
      await options.close(options.next);
    } catch (nextCleanupCause) {
      throw new WorkspaceReplacementCleanupError(
        replacementCause,
        nextCleanupCause,
      );
    }
    throw replacementCause;
  }
  options.activate(options.next);
  return prepared;
}
