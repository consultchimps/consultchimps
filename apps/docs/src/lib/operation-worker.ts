/**
 * The page-side half of the operation worker: one module Web Worker, shared by
 * every tool page, driven through promises.
 *
 * The worker is created on the first task rather than on import, so a visitor
 * who only reads a tool page never fetches it. Next.js resolves the
 * `new Worker(new URL(...), { type: "module" })` form at build time and emits
 * the worker as its own chunk, which keeps the pattern compatible with the
 * site's static export.
 */
import { ConsultChimpsError, type ProgressReporter } from "@consultchimps/core";

import {
  fromTransferable,
  type CancelCommand,
  type OperationTask,
  type OperationTaskResult,
  type RunCommand,
  type WorkerEvent,
} from "./operation-tasks";

export const WORKER_UNAVAILABLE = "WORKER_UNAVAILABLE";

interface PendingTask {
  readonly onProgress: ProgressReporter | undefined;
  readonly reject: (error: unknown) => void;
  readonly resolve: (value: unknown) => void;
}

const pending = new Map<number, PendingTask>();
let worker: Worker | null = null;
let lastTaskId = 0;

function settle(id: number): PendingTask | undefined {
  const task = pending.get(id);
  pending.delete(id);
  return task;
}

function handleEvent(event: MessageEvent<WorkerEvent>): void {
  const message = event.data;
  if (message.type === "progress") {
    pending.get(message.id)?.onProgress?.(message.progress);
    return;
  }
  const task = settle(message.id);
  if (!task) {
    return;
  }
  if (message.type === "failed") {
    task.reject(
      message.code === undefined
        ? new Error(message.message)
        : // Rebuilt with this bundle's class so the shared error formatting,
          // which tests for a ConsultChimpsError, still recognizes it.
          new ConsultChimpsError(message.code, message.message),
    );
    return;
  }
  task.resolve(
    message.artifacts === undefined
      ? message.value
      : {
          result: message.value,
          outputs: message.artifacts.map(fromTransferable),
        },
  );
}

function failEveryPendingTask(reason: string): void {
  const failed = [...pending.values()];
  pending.clear();
  for (const task of failed) {
    task.reject(new ConsultChimpsError(WORKER_UNAVAILABLE, reason));
  }
}

function operationWorker(): Worker {
  if (worker) {
    return worker;
  }
  const created = new Worker(
    new URL("../workers/operations.worker.ts", import.meta.url),
    { type: "module" },
  );
  created.addEventListener("message", handleEvent);
  created.addEventListener("error", () => {
    // A worker that cannot start would otherwise leave every caller waiting.
    // Dropping the instance lets the next task try again from scratch.
    worker = null;
    created.terminate();
    failEveryPendingTask(
      "The background worker that runs this task could not start.",
    );
  });
  worker = created;
  return created;
}

/**
 * Run one task in the worker.
 *
 * Progress events are forwarded as they arrive. Aborting `signal` sends a
 * cancel command; the worker aborts its own controller and the operation
 * rejects with the usual cancellation error, so callers see the same outcome
 * they saw when the work ran on the main thread.
 */
export async function runOperation<TTask extends OperationTask>(
  task: TTask,
  controls: {
    readonly onProgress?: ProgressReporter | undefined;
    readonly signal?: AbortSignal | undefined;
  } = {},
): Promise<OperationTaskResult<TTask>> {
  const instance = operationWorker();
  lastTaskId += 1;
  const id = lastTaskId;
  const { signal } = controls;

  return new Promise<OperationTaskResult<TTask>>((resolve, reject) => {
    const cancel = (): void => {
      instance.postMessage({ type: "cancel", id } satisfies CancelCommand);
    };
    const stopListening = (): void => {
      signal?.removeEventListener("abort", cancel);
    };

    pending.set(id, {
      onProgress: controls.onProgress,
      reject: (error) => {
        stopListening();
        reject(error);
      },
      // The worker answers with exactly the shape this task's result type
      // describes; the boundary is untyped, so the cast lands here alone.
      resolve: (value) => {
        stopListening();
        resolve(value as OperationTaskResult<TTask>);
      },
    });

    instance.postMessage({ type: "run", id, task } satisfies RunCommand);

    if (signal?.aborted) {
      cancel();
      return;
    }
    signal?.addEventListener("abort", cancel, { once: true });
  });
}
