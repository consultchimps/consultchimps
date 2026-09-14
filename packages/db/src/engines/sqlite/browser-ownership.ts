interface BrowserSqliteOperationQueue {
  tail: Promise<void>;
  references: number;
}

export interface BrowserSqliteOperationLease {
  run<T>(work: () => Promise<T>): Promise<T>;
  release(): void;
}

export interface BrowserSqliteStorageIdentity {
  readonly owner: object;
  readonly name: string;
}

const queuesByOwner = new WeakMap<
  object,
  Map<string, BrowserSqliteOperationQueue>
>();

export function acquireBrowserSqliteOperationLease(
  storage: BrowserSqliteStorageIdentity,
): BrowserSqliteOperationLease {
  let queues = queuesByOwner.get(storage.owner);
  if (queues === undefined) {
    queues = new Map();
    queuesByOwner.set(storage.owner, queues);
  }
  let queue = queues.get(storage.name);
  if (queue === undefined) {
    queue = { tail: Promise.resolve(), references: 0 };
    queues.set(storage.name, queue);
  }
  queue.references += 1;
  let released = false;
  return {
    async run<T>(work: () => Promise<T>): Promise<T> {
      const previous = queue.tail;
      let release: () => void = () => undefined;
      queue.tail = new Promise<void>((resolve) => {
        release = resolve;
      });
      await previous;
      try {
        return await work();
      } finally {
        release();
      }
    },
    release(): void {
      if (released) return;
      released = true;
      queue.references -= 1;
      const finalOperation = queue.tail;
      void finalOperation.then(() => {
        if (queue.references === 0 && queue.tail === finalOperation) {
          queues.delete(storage.name);
        }
      });
    },
  };
}
