import { afterEach, describe, expect, it, vi } from "vitest";

import { BrowserOpfsFile } from "./workspace-files";

function opfsFixture(options?: {
  readonly flush?: (() => void) | undefined;
  readonly close?: (() => void) | undefined;
  readonly getFile?: (() => Promise<File>) | undefined;
  readonly remove?: ((name: string) => Promise<void>) | undefined;
}) {
  let bytes = new Uint8Array(0);
  const access = {
    read(target: Uint8Array, { at }: { readonly at: number }) {
      const source = bytes.subarray(at, at + target.byteLength);
      target.set(source);
      return source.byteLength;
    },
    write(source: Uint8Array, { at }: { readonly at: number }) {
      const next = new Uint8Array(
        Math.max(bytes.byteLength, at + source.length),
      );
      next.set(bytes);
      next.set(source, at);
      bytes = next;
      return source.byteLength;
    },
    truncate(size: number) {
      bytes = bytes.slice(0, size);
    },
    flush: vi.fn(options?.flush ?? (() => undefined)),
    close: vi.fn(options?.close ?? (() => undefined)),
  };
  const handle = {
    name: "scratch",
    getFile: vi.fn(
      options?.getFile ?? (async () => new File([bytes], "scratch")),
    ),
    async createSyncAccessHandle() {
      return access;
    },
  };
  const remove = vi.fn(options?.remove ?? (async () => undefined));
  vi.stubGlobal("navigator", {
    storage: {
      async getDirectory() {
        return {
          async getFileHandle() {
            return handle;
          },
          removeEntry: remove,
        };
      },
    },
  });
  return { access, handle, remove };
}

describe("BrowserOpfsFile", () => {
  afterEach(() => vi.unstubAllGlobals());

  it("keeps one sync access handle for its bounded lifecycle", async () => {
    let bytes = new Uint8Array(0);
    let accessCreates = 0;
    let closes = 0;
    const removed: string[] = [];
    const access = {
      read(target: Uint8Array, { at }: { readonly at: number }) {
        const source = bytes.subarray(at, at + target.byteLength);
        target.set(source);
        return source.byteLength;
      },
      write(source: Uint8Array, { at }: { readonly at: number }) {
        const next = new Uint8Array(
          Math.max(bytes.byteLength, at + source.length),
        );
        next.set(bytes);
        next.set(source, at);
        bytes = next;
        return source.byteLength;
      },
      truncate(size: number) {
        const next = new Uint8Array(size);
        next.set(bytes.subarray(0, size));
        bytes = next;
      },
      flush: vi.fn(),
      close() {
        closes += 1;
      },
    };
    const handle = {
      name: "scratch",
      async getFile() {
        return new File([bytes], "scratch");
      },
      async createSyncAccessHandle() {
        accessCreates += 1;
        return access;
      },
    };
    vi.stubGlobal("navigator", {
      storage: {
        async getDirectory() {
          return {
            async getFileHandle() {
              return handle;
            },
            async removeEntry(name: string) {
              removed.push(name);
            },
          };
        },
      },
    });

    const file = await BrowserOpfsFile.open("scratch", true, true);
    await file.writeAt(0, new Uint8Array([1, 2]));
    await file.writeAt(2, new Uint8Array([3, 4]));
    expect(await file.readAt(1, 3)).toEqual(new Uint8Array([2, 3, 4]));
    await file.truncate(3);
    expect(file.size).toBe(3);
    await file.close();
    await file.close();

    expect(accessCreates).toBe(1);
    expect(closes).toBe(1);
    expect(removed).toEqual(["scratch"]);
  });

  it("releases and removes the file even when flush fails", async () => {
    const flushFailure = new Error("Injected flush failure");
    const fixture = opfsFixture({
      flush() {
        throw flushFailure;
      },
    });
    const file = await BrowserOpfsFile.open("scratch", true, true);

    await expect(file.close()).rejects.toBe(flushFailure);
    await expect(file.close()).resolves.toBeUndefined();
    expect(fixture.access.flush).toHaveBeenCalledOnce();
    expect(fixture.access.close).toHaveBeenCalledOnce();
    expect(fixture.remove).toHaveBeenCalledOnce();
  });

  it("aggregates flush and removal failures then retries removal", async () => {
    const flushFailure = new Error("Injected flush failure");
    const removalFailure = new Error("Injected removal failure");
    const remove = vi
      .fn<(name: string) => Promise<void>>()
      .mockRejectedValueOnce(removalFailure)
      .mockResolvedValueOnce(undefined);
    const fixture = opfsFixture({
      flush() {
        throw flushFailure;
      },
      remove,
    });
    const file = await BrowserOpfsFile.open("scratch", true, true);

    let rejection: unknown;
    try {
      await file.close();
    } catch (error) {
      rejection = error;
    }
    expect(rejection).toBeInstanceOf(AggregateError);
    expect((rejection as AggregateError).errors).toEqual([
      flushFailure,
      removalFailure,
    ]);
    await expect(file.close()).resolves.toBeUndefined();
    expect(fixture.access.flush).toHaveBeenCalledOnce();
    expect(fixture.access.close).toHaveBeenCalledOnce();
    expect(remove).toHaveBeenCalledTimes(2);
  });

  it("retries access release before removing the file", async () => {
    const closeFailure = new Error("Injected access close failure");
    const close = vi
      .fn<() => void>()
      .mockImplementationOnce(() => {
        throw closeFailure;
      })
      .mockImplementationOnce(() => undefined);
    const fixture = opfsFixture({ close });
    const file = await BrowserOpfsFile.open("scratch", true, true);

    await expect(file.close()).rejects.toBe(closeFailure);
    expect(fixture.remove).not.toHaveBeenCalled();
    await expect(file.close()).resolves.toBeUndefined();
    expect(fixture.access.flush).toHaveBeenCalledTimes(2);
    expect(close).toHaveBeenCalledTimes(2);
    expect(fixture.remove).toHaveBeenCalledOnce();
  });

  it("shares one close lifecycle across concurrent callers", async () => {
    let finishRemoval = (): void => undefined;
    const removal = new Promise<void>((resolve) => {
      finishRemoval = resolve;
    });
    const fixture = opfsFixture({ remove: async () => removal });
    const file = await BrowserOpfsFile.open("scratch", true, true);

    const first = file.close();
    const second = file.close();
    expect(second).toBe(first);
    finishRemoval();
    await Promise.all([first, second]);
    expect(fixture.access.close).toHaveBeenCalledOnce();
    expect(fixture.remove).toHaveBeenCalledOnce();
  });

  it("shares a pending flush failure when no removal is required", async () => {
    const flushFailure = new Error("Injected flush failure");
    const fixture = opfsFixture({
      flush() {
        throw flushFailure;
      },
    });
    const file = await BrowserOpfsFile.open("scratch", true, false);

    const first = file.close();
    const second = file.close();
    expect(second).toBe(first);
    expect(await Promise.allSettled([first, second])).toEqual([
      { status: "rejected", reason: flushFailure },
      { status: "rejected", reason: flushFailure },
    ]);
    expect(fixture.access.close).toHaveBeenCalledOnce();
    expect(fixture.remove).not.toHaveBeenCalled();
  });

  it("releases the access handle when opening cannot read file metadata", async () => {
    const openFailure = new Error("Injected metadata failure");
    const fixture = opfsFixture({
      async getFile() {
        throw openFailure;
      },
    });

    await expect(BrowserOpfsFile.open("scratch", true, true)).rejects.toBe(
      openFailure,
    );
    expect(fixture.access.close).toHaveBeenCalledOnce();
  });

  it("preserves metadata and access-release failures while opening", async () => {
    const openFailure = new Error("Injected metadata failure");
    const cleanupFailure = new Error("Injected access close failure");
    opfsFixture({
      async getFile() {
        throw openFailure;
      },
      close() {
        throw cleanupFailure;
      },
    });

    let rejection: unknown;
    try {
      await BrowserOpfsFile.open("scratch", true, true);
    } catch (error) {
      rejection = error;
    }
    expect(rejection).toBeInstanceOf(AggregateError);
    expect((rejection as AggregateError).errors).toEqual([
      openFailure,
      cleanupFailure,
    ]);
  });
});
