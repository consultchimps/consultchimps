import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { afterEach, expect, test } from "vitest";

import { NativeFileRegistry } from "../src/native-files.js";

const directories: string[] = [];

afterEach(async () => {
  await Promise.all(
    directories
      .splice(0)
      .map((directory) => rm(directory, { recursive: true, force: true })),
  );
});

async function paths(): Promise<{
  readonly first: string;
  readonly second: string;
}> {
  const directory = await mkdtemp(path.join(tmpdir(), "cc-native-identity-"));
  directories.push(directory);
  const first = path.join(directory, "first.sqlite");
  const second = path.join(directory, "second.sqlite");
  await Promise.all([writeFile(first, "first"), writeFile(second, "second")]);
  return { first, second };
}

function handle() {
  let open = true;
  return {
    get isOpen() {
      return open;
    },
    async close() {
      open = false;
    },
  };
}

test("allows case-distinct files when filesystem identities differ", async () => {
  const { first, second } = await paths();
  const identities = new Map([
    [
      path.resolve(first),
      {
        pathKey: "C:\\Sensitive\\Case.sqlite",
        device: 10n,
        inode: 20n,
      },
    ],
    [
      path.resolve(second),
      {
        pathKey: "C:\\Sensitive\\case.sqlite",
        device: 10n,
        inode: 21n,
      },
    ],
  ]);
  const registry = new NativeFileRegistry(async (filePath) => {
    const identity = identities.get(path.resolve(filePath));
    if (identity === undefined) throw new Error("Missing test identity");
    return identity;
  });
  const opened = handle();
  await registry.open(first, async () => opened);
  await expect(
    registry.planPublication({ output: second, inputs: [], overwrite: true }),
  ).resolves.toMatchObject({ output: path.resolve(second) });
  await opened.close();
});

test.each([
  {
    name: "case-insensitive alias",
    second: {
      pathKey: "C:\\Data\\case.sqlite",
      device: 10n,
      inode: 20n,
    },
  },
  {
    name: "symbolic-link alias",
    second: {
      pathKey: "C:\\Data\\Case.sqlite",
      device: 10n,
      inode: 20n,
    },
  },
  {
    name: "hard-link alias",
    second: {
      pathKey: "C:\\Data\\hard-link.sqlite",
      device: 10n,
      inode: 20n,
    },
  },
])("keeps a $name busy", async ({ second: secondIdentity }) => {
  const { first, second } = await paths();
  const identities = new Map([
    [
      path.resolve(first),
      {
        pathKey: "C:\\Data\\Case.sqlite",
        device: 10n,
        inode: 20n,
      },
    ],
    [path.resolve(second), secondIdentity],
  ]);
  const registry = new NativeFileRegistry(async (filePath) => {
    const identity = identities.get(path.resolve(filePath));
    if (identity === undefined) throw new Error("Missing test identity");
    return identity;
  });
  const opened = handle();
  await registry.open(first, async () => opened);
  await expect(
    registry.planPublication({ output: second, inputs: [], overwrite: true }),
  ).rejects.toMatchObject({ code: "DB_NATIVE_FILE_BUSY" });
  await opened.close();
});

test("keeps an exact path busy after its inode is externally replaced", async () => {
  const { first } = await paths();
  let inode = 20n;
  const registry = new NativeFileRegistry(async () => ({
    pathKey: "C:\\Data\\workspace.sqlite",
    device: 10n,
    inode,
  }));
  const opened = handle();
  await registry.open(first, async () => opened);
  inode = 30n;
  await expect(
    registry.planPublication({ output: first, inputs: [], overwrite: true }),
  ).rejects.toMatchObject({ code: "DB_NATIVE_FILE_BUSY" });
  await opened.close();
});

test("keeps unresolved case-distinct leaf names separate", async () => {
  const { first, second } = await paths();
  const identities = new Map([
    [path.resolve(first), { pathKey: "C:\\Sensitive\\Case.sqlite" }],
    [path.resolve(second), { pathKey: "C:\\Sensitive\\case.sqlite" }],
  ]);
  const registry = new NativeFileRegistry(async (filePath) => {
    const identity = identities.get(path.resolve(filePath));
    if (identity === undefined) throw new Error("Missing test identity");
    return identity;
  });
  const opened = handle();
  await registry.open(first, async () => opened);
  await expect(
    registry.planPublication({ output: second, inputs: [], overwrite: true }),
  ).resolves.toMatchObject({ output: path.resolve(second) });
  await opened.close();
});

test("blocks replacement after registration cleanup fails until the handle closes", async () => {
  const { first, second } = await paths();
  const registrationFailure = new Error("Synthetic identity lookup failure");
  let failIdentity = true;
  const registry = new NativeFileRegistry(async (filePath) => {
    if (failIdentity) throw registrationFailure;
    return { pathKey: path.resolve(filePath) };
  });
  let open = true;
  let failClose = true;
  const opened = {
    get isOpen() {
      return open;
    },
    async close() {
      if (failClose) throw undefined;
      open = false;
    },
  };
  try {
    await expect(
      registry.open(first, async () => opened),
    ).rejects.toMatchObject({
      code: "DB_NATIVE_HANDLE_CLEANUP_REQUIRED",
      details: { path: path.resolve(first) },
      cause: { errors: [registrationFailure, undefined] },
    });
    failIdentity = false;
    await expect(
      registry.planPublication({
        output: second,
        inputs: [],
        overwrite: true,
      }),
    ).rejects.toMatchObject({ code: "DB_NATIVE_HANDLE_CLEANUP_REQUIRED" });
    expect(await readFile(first, "utf8")).toBe("first");
    expect(await readFile(second, "utf8")).toBe("second");
    failClose = false;
    await opened.close();
    await expect(
      registry.planPublication({
        output: second,
        inputs: [],
        overwrite: true,
      }),
    ).resolves.toMatchObject({ output: path.resolve(second) });
  } finally {
    failClose = false;
    await opened.close();
  }
});

test("reports a published output when reopening it fails", async () => {
  const { first, second } = await paths();
  const registry = new NativeFileRegistry();
  const plan = await registry.planPublication({
    output: first,
    inputs: [],
    overwrite: true,
  });
  const failure = new Error("Synthetic reopen failure");
  await expect(
    registry.publishAndOpen({
      temporary: second,
      plan,
      async open() {
        throw failure;
      },
    }),
  ).rejects.toMatchObject({
    code: "DB_NATIVE_PUBLISHED_OPEN_FAILED",
    details: { path: path.resolve(first), published: true },
    cause: failure,
  });
  expect(await readFile(first, "utf8")).toBe("second");
});

test.each([false, true])(
  "preserves publication status when registration fails, failed close=%s",
  async (failClose) => {
    const { first, second } = await paths();
    const registrationFailure = new Error("Synthetic registration failure");
    const cleanupFailure = new Error("Synthetic handle close failure");
    let failIdentity = false;
    const registry = new NativeFileRegistry(async (filePath) => {
      if (failIdentity) throw registrationFailure;
      return { pathKey: path.resolve(filePath) };
    });
    let open = true;
    const opened = {
      get isOpen() {
        return open;
      },
      async close() {
        if (failClose) throw cleanupFailure;
        open = false;
      },
    };
    const plan = await registry.planPublication({
      output: first,
      inputs: [],
      overwrite: true,
    });
    try {
      await expect(
        registry.publishAndOpen({
          temporary: second,
          plan,
          async open() {
            failIdentity = true;
            return opened;
          },
        }),
      ).rejects.toMatchObject({
        code: failClose
          ? "DB_NATIVE_HANDLE_CLEANUP_REQUIRED"
          : "DB_NATIVE_PUBLISHED_OPEN_FAILED",
        details: { path: path.resolve(first), published: true },
        cause: failClose
          ? { errors: [registrationFailure, cleanupFailure] }
          : registrationFailure,
      });
      expect(await readFile(first, "utf8")).toBe("second");
      expect(opened.isOpen).toBe(failClose);
    } finally {
      failClose = false;
      await opened.close();
    }
  },
);
