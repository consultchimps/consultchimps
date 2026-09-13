import type { RandomAccessFile } from "@consultchimps/core";
import type { FileEntry } from "@zip.js/zip.js";
import { SaxesParser } from "saxes";

import { BoundedXmlText, localName } from "./xml.js";
import type { ScratchFactory, WorkbookStreamOptions } from "./types.js";
import {
  forEachEntryChunk,
  type StreamLimits,
  WorkbookOperationCleanupFailure,
} from "./zip.js";

export class SharedStrings {
  #closed = false;
  #closeAttempt: Promise<void> | undefined;
  #closeRequested = false;
  #indexClosed = false;
  #payloadClosed = false;
  #count = 0;
  #payloadBytes = 0;
  readonly #decoder = new TextDecoder();
  readonly #encoder = new TextEncoder();

  private constructor(
    private readonly payload: RandomAccessFile,
    private readonly index: RandomAccessFile,
    private readonly limits: StreamLimits,
  ) {}

  static async create(
    scratch: ScratchFactory,
    limits: StreamLimits,
    signal: AbortSignal | undefined,
  ): Promise<SharedStrings> {
    const createOptions: Parameters<ScratchFactory["create"]>[0] = {
      purpose: "xlsx-shared-strings",
      ...(signal === undefined ? {} : { signal }),
    };
    const payload = await scratch.create(createOptions);
    let index: RandomAccessFile | undefined;
    try {
      index = await scratch.create(createOptions);
      await Promise.all([payload.truncate(0), index.truncate(0)]);
      return new SharedStrings(payload, index, limits);
    } catch (cause) {
      const cleanupResults = await Promise.allSettled([
        Promise.resolve().then(() => payload.close()),
        ...(index ? [Promise.resolve().then(() => index?.close())] : []),
      ]);
      const cleanupFailures = cleanupResults.flatMap((result) =>
        result.status === "rejected" ? [result.reason] : [],
      );
      if (cleanupFailures.length > 0) {
        throw new WorkbookOperationCleanupFailure(
          cause,
          cleanupFailures.length === 1
            ? cleanupFailures[0]
            : new AggregateError(
                cleanupFailures,
                "More than one shared-string scratch file could not be closed after initialization failed.",
              ),
        );
      }
      throw cause;
    }
  }

  get count(): number {
    return this.#count;
  }

  async append(value: string): Promise<void> {
    if (this.#closeRequested) {
      throw new Error("The shared-string store is closing.");
    }
    if (this.#count >= this.limits.maximumSharedStrings) {
      throw new Error(
        `The workbook has more than ${this.limits.maximumSharedStrings} shared strings.`,
      );
    }
    const bytes = this.#encoder.encode(value);
    if (bytes.byteLength > this.limits.maximumCellBytes) {
      throw new Error(
        `A shared string exceeds the configured ${this.limits.maximumCellBytes}-byte cell limit.`,
      );
    }
    if (!Number.isSafeInteger(this.#payloadBytes + bytes.byteLength)) {
      throw new Error("The shared-string scratch offset is too large.");
    }
    await this.payload.writeAt(this.#payloadBytes, bytes);
    const record = new Uint8Array(12);
    const view = new DataView(record.buffer);
    view.setFloat64(0, this.#payloadBytes, true);
    view.setUint32(8, bytes.byteLength, true);
    await this.index.writeAt(this.#count * 12, record);
    this.#payloadBytes += bytes.byteLength;
    this.#count += 1;
  }

  async value(index: number): Promise<string> {
    if (this.#closeRequested) {
      throw new Error("The shared-string store is closing.");
    }
    if (!Number.isSafeInteger(index) || index < 0 || index >= this.#count) {
      throw new Error(
        `Shared-string index ${index} is outside the workbook table of ${this.#count} strings.`,
      );
    }
    const record = await this.index.readAt(index * 12, 12);
    if (record.byteLength !== 12) {
      throw new Error("The shared-string scratch index returned a short read.");
    }
    const view = new DataView(
      record.buffer,
      record.byteOffset,
      record.byteLength,
    );
    const offset = view.getFloat64(0, true);
    const length = view.getUint32(8, true);
    if (!Number.isSafeInteger(offset) || offset < 0) {
      throw new Error(
        "The shared-string scratch index contains an invalid offset.",
      );
    }
    const bytes = await this.payload.readAt(offset, length);
    if (bytes.byteLength !== length) {
      throw new Error(
        "The shared-string scratch payload returned a short read.",
      );
    }
    return this.#decoder.decode(bytes);
  }

  close(): Promise<void> {
    if (this.#closeAttempt) return this.#closeAttempt;
    if (this.#closed) return Promise.resolve();
    this.#closeRequested = true;
    const attempt = this.#closeRemaining();
    this.#closeAttempt = attempt;
    void attempt.then(
      () => {
        if (this.#closeAttempt === attempt) this.#closeAttempt = undefined;
      },
      () => {
        if (this.#closeAttempt === attempt) this.#closeAttempt = undefined;
      },
    );
    return attempt;
  }

  async #closeRemaining(): Promise<void> {
    const stages = [
      ...(!this.#payloadClosed
        ? [
            {
              close: () => this.payload.close(),
              complete: () => {
                this.#payloadClosed = true;
              },
            },
          ]
        : []),
      ...(!this.#indexClosed
        ? [
            {
              close: () => this.index.close(),
              complete: () => {
                this.#indexClosed = true;
              },
            },
          ]
        : []),
    ];
    const results = await Promise.allSettled(
      stages.map(async ({ close }) => close()),
    );
    const failures: unknown[] = [];
    results.forEach((result, index) => {
      if (result.status === "fulfilled") stages[index]?.complete();
      else failures.push(result.reason);
    });
    if (failures.length === 1) throw failures[0];
    if (failures.length > 1) {
      throw new AggregateError(
        failures,
        "More than one shared-string scratch file could not be closed.",
      );
    }
    this.#closed = this.#payloadClosed && this.#indexClosed;
  }
}

export async function loadSharedStrings(
  entry: FileEntry | undefined,
  options: WorkbookStreamOptions,
  limits: StreamLimits,
): Promise<SharedStrings> {
  const store = await SharedStrings.create(
    options.scratch,
    limits,
    options.signal,
  );
  if (!entry) return store;
  const parser = new SaxesParser();
  const decoder = new TextDecoder("utf-8", { fatal: true });
  const pending: string[] = [];
  let inItem = false;
  let inText = false;
  let depth = 0;
  let rootName: string | undefined;
  let phoneticDepth = 0;
  let value = "";
  let valueBytes = 0;
  const encoder = new TextEncoder();
  const textLimit = new BoundedXmlText(
    new Set(["t"]),
    limits.maximumCellBytes,
    () =>
      `A shared string exceeds the configured ${limits.maximumCellBytes}-byte cell limit.`,
  );
  parser.on("opentag", (tag) => {
    const name = localName(tag.name);
    if (depth === 0) rootName = name;
    if (name === "si") {
      if (inItem || depth !== 1 || rootName !== "sst") {
        throw new Error(
          "A shared-string item must be a direct child of the shared-string table and cannot contain another item.",
        );
      }
      inItem = true;
      inText = false;
      phoneticDepth = 0;
      value = "";
      valueBytes = 0;
    } else if (inItem && name === "rPh") {
      phoneticDepth += 1;
    } else if (inItem && name === "t" && phoneticDepth === 0) {
      inText = true;
    }
    depth += 1;
  });
  const appendText = (text: string) => {
    if (!inText) return;
    valueBytes += encoder.encode(text).byteLength;
    if (valueBytes > limits.maximumCellBytes) {
      throw new Error(
        `A shared string exceeds the configured ${limits.maximumCellBytes}-byte cell limit.`,
      );
    }
    value += text;
  };
  parser.on("text", appendText);
  parser.on("cdata", appendText);
  parser.on("closetag", (tag) => {
    const name = localName(tag.name);
    if (name === "t") inText = false;
    else if (name === "rPh") phoneticDepth -= 1;
    else if (name === "si") {
      inItem = false;
      pending.push(value);
    }
    depth -= 1;
  });
  try {
    await forEachEntryChunk(entry, {
      signal: options.signal,
      onProgress: options.onProgress,
      stage: "shared-strings",
      async consume(chunk) {
        textLimit.consume(chunk);
        parser.write(decoder.decode(chunk, { stream: true }));
        while (pending.length > 0) {
          const next = pending.shift();
          if (next !== undefined) await store.append(next);
        }
      },
    });
    parser.write(decoder.decode());
    parser.close();
    while (pending.length > 0) {
      const next = pending.shift();
      if (next !== undefined) await store.append(next);
    }
    return store;
  } catch (cause) {
    try {
      await store.close();
    } catch (cleanupCause) {
      throw new WorkbookOperationCleanupFailure(cause, cleanupCause);
    }
    throw cause;
  }
}
