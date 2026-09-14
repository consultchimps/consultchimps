import { sha256 } from "@noble/hashes/sha2.js";
import { bytesToHex } from "@noble/hashes/utils.js";

const MAX_SOURCE_ROW = BigInt(Number.MAX_SAFE_INTEGER);
const ROW_FRAME_BYTES = 17;
const textEncoder = new TextEncoder();
const DOMAIN = textEncoder.encode("consultchimps.capture-rows.sha256.v1\0");

interface CaptureRowChecksum {
  update(sourceRow: bigint, valuesJson: string): void;
  digest(): string;
}

export function createCaptureRowChecksum(): CaptureRowChecksum {
  const hash = sha256.create().update(DOMAIN);
  let digest: string | undefined;

  return {
    update(sourceRow, valuesJson) {
      if (digest !== undefined) {
        throw new Error("The capture row checksum is already finalized.");
      }
      if (sourceRow < 1n || sourceRow > MAX_SOURCE_ROW) {
        throw new RangeError(
          "The capture row checksum requires a positive safe source row.",
        );
      }
      const values = textEncoder.encode(valuesJson);
      const frame = new Uint8Array(ROW_FRAME_BYTES);
      const view = new DataView(frame.buffer);
      frame[0] = 1;
      view.setBigUint64(1, sourceRow);
      view.setBigUint64(9, BigInt(values.byteLength));
      hash.update(frame);
      hash.update(values);
    },
    digest() {
      digest ??= bytesToHex(hash.digest());
      return digest;
    },
  };
}
