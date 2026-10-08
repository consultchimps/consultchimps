/**
 * The part of pako 1.0 the package writer uses. JSZip compresses through the
 * same class, which is why the writer uses it too (see `jszip-writer.ts`).
 * pako is CommonJS, so it is imported whole.
 */
declare module "pako" {
  interface Deflate {
    onData: (chunk: Uint8Array) => void;
    push(data: Uint8Array | readonly number[], final: boolean): boolean;
    err: number;
    msg: string;
  }
  const pako: {
    Deflate: new (options: { raw: boolean; level: number }) => Deflate;
  };
  export default pako;
}
