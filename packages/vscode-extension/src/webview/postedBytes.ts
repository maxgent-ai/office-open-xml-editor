/**
 * The document bytes posted in `ooxml-init`, as an exact-length ArrayBuffer.
 * The webview message channel revives a typed array as a view that may cover
 * only part of its underlying buffer, so copy in that case.
 */
export function exactArrayBuffer(data: unknown): ArrayBuffer | null {
  if (data instanceof ArrayBuffer) return data;
  if (!ArrayBuffer.isView(data)) return null;
  const bytes = new Uint8Array(data.buffer, data.byteOffset, data.byteLength);
  return bytes.byteOffset === 0 && bytes.byteLength === bytes.buffer.byteLength
    ? (bytes.buffer as ArrayBuffer)
    : (bytes.slice().buffer as ArrayBuffer);
}
