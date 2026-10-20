const SHA256_K = new Uint32Array([
  0x428a2f98, 0x71374491, 0xb5c0fbcf, 0xe9b5dba5, 0x3956c25b, 0x59f111f1, 0x923f82a4, 0xab1c5ed5,
  0xd807aa98, 0x12835b01, 0x243185be, 0x550c7dc3, 0x72be5d74, 0x80deb1fe, 0x9bdc06a7, 0xc19bf174,
  0xe49b69c1, 0xefbe4786, 0x0fc19dc6, 0x240ca1cc, 0x2de92c6f, 0x4a7484aa, 0x5cb0a9dc, 0x76f988da,
  0x983e5152, 0xa831c66d, 0xb00327c8, 0xbf597fc7, 0xc6e00bf3, 0xd5a79147, 0x06ca6351, 0x14292967,
  0x27b70a85, 0x2e1b2138, 0x4d2c6dfc, 0x53380d13, 0x650a7354, 0x766a0abb, 0x81c2c92e, 0x92722c85,
  0xa2bfe8a1, 0xa81a664b, 0xc24b8b70, 0xc76c51a3, 0xd192e819, 0xd6990624, 0xf40e3585, 0x106aa070,
  0x19a4c116, 0x1e376c08, 0x2748774c, 0x34b0bcb5, 0x391c0cb3, 0x4ed8aa4a, 0x5b9cca4f, 0x682e6ff3,
  0x748f82ee, 0x78a5636f, 0x84c87814, 0x8cc70208, 0x90befffa, 0xa4506ceb, 0xbef9a3f7, 0xc67178f2,
]);

/**
 * Synchronous SHA-256 (FIPS 180-4) of `bytes`, returned as the 32-byte digest.
 * For callers that must stay synchronous or run where SubtleCrypto is absent
 * (ordinary HTTP origins); the digest equals `crypto.subtle.digest('SHA-256')`.
 */
export function sha256(bytes: Uint8Array): Uint8Array {
  const h = new Uint32Array([
    0x6a09e667, 0xbb67ae85, 0x3c6ef372, 0xa54ff53a,
    0x510e527f, 0x9b05688c, 0x1f83d9ab, 0x5be0cd19,
  ]);
  const words = new Uint32Array(64);
  const total = Math.ceil((bytes.length + 9) / 64) * 64;
  const bitLength = bytes.length * 8;
  const paddedByte = (index: number): number => index < bytes.length ? bytes[index]
    : index === bytes.length ? 0x80
      : index >= total - 8 ? Math.floor(bitLength / 2 ** ((total - index - 1) * 8)) & 0xff : 0;
  const rotate = (value: number, bits: number) => (value >>> bits) | (value << (32 - bits));
  for (let offset = 0; offset < total; offset += 64) {
    for (let i = 0; i < 16; i++) {
      const start = offset + i * 4;
      words[i] = (paddedByte(start) << 24) | (paddedByte(start + 1) << 16)
        | (paddedByte(start + 2) << 8) | paddedByte(start + 3);
    }
    for (let i = 16; i < 64; i++) {
      const p = words[i - 15];
      const q = words[i - 2];
      const s0 = rotate(p, 7) ^ rotate(p, 18) ^ (p >>> 3);
      const s1 = rotate(q, 17) ^ rotate(q, 19) ^ (q >>> 10);
      words[i] = (words[i - 16] + s0 + words[i - 7] + s1) >>> 0;
    }
    let [a, b, c, d, e, f, g, j] = h;
    for (let i = 0; i < 64; i++) {
      const s1 = rotate(e, 6) ^ rotate(e, 11) ^ rotate(e, 25);
      const choose = (e & f) ^ (~e & g);
      const t1 = (j + s1 + choose + SHA256_K[i] + words[i]) >>> 0;
      const s0 = rotate(a, 2) ^ rotate(a, 13) ^ rotate(a, 22);
      const majority = (a & b) ^ (a & c) ^ (b & c);
      const t2 = (s0 + majority) >>> 0;
      j = g; g = f; f = e; e = (d + t1) >>> 0;
      d = c; c = b; b = a; a = (t1 + t2) >>> 0;
    }
    h[0] = (h[0] + a) >>> 0; h[1] = (h[1] + b) >>> 0;
    h[2] = (h[2] + c) >>> 0; h[3] = (h[3] + d) >>> 0;
    h[4] = (h[4] + e) >>> 0; h[5] = (h[5] + f) >>> 0;
    h[6] = (h[6] + g) >>> 0; h[7] = (h[7] + j) >>> 0;
  }
  const result = new Uint8Array(32);
  const view = new DataView(result.buffer);
  h.forEach((value, index) => view.setUint32(index * 4, value));
  return result;
}
