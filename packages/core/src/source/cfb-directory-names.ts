import { collectCfbFatSectors } from '../errors/cfb-read.js';
import { enumerateDirectoryNames } from '../errors/cfb-sniff.js';

const CFB_SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];

/**
 * [MS-CFB] §2.6 names for an opt-in classifier. This separate entry leaves
 * the ordinary OOXML container sniffer and its allocation path unchanged.
 * A malformed CFB has no trustworthy names; non-CFB bytes return null.
 */
export function cfbDirectoryNames(bytes: Uint8Array): ReadonlySet<string> | null {
  if (bytes.length < 512) return null;
  for (let i = 0; i < CFB_SIGNATURE.length; i++) {
    if (bytes[i] !== CFB_SIGNATURE[i]) return null;
  }
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const sectorShift = view.getUint16(0x1e, true);
  if (sectorShift !== 9 && sectorShift !== 12) return new Set();
  const sectorSize = 1 << sectorShift;
  const fatSectors = collectCfbFatSectors(view, bytes.length, {
    sectorSize,
    numFatSectors: view.getUint32(0x2c, true),
    firstDifatSector: view.getUint32(0x44, true),
    numDifatSectors: view.getUint32(0x48, true),
  });
  if (fatSectors === null) return new Set();
  return enumerateDirectoryNames(
    view, bytes.length, sectorSize, view.getUint32(0x30, true), fatSectors,
  ) ?? new Set();
}
