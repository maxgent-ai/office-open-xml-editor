import type { DeepReadonly, LayoutDiagnostic, SourceRef } from './types.js';
import { stableFingerprint } from './fingerprint.js';
import type { BodyElement, DocParagraph, DocxDocumentModel } from '../types.js';
import type { BodyAcquisitionInputProjections } from './acquisition-input-projections.js';
import type { MathNode } from '@silurus/ooxml-core';
import { rasterExceedsBudget, sniffRasterDimensions } from '@silurus/ooxml-core';
import { mathResourceKey } from './source-key.js';
import { projectDocumentSnapshotResources } from './production-paint-resources.js';

export { imageResourceKey, mathResourceKey } from './source-key.js';

export interface ImageLayoutResource {
  readonly widthPt: number;
  readonly heightPt: number;
  readonly mimeType: string;
}

export interface ImageMetadataRecord extends ImageLayoutResource {
  readonly resourceKey: string;
}

export interface MathLayoutResource {
  readonly resourceKey: string;
  readonly widthEm: number;
  readonly ascentEm: number;
  readonly descentEm: number;
  readonly diagnostics: readonly LayoutDiagnostic[];
  readonly available?: boolean;
}

export interface MathOccurrence {
  readonly nodes: MathNode[];
  readonly display: boolean;
  readonly source: SourceRef;
  readonly resourceKey: string;
}

export interface ImageMetadataService {
  readonly fingerprint: string;
  resolve(resourceKey: string): Readonly<ImageLayoutResource>;
}

export interface MathMetadataService {
  readonly fingerprint: string;
  resolve(resourceKey: string): DeepReadonly<MathLayoutResource>;
}

export function bodyMathOccurrences(
  body: BodyElement[],
  story: SourceRef['story'] = 'body',
  storyInstance = 'body',
): MathOccurrence[] {
  const found: MathOccurrence[] = [];
  const visit = (elements: BodyElement[], prefix: number[] = []): void => {
    elements.forEach((element, elementIndex) => {
      const path = [...prefix, elementIndex];
      if (element.type === 'paragraph') {
        element.runs.forEach((run, runIndex) => {
          if (run.type === 'math') found.push({
            nodes: run.nodes,
            display: run.display,
            source: { story, storyInstance, path: [...path, runIndex] },
            resourceKey: mathResourceKey(
              { story, storyInstance, path: [...path, runIndex] },
              run.display ? 'display' : 'inline',
            ),
          });
        });
      } else if (element.type === 'table') {
        element.rows.forEach((row, rowIndex) => row.cells.forEach((cell, cellIndex) => {
          visit(cell.content as BodyElement[], [...path, rowIndex, cellIndex]);
        }));
      }
    });
  };
  visit(body);
  return found;
}

function finiteNonNegative(value: number, name: string): number {
  if (!Number.isFinite(value) || value < 0) throw new RangeError(`${name} must be finite and non-negative`);
  return value;
}

export function createImageMetadataService(records: readonly ImageMetadataRecord[]): ImageMetadataService {
  const snapshot = [...records]
    .map((record) => Object.freeze({
      resourceKey: record.resourceKey,
      widthPt: finiteNonNegative(record.widthPt, 'widthPt'),
      heightPt: finiteNonNegative(record.heightPt, 'heightPt'),
      mimeType: record.mimeType,
    }))
    .sort((a, b) => a.resourceKey.localeCompare(b.resourceKey));
  const byKey = new Map(snapshot.map(({ resourceKey, ...metadata }) => [resourceKey, Object.freeze(metadata)]));
  if (byKey.size !== snapshot.length) throw new Error('Duplicate image resource key');
  return Object.freeze({
    fingerprint: stableFingerprint('images', snapshot),
    resolve(resourceKey: string): Readonly<ImageLayoutResource> {
      const resource = byKey.get(resourceKey);
      if (!resource) throw new Error(`Unknown image resource: ${resourceKey}`);
      return resource;
    },
  });
}

export function rasterImageMetadataRecord(
  resourceKey: string,
  bytes: Uint8Array,
  mimeType: string,
  dpi: number,
): ImageMetadataRecord {
  const dimensions = sniffRasterDimensions(bytes);
  if (!dimensions || rasterExceedsBudget(dimensions)) {
    throw new Error(`Raster dimensions are unavailable or unsafe for ${resourceKey}`);
  }
  if (!Number.isFinite(dpi) || dpi <= 0) throw new RangeError('dpi must be positive');
  return {
    resourceKey,
    widthPt: dimensions.width * 72 / dpi,
    heightPt: dimensions.height * 72 / dpi,
    mimeType,
  };
}

export function createMathMetadataService(records: readonly MathLayoutResource[]): MathMetadataService {
  const snapshot = [...records]
    .map((record) => Object.freeze({
      resourceKey: record.resourceKey,
      widthEm: finiteNonNegative(record.widthEm, 'widthEm'),
      ascentEm: finiteNonNegative(record.ascentEm, 'ascentEm'),
      descentEm: finiteNonNegative(record.descentEm, 'descentEm'),
      diagnostics: Object.freeze(record.diagnostics.map((diagnostic) => Object.freeze({ ...diagnostic }))),
      ...(record.available === false ? { available: false } : {}),
    }))
    .sort((a, b) => a.resourceKey.localeCompare(b.resourceKey));
  const byKey = new Map(snapshot.map((resource) => [resource.resourceKey, resource]));
  if (byKey.size !== snapshot.length) throw new Error('Duplicate math resource key');
  return Object.freeze({
    fingerprint: stableFingerprint('math', snapshot),
    resolve(resourceKey: string): DeepReadonly<MathLayoutResource> {
      const resource = byKey.get(resourceKey);
      if (!resource) throw new Error(`Unknown math resource: ${resourceKey}`);
      return resource;
    },
  });
}

export type PictureBulletSizeResolver = (
  paragraph: DocParagraph,
) => Readonly<{ widthPt: number; heightPt: number }>;

export function documentImageMetadataRecords(
  doc: DocxDocumentModel,
  resolvePictureBulletSize?: PictureBulletSizeResolver,
  acquisitionInputs?: BodyAcquisitionInputProjections,
): ImageMetadataRecord[] {
  return [...projectDocumentSnapshotResources(
    doc,
    acquisitionInputs,
    [],
    resolvePictureBulletSize,
  ).imageMetadata];
}

/** Atomic upright OMML in a quarter-turn text frame. Library layout policy:
 * §22.1.2.77 keeps the math zone intact; the rendered height advances along
 * the column, and rendered width reserves its cross extent. Paint already
 * counter-rotates resources and swaps these retained logical axes back. */
export function quarterTurnMathMetadataService(math: MathMetadataService): MathMetadataService {
  return Object.freeze({
    fingerprint: stableFingerprint('quarter-turn-math', math.fingerprint),
    resolve(resourceKey: string): DeepReadonly<MathLayoutResource> {
      const resource = math.resolve(resourceKey);
      return Object.freeze({ ...resource, widthEm: resource.ascentEm + resource.descentEm,
        ascentEm: resource.widthEm / 2, descentEm: resource.widthEm / 2 });
    },
  });
}
