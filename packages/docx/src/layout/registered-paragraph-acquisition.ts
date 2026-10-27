import {
  acquireParagraphResult,
  type ParagraphAcquisitionOptions,
} from './paragraph.js';
import {
  accessParagraphWrapRegistry,
  commitOwnedParagraphWrapRegistry,
} from './paragraph-wrap-registry.js';
import type {
  DrawingMLCollisionEntryPt,
  WrapExclusion,
} from './types.js';

export interface InheritedParagraphAuthority {
  readonly exclusions: readonly WrapExclusion[];
  readonly collisions: readonly DrawingMLCollisionEntryPt[];
}

/** Acquire and commit one sequential paragraph as one transaction. Registry
 * authority is read before acquisition and replaced only after the retained
 * paragraph succeeds, so rejected candidates cannot leak wrap geometry. */
export function acquireRegisteredParagraph(
  owner: object,
  input: Parameters<typeof acquireParagraphResult>[0],
  options: ParagraphAcquisitionOptions,
  inherited?: InheritedParagraphAuthority,
  continuation?: Parameters<typeof acquireParagraphResult>[2],
): ReturnType<typeof acquireParagraphResult> {
  const registry = accessParagraphWrapRegistry(owner, options.flowDomainId);
  const acquired = acquireParagraphResult(input, {
    ...options,
    exclusions: Object.freeze([
      ...options.exclusions,
      ...registry.exclusions,
      ...(inherited?.exclusions ?? []),
    ]),
    anchorCollisions: Object.freeze([
      ...(options.anchorCollisions ?? []),
      ...registry.collisions,
      ...(inherited?.collisions ?? []),
    ]),
  }, continuation);
  commitOwnedParagraphWrapRegistry(owner, registry, acquired.layout);
  return acquired;
}
