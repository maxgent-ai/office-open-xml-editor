import type {
  NoteSeparatorInput,
  SelectedNoteSeparatorDefinitionInput,
  SelectedNoteSeparatorParagraphsInput,
  SelectedNoteSeparatorRole,
} from './body-layout-input.js';
import type { ParagraphAcquisitionInput } from './text.js';
import type { DeepReadonly, SourceRef } from './types.js';
import type { DocParagraph } from '../types.js';

/**
 * Ordinary DOCX selected footnote separator stories of the formatted class:
 * a document-listed (ECMA-376 §17.11.9) `separator` or `continuationSeparator`
 * story whose one paragraph formats its single §17.11.23/§17.11.1 mark run.
 *
 * Identity: one canonical `selected:*` root per role inside the `footnote`
 * story kind. These are library identities, distinct from native MS-DOC
 * `reserved:*` stories and never DOCX note ids. A root is shared by every
 * page occurrence and never joins note numbering.
 *
 * Geometry: the story paragraph's acquired advance replaces only the scalar
 * band height of its role; the mark (Short/Full) and rule ink keep the
 * existing scalar policy. Bare, missing and endnote stories keep the scalar
 * band. The parser admits only a mark run whose rPr equals its paragraph
 * mark's, so the run-free paragraph line box is the mark's line box.
 */

const SELECTED_INSTANCES: Readonly<Record<SelectedNoteSeparatorRole, string>> = Object.freeze({
  separator: 'selected:separator',
  continuationSeparator: 'selected:continuation-separator',
});
const ROLES = Object.freeze(Object.keys(SELECTED_INSTANCES) as SelectedNoteSeparatorRole[]);
const SELECTED_ROLE_BY_INSTANCE: ReadonlyMap<string, SelectedNoteSeparatorRole> = new Map(
  ROLES.map((role) => [SELECTED_INSTANCES[role], role]),
);

export type SelectedNoteSeparatorParagraphAcquirer = (
  paragraph: DeepReadonly<DocParagraph>,
  source: SourceRef,
) => ParagraphAcquisitionInput;

export interface NormalizedSelectedNoteSeparators {
  readonly definitions: Readonly<Partial<Record<SelectedNoteSeparatorRole, SelectedNoteSeparatorDefinitionInput>>>;
  readonly stories: readonly Readonly<{
    source: SourceRef;
    body: readonly ParagraphAcquisitionInput[];
  }>[];
}

/** The one identity factory for selected roots and their paragraph paths. */
export function selectedNoteSeparatorSource(
  role: SelectedNoteSeparatorRole,
  path: readonly number[] = [],
): SourceRef {
  return Object.freeze({
    story: 'footnote',
    storyInstance: SELECTED_INSTANCES[role],
    path: Object.freeze([...path]),
  });
}

export function selectedNoteSeparatorRole(
  source: Pick<SourceRef, 'story' | 'storyInstance'>,
): SelectedNoteSeparatorRole | undefined {
  return source.story === 'footnote' ? SELECTED_ROLE_BY_INSTANCE.get(source.storyInstance) : undefined;
}

export function isSelectedNoteSeparatorInstance(storyInstance: string): boolean {
  return SELECTED_ROLE_BY_INSTANCE.has(storyInstance);
}

function reject(reason: string): never {
  throw new Error(`Unsupported selected note separator story: ${reason}`);
}

/** Validate and normalize the parser-boundary paragraphs once, before the
 * layout source is sealed. Returns at most two definitions. */
export function normalizeSelectedNoteSeparators(
  raw: SelectedNoteSeparatorParagraphsInput | undefined,
  marks: Readonly<Record<SelectedNoteSeparatorRole, NoteSeparatorInput | undefined>>,
  acquire: SelectedNoteSeparatorParagraphAcquirer,
): NormalizedSelectedNoteSeparators | undefined {
  if (raw === undefined) return undefined;
  const definitions: Partial<Record<SelectedNoteSeparatorRole, SelectedNoteSeparatorDefinitionInput>> = {};
  const stories: { source: SourceRef; body: readonly ParagraphAcquisitionInput[] }[] = [];
  for (const key of Object.keys(raw)) {
    if (!(ROLES as readonly string[]).includes(key)) reject(`unknown role ${key}`);
  }
  for (const role of ROLES) {
    const paragraph = raw[role];
    if (paragraph === undefined) continue;
    const mark = marks[role];
    if (mark !== 'short' && mark !== 'full') reject(`${role} paragraph has no Short/Full mark`);
    // The closed shape the parser admits: no runs besides the omitted mark,
    // and no paragraph ink, numbering, frame or hidden mark.
    if ((paragraph.runs ?? []).length !== 0) reject(`${role} paragraph carries runs`);
    if (paragraph.numbering != null || paragraph.framePr != null || paragraph.borders != null
      || paragraph.shading != null || paragraph.markVanish === true) {
      reject(`${role} paragraph carries unrepresented properties`);
    }
    const root = selectedNoteSeparatorSource(role);
    const paragraphSource = selectedNoteSeparatorSource(role, [0]);
    definitions[role] = Object.freeze({ role, root, paragraph: paragraphSource });
    stories.push(Object.freeze({ source: root, body: Object.freeze([acquire(paragraph, paragraphSource)]) }));
  }
  return Object.freeze({ definitions: Object.freeze(definitions), stories: Object.freeze(stories) });
}
