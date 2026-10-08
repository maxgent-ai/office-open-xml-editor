import type {
  NativeNoteSeparatorCharacterInput,
  NativeNoteSeparatorDefinitionInput,
  NativeNoteSeparatorDefinitionsInput,
  NativeNoteSeparatorRole,
  NativeNoteSeparatorRolesInput,
  NativeNoteSeparatorSourceInput,
  NativeNoteSeparatorStoryInput,
  NativeNoteSeparatorsInput,
} from './body-layout-input.js';
import { projectBodyOccurrence } from './occurrence-projection.js';
import { translateBorder, translateRect } from './retained-geometry-translation.js';
import type { ParagraphAcquisitionInput } from './text.js';
import type {
  DeepReadonly,
  NoteLayout,
  NoteSeparatorLayout,
  PaintNode,
  SourceRef,
  StoryLayout,
} from './types.js';
import type { DocParagraph, DocxTextRun } from '../types.js';

/**
 * Native (MS-DOC 2.3.3) reserved note separator stories: the separator, the
 * continuation separator and the continuation notice of each note kind.
 *
 * Identity: each definition is one canonical source root per note kind and
 * role, `reserved:*` inside the existing `footnote`/`endnote` story kinds.
 * These are library identities, never DOCX special note ids or native CP/FC;
 * a numbered note using one of them is rejected rather than shadowed. A root
 * is shared by every page occurrence and never joins note numbering.
 *
 * Validation covers only what the wire represents. The native producer owns
 * effective sprmCFSpec (MS-DOC 2.6.1) and the byte-level story shape; this
 * consumer cannot reread native bytes. It checks closed class/range/address
 * consistency and refuses facts it cannot represent (presence-only frame or
 * numbering cascades, revision-marked characters, authored runs) before
 * layout, instead of approximating them.
 *
 * Page dependence: an admitted story holds no runs other than its control and
 * content-mark metric participants, so its occurrences carry no fields,
 * drawings or text boxes. That closed shape is why note reuse walks need not
 * visit separator occurrences for page-dependent content.
 */

const RESERVED_INSTANCES: Readonly<Record<NativeNoteSeparatorRole, string>> = Object.freeze({
  separator: 'reserved:separator',
  continuationSeparator: 'reserved:continuation-separator',
  continuationNotice: 'reserved:continuation-notice',
});
const ROLES = Object.freeze(Object.keys(RESERVED_INSTANCES) as NativeNoteSeparatorRole[]);
const RESERVED_ROLE_BY_INSTANCE: ReadonlyMap<string, NativeNoteSeparatorRole> = new Map(
  ROLES.map((role) => [RESERVED_INSTANCES[role], role]),
);
const MAX_PRM = 0xffff;

export type NoteSeparatorParticipantRole = 'rule-control' | 'paragraph-mark';

export interface NoteSeparatorParticipant {
  readonly role: NoteSeparatorParticipantRole;
  readonly run: DeepReadonly<DocxTextRun>;
}

/** Builds the immutable paragraph acquisition input through the parser-model
 * boundary; this module never imports the parser projection itself. */
export type NativeNoteSeparatorParagraphAcquirer = (
  paragraph: DeepReadonly<DocParagraph>,
  participants: readonly NoteSeparatorParticipant[],
  source: SourceRef,
) => ParagraphAcquisitionInput;

export interface NormalizedNativeNoteSeparators {
  readonly roles: NativeNoteSeparatorRolesInput;
  readonly stories: readonly Readonly<{
    source: SourceRef;
    body: readonly ParagraphAcquisitionInput[];
  }>[];
}

/** The one identity factory for reserved roots and their structural paths. */
export function reservedNoteSeparatorSource(
  kind: 'footnote' | 'endnote',
  role: NativeNoteSeparatorRole,
  path: readonly number[] = [],
): SourceRef {
  return Object.freeze({
    story: kind,
    storyInstance: RESERVED_INSTANCES[role],
    path: Object.freeze([...path]),
  });
}

export function reservedNoteSeparatorRole(
  source: Pick<SourceRef, 'story' | 'storyInstance'>,
): NativeNoteSeparatorRole | undefined {
  if (source.story !== 'footnote' && source.story !== 'endnote') return undefined;
  return RESERVED_ROLE_BY_INSTANCE.get(source.storyInstance);
}

export function isReservedNoteSeparatorInstance(storyInstance: string): boolean {
  return RESERVED_ROLE_BY_INSTANCE.has(storyInstance);
}

function reject(reason: string): never {
  throw new Error(`Unsupported native note separator story: ${reason}`);
}

function count(value: unknown, label: string): number {
  if (typeof value !== 'number' || !Number.isSafeInteger(value) || value < 0) {
    reject(`${label} is not a non-negative safe integer`);
  }
  return value;
}

function validateAddress(source: NativeNoteSeparatorSourceInput | undefined, label: string): number {
  if (!source) reject(`${label} source is missing`);
  count(source.fc, `${label} FC`);
  count(source.paragraphStyle, `${label} paragraph style`);
  if (count(source.prm, `${label} PRM`) > MAX_PRM) reject(`${label} PRM exceeds u16`);
  return count(source.headerCp, `${label} CP`);
}

function validateCharacter(
  character: NativeNoteSeparatorCharacterInput | undefined,
  label: string,
): void {
  if (!character) reject(`${label} is missing`);
  // sprmCFRMark revision metadata is not retained by the wire.
  if (character.insertion === true) reject(`${label} carries an unretained tracked insertion`);
  // MS-DOC §2.6.1 sprmCSymbol can replace even the producer's empty
  // character payload with a glyph. This closed rule/mark contract has only
  // metric participants; authored glyphs need a separate content contract.
  // Reject them before normalization can erase text or omit a notice mark.
  if (character.run && character.run.text !== '') {
    reject(`${label} carries unsupported authored text`);
  }
}

function validateParagraph(story: NativeNoteSeparatorStoryInput, markCp: number): void {
  const authored = story.paragraph;
  if (!authored) reject(`${story.class} story has no paragraph`);
  if (validateAddress(authored.source, 'paragraph mark') !== markCp) {
    reject('paragraph mark is not the last content CP');
  }
  const paragraph = authored.paragraph;
  // Presence-only flags: the owned frame/numbering cascades are absent.
  if (authored.numbered === true || authored.framed === true
    || paragraph.numbering != null || paragraph.framePr != null) {
    reject('numbered or framed paragraph cascades are not represented');
  }
  if ((paragraph.runs ?? []).length !== 0) reject('authored paragraph runs are not represented');
  validateCharacter(authored.contentMark, 'content paragraph mark');
  // No mark run is an effective vanish of the same CHPX as the paragraph mark.
  if ((authored.contentMark!.run === undefined) !== (paragraph.markVanish === true)) {
    reject('content mark visibility disagrees with its paragraph');
  }
}

/** Closed structural validation of one retained native story. */
export function validateNativeNoteSeparatorStory(story: NativeNoteSeparatorStoryInput): void {
  const start = count(story.contentStartCp, 'content start');
  const end = count(story.contentEndCp, 'content end');
  if (end < start) reject('content range is reversed');
  const guard = story.guardCp === undefined ? undefined : count(story.guardCp, 'guard');
  if (guard !== undefined && guard !== end) reject('guard is not the terminal story mark');
  if (story.class !== 'rule' && story.rule !== undefined) reject(`${story.class} story has a rule`);
  switch (story.class) {
    case 'empty':
      if (end !== start || guard !== undefined || story.paragraph) reject('empty story has content');
      return;
    case 'guardOnly':
      if (end !== start || guard === undefined || story.paragraph) reject('guard-only story has content');
      return;
    case 'paragraphOnly':
      if (end - start !== 1 || guard === undefined) reject('paragraph-only story is not one mark before its guard');
      validateParagraph(story, end - 1);
      return;
    case 'rule': {
      // Whole rule only: control, content paragraph mark, guard. Partial
      // rules and other shapes are a separate unsupported class.
      if (end - start !== 2 || guard === undefined) reject('rule story is not control, mark and guard');
      const rule = story.rule;
      if (!rule) reject('rule story has no control');
      if (rule.mark !== 'short' && rule.mark !== 'full') reject('rule control lacks a special-character mark');
      if (validateAddress(rule.source, 'rule control') !== start) reject('rule control is not the first content CP');
      validateCharacter(rule.control, 'rule control');
      if (!rule.control.run) reject('rule control is hidden');
      validateParagraph(story, end - 1);
      if (rule.source.paragraphStyle !== story.paragraph!.source.paragraphStyle) {
        reject('rule control and paragraph mark disagree on their paragraph style');
      }
      return;
    }
    default:
      reject('unknown story class');
  }
}

function explicitColor(run: DeepReadonly<DocxTextRun>): string | undefined {
  // ECMA-376 §17.3.2.6: an explicit hex colour; auto/absent keeps the
  // existing library rule ink colour.
  return typeof run.color === 'string' && /^[0-9A-Fa-f]{6}$/.test(run.color)
    ? `#${run.color}` : undefined;
}

function normalizeStory(
  kind: 'footnote' | 'endnote',
  role: NativeNoteSeparatorRole,
  story: NativeNoteSeparatorStoryInput,
  acquire: NativeNoteSeparatorParagraphAcquirer,
): Readonly<{ definition: NativeNoteSeparatorDefinitionInput; body: readonly ParagraphAcquisitionInput[] }> {
  validateNativeNoteSeparatorStory(story);
  const root = reservedNoteSeparatorSource(kind, role);
  const authored = story.paragraph;
  if (!authored) return Object.freeze({ definition: Object.freeze({ role, root }), body: Object.freeze([]) });
  const paragraphSource = reservedNoteSeparatorSource(kind, role, [0]);
  // CP order: the rule control precedes the content paragraph mark. A
  // visible content mark is a participant in every class, so its complete
  // effective CHPX (including §17.3.2.24 position and §17.3.2.42 vertAlign,
  // which paragraph-mark font facts do not carry) reaches its line box. An
  // effectively vanished mark has no run: a paragraph-only story is then
  // the shared fully hidden paragraph, and a rule keeps only its control.
  const participants: NoteSeparatorParticipant[] = [];
  if (story.rule?.control.run) participants.push({ role: 'rule-control', run: story.rule.control.run });
  if (authored.contentMark?.run) {
    participants.push({ role: 'paragraph-mark', run: authored.contentMark.run });
  }
  const input = acquire(authored.paragraph, Object.freeze(participants), paragraphSource);
  const color = story.rule?.control.run ? explicitColor(story.rule.control.run) : undefined;
  const definition: NativeNoteSeparatorDefinitionInput = Object.freeze({
    role,
    root,
    paragraph: Object.freeze({
      source: paragraphSource,
      hidden: story.class === 'paragraphOnly' && authored.paragraph.markVanish === true,
    }),
    ...(story.rule ? {
      rule: Object.freeze({
        mark: story.rule.mark!,
        control: reservedNoteSeparatorSource(kind, role, [0, 0]),
        ...(color === undefined ? {} : { color }),
      }),
    } : {}),
  });
  return Object.freeze({ definition, body: Object.freeze([input]) });
}

/** Validate and normalize the raw wire once, before the layout source is
 * sealed. Returns at most six definitions; raw native addresses are dropped. */
export function normalizeNativeNoteSeparators(
  raw: NativeNoteSeparatorsInput | undefined,
  acquire: NativeNoteSeparatorParagraphAcquirer,
): NormalizedNativeNoteSeparators | undefined {
  if (raw === undefined) return undefined;
  const stories: { source: SourceRef; body: readonly ParagraphAcquisitionInput[] }[] = [];
  const roles: { footnote?: NativeNoteSeparatorDefinitionsInput; endnote?: NativeNoteSeparatorDefinitionsInput } = {};
  for (const kind of ['footnote', 'endnote'] as const) {
    const kindStories = raw[kind];
    if (!kindStories) continue;
    const definitions = {} as Record<NativeNoteSeparatorRole, NativeNoteSeparatorDefinitionInput>;
    for (const role of ROLES) {
      const normalized = normalizeStory(kind, role, kindStories[role], acquire);
      definitions[role] = normalized.definition;
      stories.push(Object.freeze({ source: normalized.definition.root, body: normalized.body }));
    }
    roles[kind] = Object.freeze(definitions);
  }
  return Object.freeze({ roles: Object.freeze(roles), stories: Object.freeze(stories) });
}

/**
 * Build one occurrence from the acquired story in its acquisition frame.
 *
 * Flow: the story paragraph's acquired advance (lines plus authored spacing)
 * is the occurrence's whole reservation; no scalar band is added beside it.
 *
 * Rule ink, library policy separate from paragraph geometry:
 * - span: Full covers the main-story text extent (ECMA-376 §17.11.1; MS-DOC
 *   2.6.1 U+0004). Short keeps the existing one-third library span; §17.11.23
 *   specifies a partial rule but no fraction. Neither span follows paragraph
 *   indentation or alignment: no Office alignment/indent rule is settled.
 * - placement: centred on the control's acquired line box, as the scalar
 *   path centres its rule in its band. Not an Office baseline/gap observation.
 * - stroke: the existing 0.5pt single rule in the control's effective explicit
 *   colour (character properties of the displayed control), else black.
 */
export function noteSeparatorOccurrence(
  definition: NativeNoteSeparatorDefinitionInput,
  story: StoryLayout,
  inline: Readonly<{ xPt: number; widthPt: number }>,
): NoteSeparatorLayout {
  const [paragraph, ...rest] = story.blocks;
  if (!paragraph || paragraph.kind !== 'paragraph' || rest.length !== 0) {
    throw new Error('A native note separator story must acquire exactly one paragraph');
  }
  const flowBounds = Object.freeze({
    xPt: inline.xPt, yPt: story.flowBounds.yPt, widthPt: inline.widthPt, heightPt: story.advancePt,
  });
  let rule: NoteSeparatorLayout['rule'];
  if (definition.rule) {
    const line = paragraph.lines[0];
    if (!line) throw new Error('A native note separator rule has no acquired control line');
    const yPt = line.bounds.yPt + line.bounds.heightPt / 2;
    rule = Object.freeze({
      mark: definition.rule.mark,
      source: definition.rule.control,
      segment: Object.freeze({
        edge: 'top' as const,
        from: Object.freeze({ xPt: inline.xPt, yPt }),
        to: Object.freeze({
          xPt: inline.xPt + inline.widthPt * (definition.rule.mark === 'full' ? 1 : 1 / 3),
          yPt,
        }),
        color: definition.rule.color ?? '#000000',
        widthPt: 0.5,
        authoredStyle: 'single',
        style: 'solid' as const,
      }),
    });
  }
  return Object.freeze({
    role: definition.role,
    source: definition.root,
    flowBounds,
    advancePt: story.advancePt,
    paragraph,
    ...(rule ? { rule } : {}),
  });
}

/** Give an occurrence its destination page identity, domain and block offset.
 * The canonical definition source is unchanged. */
export function placeNoteSeparatorOccurrence(
  occurrence: NoteSeparatorLayout,
  destination: Readonly<{ occurrenceId: string; flowDomainId: string; yPt: number }>,
): NoteSeparatorLayout {
  const delta = { xPt: 0, yPt: destination.yPt - occurrence.flowBounds.yPt };
  return Object.freeze({
    ...occurrence,
    flowBounds: translateRect(occurrence.flowBounds, delta),
    ...(occurrence.paragraph ? {
      paragraph: projectBodyOccurrence(occurrence.paragraph, {
        occurrenceId: destination.occurrenceId,
        destination: {
          coordinateSpace: 'logical-page-points',
          flowDomainId: destination.flowDomainId,
          translation: delta,
        },
      }),
    } : {}),
    ...(occurrence.rule ? {
      rule: Object.freeze({ ...occurrence.rule, segment: translateBorder(occurrence.rule.segment, delta) }),
    } : {}),
  });
}

/** Retained blocks a note owns, in paint and reading order. */
export function noteOwnedBlocks(note: NoteLayout): readonly PaintNode[] {
  return [
    ...(note.leading?.paragraph ? [note.leading.paragraph] : []),
    ...note.story.blocks,
    ...(note.trailing?.paragraph ? [note.trailing.paragraph] : []),
  ];
}
