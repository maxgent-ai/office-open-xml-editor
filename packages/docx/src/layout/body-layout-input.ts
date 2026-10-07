import type { SectionLayoutContext } from '../layout-context.js';
import type {
  DocParagraph,
  DocxTextRun,
  HeadersFooters,
  PageBorders,
  SectionProps,
} from '../types.js';
import {
  resolveAcquiredSectionLayoutContext,
  type BodySectionIndexInput,
  type BodySectionOccurrence,
  type SectionPageLayoutPolicy,
} from './context.js';
import type { SectionStartType, AuthoredBreak } from './paginator.js';
import { snapshotPlainData } from './plain-data.js';
import type { DeepReadonly, LayoutDiagnostic, SourceRef } from './types.js';
import { wordContinuousSectionRole } from './body-pagination-compatibility.js';

/** ECMA-376 §17.11.1/.23 and MS-DOC §2.6.1 sprmCFSpec define the
 * mark kind independently of the separator story's ordinary/continuation role.
 * Default is an absent producer fact, not an explicitly authored mark. Native
 * producers must validate control semantics/formatting before emitting a mode. */
export type NoteSeparatorMark = 'short' | 'full' | 'none';
export type NoteSeparatorInput = NoteSeparatorMark | 'default';

/** Native (MS-DOC 2.3.3) reserved separator story shape. The guard mark is
 * never authored content. */
export type NativeNoteSeparatorStoryClass = 'empty' | 'guardOnly' | 'paragraphOnly' | 'rule';

/** Native-local header CP/FC/PRM/paragraph-style address; not a SourceRef. */
export interface NativeNoteSeparatorSourceInput {
  readonly headerCp: number;
  readonly fc: number;
  readonly prm: number;
  readonly paragraphStyle: number;
}

/** Effective supported CHPX of one character. No run means an effective
 * vanish, which differs from an absent character. Parser-private run keys
 * (for example typography acquisition) are retained verbatim. */
export interface NativeNoteSeparatorCharacterInput {
  readonly run?: DeepReadonly<DocxTextRun>;
  readonly insertion?: boolean;
}

export interface NativeNoteSeparatorStoryInput {
  readonly class: NativeNoteSeparatorStoryClass;
  readonly contentStartCp: number;
  readonly contentEndCp: number;
  readonly guardCp?: number;
  readonly rule?: Readonly<{
    /** Present only when effective sprmCFSpec gives the control its meaning. */
    mark?: Exclude<NoteSeparatorMark, 'none'>;
    control: NativeNoteSeparatorCharacterInput;
    source: NativeNoteSeparatorSourceInput;
  }>;
  readonly paragraph?: Readonly<{
    paragraph: DeepReadonly<DocParagraph>;
    /** Absent when the story authors no paragraph mark (a partial rule). */
    contentMark?: NativeNoteSeparatorCharacterInput;
    /** Presence only; numbering/frame cascades are not retained. */
    numbered?: boolean;
    framed?: boolean;
    source: NativeNoteSeparatorSourceInput;
  }>;
}

export interface NativeNoteSeparatorStoriesInput {
  readonly separator: NativeNoteSeparatorStoryInput;
  readonly continuationSeparator: NativeNoteSeparatorStoryInput;
  readonly continuationNotice: NativeNoteSeparatorStoryInput;
}

/** Raw parser-boundary facts retained once per document for the separator
 * stories a native producer admitted; absent for DOCX. The source-model
 * adapter replaces them with {@link NativeNoteSeparatorRolesInput} before the
 * layout source is sealed, so layout never reads native CP/FC/PAPX wire. */
export interface NativeNoteSeparatorsInput {
  readonly footnote?: NativeNoteSeparatorStoriesInput;
  readonly endnote?: NativeNoteSeparatorStoriesInput;
}

export type NativeNoteSeparatorRole = keyof NativeNoteSeparatorStoriesInput;

/** Canonical definition of one reserved separator story: a source definition
 * shared by every page occurrence, never a numbered note. */
export interface NativeNoteSeparatorDefinitionInput {
  readonly role: NativeNoteSeparatorRole;
  /** Closed reserved story root (`footnote`/`endnote` + `reserved:*`). */
  readonly root: SourceRef;
  /** The story's authored paragraph `[0]`. `hidden` follows the shared
   * §17.3.1.29/§17.3.2.41 policy for an inkless paragraph whose mark vanishes:
   * it stays source-addressable but owns neither flow nor ink. */
  readonly paragraph?: Readonly<{ source: SourceRef; hidden: boolean }>;
  /** The U+0003/U+0004 control at `[0, 0]`; its mark selects Short/Full. */
  readonly rule?: Readonly<{
    mark: Exclude<NoteSeparatorMark, 'none'>;
    control: SourceRef;
    /** Effective explicit control colour (`#RRGGBB`), absent for auto. */
    color?: string;
  }>;
}

export type NativeNoteSeparatorDefinitionsInput = Readonly<
  Record<NativeNoteSeparatorRole, NativeNoteSeparatorDefinitionInput>
>;

/** At most six canonical definitions per document (two note kinds, three roles). */
export interface NativeNoteSeparatorRolesInput {
  readonly footnote?: NativeNoteSeparatorDefinitionsInput;
  readonly endnote?: NativeNoteSeparatorDefinitionsInput;
}

/** Ordinary DOCX listed footnote story roles that can retain a formatted
 * paragraph. DOCX continuation notices keep no story geometry. */
export type SelectedNoteSeparatorRole = 'separator' | 'continuationSeparator';

/** Raw parser-boundary effective paragraphs of document-listed (§17.11.9)
 * formatted footnote separator stories. The source-model adapter replaces them
 * with {@link SelectedNoteSeparatorDefinitionsInput} before sealing. */
export type SelectedNoteSeparatorParagraphsInput = Readonly<
  Partial<Record<SelectedNoteSeparatorRole, DeepReadonly<DocParagraph>>>
>;

/** Canonical definition of one ordinary DOCX selected story: a `selected:*`
 * root (never native `reserved:*` provenance or a DOCX note id) shared by
 * every page occurrence, and its single paragraph `[0]`. Its mark stays the
 * role's existing Short/Full separator mode. */
export interface SelectedNoteSeparatorDefinitionInput {
  readonly role: SelectedNoteSeparatorRole;
  readonly root: SourceRef;
  readonly paragraph: SourceRef;
}

export type SelectedNoteSeparatorDefinitionsInput = Readonly<
  Partial<Record<SelectedNoteSeparatorRole, SelectedNoteSeparatorDefinitionInput>>
>;

export interface BodyParagraphSourceInput {
  readonly kind: 'paragraph';
  readonly source: SourceRef;
  readonly pageBreakBefore: boolean;
  readonly keepLines: boolean;
  readonly keepNext: boolean;
  readonly widowControl: boolean;
  readonly spaceBeforePt: number;
  readonly spaceAfterPt: number;
  readonly contextualSpacing: boolean;
  readonly styleId: string | null;
  /** Source-level visibility used only for pagination look-ahead across unmeasured blocks. */
  readonly inkless?: boolean;
  /** Parser-established ordinary text, excluding inline objects and mark-only paragraphs. */
  readonly onlyVisibleText?: boolean;
  /** Mutually exclusive Word/LibreOffice section-mark spacing interop role. */
  readonly continuousSectionRole?:
    | 'suppress-before'
    | 'collapse-mark'
    | 'drop-previous-after';
  readonly pageOwnedAnchorOccurrenceIds?: readonly string[];
}

export interface BodyTableSourceInput {
  readonly kind: 'table';
  readonly source: SourceRef;
  readonly rowCount?: number;
  /** §17.4.57 page/margin-positioned table whose exclusion can affect earlier text. */
  readonly pageOwnedFloatingTable?: boolean;
}

export interface BodyAdjacentTableGroupInput {
  readonly kind: 'adjacent-table-group';
  readonly logicalSequenceId: string;
  readonly source: SourceRef;
  readonly tables: readonly BodyTableSourceInput[];
}

export type BodyStoryReferenceSet = Readonly<{
  default: SourceRef | null;
  first: SourceRef | null;
  even: SourceRef | null;
}>;

export interface BodySectionLayoutInput {
  readonly sectionOccurrenceId: string;
  readonly source: SourceRef;
  readonly startType: SectionStartType;
  readonly context: DeepReadonly<SectionLayoutContext>;
  readonly pageNumbering: Readonly<{ start: number | null; format: string | null }>;
  readonly titlePage: boolean;
  readonly evenAndOddHeaders: boolean;
  readonly headers: BodyStoryReferenceSet;
  readonly footers: BodyStoryReferenceSet;
  readonly pageBordersAuthored: boolean;
  readonly pageBorders: DeepReadonly<PageBorders> | null;
  readonly pageLayout: DeepReadonly<SectionPageLayoutPolicy>;
}

export type BodyLayoutSequenceEntryFor<TSection> =
  | Readonly<{ kind: 'body-block'; block: BodyParagraphSourceInput | BodyTableSourceInput }>
  | BodyAdjacentTableGroupInput
  | Readonly<{
      kind: 'authored-break';
      source: SourceRef;
      break: AuthoredBreak;
      origin?: 'authored' | 'coverPageSynthetic';
      parity?: 'odd' | 'even';
      sameSourceParagraphAsPrevious?: boolean;
    }>
  | Readonly<{ kind: 'begin-section'; source: SourceRef; section: TSection }>
  | Readonly<{ kind: 'consume-source'; source: SourceRef; reason: 'hidden-paragraph' }>;

export type BodyLayoutSequenceEntry = BodyLayoutSequenceEntryFor<BodySectionLayoutInput>;

export interface BodyLayoutInput {
  readonly source: SourceRef;
  readonly initialSection: BodySectionLayoutInput;
  readonly sequence: readonly BodyLayoutSequenceEntry[];
  /** Immutable parse facts; attached only after geometry convergence. */
  readonly parserDiagnostics?: readonly LayoutDiagnostic[];
  /** §17.11 document-end note stories, in authored numbering order. */
  readonly endnoteIds?: readonly string[];
  readonly noteLayoutSettings?: Readonly<{
    footnotePosition: string;
    endnotePosition: string;
    footnoteNumbering?: NoteNumberingInput;
    endnoteNumbering?: NoteNumberingInput;
    footnoteSeparator?: NoteSeparatorInput;
    endnoteSeparator?: NoteSeparatorInput;
    footnoteContinuationSeparator?: NoteSeparatorInput;
    /** Parser-boundary only; rejected by the sealed layout source. */
    footnoteSeparatorParagraphs?: SelectedNoteSeparatorParagraphsInput;
    /** Canonical ordinary DOCX formatted story definitions; each replaces
     * only the scalar band height of its role, never its mark. */
    footnoteSeparatorStories?: SelectedNoteSeparatorDefinitionsInput;
    /** Parser-boundary only; rejected by the sealed layout source. */
    nativeSeparators?: NativeNoteSeparatorsInput;
    /** Canonical native separator definitions; they replace the scalar
     * separator modes for their note kind. */
    nativeSeparatorRoles?: NativeNoteSeparatorRolesInput;
  }>;
}

/** ECMA-376 §17.11.17/.18 numFmt and §17.11.20 numStart for one note kind. */
export interface NoteNumberingInput {
  /** ST_NumberFormat (§17.18.59); `decimal` when not authored. */
  readonly format: string;
  /** First automatic note number; 1 when not authored. */
  readonly start: number;
}

export interface BodyLayoutAcquisitionInput {
  readonly sectionIndex: BodySectionIndexInput;
  readonly evenAndOddHeaders: boolean;
  readonly parserDiagnostics?: readonly LayoutDiagnostic[];
  readonly endnoteIds?: readonly string[];
  readonly noteLayoutSettings: Readonly<{
    footnotePosition: string;
    endnotePosition: string;
    footnoteNumbering?: NoteNumberingInput;
    endnoteNumbering?: NoteNumberingInput;
    footnoteSeparator?: NoteSeparatorInput;
    endnoteSeparator?: NoteSeparatorInput;
    footnoteContinuationSeparator?: NoteSeparatorInput;
    footnoteSeparatorParagraphs?: SelectedNoteSeparatorParagraphsInput;
    nativeSeparators?: NativeNoteSeparatorsInput;
  }>;
  readonly pageLayoutSettings: Readonly<{
    mirrorMargins: boolean;
    gutterAtTop: boolean;
    bookFoldPrinting: boolean;
    bookFoldRevPrinting: boolean;
    printTwoOnOne: boolean;
  }>;
  readonly sequence: readonly BodyLayoutSequenceEntryFor<Readonly<{
    sectionOccurrenceId: string;
    startType: string;
  }>>[];
}

export function normalizeSectionStartType(value: string | null | undefined): SectionStartType {
  switch (value) {
    case 'continuous':
    case 'nextColumn':
    case 'nextPage':
    case 'oddPage':
    case 'evenPage':
      return value;
    default:
      return 'nextPage';
  }
}

export function bodyStoryReferences(
  stories: HeadersFooters,
  story: 'header' | 'footer',
  markerBodyIndex: number | null,
): BodyStoryReferenceSet {
  const prefix = markerBodyIndex === null ? null : `section:${markerBodyIndex}`;
  const reference = (kind: 'default' | 'first' | 'even'): SourceRef | null => (
    stories[kind] === null
      ? null
      : {
          story,
          storyInstance: prefix === null ? kind : `${prefix}:${kind}`,
          path: [],
        }
  );
  return Object.freeze({
    default: reference('default'),
    first: reference('first'),
    even: reference('even'),
  });
}

function sectionProps(occurrence: BodySectionOccurrence): SectionProps {
  return {
    ...occurrence.geometry,
    titlePage: occurrence.titlePage,
    evenAndOddHeaders: false,
    sectionStart: occurrence.startType,
    columns: occurrence.columns,
    textDirection: occurrence.textDirection,
    docGridType: occurrence.docGridType,
    docGridLinePitch: occurrence.docGridLinePitch,
    docGridCharSpace: occurrence.docGridCharSpace,
    pageNumType: occurrence.pageNumType,
    vAlign: occurrence.vAlign,
    lineNumbering: occurrence.lineNumbering,
  };
}

function sectionInput(
  occurrence: BodySectionOccurrence,
  acquired: BodyLayoutAcquisitionInput,
): BodySectionLayoutInput {
  const markerBodyIndex = occurrence.markerBodyIndex;
  return Object.freeze({
    sectionOccurrenceId: occurrence.sectionOccurrenceId,
    source: markerBodyIndex === null
      ? Object.freeze({ story: 'body' as const, storyInstance: 'body', path: Object.freeze([]) })
      : Object.freeze({
          story: 'body' as const,
          storyInstance: 'body',
          path: Object.freeze([markerBodyIndex]),
        }),
    startType: normalizeSectionStartType(occurrence.startType),
    context: Object.freeze(resolveAcquiredSectionLayoutContext(
      sectionProps(occurrence),
      occurrence.sectionBidi,
      occurrence.nativeSectionFlow,
    )),
    pageNumbering: Object.freeze({
      start: occurrence.pageNumType?.start ?? null,
      format: occurrence.pageNumType?.fmt ?? null,
    }),
    titlePage: occurrence.titlePage,
    evenAndOddHeaders: acquired.evenAndOddHeaders,
    headers: bodyStoryReferences(occurrence.headers, 'header', markerBodyIndex),
    footers: bodyStoryReferences(occurrence.footers, 'footer', markerBodyIndex),
    pageBordersAuthored: occurrence.pageBordersAuthored,
    pageBorders: occurrence.pageBorders,
    pageLayout: Object.freeze({
      physicalGeometry: Object.freeze({ ...occurrence.geometry }),
      columns: occurrence.columns,
      textDirection: occurrence.textDirection ?? 'lrTb',
      gutterPt: occurrence.gutterPt,
      rtlGutter: occurrence.rtlGutter,
      ...acquired.pageLayoutSettings,
    }),
  });
}

export function projectBodyLayoutInput(acquired: BodyLayoutAcquisitionInput): BodyLayoutInput {
  const sections = new Map(acquired.sectionIndex.occurrences.map((occurrence) => [
    occurrence.sectionOccurrenceId,
    sectionInput(occurrence, acquired),
  ]));
  const initialOccurrence = acquired.sectionIndex.occurrences[0];
  if (!initialOccurrence) throw new Error('DOCX body requires a final section owner');
  const initialSection = sections.get(initialOccurrence.sectionOccurrenceId)!;
  // The spacing role depends only on the authored section start type, already
  // present in the acquisition sequence. Resolve section owners and roles in
  // one pass so no second body-sized sequence is retained transiently.
  const sequence: BodyLayoutSequenceEntry[] = acquired.sequence.map((entry, index) => {
    if (entry.kind === 'begin-section') {
      const section = sections.get(entry.section.sectionOccurrenceId);
      if (!section) throw new Error(`Missing body section owner: ${entry.section.sectionOccurrenceId}`);
      return Object.freeze({ ...entry, section });
    }
    if (entry.kind !== 'body-block' || entry.block.kind !== 'paragraph') return entry;
    const continuousSectionRole = wordContinuousSectionRole(acquired.sequence, index);
    if (continuousSectionRole === undefined) return entry;
    return Object.freeze({
      ...entry,
      block: Object.freeze({
        ...entry.block,
        continuousSectionRole,
      }),
    });
  });
  return snapshotPlainData({
    source: { story: 'body', storyInstance: 'body', path: [] },
    initialSection,
    sequence,
    parserDiagnostics: acquired.parserDiagnostics ?? [],
    endnoteIds: acquired.endnoteIds ?? [],
    noteLayoutSettings: acquired.noteLayoutSettings,
  }, 'DOCX body layout input') as BodyLayoutInput;
}
