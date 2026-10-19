import type { CompatibilityRule, DeepReadonly } from './types.js';

function requireText(value: string, field: string): void {
  if (value.trim() === '') throw new Error(`CompatibilityRule.${field} must not be empty`);
}

/** Source-only evidence validation; production builds validate declarations and
 * emit frozen IDs via scripts/docx-compatibility-build.mjs. Rendering logic must
 * consume the ID, and must not depend on metadata or factory side effects. */
export function defineCompatibilityRule<const Rule extends CompatibilityRule>(
  rule: Rule,
): DeepReadonly<Rule> {
  requireText(rule.id, 'id');
  requireText(rule.description, 'description');
  if (!/^[a-z0-9]+(?:-[a-z0-9]+)*$/.test(rule.id)) {
    throw new Error('CompatibilityRule.id must be a stable kebab-case identifier');
  }
  if (rule.evidence.kind === 'microsoft-note') {
    requireText(rule.evidence.reference, 'evidence.reference');
    if (!/^\[MS-[A-Z0-9]+\] §§?\d/.test(rule.evidence.reference)) {
      throw new Error('CompatibilityRule.evidence.reference must identify a Microsoft specification section');
    }
  } else if (rule.evidence.kind === 'regression-test' || rule.evidence.kind === 'library-policy') {
    requireText(rule.evidence.reference, 'evidence.reference');
    if (!/^packages\/docx\/src\/.+\.(?:test|spec)\.tsx?#[^#]+$/.test(
      rule.evidence.reference,
    )) {
      throw new Error('CompatibilityRule.evidence.reference must use DOCX path#test-title');
    }
  } else {
    requireText(rule.evidence.syntheticFixtureId, 'evidence.syntheticFixtureId');
    requireText(rule.evidence.application, 'evidence.application');
    requireText(rule.evidence.version, 'evidence.version');
    requireText(rule.evidence.platform, 'evidence.platform');
    if (!/^[a-z0-9]+(?:-[a-z0-9]+)*$/.test(rule.evidence.syntheticFixtureId)) {
      throw new Error('CompatibilityRule.evidence.syntheticFixtureId must be kebab-case');
    }
  }
  Object.freeze(rule.evidence);
  return Object.freeze(rule) as DeepReadonly<Rule>;
}

export const WORD_SECTION_BTLR_TBRL_PAGE_FRAME = defineCompatibilityRule({
  id: 'word-section-btlr-tbrl-page-frame',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/layout/coordinate-space.test.ts#maps Transitional text direction %s to %s',
  },
  description: 'Issue #988 comment 4950296007 records that, unlike the normative ECMA-376 Part 4 §14.11.7 equivalence to lr, Word uses the tbRl page frame for section-level btLr; this rule covers only the page frame, while glyph orientation is paint-owned. Evidence (observed Office behaviour): Word output adjudicated in issue #988 and PR #996; [MS-OI29500] §2.1.558(b) states that Word ignores section-level btLr but not the tbRl mapping.',
});

export const WORD_PAGE_ANCHORED_TABLE_TEXT_AREA_MOVE = defineCompatibilityRule({
  id: 'word-page-anchored-table-text-area-move',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'issue-1902-page-anchored-tables',
    application: 'Microsoft Word',
    version: '16.0.20430',
    platform: 'Windows',
  },
  description: 'A page- or margin-anchored floating table whose box ends below the bottom margin, or that tblOverlap="never" displacement set above the top margin, moves to the next page with the flow after it, where its authored position applies; when nothing precedes it in its region it stays. Its box is clamped only to the page edges, not to the margin band. Evidence: issue #1902 Windows controls with a full-width table at y=149pt ending at 700/737 (kept) and 740/760/790/800pt (moved; 800 clamped to the page bottom) after a paragraph, the same table first on the page (kept, clamped), margin-anchored ends 737/760pt, and a never-displaced table set at y=29pt (moved). The overlap-permitted collision deferral this replaces (word-page-anchored-table-collision-deferral) did not occur in any control.',
});

export const WORD_NEVER_TABLE_PAGE_DISPLACEMENT = defineCompatibilityRule({
  id: 'word-never-table-page-displacement',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'issue-1902-page-anchored-tables',
    application: 'Microsoft Word',
    version: '16.0.20430',
    platform: 'Windows',
  },
  description: 'A floating table that must avoid another (§17.4.56 tblOverlap="never" on either) moves right, then down; when that leaves the page it sits directly above the tables it meets, and when that leaves the page too it keeps its authored position. Evidence: issue #1902 Windows controls with a full-width never table at y=98pt over a 72pt table at y=149pt (moved below), over a table reaching y=779pt (72pt table set above it; 120pt table set above the top margin and moved by word-page-anchored-table-text-area-move; 160pt table kept at y=98pt).',
});

export const WORD_PAGE_TABLE_LATER_YIELDS = defineCompatibilityRule({
  id: 'word-page-table-later-yields',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'issue-1902-page-anchored-tables',
    application: 'Microsoft Word',
    version: '16.0.20430',
    platform: 'Windows',
  },
  description: 'When placing page-positioned tables oscillates between pages, the later table in source order moves to the later page and the earlier one keeps its page, as Word lays tables out in document order. Evidence: issue #1902 Windows controls (a page-anchored table whose exclusion would push its own preceding text off the page moves to the next page with it) and a corpus document whose second full-width page-anchored table leaves the first on page 1.',
});

export const WORD_EMPTY_MARK_FLOAT_SIDE_GAP = defineCompatibilityRule({
  id: 'word-empty-mark-float-side-gap',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/float-line-start-one-inch.test.ts#keeps an anchor-host metric-only line on the paragraph-mark threshold',
  },
  description: 'An empty or anchor-only paragraph-mark line may start beside a square-wrapped object when the available side gap can hold the paragraph mark em; Word controls also admit visible text by its next atom, rather than a universal width. Evidence (observed Office behaviour): issue #676 / PR #696 and PR #1035; the issue #1670 controls of word-float-gap-flow (Word 16.113.2 macOS, Windows Word 16.0.20430) cover the empty-mark threshold.',
});

export function wordEmptyMarkMinimumStartWidthPx(
  paragraphMarkEmPt: number,
  scale: number,
): number {
  return paragraphMarkEmPt * scale;
}

export const WORD_FLOAT_GAP_FLOW = defineCompatibilityRule({
  id: 'word-float-gap-flow',
  evidence: { kind: 'office-observation', syntheticFixtureId: 'float-gap-continuation',
    application: 'Microsoft Word', version: '16.113.2', platform: 'macOS 27.0' },
  description: 'Issue #1670 controlled exports in modes 14 and 15 fill successive gaps on one baseline in paragraph reading order, then restart at the leading gap of the next baseline. Two/three-gap square and rectangular tight/through controls, left/right/largest restrictions, 20–100pt gaps, 10/20pt text, alignment, indents, padding, empty marks and Hebrew RTL establish atom-fit admission, including 40pt gaps. An unbroken Latin word too wide for every gap moves below the object. These controls supersede the one-inch inference from issue #676. Geometry remains fail-closed on non-finite input and unavailable polygon contours; emergency splitting is retained in a full paragraph band. Windows Word 16.0 build 16.0.20430 retest is consistent within the tested subset; see the officeObservations entry for this rule in scripts/docx-compatibility-microsoft-evidence.json for values and evidence limits.',
});

/** Word for Mac issue #1668 controls (527 cases including horizontal twins):
 * both stacked WordArt values use continuous clockwise sideways Latin with
 * ordinary horizontal advances, upright East Asian/emoji clusters, and left-to-right
 * columns. Sizes 12–48 pt, face/style and character sweeps, wrapping, spacing,
 * anchors, transforms and split graphemes showed no direction-mode exception.
 * This is host compatibility, not ECMA-376 §20.1.10.83's stacked-letter rule.
 * Reuse the existing mixed-orientation vertical pipeline; native glyph metrics
 * and vertical glyph designs remain the text service's responsibility.
 */
export function wordTextBoxVerticalMode(
  value: string | null | undefined,
): 'vert' | 'vert270' | 'eaVert' | 'mongolianVert' | undefined {
  if (value === 'wordArtVert' || value === 'wordArtVertRtl') return 'mongolianVert';
  return value === 'vert' || value === 'vert270' || value === 'eaVert' || value === 'mongolianVert'
    ? value : undefined;
}
