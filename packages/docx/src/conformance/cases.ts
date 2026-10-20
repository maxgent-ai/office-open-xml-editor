/**
 * Redistributable synthetic DOCX coverage matrix.
 *
 * The axes below describe OOXML sources, not renderer branches.  A row is valid
 * only when every selected value can be represented by the same minimal
 * document.  Coverage therefore means every *feasible* value pair appears at
 * least once; impossible pairs (for example an inline object with an anchor
 * reference frame) are deliberately excluded.
 */

const CONFORMANCE_AXES = {
  story: ['body', 'header', 'footer'],
  container: ['paragraph', 'table', 'nestedTable'],
  paragraph: ['single', 'wrapped'],
  object: ['none', 'inline', 'floating'],
  direction: ['ltr', 'rtl'],
  spacing: ['auto', 'exact'],
  styleSource: ['direct', 'paragraphStyle', 'documentDefault'],
  fontSource: ['direct', 'theme', 'documentDefault'],
  anchorReference: ['none', 'margin', 'page', 'paragraph'],
} as const;

export type ConformanceAxisName = keyof typeof CONFORMANCE_AXES;

export type ConformanceAxisValues = {
  readonly [K in ConformanceAxisName]: (typeof CONFORMANCE_AXES)[K][number];
};

export interface ConformanceExpectation {
  readonly pageCount: number;
  readonly pageWidthPt: number;
  readonly pageHeightPt: number;
  readonly targetText: string;
  readonly targetStory: ConformanceAxisValues['story'];
  readonly tableDepth: 0 | 1 | 2;
  readonly drawingCount: 0 | 1;
}

export interface ConformanceCase {
  readonly id: string;
  readonly axes: ConformanceAxisValues;
  readonly expected: ConformanceExpectation;
  readonly seed: boolean;
}

const AXIS_NAMES = Object.keys(CONFORMANCE_AXES) as ConformanceAxisName[];

function compareCodeUnits(left: string, right: string): number {
  return left < right ? -1 : left > right ? 1 : 0;
}

function isValidAxes(axes: ConformanceAxisValues): boolean {
  if (axes.object === 'floating') {
    if (axes.anchorReference === 'none') return false;
  } else if (axes.anchorReference !== 'none') {
    return false;
  }

  // Header/footer fixtures isolate story routing. Tables and drawings remain in
  // body cases so a failure cannot be masked by header/footer repetition.
  if (axes.story !== 'body') {
    return axes.container === 'paragraph' && axes.object === 'none';
  }
  return true;
}

function isCompleteAxes(
  row: Partial<ConformanceAxisValues>,
): row is ConformanceAxisValues {
  return AXIS_NAMES.every((axis) => row[axis] !== undefined);
}

function cartesianRows(): ConformanceAxisValues[] {
  let rows: Partial<ConformanceAxisValues>[] = [{}];
  for (const axis of AXIS_NAMES) {
    rows = rows.flatMap((row) =>
      CONFORMANCE_AXES[axis].map<Partial<ConformanceAxisValues>>((value) => ({
        ...row,
        [axis]: value,
      })),
    );
  }
  return rows
    .filter(isCompleteAxes)
    .filter(isValidAxes);
}

function valueKey(axis: ConformanceAxisName, value: string): string {
  return `${axis}=${value}`;
}

function pairKey(
  leftAxis: ConformanceAxisName,
  leftValue: string,
  rightAxis: ConformanceAxisName,
  rightValue: string,
): string {
  return `${valueKey(leftAxis, leftValue)}|${valueKey(rightAxis, rightValue)}`;
}

function pairsOf(row: ConformanceAxisValues): string[] {
  const pairs: string[] = [];
  for (let left = 0; left < AXIS_NAMES.length; left += 1) {
    for (let right = left + 1; right < AXIS_NAMES.length; right += 1) {
      const leftAxis = AXIS_NAMES[left]!;
      const rightAxis = AXIS_NAMES[right]!;
      pairs.push(pairKey(leftAxis, row[leftAxis], rightAxis, row[rightAxis]));
    }
  }
  return pairs;
}

function rowKey(row: ConformanceAxisValues): string {
  return AXIS_NAMES.map((axis) => valueKey(axis, row[axis])).join('|');
}

const SEED_ROWS: readonly ConformanceAxisValues[] = [
  {
    story: 'body',
    container: 'nestedTable',
    paragraph: 'wrapped',
    object: 'floating',
    direction: 'rtl',
    spacing: 'exact',
    styleSource: 'paragraphStyle',
    fontSource: 'theme',
    anchorReference: 'page',
  },
  {
    story: 'body',
    container: 'table',
    paragraph: 'single',
    object: 'inline',
    direction: 'ltr',
    spacing: 'auto',
    styleSource: 'direct',
    fontSource: 'direct',
    anchorReference: 'none',
  },
  {
    story: 'header',
    container: 'paragraph',
    paragraph: 'wrapped',
    object: 'none',
    direction: 'rtl',
    spacing: 'exact',
    styleSource: 'documentDefault',
    fontSource: 'documentDefault',
    anchorReference: 'none',
  },
  {
    story: 'footer',
    container: 'paragraph',
    paragraph: 'single',
    object: 'none',
    direction: 'ltr',
    spacing: 'auto',
    styleSource: 'paragraphStyle',
    fontSource: 'theme',
    anchorReference: 'none',
  },
];

/**
 * Deterministic greedy covering-array construction.
 *
 * Exhaustive candidate enumeration is intentionally acceptable here: the
 * constrained matrix has fewer than two thousand rows.  At each step the
 * lexicographically first row covering the most still-uncovered feasible pairs
 * wins.  Reverse elimination then removes redundant non-seed rows.
 */
function generatePairwiseAxes(): readonly {
  readonly axes: ConformanceAxisValues;
  readonly seed: boolean;
}[] {
  const candidates = cartesianRows().sort((left, right) =>
    compareCodeUnits(rowKey(left), rowKey(right)));
  const required = new Set(candidates.flatMap(pairsOf));
  const selected: Array<{ axes: ConformanceAxisValues; seed: boolean }> = [];
  const selectedKeys = new Set<string>();

  const add = (axes: ConformanceAxisValues, seed: boolean): void => {
    if (!isValidAxes(axes)) throw new Error(`invalid conformance seed: ${rowKey(axes)}`);
    const key = rowKey(axes);
    if (selectedKeys.has(key)) return;
    selected.push({ axes, seed });
    selectedKeys.add(key);
    for (const pair of pairsOf(axes)) required.delete(pair);
  };

  for (const seed of SEED_ROWS) add(seed, true);

  while (required.size > 0) {
    let best: ConformanceAxisValues | undefined;
    let bestScore = -1;
    for (const candidate of candidates) {
      if (selectedKeys.has(rowKey(candidate))) continue;
      let score = 0;
      for (const pair of pairsOf(candidate)) {
        if (required.has(pair)) score += 1;
      }
      if (score > bestScore) {
        best = candidate;
        bestScore = score;
      }
    }
    if (!best || bestScore <= 0) {
      throw new Error(`pairwise generation stalled with ${required.size} pairs`);
    }
    add(best, false);
  }

  // Remove a non-seed row when every pair it owns is still covered elsewhere.
  for (let index = selected.length - 1; index >= 0; index -= 1) {
    const row = selected[index]!;
    if (row.seed) continue;
    const others = selected.filter((_, otherIndex) => otherIndex !== index);
    const covered = new Set(others.flatMap(({ axes }) => pairsOf(axes)));
    if (pairsOf(row.axes).every((pair) => covered.has(pair))) {
      selected.splice(index, 1);
    }
  }
  return selected;
}

type AxisAbbreviations = {
  readonly [K in ConformanceAxisName]: Readonly<Record<ConformanceAxisValues[K], string>>;
};

const AXIS_ABBREVIATIONS: AxisAbbreviations = {
  story: { body: 'body', header: 'hdr', footer: 'ftr' },
  container: { paragraph: 'p', table: 'tbl', nestedTable: 'ntbl' },
  paragraph: { single: 'one', wrapped: 'wrap' },
  object: { none: 'noobj', inline: 'inl', floating: 'float' },
  direction: { ltr: 'ltr', rtl: 'rtl' },
  spacing: { auto: 'auto', exact: 'exact' },
  styleSource: { direct: 'direct', paragraphStyle: 'pstyle', documentDefault: 'docdef' },
  fontSource: { direct: 'font', theme: 'theme', documentDefault: 'fontdef' },
  anchorReference: { none: 'noref', margin: 'margin', page: 'page', paragraph: 'para' },
};

function axisAbbreviation<K extends ConformanceAxisName>(
  axis: K,
  value: ConformanceAxisValues[K],
): string {
  return AXIS_ABBREVIATIONS[axis][value];
}

function slug(axes: ConformanceAxisValues): string {
  return AXIS_NAMES.map((axis) => axisAbbreviation(axis, axes[axis])).join('-');
}

export const CONFORMANCE_CASES: readonly ConformanceCase[] = generatePairwiseAxes()
  .map(({ axes, seed }) => {
    const id = slug(axes);
    return {
      id,
      axes,
      seed,
      expected: {
        pageCount: 1,
        pageWidthPt: 612,
        pageHeightPt: 792,
        targetText: `CASE_${id.toUpperCase().replaceAll('-', '_')}`,
        targetStory: axes.story,
        tableDepth: axes.container === 'nestedTable' ? 2 : axes.container === 'table' ? 1 : 0,
        drawingCount: axes.object === 'none' ? 0 : 1,
      },
    } satisfies ConformanceCase;
  });

export function feasiblePairKeys(): ReadonlySet<string> {
  return new Set(cartesianRows().flatMap(pairsOf));
}

export function coveredPairKeys(
  cases: readonly Pick<ConformanceCase, 'axes'>[],
): ReadonlySet<string> {
  return new Set(cases.flatMap(({ axes }) => pairsOf(axes)));
}
