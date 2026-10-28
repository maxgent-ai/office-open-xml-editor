import type { Paragraph, TextBody, TextRun, TextRunData } from '@maxgent/ooxml/pptx';

import { replaceTextBodyParagraphPlainText } from '../../adapters/pptx-json-adapter';

/**
 * 相对某一 plain-text 基线的 0-based half-open 字符选区。
 * 基线为 OfficeCLI 同款：只拼接 text run 的 `text`（不含段落间额外 `\n`）。
 */
export interface TextSpan {
  readonly start: number;
  readonly end: number;
}

/**
 * 文本作用域。字符偏移相对该作用域的 {@link runPlainText}。
 */
export type TextScope =
  | { readonly kind: 'shape' }
  | {
    readonly kind: 'spans';
    readonly spans: readonly TextSpan[];
  }
  | {
    readonly kind: 'paragraph';
    /** 0-based，对应 `textBody.paragraphs`。 */
    readonly paragraphIndex: number;
    /** 缺省表示整段；有则相对该段 run 拼接文本。 */
    readonly spans?: readonly TextSpan[];
  };

/**
 * 一条增量文本编辑：可改整段文案、作用域样式，或两者同时。
 *
 * - `text`：整段纯文本替换（无 `\n`）；仅 `paragraph` 且无 `spans`。
 * - `style`：字段级样式补丁；与 `text` 至少提供一类。
 * - `spans` scope 仍仅支持 style（不含选区改字）。
 */
export interface TextStyleEdit {
  readonly scope: Exclude<TextScope, { kind: 'shape' }>;
  /** 整段纯文本；仅 paragraph 全段。 */
  readonly text?: string;
  /** 作用域样式补丁。 */
  readonly style?: TextStylePatch;
}

/**
 * Shape / 选区文本样式补丁。未出现的键表示不改。
 */
export interface TextStylePatch {
  /** 加粗；`null` 表示清除显式值、改回继承。 */
  readonly bold?: boolean | null;
  /** 斜体；`null` 表示清除显式值、改回继承。 */
  readonly italic?: boolean | null;
  /** 下划线：关闭 / 单线 / 双线。 */
  readonly underline?: false | 'single' | 'double';
  /** 删除线：关闭 / 单线 / 双线。 */
  readonly strikethrough?: false | 'single' | 'double';
  /** 字号（pt）；`null` 表示清除显式值。 */
  readonly fontSize?: number | null;
  /** 文本颜色（RRGGBB 或 RRGGBBAA）；`null` 表示清除显式值。 */
  readonly color?: string | null;
  /** Latin/默认字体；`null` 表示清除显式值。 */
  readonly fontFamily?: string | null;
  /** 东亚字体（a:ea）；`null` 表示清除显式值。 */
  readonly fontFamilyEa?: string | null;
  /** 字母大小写变换。 */
  readonly caps?: 'none' | 'small' | 'all';
  /**
   * 字距，单位为 pt（与 parser / renderer 的 `TextRun.letterSpacing` 一致）。
   * OOXML `rPr@spc` 为 1/100 pt；parser 会 `/100`，editor 侧始终用 pt。
   */
  readonly letterSpacing?: number | null;
  /** 文本高亮色（RRGGBB[AA]）；`null` 表示清除。 */
  readonly highlight?: string | null;
  /**
   * 段落对齐。整框 / 整段 scope 可用；字符 spans 上若出现则作用于与选区相交的段落。
   */
  readonly align?: 'l' | 'ctr' | 'r' | 'just';
  /** 文本框垂直锚点；仅整框 style 有意义。 */
  readonly verticalAlign?: 't' | 'ctr' | 'b';
}

/** 与 OfficeCLI shape/paragraph `range` 对齐的纯文本：仅拼接 text run。 */
export function runPlainText(textBody: TextBody): string {
  let text = '';
  for (const paragraph of textBody.paragraphs) {
    text += paragraphRunPlainText(paragraph);
  }
  return text;
}

export function paragraphRunPlainText(paragraph: Paragraph): string {
  let text = '';
  for (const run of paragraph.runs) {
    if (run.type === 'text') text += run.text;
  }
  return text;
}

export function hasStyleKeys(style: TextStylePatch): boolean {
  return Object.keys(style).length > 0;
}

/** OfficeCLI 无法表达的“清回继承”键；翻译期需 resolve-then-set。 */
export function hasNullClearableStyleKeys(style: TextStylePatch): boolean {
  return (
    style.bold === null
    || style.italic === null
    || style.fontSize === null
    || style.color === null
    || style.fontFamily === null
    || style.fontFamilyEa === null
    || style.letterSpacing === null
  );
}

/**
 * 解析 run 级继承时使用的上下文（对齐 pptx renderer：run → para def* → body default）。
 * 主题色/字体仅能用 presentation / shape 上已物化的字段近似。
 */
export interface TextStyleInheritanceContext {
  readonly textBody: TextBody;
  readonly paragraph: Paragraph;
  readonly shapeDefaultTextColor?: string | null;
  readonly presentationDefaultTextColor?: string | null;
  readonly presentationMinorFont?: string | null;
  readonly presentationMajorFont?: string | null;
}

const DEFAULT_INHERITED_FONT_SIZE_PT = 18;
const DEFAULT_INHERITED_COLOR = '000000';

/**
 * 将 style 中的 `null`（清回继承）解析为可下发 OfficeCLI 的显式值。
 * 这是同步侧近似：包内会写入显式 rPr，而非真正删除属性。
 */
export function resolveInheritedStylePatch(
  style: TextStylePatch,
  context: TextStyleInheritanceContext,
): TextStylePatch {
  if (!hasNullClearableStyleKeys(style)) return style;

  const { paragraph, textBody } = context;
  const next: Record<string, unknown> = { ...style };

  if (style.bold === null) {
    next.bold = paragraph.defBold ?? textBody.defaultBold ?? false;
  }
  if (style.italic === null) {
    next.italic = paragraph.defItalic ?? textBody.defaultItalic ?? false;
  }
  if (style.fontSize === null) {
    next.fontSize = paragraph.defFontSize
      ?? textBody.defaultFontSize
      ?? DEFAULT_INHERITED_FONT_SIZE_PT;
  }
  if (style.color === null) {
    next.color = paragraph.defColor
      ?? context.shapeDefaultTextColor
      ?? context.presentationDefaultTextColor
      ?? DEFAULT_INHERITED_COLOR;
  }
  if (style.fontFamily === null) {
    const font = paragraph.defFontFamily
      ?? context.presentationMinorFont
      ?? context.presentationMajorFont;
    if (!font) {
      throw new TypeError(
        'Cannot resolve inherited fontFamily: no paragraph/theme font is available',
      );
    }
    next.fontFamily = font;
  }
  if (style.fontFamilyEa === null) {
    // 无独立 ea 继承槽时，用拉丁有效字体近似“与拉丁同一 typeface”。
    const ea = paragraph.defFontFamily
      ?? context.presentationMinorFont
      ?? context.presentationMajorFont;
    if (!ea) {
      throw new TypeError(
        'Cannot resolve inherited fontFamilyEa: no paragraph/theme font is available',
      );
    }
    next.fontFamilyEa = ea;
  }
  if (style.letterSpacing === null) {
    next.letterSpacing = 0;
  }

  return freezeStyle(next as TextStylePatch);
}

/**
 * 整框 style 在含 null 且各段继承不同时，展开为按段 edits；否则返回单一 resolved style。
 */
export function materializeShapeStyleForOfficeCli(
  textBody: TextBody,
  style: TextStylePatch,
  inheritance: Omit<TextStyleInheritanceContext, 'paragraph' | 'textBody'>,
): { style: TextStylePatch } | { edits: readonly TextStyleEdit[] } {
  if (!hasNullClearableStyleKeys(style)) return { style };
  if (textBody.paragraphs.length === 0) {
    throw new TypeError('Cannot resolve inherited text style without paragraphs');
  }

  const resolved = textBody.paragraphs.map((paragraph, paragraphIndex) => ({
    paragraphIndex,
    style: resolveInheritedStylePatch(style, {
      textBody,
      paragraph,
      ...inheritance,
    }),
  }));

  const firstKey = stableStyleKey(resolved[0]!.style);
  if (resolved.every((entry) => stableStyleKey(entry.style) === firstKey)) {
    return { style: resolved[0]!.style };
  }

  return {
    edits: resolved.map((entry) => freezeTextStyleEdit({
      scope: { kind: 'paragraph', paragraphIndex: entry.paragraphIndex },
      style: entry.style,
    })),
  };
}

/**
 * 将一条 TextStyleEdit 中的 null 解析为显式值；跨段且继承不同时按段拆开。
 */
export function materializeTextStyleEditForOfficeCli(
  textBody: TextBody,
  edit: TextStyleEdit,
  inheritance: Omit<TextStyleInheritanceContext, 'paragraph' | 'textBody'>,
): readonly TextStyleEdit[] {
  if (!edit.style || !hasNullClearableStyleKeys(edit.style)) return [edit];

  if (edit.scope.kind === 'paragraph') {
    const paragraph = textBody.paragraphs[edit.scope.paragraphIndex];
    if (!paragraph) {
      throw new TypeError(
        `paragraphIndex ${edit.scope.paragraphIndex} is out of range`,
      );
    }
    return [freezeTextStyleEdit({
      scope: edit.scope,
      ...(edit.text !== undefined ? { text: edit.text } : {}),
      style: resolveInheritedStylePatch(edit.style, {
        textBody,
        paragraph,
        ...inheritance,
      }),
    })];
  }

  const absoluteSpans = normalizeSpans(edit.scope.spans);
  const pieces: TextStyleEdit[] = [];
  let cursor = 0;
  for (const [paragraphIndex, paragraph] of textBody.paragraphs.entries()) {
    const paragraphStart = cursor;
    const paragraphEnd = cursor + paragraphRunPlainText(paragraph).length;
    cursor = paragraphEnd;

    const localSpans: TextSpan[] = [];
    for (const span of absoluteSpans) {
      const overlapStart = Math.max(paragraphStart, span.start);
      const overlapEnd = Math.min(paragraphEnd, span.end);
      if (overlapStart >= overlapEnd) continue;
      localSpans.push({
        start: overlapStart - paragraphStart,
        end: overlapEnd - paragraphStart,
      });
    }
    if (localSpans.length === 0) continue;

    const resolved = resolveInheritedStylePatch(edit.style, {
      textBody,
      paragraph,
      ...inheritance,
    });
    pieces.push(freezeTextStyleEdit({
      scope: {
        kind: 'paragraph',
        paragraphIndex,
        spans: Object.freeze(localSpans.map((span) => Object.freeze({ ...span }))),
      },
      style: resolved,
    }));
  }

  if (pieces.length === 0) {
    // 空选区：退回首段解析，保持与单点 scope 行为一致。
    const paragraph = textBody.paragraphs[0]!;
    return [freezeTextStyleEdit({
      scope: edit.scope,
      style: resolveInheritedStylePatch(edit.style, {
        textBody,
        paragraph,
        ...inheritance,
      }),
    })];
  }

  return pieces;
}

export function freezeStyle(style: TextStylePatch): TextStylePatch {
  return Object.freeze({ ...style });
}

export function freezeScope(scope: TextScope): TextScope {
  if (scope.kind === 'shape') return Object.freeze({ kind: 'shape' });
  if (scope.kind === 'spans') {
    return Object.freeze({
      kind: 'spans',
      spans: Object.freeze(scope.spans.map((span) => Object.freeze({ ...span }))),
    });
  }
  return Object.freeze({
    kind: 'paragraph',
    paragraphIndex: scope.paragraphIndex,
    spans: scope.spans === undefined
      ? undefined
      : Object.freeze(scope.spans.map((span) => Object.freeze({ ...span }))),
  });
}

export function freezeTextStyleEdit(edit: TextStyleEdit): TextStyleEdit {
  const hasText = edit.text !== undefined;
  const hasStyle = edit.style !== undefined && hasStyleKeys(edit.style);
  if (!hasText && !hasStyle) {
    throw new TypeError('TextStyleEdit requires text and/or a non-empty style patch');
  }
  return Object.freeze({
    scope: freezeScope(edit.scope) as TextStyleEdit['scope'],
    ...(hasText ? { text: edit.text } : {}),
    ...(hasStyle ? { style: freezeStyle(edit.style!) } : {}),
  });
}

export function editHasStyle(edit: TextStyleEdit): edit is TextStyleEdit & {
  readonly style: TextStylePatch;
} {
  return edit.style !== undefined && hasStyleKeys(edit.style);
}

export function assertValidSpan(span: TextSpan, textLength: number, label: string): void {
  if (
    !Number.isInteger(span.start)
    || !Number.isInteger(span.end)
    || span.start < 0
    || span.end < span.start
    || span.end > textLength
  ) {
    throw new TypeError(
      `${label} span [${span.start}, ${span.end}) is invalid for text length ${textLength}`,
    );
  }
}

export function assertValidStyleEdit(edit: TextStyleEdit, textBody: TextBody): void {
  const hasText = edit.text !== undefined;
  const hasStyle = editHasStyle(edit);
  if (!hasText && !hasStyle) {
    throw new TypeError('TextStyleEdit requires text and/or a non-empty style patch');
  }
  if (hasText) {
    if (edit.scope.kind !== 'paragraph') {
      throw new TypeError('TextStyleEdit text is only supported on paragraph scope');
    }
    if (edit.scope.spans) {
      throw new TypeError('TextStyleEdit text cannot be combined with paragraph spans');
    }
    if (/[\r\n]/.test(edit.text!)) {
      throw new TypeError('TextStyleEdit text must be a single paragraph (no newlines)');
    }
  }
  if (edit.scope.kind === 'spans') {
    if (edit.scope.spans.length === 0) {
      throw new TypeError('TextStyleEdit spans scope requires at least one span');
    }
    if (hasStyle && 'verticalAlign' in edit.style) {
      throw new TypeError('verticalAlign is only valid on whole-shape text style updates');
    }
    const length = runPlainText(textBody).length;
    for (const [index, span] of edit.scope.spans.entries()) {
      assertValidSpan(span, length, `spans[${index}]`);
    }
    return;
  }

  const paragraph = textBody.paragraphs[edit.scope.paragraphIndex];
  if (!paragraph) {
    throw new TypeError(
      `paragraphIndex ${edit.scope.paragraphIndex} is out of range`,
    );
  }
  if (hasStyle && 'verticalAlign' in edit.style) {
    throw new TypeError('verticalAlign is only valid on whole-shape text style updates');
  }
  if (edit.scope.spans) {
    if (edit.scope.spans.length === 0) {
      throw new TypeError('paragraph spans requires at least one span when provided');
    }
    const length = paragraphRunPlainText(paragraph).length;
    for (const [index, span] of edit.scope.spans.entries()) {
      assertValidSpan(span, length, `paragraph.spans[${index}]`);
    }
  }
}

/** 整框应用样式（所有 text run + 可选段落对齐 / 垂直锚点）。 */
export function applyTextStylePatch(textBody: TextBody, style: TextStylePatch): TextBody {
  return {
    ...textBody,
    verticalAnchor: style.verticalAlign ?? textBody.verticalAnchor,
    paragraphs: textBody.paragraphs.map((paragraph) => patchParagraph(paragraph, style)),
  };
}

export function applyTextStyleEdit(textBody: TextBody, edit: TextStyleEdit): TextBody {
  assertValidStyleEdit(edit, textBody);

  let next = textBody;
  if (edit.text !== undefined && edit.scope.kind === 'paragraph') {
    const replaced = replaceTextBodyParagraphPlainText(
      next,
      edit.scope.paragraphIndex,
      edit.text,
    );
    if (!replaced) {
      throw new TypeError(`paragraphIndex ${edit.scope.paragraphIndex} is out of range`);
    }
    next = replaced;
  }

  if (!editHasStyle(edit)) return next;

  if (edit.scope.kind === 'spans') {
    return applySpansToTextBody(next, edit.scope.spans, edit.style, 0);
  }

  const paragraphIndex = edit.scope.paragraphIndex;
  const spans = edit.scope.spans;
  if (!spans) {
    return {
      ...next,
      paragraphs: next.paragraphs.map((paragraph, index) => (
        index === paragraphIndex ? patchParagraph(paragraph, edit.style) : paragraph
      )),
    };
  }

  let offset = 0;
  for (let index = 0; index < paragraphIndex; index += 1) {
    offset += paragraphRunPlainText(next.paragraphs[index]).length;
  }
  const absoluteSpans = spans.map((span) => ({
    start: offset + span.start,
    end: offset + span.end,
  }));
  return applySpansToTextBody(next, absoluteSpans, edit.style, paragraphIndex);
}

/**
 * 按修改前坐标把 style 打到与 spans 相交的 text run 上（必要时拆 run）。
 * `align` 作用于与选区相交的段落；`verticalAlign` 忽略（调用方应已拒绝）。
 */
export function applySpansToTextBody(
  textBody: TextBody,
  spans: readonly TextSpan[],
  style: TextStylePatch,
  /** 若只允许改某一段的 align，传入该段 index；否则相交段都可改 align。 */
  alignParagraphIndex?: number,
): TextBody {
  const normalized = normalizeSpans(spans);
  if (normalized.length === 0) return textBody;

  const touchedParagraphs = new Set<number>();
  let cursor = 0;
  const paragraphs = textBody.paragraphs.map((paragraph, paragraphIndex) => {
    const nextRuns: TextRun[] = [];
    for (const run of paragraph.runs) {
      if (run.type !== 'text') {
        nextRuns.push(run);
        continue;
      }
      const runStart = cursor;
      const runEnd = cursor + run.text.length;
      cursor = runEnd;

      let sliceStart = 0;
      const pieces: TextRunData[] = [];
      for (const span of normalized) {
        const overlapStart = Math.max(runStart, span.start);
        const overlapEnd = Math.min(runEnd, span.end);
        if (overlapStart >= overlapEnd) continue;
        touchedParagraphs.add(paragraphIndex);
        const localStart = overlapStart - runStart;
        const localEnd = overlapEnd - runStart;
        if (localStart > sliceStart) {
          pieces.push(sliceTextRun(run, sliceStart, localStart));
        }
        pieces.push(patchTextRun(sliceTextRun(run, localStart, localEnd), style));
        sliceStart = localEnd;
      }
      if (pieces.length === 0) {
        nextRuns.push(run);
        continue;
      }
      if (sliceStart < run.text.length) {
        pieces.push(sliceTextRun(run, sliceStart, run.text.length));
      }
      nextRuns.push(...pieces);
    }

    const align = style.align !== undefined
      && (alignParagraphIndex === undefined || alignParagraphIndex === paragraphIndex)
      && touchedParagraphs.has(paragraphIndex)
      ? style.align
      : paragraph.alignment;

    return {
      ...paragraph,
      alignment: align,
      runs: nextRuns,
    };
  });

  return { ...textBody, paragraphs };
}

function normalizeSpans(spans: readonly TextSpan[]): TextSpan[] {
  const sorted = spans
    .filter((span) => span.end > span.start)
    .slice()
    .sort((a, b) => a.start - b.start || a.end - b.end);
  if (sorted.length === 0) return sorted;

  const merged: TextSpan[] = [{ ...sorted[0]! }];
  for (let index = 1; index < sorted.length; index += 1) {
    const current = sorted[index]!;
    const last = merged[merged.length - 1]!;
    if (current.start <= last.end) {
      merged[merged.length - 1] = {
        start: last.start,
        end: Math.max(last.end, current.end),
      };
      continue;
    }
    merged.push({ ...current });
  }
  return merged;
}

function sliceTextRun(run: TextRunData, start: number, end: number): TextRunData {
  return { ...run, text: run.text.slice(start, end) };
}

export function patchParagraph(paragraph: Paragraph, style: TextStylePatch): Paragraph {
  return {
    ...paragraph,
    alignment: style.align ?? paragraph.alignment,
    runs: paragraph.runs.map((run) => {
      if (run.type !== 'text') return run;
      return patchTextRun(run, style);
    }),
  };
}

export function patchTextRun(run: TextRunData, style: TextStylePatch): TextRunData {
  const next: TextRunData = { ...run };

  if ('bold' in style) next.bold = style.bold ?? null;
  if ('italic' in style) next.italic = style.italic ?? null;
  if ('fontSize' in style) next.fontSize = style.fontSize ?? null;
  if ('color' in style) next.color = style.color ?? null;
  if ('fontFamily' in style) next.fontFamily = style.fontFamily ?? null;
  if ('fontFamilyEa' in style) {
    if (style.fontFamilyEa == null) delete next.fontFamilyEa;
    else next.fontFamilyEa = style.fontFamilyEa;
  }
  if (style.caps !== undefined) {
    if (style.caps === 'none') delete next.caps;
    else next.caps = style.caps;
  }
  if ('letterSpacing' in style) {
    if (style.letterSpacing == null) delete next.letterSpacing;
    else next.letterSpacing = style.letterSpacing;
  }
  if ('highlight' in style) {
    if (style.highlight == null) delete next.highlight;
    else next.highlight = style.highlight;
  }
  if (style.underline !== undefined) {
    if (style.underline === false) {
      next.underline = false;
      delete next.underlineStyle;
    } else if (style.underline === 'single') {
      next.underline = true;
      delete next.underlineStyle;
    } else {
      next.underline = true;
      next.underlineStyle = 'dbl';
    }
  }
  if (style.strikethrough !== undefined) {
    if (style.strikethrough === false) {
      next.strikethrough = false;
      delete next.strikeDouble;
    } else if (style.strikethrough === 'single') {
      next.strikethrough = true;
      delete next.strikeDouble;
    } else {
      next.strikethrough = true;
      next.strikeDouble = true;
    }
  }

  return next;
}

/** 按 patch 键从指定作用域采集可逆样式快照（取选区起点 / 段首 text run）。 */
export function captureTextStylePatchAtScope(
  textBody: TextBody,
  scope: TextScope,
  patch: TextStylePatch,
): TextStylePatch | undefined {
  const run = findRunAtScopeStart(textBody, scope);
  const fromRun = captureStyleFromRun(run, patch);
  if (fromRun === undefined) return undefined;

  const captured: Record<string, unknown> = { ...fromRun };
  if ('align' in patch) {
    const paragraphIndex = scope.kind === 'paragraph'
      ? scope.paragraphIndex
      : 0;
    const alignment = textBody.paragraphs[paragraphIndex]?.alignment;
    if (alignment !== 'l' && alignment !== 'ctr' && alignment !== 'r' && alignment !== 'just') {
      return undefined;
    }
    captured.align = alignment;
  }
  if ('verticalAlign' in patch) {
    const anchor = textBody.verticalAnchor;
    if (anchor !== 't' && anchor !== 'ctr' && anchor !== 'b') return undefined;
    captured.verticalAlign = anchor;
  }

  return freezeStyle(captured as TextStylePatch);
}

/**
 * 为多样式选区编辑生成可逆 edits：按与选区相交的每个 text-run 切片分别保存
 * 修改前样式（而非只取选区起点一个值）。返回顺序为 forward edits 的逆序。
 */
export function captureInverseTextStyleEdits(
  textBody: TextBody,
  edits: readonly TextStyleEdit[],
): TextStyleEdit[] | undefined {
  const inverse: TextStyleEdit[] = [];
  for (let index = edits.length - 1; index >= 0; index -= 1) {
    const pieces = captureInversePiecesForEdit(textBody, edits[index]!);
    if (pieces === undefined) return undefined;
    inverse.push(...pieces);
  }
  return inverse;
}

/**
 * 整框 style 的可逆快照：run / 段落对齐按切片分别保存；若全部同质则折叠为顶层 style。
 * `verticalAlign` 只能出现在返回的顶层 `style` 上（可与 edits 并存）。
 */
export function captureInverseShapeStylePatch(
  textBody: TextBody,
  patch: TextStylePatch,
): { style?: TextStylePatch; edits?: readonly TextStyleEdit[] } | undefined {
  const withoutVertical = omitVerticalAlign(patch);
  let edits: TextStyleEdit[] | undefined;

  if (hasStyleKeys(withoutVertical)) {
    const length = runPlainText(textBody).length;
    if (hasRunScopedKeys(withoutVertical) && length === 0) return undefined;

    const forwardEdits: TextStyleEdit[] = length > 0
      ? [{
        scope: { kind: 'spans', spans: [{ start: 0, end: length }] },
        style: withoutVertical,
      }]
      : textBody.paragraphs.map((_, paragraphIndex) => ({
        scope: { kind: 'paragraph' as const, paragraphIndex },
        style: withoutVertical,
      }));
    if (forwardEdits.length === 0) return undefined;

    edits = captureInverseTextStyleEdits(textBody, forwardEdits);
    if (edits === undefined) return undefined;
  }

  let verticalAlign: TextStylePatch['verticalAlign'];
  if ('verticalAlign' in patch) {
    const anchor = textBody.verticalAnchor;
    if (anchor !== 't' && anchor !== 'ctr' && anchor !== 'b') return undefined;
    verticalAlign = anchor;
  }

  const collapsed = tryCollapseInverseShapeStyle(
    textBody,
    patch,
    edits ?? [],
    verticalAlign,
  );
  if (collapsed) return { style: collapsed };

  if ((!edits || edits.length === 0) && verticalAlign === undefined) return undefined;
  return {
    ...(edits && edits.length > 0 ? { edits } : {}),
    ...(verticalAlign !== undefined
      ? { style: freezeStyle({ verticalAlign }) }
      : {}),
  };
}

function omitVerticalAlign(patch: TextStylePatch): TextStylePatch {
  if (!('verticalAlign' in patch)) return patch;
  const { verticalAlign: _verticalAlign, ...rest } = patch;
  return freezeStyle(rest);
}

function tryCollapseInverseShapeStyle(
  textBody: TextBody,
  originalPatch: TextStylePatch,
  edits: readonly TextStyleEdit[],
  verticalAlign: TextStylePatch['verticalAlign'],
): TextStylePatch | undefined {
  const runEdits = edits.filter((edit) => editHasStyle(edit) && hasRunScopedKeys(edit.style));
  const alignEdits = edits.filter((edit) => editHasStyle(edit) && 'align' in edit.style);

  if (hasRunScopedKeys(originalPatch)) {
    if (runEdits.length !== 1) return undefined;
    if (!scopeCoversEntirePlainText(textBody, runEdits[0]!.scope)) return undefined;
  } else if (runEdits.length > 0) {
    return undefined;
  }

  let align: TextStylePatch['align'];
  if ('align' in originalPatch) {
    if (alignEdits.length !== textBody.paragraphs.length || textBody.paragraphs.length === 0) {
      return undefined;
    }
    const values = new Set(alignEdits.map((edit) => edit.style!.align));
    if (values.size !== 1) return undefined;
    align = alignEdits[0]!.style!.align;
  } else if (alignEdits.length > 0) {
    return undefined;
  }

  const captured: Record<string, unknown> = {
    ...(runEdits[0]?.style ?? {}),
  };
  if (align !== undefined) captured.align = align;
  if (verticalAlign !== undefined) captured.verticalAlign = verticalAlign;
  if (!hasStyleKeys(captured as TextStylePatch)) return undefined;
  return freezeStyle(captured as TextStylePatch);
}

function scopeCoversEntirePlainText(
  textBody: TextBody,
  scope: TextStyleEdit['scope'],
): boolean {
  const length = runPlainText(textBody).length;
  if (scope.kind === 'spans') {
    const normalized = normalizeSpans(scope.spans);
    return normalized.length === 1
      && normalized[0]!.start === 0
      && normalized[0]!.end === length;
  }
  if (scope.spans) return false;
  return textBody.paragraphs.length === 1 && scope.paragraphIndex === 0;
}

function captureInversePiecesForEdit(
  textBody: TextBody,
  edit: TextStyleEdit,
): TextStyleEdit[] | undefined {
  if (edit.text !== undefined) {
    if (edit.scope.kind !== 'paragraph' || edit.scope.spans) return undefined;
    const paragraph = textBody.paragraphs[edit.scope.paragraphIndex];
    if (!paragraph || !canInvertParagraphTextReplacement(paragraph)) return undefined;

    const priorText = paragraphRunPlainText(paragraph);
    let style: TextStylePatch | undefined;
    if (editHasStyle(edit)) {
      const captured = captureTextStylePatchAtScope(
        textBody,
        { kind: 'paragraph', paragraphIndex: edit.scope.paragraphIndex },
        edit.style,
      );
      if (captured === undefined) return undefined;
      style = captured;
    }

    return [freezeTextStyleEdit({
      scope: { kind: 'paragraph', paragraphIndex: edit.scope.paragraphIndex },
      text: priorText,
      ...(style ? { style } : {}),
    })];
  }

  if (!editHasStyle(edit)) return undefined;

  const pieces: TextStyleEdit[] = [];

  if (hasRunScopedKeys(edit.style)) {
    const absoluteSpans = resolveAbsoluteSpans(textBody, edit.scope);
    const slices = collectOverlappingRunSlices(textBody, absoluteSpans);
    let pending: {
      spans: TextSpan[];
      style: TextStylePatch;
      styleKey: string;
    } | undefined;

    const flush = (): void => {
      if (!pending) return;
      pieces.push(freezeTextStyleEdit({
        scope: toInverseRunScope(edit.scope, pending.spans, textBody),
        style: pending.style,
      }));
      pending = undefined;
    };

    for (const slice of slices) {
      const style = captureStyleFromRun(slice.run, edit.style);
      if (style === undefined) return undefined;
      const styleKey = stableStyleKey(style);
      const localSpan = toLocalSpan(edit.scope, slice);
      if (pending && pending.styleKey === styleKey && spansAreAdjacent(pending.spans, localSpan)) {
        const last = pending.spans[pending.spans.length - 1]!;
        pending.spans = [
          ...pending.spans.slice(0, -1),
          { start: last.start, end: localSpan.end },
        ];
        continue;
      }
      if (pending && pending.styleKey === styleKey) {
        pending.spans = [...pending.spans, localSpan];
        continue;
      }
      flush();
      pending = { spans: [localSpan], style, styleKey };
    }
    flush();
  }

  if ('align' in edit.style) {
    for (const paragraphIndex of affectedParagraphIndexes(textBody, edit.scope)) {
      const alignment = textBody.paragraphs[paragraphIndex]?.alignment;
      if (alignment !== 'l' && alignment !== 'ctr' && alignment !== 'r' && alignment !== 'just') {
        return undefined;
      }
      pieces.push(freezeTextStyleEdit({
        scope: { kind: 'paragraph', paragraphIndex },
        style: { align: alignment },
      }));
    }
  }

  return pieces;
}

function canInvertParagraphTextReplacement(paragraph: Paragraph): boolean {
  if (paragraph.runs.length !== 1) return false;
  const run = paragraph.runs[0];
  return run?.type === 'text'
    && run.fieldType == null
    && run.hyperlink == null
    && run.hyperlinkAction == null;
}

export function canInvertTextBodyPlainTextReplacement(textBody: TextBody): boolean {
  const paragraph = textBody.paragraphs[0];
  return textBody.paragraphs.length === 1
    && paragraph !== undefined
    && canInvertParagraphTextReplacement(paragraph);
}

function captureStyleFromRun(
  run: TextRunData | undefined,
  patch: TextStylePatch,
): TextStylePatch | undefined {
  if (!run && hasRunScopedKeys(patch)) return undefined;

  // ast-grep-ignore: no-plain-object-string-map -- captured keys are the explicit TextStylePatch properties below.
  const captured: Record<string, unknown> = {};
  if ('bold' in patch) captured.bold = run?.bold ?? null;
  if ('italic' in patch) captured.italic = run?.italic ?? null;
  if ('fontSize' in patch) captured.fontSize = run?.fontSize ?? null;
  if ('color' in patch) captured.color = run?.color ?? null;
  if ('fontFamily' in patch) captured.fontFamily = run?.fontFamily ?? null;
  if ('fontFamilyEa' in patch) captured.fontFamilyEa = run?.fontFamilyEa ?? null;
  if ('caps' in patch) captured.caps = run?.caps ?? 'none';
  if ('letterSpacing' in patch) captured.letterSpacing = run?.letterSpacing ?? null;
  if ('highlight' in patch) captured.highlight = run?.highlight ?? null;
  if ('underline' in patch) {
    captured.underline = !run?.underline
      ? false
      : run.underlineStyle === 'dbl'
        ? 'double'
        : 'single';
  }
  if ('strikethrough' in patch) {
    captured.strikethrough = !run?.strikethrough
      ? false
      : run.strikeDouble
        ? 'double'
        : 'single';
  }

  return freezeStyle(captured as TextStylePatch);
}

interface AbsoluteRunSlice {
  readonly absoluteStart: number;
  readonly absoluteEnd: number;
  readonly paragraphIndex: number;
  readonly paragraphAbsoluteStart: number;
  readonly run: TextRunData;
}

function resolveAbsoluteSpans(
  textBody: TextBody,
  scope: TextStyleEdit['scope'],
): TextSpan[] {
  if (scope.kind === 'spans') return normalizeSpans(scope.spans);

  let paragraphStart = 0;
  for (let index = 0; index < scope.paragraphIndex; index += 1) {
    paragraphStart += paragraphRunPlainText(textBody.paragraphs[index]!).length;
  }
  const paragraph = textBody.paragraphs[scope.paragraphIndex]!;
  if (!scope.spans) {
    const length = paragraphRunPlainText(paragraph).length;
    return [{ start: paragraphStart, end: paragraphStart + length }];
  }
  return normalizeSpans(scope.spans.map((span) => ({
    start: paragraphStart + span.start,
    end: paragraphStart + span.end,
  })));
}

function collectOverlappingRunSlices(
  textBody: TextBody,
  absoluteSpans: readonly TextSpan[],
): AbsoluteRunSlice[] {
  const slices: AbsoluteRunSlice[] = [];
  let cursor = 0;
  for (const [paragraphIndex, paragraph] of textBody.paragraphs.entries()) {
    const paragraphAbsoluteStart = cursor;
    for (const run of paragraph.runs) {
      if (run.type !== 'text') continue;
      const runStart = cursor;
      const runEnd = cursor + run.text.length;
      cursor = runEnd;
      for (const span of absoluteSpans) {
        const overlapStart = Math.max(runStart, span.start);
        const overlapEnd = Math.min(runEnd, span.end);
        if (overlapStart >= overlapEnd) continue;
        slices.push({
          absoluteStart: overlapStart,
          absoluteEnd: overlapEnd,
          paragraphIndex,
          paragraphAbsoluteStart,
          run,
        });
      }
    }
  }
  return mergeAdjacentSameRunSlices(slices);
}

function mergeAdjacentSameRunSlices(slices: readonly AbsoluteRunSlice[]): AbsoluteRunSlice[] {
  const merged: AbsoluteRunSlice[] = [];
  for (const slice of slices) {
    const last = merged[merged.length - 1];
    if (
      last
      && last.run === slice.run
      && last.absoluteEnd === slice.absoluteStart
    ) {
      merged[merged.length - 1] = {
        ...last,
        absoluteEnd: slice.absoluteEnd,
      };
      continue;
    }
    merged.push(slice);
  }
  return merged;
}

function toLocalSpan(
  scope: TextStyleEdit['scope'],
  slice: AbsoluteRunSlice,
): TextSpan {
  if (scope.kind === 'paragraph') {
    return {
      start: slice.absoluteStart - slice.paragraphAbsoluteStart,
      end: slice.absoluteEnd - slice.paragraphAbsoluteStart,
    };
  }
  return { start: slice.absoluteStart, end: slice.absoluteEnd };
}

function toInverseRunScope(
  original: TextStyleEdit['scope'],
  spans: readonly TextSpan[],
  textBody: TextBody,
): TextStyleEdit['scope'] {
  if (original.kind === 'paragraph') {
    const paragraphLength = paragraphRunPlainText(
      textBody.paragraphs[original.paragraphIndex]!,
    ).length;
    const coversWhole = spans.length === 1
      && spans[0]!.start === 0
      && spans[0]!.end === paragraphLength;
    if (coversWhole && !hasOnlyPartialSpans(original)) {
      return { kind: 'paragraph', paragraphIndex: original.paragraphIndex };
    }
    return {
      kind: 'paragraph',
      paragraphIndex: original.paragraphIndex,
      spans: Object.freeze(spans.map((span) => Object.freeze({ ...span }))),
    };
  }
  return {
    kind: 'spans',
    spans: Object.freeze(spans.map((span) => Object.freeze({ ...span }))),
  };
}

function hasOnlyPartialSpans(scope: TextStyleEdit['scope']): boolean {
  return scope.kind === 'paragraph' && scope.spans !== undefined;
}

function spansAreAdjacent(spans: readonly TextSpan[], next: TextSpan): boolean {
  const last = spans[spans.length - 1];
  return last !== undefined && last.end === next.start;
}

function stableStyleKey(style: TextStylePatch): string {
  return JSON.stringify(style, Object.keys(style).sort());
}

function affectedParagraphIndexes(
  textBody: TextBody,
  scope: TextStyleEdit['scope'],
): number[] {
  if (scope.kind === 'paragraph') return [scope.paragraphIndex];
  const indexes = new Set<number>();
  let cursor = 0;
  for (const [paragraphIndex, paragraph] of textBody.paragraphs.entries()) {
    const start = cursor;
    const end = cursor + paragraphRunPlainText(paragraph).length;
    cursor = end;
    for (const span of normalizeSpans(scope.spans)) {
      if (span.start < end && span.end > start) indexes.add(paragraphIndex);
    }
  }
  return [...indexes].sort((a, b) => a - b);
}

function hasRunScopedKeys(patch: TextStylePatch): boolean {
  return (
    'bold' in patch
    || 'italic' in patch
    || 'underline' in patch
    || 'strikethrough' in patch
    || 'fontSize' in patch
    || 'color' in patch
    || 'fontFamily' in patch
    || 'fontFamilyEa' in patch
    || 'caps' in patch
    || 'letterSpacing' in patch
    || 'highlight' in patch
  );
}

function findRunAtScopeStart(
  textBody: TextBody,
  scope: TextScope,
): TextRunData | undefined {
  if (scope.kind === 'shape') {
    return findFirstTextRun(textBody);
  }
  if (scope.kind === 'paragraph') {
    const paragraph = textBody.paragraphs[scope.paragraphIndex];
    if (!paragraph) return undefined;
    if (!scope.spans || scope.spans.length === 0) {
      return paragraph.runs.find((run): run is TextRunData => run.type === 'text');
    }
    return findRunAtParagraphOffset(paragraph, scope.spans[0].start);
  }
  if (scope.spans.length === 0) return findFirstTextRun(textBody);
  return findRunAtShapeOffset(textBody, scope.spans[0].start);
}

function findRunAtShapeOffset(textBody: TextBody, offset: number): TextRunData | undefined {
  let cursor = 0;
  let last: TextRunData | undefined;
  for (const paragraph of textBody.paragraphs) {
    for (const run of paragraph.runs) {
      if (run.type !== 'text') continue;
      last = run;
      const next = cursor + run.text.length;
      if (offset >= cursor && offset < next) return run;
      cursor = next;
    }
  }
  return offset === cursor ? last : undefined;
}

function findRunAtParagraphOffset(
  paragraph: Paragraph,
  offset: number,
): TextRunData | undefined {
  let cursor = 0;
  let last: TextRunData | undefined;
  for (const run of paragraph.runs) {
    if (run.type !== 'text') continue;
    last = run;
    const next = cursor + run.text.length;
    if (offset >= cursor && offset < next) return run;
    cursor = next;
  }
  return offset === cursor ? last : undefined;
}

function findFirstTextRun(textBody: TextBody): TextRunData | undefined {
  for (const paragraph of textBody.paragraphs) {
    const run = paragraph.runs.find((candidate): candidate is TextRunData => candidate.type === 'text');
    if (run) return run;
  }
  return undefined;
}

/** 编码 OfficeCLI `range` prop（可多段）。 */
export function formatOfficeCliRange(spans: readonly TextSpan[]): string {
  return spans.map((span) => `${span.start}:${span.end}`).join(',');
}
