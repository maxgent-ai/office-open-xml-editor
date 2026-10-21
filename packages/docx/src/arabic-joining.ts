import { joiningType, joinsToFollowing, joinsToPreceding } from '@silurus/ooxml-core/internal/joining-type';
import {
  KASHIDA_ALEF_TAH_CUR,
  KASHIDA_BEH_PREV,
  KASHIDA_HAH_PREV,
  KASHIDA_REH_YEH_CUR,
  KASHIDA_SEEN_PREV,
  KASHIDA_WAW_AIN_CUR,
} from './arabic-joining.generated.js';
import { wordKashidaFinalFormApplies } from './layout/script-compatibility.js';

const LAM = 0x0644;
const LAM_ALEF_PARTNERS = new Set([0x0627, 0x0622, 0x0623, 0x0625, 0x0671]);
const TATWEEL = 0x0640;

/*
 * Priority values follow usp10.h SCRIPT_JUSTIFY and Qt's qtextengine.cpp
 * JustificationClass ranking. Letter families follow the ArabicGroup aliases
 * in old HarfBuzz's harfbuzz-arabic.c getArabicProperties. Our class position
 * rules intentionally simplify Qt's shape-form conditions under
 * `word-kashida-final-form-priority`. Keep final-letter classes gated by word
 * position below.
 */
const SEEN_FAMILY = new Set(KASHIDA_SEEN_PREV);
const HAH_FAMILY = new Set(KASHIDA_HAH_PREV);
const BEH_FAMILY = new Set(KASHIDA_BEH_PREV);
const RAH_OR_YEH = new Set(KASHIDA_REH_YEH_CUR);
const FINAL_ALEF_LAM_KAF = new Set(KASHIDA_ALEF_TAH_CUR);
const FINAL_WAW = new Set(KASHIDA_WAW_AIN_CUR);

const enum JustificationPriority {
  Normal = 7,
  Waw = 8,
  BaRa = 9,
  Alef = 10,
  HahDal = 11,
  Seen = 12,
  Kashida = 13,
}

/**
 * Return original code-point offsets before which U+0640 may be inserted.
 * Transparent combining marks remain attached to the preceding joining letter.
 */
export function kashidaInsertionPoints(text: string): number[] {
  const cps = [...text].map((ch) => ch.codePointAt(0)!);
  const points: number[] = [];
  let preceding = cps.length > 0 && joiningType(cps[0]) !== 'T' ? 0 : -1;

  for (let k = 1; k < cps.length; k++) {
    const current = cps[k];
    if (joiningType(current) === 'T') continue;

    if (preceding >= 0) {
      const previous = cps[preceding];
      const lamAlef = previous === LAM && LAM_ALEF_PARTNERS.has(current);
      if (
        !lamAlef &&
        joinsToFollowing(previous) &&
        joinsToPreceding(current)
      ) {
        points.push(k);
      }
    }
    preceding = k;
  }

  return points;
}

export interface KashidaPriorityPoint {
  beforeCp: number;
  priority: number;
}

function candidatePriority(
  cps: readonly number[],
  beforeCp: number,
  lastLetterCp: number,
): JustificationPriority {
  let previousCp = beforeCp - 1;
  while (previousCp >= 0 && joiningType(cps[previousCp]) === 'T') previousCp--;

  const previous = cps[previousCp];
  const current = cps[beforeCp];
  if (previous === TATWEEL) return JustificationPriority.Kashida;
  if (SEEN_FAMILY.has(previous)) return JustificationPriority.Seen;
  if (HAH_FAMILY.has(previous)) return JustificationPriority.HahDal;

  // Alef and Waw classes apply only to a word-final following letter. This
  // deliberately prevents a medial alef (for example in الكتابة) outranking a
  // later Normal join.
  if (wordKashidaFinalFormApplies(beforeCp, lastLetterCp)
    && FINAL_ALEF_LAM_KAF.has(current)) {
    return JustificationPriority.Alef;
  }
  if (BEH_FAMILY.has(previous) && RAH_OR_YEH.has(current)) {
    return JustificationPriority.BaRa;
  }
  if (wordKashidaFinalFormApplies(beforeCp, lastLetterCp)
    && FINAL_WAW.has(current)) {
    return JustificationPriority.Waw;
  }
  return JustificationPriority.Normal;
}

/**
 * Choose Word's single highest-priority kashida opportunity in each word.
 * Words are maximal non-whitespace runs. Equal classes choose the opportunity
 * nearest the word end, matching Qt's `>=` replacement tie-break.
 */
export function priorityKashidaInsertionCandidates(text: string): KashidaPriorityPoint[] {
  const cps = [...text];
  const result: KashidaPriorityPoint[] = [];

  for (let wordStart = 0; wordStart < cps.length;) {
    while (wordStart < cps.length && /\s/u.test(cps[wordStart])) wordStart++;
    if (wordStart >= cps.length) break;

    let wordEnd = wordStart + 1;
    while (wordEnd < cps.length && !/\s/u.test(cps[wordEnd])) wordEnd++;

    const word = cps.slice(wordStart, wordEnd);
    const wordCodePoints = word.map((ch) => ch.codePointAt(0)!);
    let lastLetterCp = wordCodePoints.length - 1;
    while (lastLetterCp >= 0 && joiningType(wordCodePoints[lastLetterCp]) === 'T') {
      lastLetterCp--;
    }

    let chosen = -1;
    let chosenPriority = -1;
    for (const beforeCp of kashidaInsertionPoints(word.join(''))) {
      const priority = candidatePriority(wordCodePoints, beforeCp, lastLetterCp);
      if (priority >= chosenPriority) {
        chosen = beforeCp;
        chosenPriority = priority;
      }
    }
    if (chosen >= 0) {
      result.push({ beforeCp: wordStart + chosen, priority: chosenPriority });
    }
    wordStart = wordEnd;
  }

  return result;
}

/** Return only the selected code-point offsets for callers that do not allocate slack. */
export function priorityKashidaInsertionPoints(text: string): number[] {
  return priorityKashidaInsertionCandidates(text).map(({ beforeCp }) => beforeCp);
}
