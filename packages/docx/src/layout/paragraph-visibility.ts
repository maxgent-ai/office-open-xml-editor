import type { DocParagraph } from '../types.js';

export function isInklessParagraph(paragraph: Readonly<DocParagraph>): boolean {
  return !(paragraph.runs ?? []).some((candidate) => {
    if (candidate.type === 'text') return candidate.text.length > 0;
    return true;
  });
}
