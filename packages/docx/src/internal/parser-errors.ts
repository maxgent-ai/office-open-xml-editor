// This adapter also runs in workers and the Node/Markdown CLI source graphs.
import { OoxmlError } from '@silurus/ooxml-core/worker';

/** DOCX Word-compatibility rejection emitted by its numbering parser. */
const DOCX_ILVL_PREFIX = 'OOXML_DOCX_ILVL:';

/** Decode only the numbering parser's exact known reasons; preserve the
 * established messages and leave unrelated or future parser errors alone. */
function parseDocxIlvlError(error: unknown): OoxmlError | undefined {
  const text = error instanceof Error ? error.message : String(error);
  if (!text.startsWith(DOCX_ILVL_PREFIX)) return undefined;
  const reason = text.slice(DOCX_ILVL_PREFIX.length);
  if (reason === 'repair-required:level-definition') {
    return new OoxmlError(
      'invalid-numbering',
      'The DOCX defines a numbering level outside Word’s supported 0–8 range and requires repair.',
    );
  }
  if (reason === 'cannot-open:whitespace' || reason === 'cannot-open:non-decimal') {
    return new OoxmlError('invalid-numbering', `Word cannot open this DOCX numbering level (${reason.slice(12)}).`);
  }
  return undefined;
}

/** Keep unreadable thrown values intact so the generic transport retains its
 * existing defensive fallback instead of failing inside the DOCX adapter. */
export function normalizeDocxParserError(error: unknown): unknown {
  try {
    return parseDocxIlvlError(error) ?? error;
  } catch {
    return error;
  }
}

/** Normalize before shared pull/worker transport serializes a parser failure. */
export function withDocxParserErrors<T>(operation: () => T): T {
  try {
    return operation();
  } catch (error) {
    throw normalizeDocxParserError(error);
  }
}
