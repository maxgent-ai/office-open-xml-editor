/** A rejected caller request (for example an out-of-range page index or an
 * unusable caller canvas). Keep this exact code across the ordinary worker
 * error wire; parser/resource RangeError and TypeError failures retain the
 * terminal acquisition policy. */
export class DocxCallerInputError extends RangeError {
  readonly code = 'docx-caller-input';
}
