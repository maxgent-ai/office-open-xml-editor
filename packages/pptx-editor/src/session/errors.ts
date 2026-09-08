export type PptxEditorSessionErrorCode = 'session.disposed' | 'session.locked' | 'session.historyEmpty' | 'session.invalidResolution' | 'session.invalidCommand';

export class PptxEditorSessionError extends Error {
  readonly name = 'PptxEditorSessionError';

  constructor(
    readonly code: PptxEditorSessionErrorCode,
    message: string,
    options?: ErrorOptions,
  ) {
    super(message, options);
  }
}
