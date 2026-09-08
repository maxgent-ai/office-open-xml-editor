import type { Presentation } from '@maxgent/ooxml/pptx';
import type { ElementRef } from '../domain/mutation';
import type { OfficeCliBatchSender, OfficeCliBatchSendResult } from '../submission/types';
import type { EDITOR_SESSION_CHANGE_REASONS } from './constants';

export type PptxEditorSessionChangeReason =
  (typeof EDITOR_SESSION_CHANGE_REASONS)[keyof typeof EDITOR_SESSION_CHANGE_REASONS];
export type PptxEditorSaveStatus = 'idle' | 'saving' | 'failed' | 'unknown';
export type PptxEditorSaveResult =
  | { readonly status: 'unchanged' }
  | (OfficeCliBatchSendResult & { readonly commandId: string });

export interface PptxEditorSessionOptions {
  readonly presentation: Presentation;
  readonly sendBatch: OfficeCliBatchSender;
  /** A fresh identifier for each explicit save attempt; this is not an idempotency guarantee. */
  readonly createSaveId: () => string;
  readonly onListenerError?: PptxEditorSessionListenerErrorHandler;
}

export interface PptxEditorSessionSnapshot {
  readonly presentation: Presentation;
  readonly dirty: boolean;
  readonly saveStatus: PptxEditorSaveStatus;
  readonly saveCommandId?: string;
  readonly saveError?: unknown;
  readonly canEdit: boolean;
  readonly undoDepth: number;
  readonly redoDepth: number;
  readonly canUndo: boolean;
  readonly canRedo: boolean;
}

export interface PptxEditorSessionChange {
  readonly reason: PptxEditorSessionChangeReason;
  readonly snapshot: PptxEditorSessionSnapshot;
  readonly commandId?: string;
  readonly changedSlideIds: readonly string[];
  readonly changedElements: readonly ElementRef[];
}

export type PptxEditorSessionListener = (change: PptxEditorSessionChange) => void;
export type PptxEditorSessionListenerErrorHandler = (
  cause: unknown,
  change: PptxEditorSessionChange,
) => void;
