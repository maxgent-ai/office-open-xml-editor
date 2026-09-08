import type { Presentation } from '@maxgent/ooxml/pptx';
import { getSlideMutationId } from '../adapters/pptx-json-adapter';
import type { Command, NonEmptyReadonlyArray } from '../domain/command';
import { MUTATION_TYPES } from '../domain/mutation-types';
import type { Mutation } from '../domain/mutation';
import { applyCommand } from '../engine/mutation-engine';
import type { CommandExecutionResult } from '../engine/types';
import { createUndoRedoEntry } from '../history/command-inverter';
import type { OfficeCliBatchSendResult } from '../submission/types';
import { toOfficeCliBatch } from '../transport/officecli/officecli-translator';
import { EDITOR_SESSION_CHANGE_REASONS as REASONS } from './constants';
import { PptxEditorSessionError } from './errors';
import type {
  PptxEditorSessionChange,
  PptxEditorSessionListener,
  PptxEditorSessionOptions,
  PptxEditorSessionSnapshot,
  PptxEditorSaveResult,
  PptxEditorSaveStatus,
} from './types';

interface HistoryNode {
  readonly presentation: Presentation;
  readonly parent?: HistoryNode;
  readonly command?: Command;
  readonly inverse?: NonEmptyReadonlyArray<Mutation>;
  readonly change?: CommandExecutionResult;
}

/** Local history and the saved position share the same immutable history nodes. */
export class PptxEditorSession {
  readonly #options: PptxEditorSessionOptions;
  readonly #listeners = new Set<PptxEditorSessionListener>();
  readonly #redo: HistoryNode[] = [];
  #current: HistoryNode;
  #saved: HistoryNode;
  #savedAncestors = new Set<HistoryNode>();
  #undoDepth = 0;
  #saveStatus: PptxEditorSaveStatus = 'idle';
  #saveCommandId: string | undefined;
  #saveError: unknown;
  #snapshot: PptxEditorSessionSnapshot;
  #disposed = false;

  constructor(options: PptxEditorSessionOptions) {
    this.#options = options;
    this.#current = this.#saved = { presentation: options.presentation };
    this.#savedAncestors.add(this.#saved);
    this.#snapshot = this.#createSnapshot();
  }

  getSnapshot(): PptxEditorSessionSnapshot {
    this.#assertActive();
    return this.#snapshot;
  }

  subscribe(listener: PptxEditorSessionListener): () => void {
    this.#assertActive();
    this.#listeners.add(listener);
    return () => { this.#listeners.delete(listener); };
  }

  apply(input: Command): PptxEditorSessionChange {
    this.#assertEditable();
    if (!input.id || input.mutations.length === 0) {
      throw new PptxEditorSessionError('session.invalidCommand', 'A command needs an id and at least one mutation');
    }
    const command: Command = Object.freeze({
      ...input,
      mutations: Object.freeze([...input.mutations]) as NonEmptyReadonlyArray<Mutation>,
    });
    const presentation = this.#current.presentation;
    // Validate translation before accepting the edit. Nothing is sent here.
    toOfficeCliBatch(presentation, command);
    const change = applyCommand(presentation, command);
    // Existing shape restoration drops rich formatting and geometry. Do not
    // claim a lossless saved undo for deletion; local undo still restores JSON.
    const inverse = command.mutations.some((mutation) => mutation.type === MUTATION_TYPES.REMOVE_ELEMENT)
      ? undefined
      : createUndoRedoEntry(presentation, command)?.inverseMutations;
    this.#current = {
      presentation: change.presentation,
      parent: this.#current,
      command,
      inverse,
      change,
    };
    this.#redo.length = 0;
    this.#undoDepth++;
    return this.#publish(REASONS.LOCAL_APPLIED, change);
  }

  undo(): PptxEditorSessionChange {
    this.#assertEditable();
    if (!this.#canUndo(this.#current)) throw historyEmpty();
    const node = this.#current;
    this.#redo.push(node);
    this.#current = node.parent as HistoryNode;
    this.#undoDepth--;
    return this.#publish(REASONS.HISTORY_CHANGED, node.change);
  }

  redo(): PptxEditorSessionChange {
    this.#assertEditable();
    const node = this.#redo.pop();
    if (!node) throw historyEmpty();
    this.#current = node;
    this.#undoDepth++;
    return this.#publish(REASONS.HISTORY_CHANGED, node.change);
  }

  async save(): Promise<PptxEditorSaveResult> {
    this.#assertEditable();
    if (this.#current === this.#saved) return { status: 'unchanged' };
    // Lock before invoking caller code, including the id factory and listeners.
    this.#saveStatus = 'saving';
    this.#saveError = undefined;
    this.#saveCommandId = undefined;
    let batch;
    try {
      this.#saveCommandId = this.#options.createSaveId();
      if (!this.#saveCommandId) throw new TypeError('A save needs a nonempty id');
      const mutations = this.#unsavedMutations();
      const first = mutations[0];
      if (!first) throw historyEmpty();
      batch = toOfficeCliBatch(this.#saved.presentation, {
        id: this.#saveCommandId,
        mutations: [first, ...mutations.slice(1)],
      });
    } catch (cause) {
      this.#saveStatus = 'failed';
      this.#saveError = cause;
      this.#publish(REASONS.SAVE_CHANGED);
      throw cause;
    }
    this.#publish(REASONS.SAVE_CHANGED);
    let result: OfficeCliBatchSendResult;
    try {
      result = await this.#options.sendBatch(batch);
      if (!result || !['confirmed', 'rejected', 'unknown'].includes(result.status)) {
        throw new TypeError('Invalid OfficeCLI send result');
      }
    } catch (cause) {
      result = { status: 'unknown', cause };
    }
    this.#settle(result);
    return Object.freeze({ ...result, commandId: batch.commandId });
  }

  /** Only use after independently verifying the outcome of this exact batch. */
  resolveUnknown(
    commandId: string,
    result: Exclude<OfficeCliBatchSendResult, { status: 'unknown' }>,
  ): void {
    this.#assertActive();
    if (this.#saveStatus !== 'unknown' || commandId !== this.#saveCommandId
      || !['confirmed', 'rejected'].includes(result.status)) {
      throw new PptxEditorSessionError('session.invalidResolution', 'No matching unknown save to resolve');
    }
    this.#settle(result);
  }

  /** Explicitly discard local history and adopt an authoritative document. */
  resync(presentation: Presentation): PptxEditorSessionChange {
    this.#assertActive();
    if (this.#saveStatus === 'saving') throw locked();
    const previous = this.#current.presentation;
    this.#current = this.#saved = { presentation };
    this.#savedAncestors = new Set([this.#saved]);
    this.#undoDepth = 0;
    this.#redo.length = 0;
    this.#saveStatus = 'idle';
    this.#saveCommandId = undefined;
    this.#saveError = undefined;
    return this.#publish(REASONS.PRESENTATION_RESYNCED, {
      changedSlideIds: [...new Set([...previous.slides, ...presentation.slides].map(getSlideMutationId))],
      changedElements: [],
    });
  }

  dispose(): void {
    this.#disposed = true;
    this.#listeners.clear();
    this.#redo.length = 0;
    this.#current = this.#saved = { presentation: this.#current.presentation };
    this.#savedAncestors.clear();
  }

  #settle(result: OfficeCliBatchSendResult): void {
    if (result.status === 'confirmed') {
      this.#saved = this.#current;
      this.#savedAncestors.clear();
      for (let node: HistoryNode | undefined = this.#saved; node; node = node.parent) {
        this.#savedAncestors.add(node);
      }
      this.#undoDepth = 0;
      for (let node = this.#current; this.#canUndo(node); node = node.parent as HistoryNode) {
        this.#undoDepth++;
      }
      this.#saveStatus = 'idle';
      this.#saveError = undefined;
    } else {
      this.#saveStatus = result.status === 'rejected' ? 'failed' : 'unknown';
      this.#saveError = result.cause;
    }
    if (!this.#disposed) this.#publish(REASONS.SAVE_CHANGED);
  }

  #unsavedMutations(): Mutation[] {
    const currentAncestors = new Set<HistoryNode>();
    for (let node: HistoryNode | undefined = this.#current; node; node = node.parent) {
      currentAncestors.add(node);
    }
    const mutations: Mutation[] = [];
    let common = this.#saved;
    while (!currentAncestors.has(common)) {
      if (!common.inverse || !common.parent) throw historyEmpty();
      mutations.push(...common.inverse);
      common = common.parent;
    }
    const forward: HistoryNode[] = [];
    for (let node = this.#current; node !== common; node = node.parent as HistoryNode) forward.push(node);
    for (const node of forward.reverse()) mutations.push(...(node.command as Command).mutations);
    return mutations;
  }

  #canUndo(node: HistoryNode): boolean {
    if (!node.parent) return false;
    if (node.inverse) return true;
    // Local-only deletes can be undone by restoring the previous snapshot.
    // Once saved, operations without an OfficeCLI inverse form a history barrier.
    return !this.#savedAncestors.has(node);
  }

  #createSnapshot(): PptxEditorSessionSnapshot {
    const canEdit = this.#saveStatus !== 'saving' && this.#saveStatus !== 'unknown';
    return Object.freeze({
      presentation: this.#current.presentation,
      dirty: this.#current !== this.#saved,
      saveStatus: this.#saveStatus,
      saveCommandId: this.#saveCommandId,
      saveError: this.#saveError,
      canEdit,
      undoDepth: this.#undoDepth,
      redoDepth: this.#redo.length,
      canUndo: canEdit && this.#undoDepth > 0,
      canRedo: canEdit && this.#redo.length > 0,
    });
  }

  #publish(
    reason: PptxEditorSessionChange['reason'],
    change?: Pick<CommandExecutionResult, 'changedSlideIds' | 'changedElements'> & { commandId?: string },
  ): PptxEditorSessionChange {
    this.#snapshot = this.#createSnapshot();
    const event = Object.freeze({
      reason,
      snapshot: this.#snapshot,
      commandId: change?.commandId ?? this.#saveCommandId,
      changedSlideIds: Object.freeze([...(change?.changedSlideIds ?? [])]),
      changedElements: Object.freeze([...(change?.changedElements ?? [])]),
    });
    for (const listener of [...this.#listeners]) {
      try {
        listener(event);
      } catch (cause) {
        try {
          (this.#options.onListenerError ?? reportListenerError)(cause, event);
        } catch (reportingCause) {
          reportListenerError(new AggregateError([cause, reportingCause]));
        }
      }
    }
    return event;
  }

  #assertActive(): void {
    if (this.#disposed) throw new PptxEditorSessionError('session.disposed', 'Cannot use a disposed PPTX editor session');
  }

  #assertEditable(): void {
    this.#assertActive();
    if (this.#saveStatus === 'saving' || this.#saveStatus === 'unknown') throw locked();
  }
}

function locked(): PptxEditorSessionError {
  return new PptxEditorSessionError('session.locked', 'Editing and saving are locked until the current save outcome is known');
}
function historyEmpty(): PptxEditorSessionError {
  return new PptxEditorSessionError('session.historyEmpty', 'No available history operation');
}
function reportListenerError(cause: unknown): void {
  console.error('PPTX editor session listener failed', cause);
}
