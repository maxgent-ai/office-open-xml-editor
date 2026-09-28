import { describe, expect, it, vi } from 'vitest';
import {
  PULL_SESSION_PROTOCOL,
  type PullSessionCommand,
  type PullSessionResponse,
  type WorkerBridgeTransport,
} from '@silurus/ooxml-core/worker';
import { XlsxWorksheetPullClient } from './worksheet-pull-client.js';

describe('XlsxWorksheetPullClient', () => {
  it('delivers the metadata preview before row chunks', async () => {
    const transport = makeTransport((command) => command.kind === 'pull'
      ? chunkResponse(command, command.sequence === 0
        ? { kind: 'preview', worksheet: terminalWorksheet(), reason: null, maxRow: 70_000, maxCol: 1 }
        : command.sequence === 1
          ? { kind: 'rows', rows: [{ index: 1, height: null, cells: [] }] }
          : { kind: 'finished', worksheet: terminalWorksheet() }, command.sequence === 2)
      : acceptedResponse(command));
    const client = new XlsxWorksheetPullClient({
      transport, sharedStrings: [], open: async () => undefined,
    });
    const kinds: string[] = [];
    for await (const unit of client.stream(0, 'Sheet1')) kinds.push(unit.kind);
    expect(kinds).toEqual(['preview', 'rows', 'finished']);
  });

  it('cancels an unknown preview blocker instead of treating it as an exact sheet', async () => {
    const commands: PullSessionCommand<number>[] = [];
    const transport = makeTransport((command) => {
      commands.push(command);
      return command.kind === 'pull'
        ? chunkResponse(command, {
            kind: 'preview', worksheet: null, reason: 'unrecognized', maxRow: 1, maxCol: 1,
          }, false)
        : acceptedResponse(command);
    });
    const client = new XlsxWorksheetPullClient({
      transport, sharedStrings: [], open: async () => undefined,
    });
    await expect(client.stream(0, 'Sheet1').next()).rejects.toThrow(/preview unit is invalid/);
    expect(commands.some((command) => command.kind === 'cancel')).toBe(true);
  });

  it('owns decode, shared-string normalization, ACK order, and transfer disposal', async () => {
    const commands: PullSessionCommand<number>[] = [];
    const disposeTransferred = vi.fn();
    const transport = makeTransport((command) => {
      commands.push(command);
      if (command.kind === 'pull') {
        const value = command.sequence === 0
          ? {
              kind: 'rows',
              rows: [{
                index: 1,
                height: null,
                cells: [{ row: 1, col: 1, value: { type: 'shared', si: 0 } }],
              }],
            }
          : { kind: 'finished', worksheet: terminalWorksheet() };
        return chunkResponse(command, value, command.sequence === 1);
      }
      return acceptedResponse(command);
    });
    const client = new XlsxWorksheetPullClient({
      transport,
      sharedStrings: [{ text: 'resolved' }],
      open: async () => undefined,
      disposeTransferred,
    });
    const iterator = client.stream(0, 'Sheet1');

    const first = await iterator.next();
    expect(first.value).toMatchObject({
      kind: 'rows',
      rows: [{ cells: [{ value: { type: 'text', text: 'resolved' } }] }],
    });
    expect(commands.filter((command) => command.kind === 'ack')).toHaveLength(0);
    expect(disposeTransferred).not.toHaveBeenCalled();

    const terminal = await iterator.next();
    expect(terminal.value).toMatchObject({ kind: 'finished', worksheet: { rows: [] } });
    expect(commands.filter((command) => command.kind === 'ack')).toHaveLength(1);
    expect(disposeTransferred).toHaveBeenCalledTimes(1);

    await expect(iterator.next()).resolves.toEqual({ done: true, value: undefined });
    expect(commands.filter((command) => command.kind === 'ack')).toHaveLength(2);
    expect(disposeTransferred).toHaveBeenCalledTimes(2);
    expect(commands.some((command) => command.kind === 'cancel')).toBe(false);
  });

  it('cancels without ACK and disposes the outstanding transfer on consumer return', async () => {
    const commands: PullSessionCommand<number>[] = [];
    const disposeTransferred = vi.fn();
    const transport = makeTransport((command) => {
      commands.push(command);
      if (command.kind === 'pull') {
        return chunkResponse(command, { kind: 'rows', rows: [] }, false);
      }
      return acceptedResponse(command);
    });
    const client = new XlsxWorksheetPullClient({
      transport,
      sharedStrings: [],
      open: async () => undefined,
      disposeTransferred,
    });
    const iterator = client.stream(0, 'Sheet1');

    await iterator.next();
    await iterator.return();

    expect(commands.filter((command) => command.kind === 'ack')).toHaveLength(0);
    expect(commands.filter((command) => command.kind === 'cancel')).toHaveLength(1);
    expect(disposeTransferred).toHaveBeenCalledOnce();
  });

  it('rejects a mismatched terminal marker through the same cancellation path', async () => {
    const commands: PullSessionCommand<number>[] = [];
    const disposeTransferred = vi.fn();
    const transport = makeTransport((command) => {
      commands.push(command);
      if (command.kind === 'pull') {
        return chunkResponse(command, { kind: 'rows', rows: [] }, true);
      }
      return acceptedResponse(command);
    });
    const client = new XlsxWorksheetPullClient({
      transport,
      sharedStrings: [],
      open: async () => undefined,
      disposeTransferred,
    });

    await expect(client.stream(0, 'Sheet1').next())
      .rejects.toThrow(/terminal marker mismatch/);
    expect(commands.filter((command) => command.kind === 'ack')).toHaveLength(0);
    expect(commands.filter((command) => command.kind === 'cancel')).toHaveLength(1);
    expect(disposeTransferred).toHaveBeenCalledOnce();
  });
});

function makeTransport(
  respond: (command: PullSessionCommand<number>) => PullSessionResponse<ArrayBuffer, number>,
): WorkerBridgeTransport<PullSessionResponse<ArrayBuffer, number>> {
  let nextRequestId = 1;
  return {
    request: async (build) => respond(build(nextRequestId++) as PullSessionCommand<number>),
    forgetOrphaned: vi.fn(),
    terminate: vi.fn(),
  };
}

function chunkResponse(
  command: Extract<PullSessionCommand<number>, { kind: 'pull' }>,
  value: unknown,
  done: boolean,
): PullSessionResponse<ArrayBuffer, number> {
  const payload = new TextEncoder().encode(JSON.stringify(value)).buffer;
  return {
    protocol: PULL_SESSION_PROTOCOL,
    kind: 'chunk',
    sessionId: command.sessionId,
    operationId: command.operationId,
    generation: command.generation,
    requestId: command.requestId,
    sequence: command.sequence,
    byteLength: payload.byteLength,
    done,
    payload,
  };
}

function acceptedResponse(
  command: Exclude<PullSessionCommand<number>, { kind: 'pull' }>,
): PullSessionResponse<ArrayBuffer, number> {
  return {
    protocol: PULL_SESSION_PROTOCOL,
    kind: 'accepted',
    sessionId: command.sessionId,
    operationId: command.operationId,
    generation: command.generation,
    requestId: command.requestId,
    command: command.kind,
  };
}

function terminalWorksheet() {
  return {
    name: 'Sheet1',
    rows: [{ index: 99, height: null, cells: [] }],
    colWidths: {},
    rowHeights: {},
    defaultColWidth: 8.43,
    defaultRowHeight: 15,
    mergeCells: [],
    freezeRows: 0,
    freezeCols: 0,
    conditionalFormats: [],
    images: [],
    charts: [],
  };
}
