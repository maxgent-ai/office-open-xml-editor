import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { SelectionContextRegistry } from '../selectionContextRegistry';

const mocks = vi.hoisted(() => ({
  resolveBinaryPath: vi.fn(),
  enabled: 'always' as 'auto' | 'always' | 'never',
  trusted: true,
  restrictToWorkspaceFolders: true,
  workspaceFolders: [] as { uri: { scheme: string; fsPath: string } }[] | undefined,
}));

vi.mock('vscode', () => ({
  EventEmitter: class {
    readonly event = () => undefined;
    fire(): void {}
  },
  McpStdioServerDefinition: class {
    constructor(
      readonly label: string,
      readonly command: string,
      readonly args: string[],
      readonly env: Record<string, string>,
      readonly version: string,
    ) {}
  },
  workspace: {
    get isTrusted() { return mocks.trusted; },
    get workspaceFolders() { return mocks.workspaceFolders; },
    getConfiguration: () => ({
      get: (key: string, fallback: unknown) => {
        if (key === 'enabled') return mocks.enabled;
        if (key === 'binaryPath') return '';
        if (key === 'restrictToWorkspaceFolders') return mocks.restrictToWorkspaceFolders;
        return fallback;
      },
    }),
    findFiles: vi.fn().mockResolvedValue([{ fsPath: '/tmp/a.xlsx' }]),
  },
}));

vi.mock('./installer', () => ({
  McpServerNotInstalledError: class extends Error {},
  resolveBinaryPath: mocks.resolveBinaryPath,
}));

import { OoxmlMcpProvider } from './provider';
import {
  ActiveContextBridge,
  ACTIVE_CONTEXT_BRIDGE_PORT_ENV,
  ACTIVE_CONTEXT_BRIDGE_TOKEN_ENV,
} from './selectionBridge';

const liveBridges: ActiveContextBridge[] = [];

afterEach(async () => {
  await Promise.all(liveBridges.splice(0).map((bridge) => bridge.close()));
});

describe('OoxmlMcpProvider selection bridge', () => {
  beforeEach(() => {
    mocks.enabled = 'always';
    mocks.trusted = true;
    mocks.restrictToWorkspaceFolders = true;
    mocks.workspaceFolders = [];
    mocks.resolveBinaryPath.mockReset().mockResolvedValue('/tmp/ooxml-mcp-server');
  });

  it('passes only the authenticated loopback endpoint credentials to the MCP child', async () => {
    const provider = new OoxmlMcpProvider(
      {
        extension: { packageJSON: { version: '0.77.0' } },
      } as never,
      {
        environment: vi.fn().mockResolvedValue({
          [ACTIVE_CONTEXT_BRIDGE_PORT_ENV]: '49152',
          [ACTIVE_CONTEXT_BRIDGE_TOKEN_ENV]: 'a'.repeat(64),
        }),
      } as never,
    );

    const definitions = await provider.provideMcpServerDefinitions({} as never);
    expect(definitions).toHaveLength(1);
    expect(definitions[0]).toMatchObject({
      command: '/tmp/ooxml-mcp-server',
      env: {
        RUST_LOG: 'warn',
        [ACTIVE_CONTEXT_BRIDGE_PORT_ENV]: '49152',
        [ACTIVE_CONTEXT_BRIDGE_TOKEN_ENV]: 'a'.repeat(64),
      },
      version: '0.77.0',
    });
  });

  it('confines file tools to the local workspace folders unless the user opts out', async () => {
    mocks.workspaceFolders = [
      { uri: { scheme: 'file', fsPath: '/work/a' } },
      { uri: { scheme: 'vscode-remote', fsPath: '/remote/b' } },
      { uri: { scheme: 'file', fsPath: '/work/c:d' } },
    ];
    const provider = new OoxmlMcpProvider(
      { extension: { packageJSON: { version: '0.77.0' } } } as never,
      { environment: vi.fn().mockResolvedValue({}) } as never,
    );

    const [scoped] = await provider.provideMcpServerDefinitions({} as never);
    expect(scoped).toMatchObject({
      args: ['--allowed-root', '/work/a', '--allowed-root', '/work/c:d'],
    });

    mocks.restrictToWorkspaceFolders = false;
    const [unscoped] = await provider.provideMcpServerDefinitions({} as never);
    expect(unscoped).toMatchObject({ args: [] });

    mocks.restrictToWorkspaceFolders = true;
    mocks.workspaceFolders = undefined;
    const [noFolder] = await provider.provideMcpServerDefinitions({} as never);
    expect(noFolder).toMatchObject({ args: [] });
  });

  it('does not start the bridge when MCP registration is disabled', async () => {
    mocks.enabled = 'never';
    const environment = vi.fn();
    const stop = vi.fn().mockResolvedValue(undefined);
    const provider = new OoxmlMcpProvider({} as never, { environment, stop } as never);

    expect(await provider.provideMcpServerDefinitions({} as never)).toEqual([]);
    expect(environment).not.toHaveBeenCalled();
    expect(stop).toHaveBeenCalledOnce();
  });

  it('does not resolve or launch an MCP executable in an untrusted workspace', async () => {
    mocks.trusted = false;
    const environment = vi.fn();
    const stop = vi.fn().mockResolvedValue(undefined);
    const provider = new OoxmlMcpProvider({} as never, { environment, stop } as never);

    expect(await provider.provideMcpServerDefinitions({} as never)).toEqual([]);
    expect(mocks.resolveBinaryPath).not.toHaveBeenCalled();
    expect(environment).not.toHaveBeenCalled();
    expect(stop).toHaveBeenCalledOnce();
  });

  it('resets a failed bridge start so a later provider refresh can retry', async () => {
    const stop = vi.fn().mockResolvedValue(undefined);
    const provider = new OoxmlMcpProvider(
      { extension: { packageJSON: { version: '0.77.0' } } } as never,
      {
        environment: vi.fn().mockRejectedValue(new Error('listen failed')),
        stop,
      } as never,
    );

    await expect(provider.provideMcpServerDefinitions({} as never))
      .rejects.toThrow('listen failed');
    expect(stop).toHaveBeenCalledOnce();
  });

  it('does not publish a stale definition when disable wins during binary resolution', async () => {
    let finishResolution!: (path: string) => void;
    mocks.resolveBinaryPath.mockReturnValue(new Promise<string>((resolve) => {
      finishResolution = resolve;
    }));
    const environment = vi.fn().mockResolvedValue({
      [ACTIVE_CONTEXT_BRIDGE_PORT_ENV]: '49152',
      [ACTIVE_CONTEXT_BRIDGE_TOKEN_ENV]: 'a'.repeat(64),
    });
    const stop = vi.fn().mockResolvedValue(undefined);
    const provider = new OoxmlMcpProvider(
      { extension: { packageJSON: { version: '0.77.0' } } } as never,
      { environment, stop } as never,
    );

    const staleEnabled = provider.provideMcpServerDefinitions({
      isCancellationRequested: false,
    } as never);
    mocks.enabled = 'never';
    const disabled = provider.provideMcpServerDefinitions({
      isCancellationRequested: false,
    } as never);
    finishResolution('/tmp/ooxml-mcp-server');

    expect(await disabled).toEqual([]);
    expect(await staleEnabled).toEqual([]);
    expect(environment).not.toHaveBeenCalled();
    expect(stop).toHaveBeenCalledOnce();
  });

  it('re-enables successfully while a previous disable is still stopping the bridge', async () => {
    const bridge = new ActiveContextBridge(new SelectionContextRegistry());
    liveBridges.push(bridge);
    const provider = new OoxmlMcpProvider(
      { extension: { packageJSON: { version: '0.77.0' } } } as never,
      bridge,
    );
    const token = { isCancellationRequested: false } as never;
    expect(await provider.provideMcpServerDefinitions(token)).toHaveLength(1);

    mocks.enabled = 'never';
    const disabling = provider.provideMcpServerDefinitions(token);
    mocks.enabled = 'always';
    const enabling = provider.provideMcpServerDefinitions(token);

    expect(await disabling).toEqual([]);
    expect(await enabling).toHaveLength(1);
  });
});
