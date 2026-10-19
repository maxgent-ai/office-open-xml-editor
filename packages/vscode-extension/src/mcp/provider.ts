import * as vscode from 'vscode';
import { McpServerNotInstalledError, resolveBinaryPath } from './installer';
import { ActiveContextBridge } from './selectionBridge';

export class OoxmlMcpProvider implements vscode.McpServerDefinitionProvider {
  private readonly _onDidChange = new vscode.EventEmitter<void>();
  readonly onDidChangeMcpServerDefinitions = this._onDidChange.event;
  private requestGeneration = 0;

  constructor(
    private readonly context: vscode.ExtensionContext,
    private readonly activeContextBridge: ActiveContextBridge,
  ) {}

  async refresh(): Promise<void> {
    this.requestGeneration++;
    const cfg = vscode.workspace.getConfiguration('ooxmlViewer.mcpServer');
    if (
      !vscode.workspace.isTrusted ||
      cfg.get<'auto' | 'always' | 'never'>('enabled', 'auto') === 'never'
    ) {
      await this.activeContextBridge.stop();
    }
    this._onDidChange.fire();
  }

  async provideMcpServerDefinitions(
    token: vscode.CancellationToken,
  ): Promise<vscode.McpServerDefinition[]> {
    const generation = ++this.requestGeneration;
    if (!vscode.workspace.isTrusted) {
      if (this.isCurrent(generation, token)) await this.activeContextBridge.stop();
      return [];
    }
    const cfg = vscode.workspace.getConfiguration('ooxmlViewer.mcpServer');
    const enabled = cfg.get<'auto' | 'always' | 'never'>('enabled', 'auto');

    if (enabled === 'never') {
      if (this.isCurrent(generation, token)) await this.activeContextBridge.stop();
      return [];
    }
    if (enabled === 'auto' && !(await workspaceHasOoxmlFiles())) {
      if (this.isCurrent(generation, token)) await this.activeContextBridge.stop();
      return [];
    }

    let binPath: string;
    try {
      binPath = await resolveBinaryPath(this.context, {
        override: cfg.get<string>('binaryPath', ''),
        consentToDownload: false,
      });
    } catch (err) {
      if (err instanceof McpServerNotInstalledError) {
        // Activation flow handles the install prompt and calls refresh().
        if (this.isCurrent(generation, token)) await this.activeContextBridge.stop();
        return [];
      }
      if (this.isCurrent(generation, token)) await this.activeContextBridge.stop();
      throw err;
    }

    if (!this.isCurrent(generation, token)) return [];
    if (!vscode.workspace.isTrusted) {
      await this.activeContextBridge.stop();
      return [];
    }
    const currentCfg = vscode.workspace.getConfiguration('ooxmlViewer.mcpServer');
    const currentEnabled = currentCfg.get<'auto' | 'always' | 'never'>('enabled', 'auto');
    if (
      currentEnabled === 'never' ||
      (currentEnabled === 'auto' && !(await workspaceHasOoxmlFiles()))
    ) {
      if (this.isCurrent(generation, token)) await this.activeContextBridge.stop();
      return [];
    }
    if (!this.isCurrent(generation, token)) return [];
    if (!vscode.workspace.isTrusted) {
      await this.activeContextBridge.stop();
      return [];
    }

    let selectionEnvironment: Readonly<Record<string, string>>;
    try {
      selectionEnvironment = await this.activeContextBridge.environment();
    } catch (error) {
      if (this.isCurrent(generation, token)) await this.activeContextBridge.stop();
      throw error;
    }
    if (!this.isCurrent(generation, token)) return [];
    if (!vscode.workspace.isTrusted) {
      await this.activeContextBridge.stop();
      return [];
    }

    return [
      new vscode.McpStdioServerDefinition(
        'ooxml-mcp-server',
        binPath,
        allowedRootArgs(),
        { RUST_LOG: 'warn', ...selectionEnvironment },
        (this.context.extension.packageJSON as { version: string }).version,
      ),
    ];
  }

  private isCurrent(generation: number, token: vscode.CancellationToken): boolean {
    return generation === this.requestGeneration && !token.isCancellationRequested;
  }

  async resolveMcpServerDefinition(
    server: vscode.McpServerDefinition,
    _token: vscode.CancellationToken,
  ): Promise<vscode.McpServerDefinition> {
    return server;
  }
}

/**
 * `--allowed-root` arguments confining the server's path-based tools to the
 * local workspace folders. Tool paths come from the agent, which can be
 * steered by document content, so by default it may only read inside the
 * workspace. The arguments are omitted — leaving the server unscoped — when
 * the user turns `restrictToWorkspaceFolders` off or the window has no local
 * folder to scope to (for example a single file opened without a folder).
 * Roots travel as separate arguments rather than a delimiter-joined
 * environment variable so no folder name can split into extra roots.
 */
export function allowedRootArgs(): string[] {
  const restrict = vscode.workspace
    .getConfiguration('ooxmlViewer.mcpServer')
    .get<boolean>('restrictToWorkspaceFolders', true);
  if (!restrict) return [];
  return (vscode.workspace.workspaceFolders ?? [])
    .filter((folder) => folder.uri.scheme === 'file')
    .flatMap((folder) => ['--allowed-root', folder.uri.fsPath]);
}

export async function workspaceHasOoxmlFiles(): Promise<boolean> {
  const found = await vscode.workspace.findFiles(
    '**/*.{xlsx,docx,pptx}',
    '{**/node_modules/**,**/.git/**,**/dist/**,**/build/**,**/target/**,**/.venv/**}',
    1,
  );
  return found.length > 0;
}
