import * as vscode from 'vscode';
import {
  McpServerNotInstalledError,
  cachedBinaryPath,
  resolveBinaryPath,
} from './installer';
import { OoxmlMcpProvider, workspaceHasOoxmlFiles } from './provider';
import { ActiveContextBridge } from './selectionBridge';
import type { SelectionContextRegistry } from '../selectionContextRegistry';

const PROVIDER_ID = 'ooxml-mcp';
const NEVER_ASK_KEY = 'ooxmlViewer.mcp.neverAskAgain';

export function activateMcp(
  context: vscode.ExtensionContext,
  selectionContexts: SelectionContextRegistry,
): void {
  // The MCP API was finalised in VS Code 1.101. If the host runs an older
  // VS Code, skip MCP integration silently — the editor features still work.
  const lm = (vscode as unknown as { lm?: typeof vscode.lm }).lm;
  if (!lm || typeof lm.registerMcpServerDefinitionProvider !== 'function') {
    return;
  }

  const activeContextBridge = new ActiveContextBridge(selectionContexts);
  const provider = new OoxmlMcpProvider(context, activeContextBridge);
  context.subscriptions.push(
    activeContextBridge,
    lm.registerMcpServerDefinitionProvider(PROVIDER_ID, provider),
  );

  context.subscriptions.push(
    vscode.workspace.onDidChangeConfiguration((e) => {
      if (e.affectsConfiguration('ooxmlViewer.mcpServer')) {
        void (async () => {
          await provider.refresh();
          await promptIfNeeded(context, provider);
        })();
      }
    }),
    // The allowed roots follow the workspace folders.
    vscode.workspace.onDidChangeWorkspaceFolders(() => {
      void provider.refresh();
    }),
    vscode.workspace.onDidGrantWorkspaceTrust(() => {
      void provider.refresh();
      void promptIfNeeded(context, provider);
    }),
    vscode.commands.registerCommand('ooxmlViewer.installMcpServer', () =>
      installCommand(context, provider),
    ),
    vscode.commands.registerCommand('ooxmlViewer.disableMcpServer', () =>
      disableCommand(provider),
    ),
  );

  void promptIfNeeded(context, provider);
}

async function promptIfNeeded(
  context: vscode.ExtensionContext,
  provider: OoxmlMcpProvider,
): Promise<void> {
  if (!vscode.workspace.isTrusted) return;
  const cfg = vscode.workspace.getConfiguration('ooxmlViewer.mcpServer');
  const enabled = cfg.get<'auto' | 'always' | 'never'>('enabled', 'auto');
  if (enabled === 'never') return;

  if (enabled === 'auto') {
    if (!(await workspaceHasOoxmlFiles())) return;
    if (context.globalState.get<boolean>(NEVER_ASK_KEY)) return;
  }

  // Already resolvable (override / cache / PATH) — just refresh and exit.
  try {
    await resolveBinaryPath(context, {
      override: cfg.get<string>('binaryPath', ''),
      consentToDownload: false,
    });
    await provider.refresh();
    return;
  } catch (err) {
    if (!(err instanceof McpServerNotInstalledError)) {
      void vscode.window.showErrorMessage(
        `OOXML MCP server: ${(err as Error).message}`,
      );
      return;
    }
  }

  const enableLabel = 'Enable';
  const notNowLabel = 'Not now';
  const neverLabel = "Don't ask again";
  const choice = await vscode.window.showInformationMessage(
    'Enable the OOXML MCP server so GitHub Copilot Chat in Agent mode can read .xlsx/.docx/.pptx files and active Viewer context? A small (~5 MB) binary will be downloaded on first use.',
    enableLabel,
    notNowLabel,
    neverLabel,
  );

  if (choice === enableLabel) {
    await installCommand(context, provider);
  } else if (choice === neverLabel) {
    await context.globalState.update(NEVER_ASK_KEY, true);
  }
}

async function installCommand(
  context: vscode.ExtensionContext,
  provider: OoxmlMcpProvider,
): Promise<void> {
  if (!vscode.workspace.isTrusted) {
    void vscode.window.showWarningMessage(
      'Trust this workspace before enabling the OOXML MCP server.',
    );
    return;
  }
  const cfg = vscode.workspace.getConfiguration('ooxmlViewer.mcpServer');
  try {
    const path = await resolveBinaryPath(context, {
      override: cfg.get<string>('binaryPath', ''),
      consentToDownload: true,
    });
    // This command is an explicit opt-in. In particular, keep the provider
    // available when the active OOXML preview is outside the current workspace;
    // `auto` intentionally gates discovery on workspace contents.
    await cfg.update('enabled', 'always', vscode.ConfigurationTarget.Global);
    await provider.refresh();
    void vscode.window.showInformationMessage(
      `OOXML MCP server is ready (${path === cachedBinaryPath(context) ? 'downloaded' : 'using existing binary'}).`,
    );
  } catch (err) {
    void vscode.window.showErrorMessage(
      `Failed to install ooxml-mcp-server: ${(err as Error).message}`,
    );
  }
}

async function disableCommand(provider: OoxmlMcpProvider): Promise<void> {
  await vscode.workspace
    .getConfiguration('ooxmlViewer.mcpServer')
    .update('enabled', 'never', vscode.ConfigurationTarget.Global);
  await provider.refresh();
  void vscode.window.showInformationMessage(
    'OOXML MCP server disabled. Re-enable it via the "ooxmlViewer.mcpServer.enabled" setting.',
  );
}
