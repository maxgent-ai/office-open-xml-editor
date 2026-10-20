import type { LayoutDiagnostic } from './layout/types.js';

// Acquisition diagnostics belong to the document owner, never to its immutable
// parser model/layout. Keep one copy per contract/message across cached renders.
const diagnostics = new WeakMap<object, readonly LayoutDiagnostic[]>();
export function recordResourceDiagnostics(owner: object, incoming: readonly LayoutDiagnostic[] = []): void {
  const prior = diagnostics.get(owner) ?? [];
  const unique = new Map(prior.map(value => [JSON.stringify(value), value]));
  for (const value of incoming) unique.set(JSON.stringify(value), Object.freeze({ ...value }));
  diagnostics.set(owner, Object.freeze([...unique.values()]));
}
export function resourceDiagnosticsOf(owner: object): readonly LayoutDiagnostic[] {
  return diagnostics.get(owner) ?? [];
}
