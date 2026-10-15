/** Resolve a bundler-emitted asset before passing it to an inline worker. */
export function resolveBundledAssetUrl(assetUrl: string, ownerUrl: string): string {
  // Bundlers may emit root-relative assets while retaining a build-time file:
  // import.meta.url (e.g. Turbopack). Use the calling page for these paths only
  // when the owner is file: or cannot base a root path (WHATWG URL semantics).
  // A valid HTTP/CDN owner keeps its own origin, including for root paths.
  if (assetUrl.startsWith('/') && globalThis.location?.href) {
    let ownerRoot: URL | undefined;
    try { ownerRoot = new URL('/', ownerUrl); } catch { /* Opaque build owner. */ }
    if (!ownerRoot || ownerRoot.protocol === 'file:') {
      try {
        return new URL(assetUrl, globalThis.location.href).href;
      } catch {
        // A blob:/data: worker or about:blank page cannot base a root path.
        // Retain module-owner resolution (or its ordinary invalid-URL error).
      }
    }
  }
  // Module-relative and explicit absolute assets, including Node file: URLs,
  // retain their original semantics. Do not reinterpret an absolute file URL.
  return new URL(assetUrl, ownerUrl).href;
}
