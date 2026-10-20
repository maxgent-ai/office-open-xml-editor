/**
 * Worker-realm pieces shared by the DOCX / XLSX / PPTX source owners. Each
 * owner holds exactly one archive — the format parser's or one opened by an
 * application model source — and keeps its own lifecycle (pending opens, view
 * defaults, host layout); only the mechanics below are format-neutral.
 */
import { withWasmResourceErrors } from '../worker/error-wire.js';
import { isWasmTrap } from '../worker/wasm-guard.js';
import { copyModelSourceBytes } from './model-source.js';

/**
 * Wrap a model-source archive so its pull-cursor `method` returns bytes copied
 * into a fresh, exactly-sized buffer. The ordinary pull coordinator transfers
 * cursor payloads; a source archive's bytes may alias memory the source still
 * owns, so normalizing here keeps the coordinator's transfer path unchanged.
 */
export function copyingCursorArchive<TArchive extends object>(
  archive: TArchive,
  method: keyof TArchive & string,
): TArchive {
  return new Proxy(archive, {
    get(target, key) {
      if (key === method) {
        return (...args: unknown[]) => new Uint8Array(copyModelSourceBytes(
          Reflect.apply(target[method] as (...callArgs: unknown[]) => Uint8Array, target, args),
        ));
      }
      const value = Reflect.get(target, key);
      return typeof value === 'function' ? value.bind(target) : value;
    },
  });
}

/**
 * Run one operation against a model-source archive under the resource-error
 * boundary. A WASM trap leaves the source unusable, so `onTrap` closes it (its
 * own failure is ignored) before the trap propagates.
 */
export function runModelSourceOperation<T>(operation: () => T, onTrap: () => void): T {
  try {
    return withWasmResourceErrors(operation);
  } catch (error) {
    if (isWasmTrap(error)) {
      try { onTrap(); } catch {}
    }
    throw error;
  }
}
