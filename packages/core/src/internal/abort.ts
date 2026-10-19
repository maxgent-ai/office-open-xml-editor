/**
 * Shared cancellation errors. Every session and pull path reports an aborted
 * operation as an `Error` named `AbortError` (the DOM `AbortSignal` convention)
 * whose message names the owning operation.
 */

/** An `AbortError` carrying the owning operation's message. */
export function createAbortError(message: string): Error {
  const error = new Error(message);
  error.name = 'AbortError';
  return error;
}

/** Throw {@link createAbortError} when `signal` is already aborted. */
export function throwIfAborted(signal: AbortSignal | undefined, message: string): void {
  if (!signal?.aborted) return;
  throw createAbortError(message);
}
