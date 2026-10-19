/** A pull unit exists but cannot be split to fit the offered byte credit. */
export const PULL_SESSION_INSUFFICIENT_CREDIT_CODE = 'ooxml-insufficient-credit' as const;

const RUST_INSUFFICIENT_CREDIT_PREFIX = 'OOXML_INSUFFICIENT_CREDIT:';

interface RustInsufficientCreditPayload {
  readonly code: typeof PULL_SESSION_INSUFFICIENT_CREDIT_CODE;
  readonly requiredBytes: number;
  readonly offeredBytes: number;
}

export interface PullSessionInsufficientCreditDetails {
  readonly requiredBytes: number;
  readonly offeredBytes: number;
}

function isPositiveSafeInteger(value: unknown): value is number {
  return typeof value === 'number' && Number.isSafeInteger(value) && value > 0;
}

export function isPullSessionInsufficientCreditDetails(
  value: unknown,
): value is PullSessionInsufficientCreditDetails {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false;
  const details = value as Partial<PullSessionInsufficientCreditDetails>;
  return isPositiveSafeInteger(details.requiredBytes)
    && isPositiveSafeInteger(details.offeredBytes)
    && details.requiredBytes > details.offeredBytes;
}

/** Stable error reconstructed on either side of a worker boundary. */
export class PullSessionInsufficientCreditError extends RangeError {
  readonly code = PULL_SESSION_INSUFFICIENT_CREDIT_CODE;
  readonly requiredBytes: number;
  readonly offeredBytes: number;

  constructor(details: PullSessionInsufficientCreditDetails) {
    super(`Pull unit requires ${details.requiredBytes} bytes but credit is ${details.offeredBytes}`);
    this.name = 'PullSessionInsufficientCreditError';
    this.requiredBytes = details.requiredBytes;
    this.offeredBytes = details.offeredBytes;
    Object.setPrototypeOf(this, PullSessionInsufficientCreditError.prototype);
  }
}

/** Decode only the exact machine envelope emitted by the shared Rust parser. */
export function parsePullSessionInsufficientCreditError(
  error: unknown,
): PullSessionInsufficientCreditError | undefined {
  if (error instanceof PullSessionInsufficientCreditError) return error;
  const text = error instanceof Error ? error.message : String(error);
  if (!text.startsWith(RUST_INSUFFICIENT_CREDIT_PREFIX)) return undefined;
  let value: unknown;
  try {
    value = JSON.parse(text.slice(RUST_INSUFFICIENT_CREDIT_PREFIX.length)) as unknown;
  } catch {
    return undefined;
  }
  if (!value || typeof value !== 'object' || Array.isArray(value)) return undefined;
  const payload = value as Partial<RustInsufficientCreditPayload>;
  if (
    payload.code !== PULL_SESSION_INSUFFICIENT_CREDIT_CODE
    || !isPullSessionInsufficientCreditDetails(payload)
  ) return undefined;
  return new PullSessionInsufficientCreditError(payload as RustInsufficientCreditPayload);
}

/** Validate a parser envelope against the request and format hard ceiling. */
export function normalizePullSessionInsufficientCreditError(
  error: unknown,
  offeredBytes: number,
  maximumRequiredBytes: number,
): PullSessionInsufficientCreditError | undefined {
  const parsed = parsePullSessionInsufficientCreditError(error);
  if (
    !parsed
    || parsed.offeredBytes !== offeredBytes
    || parsed.requiredBytes > maximumRequiredBytes
  ) return undefined;
  return parsed;
}

/** Return validated retry credit for a typed error, or `undefined`. */
export function requiredPullCredit(
  error: unknown,
  offeredBytes: number,
  maximumRequiredBytes: number,
): number | undefined {
  if (!isPullSessionInsufficientCreditError(error)) return undefined;
  return error.offeredBytes === offeredBytes
    && error.requiredBytes <= maximumRequiredBytes
    ? error.requiredBytes
    : undefined;
}

/**
 * Pull one unit with `initialBytes` of credit. When the unit is indivisible
 * and needs at most `maximumRequiredBytes`, retry once with exactly the
 * required credit (same sequence); any other failure propagates.
 */
export async function pullWithCreditRetry<TChunk>(
  session: { pull(byteCredit: number, options?: { signal?: AbortSignal }): Promise<TChunk> },
  initialBytes: number,
  maximumRequiredBytes: number,
  signal?: AbortSignal,
): Promise<TChunk> {
  try {
    return await session.pull(initialBytes, { signal });
  } catch (error) {
    const required = requiredPullCredit(error, initialBytes, maximumRequiredBytes);
    if (required === undefined) throw error;
    return session.pull(required, { signal });
  }
}

export function isPullSessionInsufficientCreditError(
  error: unknown,
): error is PullSessionInsufficientCreditError {
  return error instanceof PullSessionInsufficientCreditError
    || (!!error && typeof error === 'object'
      && (error as { code?: unknown }).code === PULL_SESSION_INSUFFICIENT_CREDIT_CODE
      && isPullSessionInsufficientCreditDetails(error));
}
