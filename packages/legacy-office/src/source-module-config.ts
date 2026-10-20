/** Whole-input limits are caller opt-in; retained/inflated work has its own budgets. */
export function readLegacyMaxInputBytes(value: unknown, label: string): number | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== 'number' || !Number.isSafeInteger(value) || value <= 0) {
    throw new RangeError(`${label} maxInputBytes is invalid`);
  }
  return value;
}

/** Read the scalar config a legacy source module receives (fail closed). */
export function readLegacySourceModuleConfig(
  config: unknown,
  label: string,
): Readonly<{ wasmUrl: string; maxInputBytes?: number }> {
  if (typeof config !== 'object' || config === null) {
    throw new TypeError(`${label} source module config must be an object`);
  }
  const record = config as Record<string, unknown>;
  for (const key of Object.keys(record)) {
    if (key !== 'wasmUrl' && key !== 'maxInputBytes') {
      throw new TypeError(`${label} source module config has an unknown field`);
    }
  }
  const maxInputBytes = readLegacyMaxInputBytes(record.maxInputBytes, label);
  if (typeof record.wasmUrl !== 'string') {
    throw new TypeError(`${label} wasmUrl must be a nonempty absolute URL`);
  }
  return { wasmUrl: record.wasmUrl, ...(maxInputBytes === undefined ? {} : { maxInputBytes }) };
}
