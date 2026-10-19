export interface Binary64Dyadic {
  readonly coefficient: bigint;
  readonly exponent: number;
}

export interface ExactRational {
  readonly numerator: bigint;
  readonly denominator: bigint;
}

function absolute(value: bigint): bigint {
  return value < 0n ? -value : value;
}

function greatestCommonDivisor(left: bigint, right: bigint): bigint {
  let a = absolute(left);
  let b = absolute(right);
  while (b !== 0n) {
    const remainder = a % b;
    a = b;
    b = remainder;
  }
  return a;
}

export function normalizeExactRational(
  numerator: bigint,
  denominator: bigint,
): ExactRational {
  if (denominator === 0n) throw new Error('Exact rational denominator must be nonzero');
  if (numerator === 0n) return Object.freeze({ numerator: 0n, denominator: 1n });
  const sign = denominator < 0n ? -1n : 1n;
  const divisor = greatestCommonDivisor(numerator, denominator);
  return Object.freeze({
    numerator: sign * numerator / divisor,
    denominator: sign * denominator / divisor,
  });
}

export function compareExactRational(
  left: ExactRational,
  right: ExactRational,
): number {
  const delta = left.numerator * right.denominator
    - right.numerator * left.denominator;
  return delta < 0n ? -1 : delta > 0n ? 1 : 0;
}

export function addExactRational(
  left: ExactRational,
  right: ExactRational,
): ExactRational {
  return normalizeExactRational(
    left.numerator * right.denominator + right.numerator * left.denominator,
    left.denominator * right.denominator,
  );
}

export function subtractExactRational(
  left: ExactRational,
  right: ExactRational,
): ExactRational {
  return normalizeExactRational(
    left.numerator * right.denominator - right.numerator * left.denominator,
    left.denominator * right.denominator,
  );
}

export function midpointExactRational(
  left: ExactRational,
  right: ExactRational,
): ExactRational {
  return normalizeExactRational(
    left.numerator * right.denominator + right.numerator * left.denominator,
    2n * left.denominator * right.denominator,
  );
}

export function exactRationalKey(value: ExactRational): string {
  return `${value.numerator}/${value.denominator}`;
}

const binary64Buffer = new ArrayBuffer(8);
const binary64View = new DataView(binary64Buffer);

export function decodeBinary64(value: number): Binary64Dyadic {
  if (!Number.isFinite(value)) throw new Error('Exact geometry requires a finite binary64 value');
  if (value === 0) return Object.freeze({ coefficient: 0n, exponent: 0 });
  binary64View.setFloat64(0, value, false);
  const bits = binary64View.getBigUint64(0, false);
  const negative = (bits >> 63n) !== 0n;
  const exponentBits = Number((bits >> 52n) & 0x7ffn);
  const fraction = bits & ((1n << 52n) - 1n);
  let coefficient = exponentBits === 0 ? fraction : (1n << 52n) | fraction;
  let exponent = exponentBits === 0 ? -1074 : exponentBits - 1023 - 52;
  while ((coefficient & 1n) === 0n) {
    coefficient >>= 1n;
    exponent += 1;
  }
  return Object.freeze({
    coefficient: negative ? -coefficient : coefficient,
    exponent,
  });
}

export function exactRationalFromNumber(value: number): ExactRational {
  const decoded = decodeBinary64(value);
  if (decoded.exponent >= 0) {
    return normalizeExactRational(
      decoded.coefficient << BigInt(decoded.exponent),
      1n,
    );
  }
  return normalizeExactRational(
    decoded.coefficient,
    1n << BigInt(-decoded.exponent),
  );
}

export function scaleExactRationalByPowerOfTwo(
  value: ExactRational,
  exponent: number,
): ExactRational {
  return exponent >= 0
    ? normalizeExactRational(value.numerator << BigInt(exponent), value.denominator)
    : normalizeExactRational(value.numerator, value.denominator << BigInt(-exponent));
}

function bitLength(value: bigint): number {
  return value === 0n ? 0 : value.toString(2).length;
}

function compareWithPowerOfTwo(
  numerator: bigint,
  denominator: bigint,
  exponent: number,
): number {
  const left = exponent >= 0 ? numerator : numerator << BigInt(-exponent);
  const right = exponent >= 0 ? denominator << BigInt(exponent) : denominator;
  return left < right ? -1 : left > right ? 1 : 0;
}

function roundedQuotient(
  numerator: bigint,
  denominator: bigint,
  shift: number,
): bigint {
  const scaledNumerator = shift >= 0
    ? numerator << BigInt(shift)
    : numerator;
  const scaledDenominator = shift >= 0
    ? denominator
    : denominator << BigInt(-shift);
  let quotient = scaledNumerator / scaledDenominator;
  const remainder = scaledNumerator % scaledDenominator;
  const doubled = remainder * 2n;
  if (doubled > scaledDenominator
    || (doubled === scaledDenominator && (quotient & 1n) !== 0n)) {
    quotient += 1n;
  }
  return quotient;
}

function binary64FromBits(bits: bigint): number {
  binary64View.setBigUint64(0, bits, false);
  return binary64View.getFloat64(0, false);
}

export function exactRationalToNumber(value: ExactRational): number {
  if (value.numerator === 0n) return 0;
  const negative = value.numerator < 0n;
  const numerator = absolute(value.numerator);
  const denominator = value.denominator;
  let exponent = bitLength(numerator) - bitLength(denominator);
  if (compareWithPowerOfTwo(numerator, denominator, exponent) < 0) exponent -= 1;
  const signBits = negative ? 1n << 63n : 0n;
  if (exponent < -1022) {
    const significand = roundedQuotient(numerator, denominator, 1074);
    if (significand === 0n) return binary64FromBits(signBits);
    if (significand >= 1n << 52n) {
      return binary64FromBits(signBits | 1n << 52n);
    }
    return binary64FromBits(signBits | significand);
  }
  let significand = roundedQuotient(numerator, denominator, 52 - exponent);
  if (significand === 1n << 53n) {
    significand >>= 1n;
    exponent += 1;
  }
  if (exponent > 1023) {
    return negative ? Number.NEGATIVE_INFINITY : Number.POSITIVE_INFINITY;
  }
  const exponentBits = BigInt(exponent + 1023) << 52n;
  const fractionBits = significand - (1n << 52n);
  return binary64FromBits(signBits | exponentBits | fractionBits);
}

function nextBinary64Up(value: number): number {
  if (value === Number.POSITIVE_INFINITY) return value;
  if (Object.is(value, -0) || value === 0) return Number.MIN_VALUE;
  binary64View.setFloat64(0, value, false);
  const bits = binary64View.getBigUint64(0, false);
  return binary64FromBits(value > 0 ? bits + 1n : bits - 1n);
}

/**
 * Convert an exact threshold to the least binary64 value not below it.
 * Unlike ties-to-even projection rounding, event search must never resume on
 * the infeasible side of an unrepresentable exact root.
 */
export function exactRationalToNumberUp(value: ExactRational): number {
  const nearest = exactRationalToNumber(value);
  if (nearest === Number.POSITIVE_INFINITY) return nearest;
  if (nearest === Number.NEGATIVE_INFINITY) return -Number.MAX_VALUE;
  const decoded = decodeBinary64(nearest);
  const exactNearest: ExactRational = decoded.exponent >= 0
    ? {
        numerator: decoded.coefficient << BigInt(decoded.exponent),
        denominator: 1n,
      }
    : {
        numerator: decoded.coefficient,
        denominator: 1n << BigInt(-decoded.exponent),
      };
  return compareExactRational(exactNearest, value) >= 0
    ? nearest
    : nextBinary64Up(nearest);
}

/** Greatest binary64 value not above an exact available extent. */
export function exactRationalToNumberDown(value: ExactRational): number {
  return -exactRationalToNumberUp({
    numerator: -value.numerator,
    denominator: value.denominator,
  });
}
