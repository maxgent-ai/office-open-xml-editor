import type { DocumentLayout } from './types.js';
import { LayoutInvariantError } from './diagnostics.js';

export interface LayoutIteration {
  readonly fingerprint: string;
  readonly pageCount: number;
  readonly layout?: DocumentLayout;
}

export class ExactConvergenceError extends LayoutInvariantError {
  readonly reason: 'cycle' | 'limit';
  readonly states: readonly string[];
  readonly passes: number;

  constructor(
    reason: 'cycle' | 'limit',
    states: readonly string[],
    passes: number,
  ) {
    super(
      'NON_CONVERGENCE',
      reason === 'cycle'
        ? `repeated exact-state cycle at ${states.at(-1) ?? '<missing>'}`
        : `hard exact-state pass limit ${passes} reached`,
    );
    this.name = 'ExactConvergenceError';
    this.reason = reason;
    this.states = Object.freeze([...states]);
    this.passes = passes;
  }
}

export interface ExactConvergenceOptions<T> {
  /** State observed before the first pass. It does not consume the pass budget. */
  readonly seedState?: string;
  /** Deterministic pass. `pass` is one-based and counts against `limit`. */
  readonly step: (previous: T | null, pass: number) => T;
  readonly stateOf: (value: T) => string;
  /** Maximum number of `step` calls, including the confirming fixed-point pass. */
  readonly limit: number;
}

/**
 * Converge any deterministic exact-state transition.
 *
 * Adjacent equality is a fixed point; a non-adjacent repeated state is a cycle.
 * The limit is an explicit resource guard, not a claim about the state-space
 * cardinality. Exhaustion fails closed and never returns the last candidate.
 */
export function convergeExactState<T>(
  options: ExactConvergenceOptions<T>,
): Readonly<{ value: T; passes: number }> {
  // One algorithm, driven to completion. The generator form below is the
  // implementation; this wrapper just refuses to suspend.
  const steps = convergeExactStateSteps<T, never>({
    ...options,
    step: function* generatorStep(previous, pass) {
      return options.step(previous, pass);
    },
  });
  let step = steps.next();
  while (!step.done) step = steps.next();
  return step.value;
}

/**
 * {@link convergeExactState} whose pass is itself suspendable.
 *
 * `step` is a generator, and its yields are forwarded to this generator's
 * consumer, so a caller driving convergence can spread each pass across
 * event-loop turns without the convergence policy — adjacent equality is a
 * fixed point, a non-adjacent repeat is a cycle, exhaustion fails closed —
 * being restated anywhere.
 *
 * `carry`, when supplied, projects each non-final value to the data the next
 * pass actually reads, and `step` receives that projection instead of the
 * value. Only the returned (fixed-point) value is ever needed whole, so a
 * caller whose passes build large graphs (a whole document layout) can let
 * the previous pass's graph die while the next one is built: this frame keeps
 * the carried projection, never the prior value itself.
 */
export function* convergeExactStateSteps<T, Y, C = T>(
  options: Omit<ExactConvergenceOptions<T>, 'step'> & {
    readonly step: (previous: C | null, pass: number) => Generator<Y, T, void>;
    readonly carry?: (value: T) => C;
  },
): Generator<Y, Readonly<{ value: T; passes: number }>, void> {
  const { seedState, step, stateOf, limit, carry } = options;
  const minimumLimit = seedState === undefined ? 2 : 1;
  if (!Number.isInteger(limit) || limit < minimumLimit) {
    throw new RangeError(
      `Exact convergence limit must be an integer >= ${minimumLimit}`,
    );
  }
  const states: string[] = seedState === undefined ? [] : [seedState];
  const seen = new Set(states);
  let previous: C | null = null;
  // One function-scoped slot, cleared before the next pass starts, so a
  // suspended frame never holds the prior value beyond its carried projection.
  let value: T | null = null;
  for (let pass = 1; pass <= limit; pass += 1) {
    value = yield* step(previous, pass);
    const state = stateOf(value);
    const priorState = states.at(-1);
    states.push(state);
    if (priorState === state) {
      return Object.freeze({ value, passes: pass });
    }
    if (seen.has(state)) {
      throw new ExactConvergenceError('cycle', states, pass);
    }
    seen.add(state);
    if (pass === limit) {
      throw new ExactConvergenceError('limit', states, pass);
    }
    previous = carry ? carry(value) : value as unknown as C;
    value = null;
  }
  throw new ExactConvergenceError('limit', states, limit);
}

export function convergeLayout<T extends LayoutIteration>(
  seed: T,
  step: (iteration: T) => T,
  limit: number,
): T {
  const steps = convergeLayoutSteps<T, never>(
    seed,
    function* generatorStep(iteration) { return step(iteration); },
    limit,
  );
  let next = steps.next();
  while (!next.done) next = steps.next();
  return next.value;
}

/**
 * {@link convergeLayout} whose step is suspendable; yields are forwarded.
 *
 * With `carry`, `step` receives `carry(iteration)` rather than the iteration,
 * and neither the seed nor any superseded iteration stays referenced by this
 * frame while a later pass builds (see {@link convergeExactStateSteps}).
 */
export function* convergeLayoutSteps<T extends LayoutIteration, Y, C = T>(
  seed: T,
  step: (carried: C) => Generator<Y, T, void>,
  limit: number,
  carry?: (iteration: T) => C,
): Generator<Y, T, void> {
  if (!Number.isInteger(limit) || limit < 1) {
    throw new LayoutInvariantError('NON_CONVERGENCE', 'limit must be a positive integer');
  }
  const seedState = seed.fingerprint;
  const seedCarried = carry ? carry(seed) : seed as unknown as C;
  // Parameters are frame slots too: drop the caller's seed so a carried
  // projection is all this suspended generator retains of it.
  seed = undefined as unknown as T;
  try {
    return (yield* convergeExactStateSteps<T, Y, C>({
      seedState,
      step: function* layoutStep(previous) { return yield* step(previous ?? seedCarried); },
      stateOf: (iteration) => iteration.fingerprint,
      ...(carry ? { carry } : {}),
      limit,
    })).value;
  } catch (error) {
    if (error instanceof ExactConvergenceError) {
      throw new LayoutInvariantError(
        'NON_CONVERGENCE',
        error.reason === 'cycle'
          ? `repeated geometry fingerprint cycle at ${error.states.at(-1) ?? '<missing>'}`
          : `hard iteration limit ${limit} reached`,
      );
    }
    throw error;
  }
}
