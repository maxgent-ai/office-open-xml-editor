import { setFlagsFromString } from 'node:v8';
import { runInNewContext } from 'node:vm';
import { describe, expect, it } from 'vitest';
import {
  convergeHeaderFooterReserveSteps,
  convergeHeaderFooterReserves,
  headerFooterOverflowReservePt,
} from './header-footer-reserve.js';

setFlagsFromString('--expose-gc');
const collectGarbage = runInNewContext('gc') as () => void;

describe('header/footer body reserve', () => {
  it('charges only overflow beyond the signed-margin allowance', () => {
    expect(headerFooterOverflowReservePt(30, 72, 36)).toBe(0);
    expect(headerFooterOverflowReservePt(48, 72, 36)).toBe(12);
    expect(headerFooterOverflowReservePt(120, -72, 36)).toBe(0);
  });

  it('returns the repaginated candidate when the exact next-pass inputs are stable', () => {
    type Candidate = Readonly<{
      geometryVersion: string;
      fieldContexts: readonly Readonly<{
        pageIndex: number;
        displayPageNumber: number;
        pageNumberFormat: string;
      }>[];
    }>;
    const fieldContexts = Object.freeze([
      Object.freeze({ pageIndex: 0, displayPageNumber: 1, pageNumberFormat: 'decimal' }),
    ]);
    const seed: Candidate = Object.freeze({ geometryVersion: 'seed', fieldContexts });
    let repaginations = 0;

    const result = convergeHeaderFooterReserves({
      seed,
      measure: () => [Object.freeze({ top: 12, bottom: 0 })],
      repaginate: (_reserves, current) => {
        repaginations += 1;
        return Object.freeze({ ...current, geometryVersion: 'repaginated' });
      },
      identity: (candidate) => candidate.fieldContexts,
    });

    expect(result.result.geometryVersion).toBe('repaginated');
    expect(repaginations).toBe(1);
  });

  it('repaginates from carried data without keeping the seed pass alive', async () => {
    type Candidate = Readonly<{ version: number; pages: number; payload: object }>;
    let seedPayload: WeakRef<object> | undefined;
    const makeSeed = (): Candidate => {
      const payload = {};
      seedPayload = new WeakRef(payload);
      return Object.freeze({ version: 0, pages: 3, payload });
    };
    const carriedInputs: unknown[] = [];
    const steps = convergeHeaderFooterReserveSteps<Candidate, 'suspended', number>(makeSeed(), {
      measure: (candidate) => Array.from({ length: candidate.pages }, () => Object.freeze({ top: 12, bottom: 0 })),
      carry: (candidate) => candidate.pages,
      repaginate: function* repaginate(reserves, pages) {
        carriedInputs.push([reserves.length, pages]);
        yield 'suspended';
        return Object.freeze({ version: carriedInputs.length, pages, payload: {} });
      },
      identity: (candidate) => candidate.pages,
    });

    expect(steps.next()).toEqual({ done: false, value: 'suspended' });
    await new Promise((resolve) => setTimeout(resolve, 0));
    collectGarbage();
    expect(seedPayload!.deref()).toBeUndefined();
    let step = steps.next();
    while (!step.done) step = steps.next();

    expect(carriedInputs).toEqual([[3, 3]]);
    expect(step.value.result.version).toBe(1);
    expect(step.value.reserves).toHaveLength(3);
  });
});
