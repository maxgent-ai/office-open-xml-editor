import { describe, expect, it } from 'vitest';
import { fontStackFor } from './renderer.js';

describe('cell font family CSS serialization', () => {
  it('keeps authored and regular alias names as complete quoted families', () => {
    const families = ['Face "Quoted"', 'Regular \\31 "Quoted"'];
    const stack = fontStackFor(families[0], undefined, '', undefined, undefined, undefined, families[1]);
    const strings = stack.match(/"(?:[^"\\]|\\.)*"/gu) ?? [];
    expect(strings.slice(0, 2).map((token) => JSON.parse(token) as string)).toEqual(families);
  });
});
