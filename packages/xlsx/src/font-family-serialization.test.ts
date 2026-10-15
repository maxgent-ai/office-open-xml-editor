import { describe, expect, it } from 'vitest';
import { fontStackFor } from './renderer.js';

describe('cell font family CSS serialization', () => {
  it('keeps authored and alias names as complete quoted families', () => {
    const families = ['Face "Quoted"', 'Alias \\31', 'Regular "Quoted"'];
    const stack = fontStackFor(families[0], undefined, '', undefined, false, families[1], families[2]);
    const strings = stack.match(/"(?:[^"\\]|\\.)*"/gu) ?? [];
    expect(strings.slice(0, 3).map((token) => JSON.parse(token) as string)).toEqual(families);
  });
});
