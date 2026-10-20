import { describe, expect, it } from 'vitest';
import { announcements, latestAnnouncements } from './announcements';

describe('announcements', () => {
  it('keeps stable, unique routes in newest-first order', () => {
    expect(new Set(announcements.map(({ slug }) => slug)).size).toBe(announcements.length);
    expect(announcements.every(({ slug }) => /^[a-z0-9-]+$/.test(slug))).toBe(true);
    expect(announcements.map(({ date }) => date)).toEqual(
      [...announcements].map(({ date }) => date).sort().reverse(),
    );
  });

  it('uses the leading announcements on the home page', () => {
    expect(latestAnnouncements).toEqual(announcements.slice(0, 3));
  });
});

describe('inlineSegments', () => {
  it('renders paired backticks as code and keeps an unpaired one literal', async () => {
    const { inlineSegments } = await import('./announcements');
    expect(inlineSegments('use `not-ooxml` here')).toEqual([
      { text: 'use ', code: false }, { text: 'not-ooxml', code: true }, { text: ' here', code: false },
    ]);
    expect(inlineSegments('a ` b')).toEqual([{ text: 'a ` b', code: false }]);
    expect(inlineSegments('plain')).toEqual([{ text: 'plain', code: false }]);
  });
});
