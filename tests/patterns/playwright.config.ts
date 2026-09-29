import { defineConfig } from '@playwright/test';

export default defineConfig({
  testDir: '.',
  testMatch: '*.spec.ts',
  workers: 1,
  projects: [{
    name: 'chrome',
    use: { channel: 'chrome', deviceScaleFactor: 1 },
  }],
});
