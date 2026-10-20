import { defineConfig } from '@playwright/test';

export default defineConfig({
  testDir: '.',
  testMatch: '*.spec.ts',
  workers: 1,
  projects: [{
    name: 'chrome',
    use: {
      channel: 'chrome',
      deviceScaleFactor: 1,
      // Main-thread and worker canvases are compared exactly. GPU raster can
      // classify a few antialiased edge pixels differently between the two
      // (and between runs), as the package VRT configs also note.
      launchOptions: { args: ['--disable-gpu'] },
    },
  }],
});
