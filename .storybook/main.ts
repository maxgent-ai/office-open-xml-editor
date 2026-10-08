import type { StorybookConfig } from '@storybook/html-vite';

const config: StorybookConfig = {
  stories: [
    './**/*.mdx',
    '../packages/*/src/**/*.mdx',
    '../packages/*/src/**/*.stories.@(js|jsx|mjs|ts|tsx)',
  ],
  addons: [
    '@storybook/addon-a11y',
    '@storybook/addon-docs',
    '@chromatic-com/storybook',
  ],
  framework: '@storybook/html-vite',
  staticDirs: [
    { from: '../packages/docx/public', to: '/docx' },
    { from: '../packages/xlsx/public', to: '/xlsx' },
    { from: '../packages/pptx/public', to: '/pptx' },
  ],
  async viteFinal(config) {
    const { default: wasm } = await import('vite-plugin-wasm');
    return {
      ...config,
      // Default '/' for local dev and standalone deploys; the unified Pages
      // workflow sets STORYBOOK_BASE=/storybook/ to nest it under the site.
      base: process.env.STORYBOOK_BASE ?? '/',
      plugins: [...(config.plugins ?? []), wasm()],
      worker: {
        format: 'es' as const,
        plugins: () => [wasm()],
        // Inline parser workers cannot resolve relative chunks from Blob URLs.
        rolldownOptions: { output: { codeSplitting: false } },
      },
    };
  },
};
export default config;
