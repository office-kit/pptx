import { resolve } from 'node:path';
import { defineConfig } from 'vitest/config';

export default defineConfig({
  resolve: {
    alias: {
      // Allow packages/preview/src/* to import from '@office-kit/pptx' without a
      // built dist/ directory. The alias wires the package name directly to the
      // TypeScript source so vitest resolves the same code the tests exercise.
      '@office-kit/pptx': resolve('./src/api/index.ts'),
    },
  },
  test: {
    include: ['src/**/*.test.ts', 'test/**/*.test.ts'],
    environment: 'node',
    coverage: {
      provider: 'v8',
      reporter: ['text', 'html'],
      include: ['src/**/*.ts'],
      exclude: ['src/**/*.test.ts', 'src/**/index.ts'],
    },
  },
});
