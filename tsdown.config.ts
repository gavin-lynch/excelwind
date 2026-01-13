import { defineConfig } from 'tsdown';

export default defineConfig({
  entry: [
    './src/index.ts',
    './src/jsx-runtime/jsx-runtime.ts',
    './src/jsx-runtime/jsx-dev-runtime.ts',
  ],
  outDir: './dist',
  splitting: false,
  format: ['esm'],
});
