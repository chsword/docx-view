import { defineConfig } from 'vite';

export default defineConfig({
  root: 'examples',
  base: './',
  build: {
    outDir: '../examples-dist',
    emptyOutDir: true,
  },
});
