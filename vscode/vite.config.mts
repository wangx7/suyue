// VS Code 插件的 webview 部分：一个 JS 入口加一个 CSS，扩展生成的页面直接引用这两个文件
import { defineConfig } from 'vite';
import { fileURLToPath } from 'node:url';

const r = (p: string) => fileURLToPath(new URL(p, import.meta.url));

export default defineConfig({
  base: './',
  publicDir: false,
  build: {
    outDir: r('./dist/webview'),
    emptyOutDir: true,
    target: 'es2022',
    chunkSizeWarningLimit: 2000,
    cssCodeSplit: false,
    rollupOptions: {
      input: r('./webview/main.ts'),
      output: {
        entryFileNames: 'webview.js',
        chunkFileNames: 'assets/[name]-[hash].js',
        assetFileNames: (asset) => (asset.names.some((n) => n.endsWith('.css')) ? 'webview.css' : 'assets/[name]-[hash][extname]'),
      },
    },
  },
  worker: { format: 'es' },
});
