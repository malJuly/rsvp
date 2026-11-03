import react from '@vitejs/plugin-react';
import { fileURLToPath } from 'node:url';
import { defineConfig, loadEnv } from 'vite';

export default defineConfig(({ mode }) => {
  // The live address, for the link-preview image (WhatsApp needs a full https:// address).
  // Empty means a relative path, so nothing breaks before it is set.
  const siteUrl = (loadEnv(mode, process.cwd(), 'VITE_').VITE_SITE_URL || '').trim().replace(/\/+$/, '');

  return {
    plugins: [
      react(),
      { name: 'site-url', transformIndexHtml: (html) => html.replaceAll('__SITE_URL__/', siteUrl ? `${siteUrl}/` : '') },
    ],
    // Relative paths, so the built site works from any folder on any static host.
    base: './',
    // Two pages: the guest site and the admin page (built separately, so guests never download admin code).
    build: {
      rollupOptions: {
        input: {
          main: fileURLToPath(new URL('./index.html', import.meta.url)),
          admin: fileURLToPath(new URL('./admin.html', import.meta.url)),
        },
      },
    },
    server: { port: 3000, strictPort: true },
    preview: { port: 3000, strictPort: true },
  };
});
