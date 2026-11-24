import tailwindcss from '@tailwindcss/vite';
import react from '@vitejs/plugin-react';
import fs from 'fs';
import path from 'path';
import {defineConfig} from 'vite';
import {viteSingleFile} from 'vite-plugin-singlefile';

export default defineConfig(() => {
  return {
    plugins: [
      react(),
      tailwindcss(),
      viteSingleFile(),
      // Copies src/assets into dist/ so root-relative URLs like
      // "/src/assets/images/..." (e.g. generate_image output) still resolve
      // in production builds. Do not remove.
      {
        name: 'copy-unbundled-src-assets',
        apply: 'build',
        closeBundle(error?: Error): void {
          if (error) {
            return;
          }
          const distDir = path.resolve(__dirname, 'dist');
          const srcAssets = path.resolve(__dirname, 'src/assets');
          const distAssets = path.resolve(distDir, 'src/assets');
          if (fs.existsSync(distDir) && fs.existsSync(srcAssets)) {
            fs.cpSync(srcAssets, distAssets, {recursive: true, force: true});
          }
        },
      },
      {
        name: 'serve-standalone-html',
        configureServer(server) {
          server.middlewares.use((req, res, next) => {
            const rawUrl = req.url || '';
            const [urlPath, query = ''] = rawUrl.split('?');
            if (urlPath === '/measure-calc.html' || urlPath === '/window-calc.html') {
              const filePath = path.resolve(__dirname, 'public/window-calc.html');
              if (fs.existsSync(filePath)) {
                res.statusCode = 200;
                res.setHeader('Content-Type', 'text/html; charset=utf-8');
                res.setHeader('Cache-Control', 'no-cache');
                if (query.includes('download=1')) {
                  res.setHeader(
                    'Content-Disposition',
                    'attachment; filename="measure-calc.html"; filename*=UTF-8\'\'%D8%AD%D8%A7%D8%B3%D8%A8%D8%A9_%D8%A7%D9%84%D9%82%D9%8A%D8%A7%D8%B3%D8%A7%D8%AA.html'
                  );
                }
                fs.createReadStream(filePath).pipe(res);
                return;
              }
            }
            if (urlPath === '/build-apk.yml') {
              const wfPath = path.resolve(__dirname, '.github/workflows/build-apk.yml');
              if (fs.existsSync(wfPath)) {
                res.statusCode = 200;
                res.setHeader('Content-Type', 'text/yaml; charset=utf-8');
                res.setHeader('Cache-Control', 'no-cache');
                if (query.includes('download=1')) {
                  res.setHeader('Content-Disposition', 'attachment; filename="build-apk.yml"');
                }
                fs.createReadStream(wfPath).pipe(res);
                return;
              }
            }
            if (urlPath === '/presentation.pptx') {
              const pptxPath = path.resolve(__dirname, 'public/presentation.pptx');
              if (fs.existsSync(pptxPath)) {
                res.statusCode = 200;
                res.setHeader(
                  'Content-Type',
                  'application/vnd.openxmlformats-officedocument.presentationml.presentation'
                );
                res.setHeader('Cache-Control', 'no-cache');
                if (query.includes('download=1')) {
                  res.setHeader(
                    'Content-Disposition',
                    'attachment; filename="presentation.pptx"; filename*=UTF-8\'\'%D8%B4%D8%B1%D8%AD_%D8%AA%D8%B7%D8%A8%D9%8A%D9%82_%D8%AD%D8%A7%D8%B3%D8%A8%D8%A9_%D8%A7%D9%84%D9%82%D9%8A%D8%A7%D8%B3%D8%A7%D8%AA.pptx'
                  );
                }
                fs.createReadStream(pptxPath).pipe(res);
                return;
              }
            }
            if (urlPath === '/architecture-diagram.jpg') {
              const imgPath = path.resolve(__dirname, 'public/architecture-diagram.jpg');
              if (fs.existsSync(imgPath)) {
                res.statusCode = 200;
                res.setHeader('Content-Type', 'image/jpeg');
                res.setHeader('Cache-Control', 'no-cache');
                if (query.includes('download=1')) {
                  res.setHeader(
                    'Content-Disposition',
                    'attachment; filename="architecture-diagram.jpg"; filename*=UTF-8\'\'%D9%85%D8%AE%D8%B7%D8%B7_%D8%AA%D8%B7%D8%A8%D9%8A%D9%82_%D8%AD%D8%A7%D8%B3%D8%A8%D8%A9_%D8%A7%D9%84%D9%82%D9%8A%D8%A7%D8%B3%D8%A7%D8%AA.jpg'
                  );
                }
                fs.createReadStream(imgPath).pipe(res);
                return;
              }
            }
            next();
          });
        },
      },
    ],
    resolve: {
      alias: {
        '@': path.resolve(__dirname, '.'),
      },
    },
    build: {
      assetsInlineLimit: 10 * 1024 * 1024,
    },
    server: {
      // HMR is disabled in AI Studio via DISABLE_HMR env var.
      // Do not modify—file watching is disabled to prevent flickering during agent edits.
      hmr: process.env.DISABLE_HMR !== 'true',
      // Disable file watching when DISABLE_HMR is true to save CPU during agent edits.
      watch: process.env.DISABLE_HMR === 'true' ? null : {},
    },
  };
});
