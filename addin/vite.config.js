import { defineConfig } from 'vite';
import devCerts from 'office-addin-dev-certs';

export default defineConfig(async () => {
  let httpsOptions = true;

  try {
    httpsOptions = await devCerts.getHttpsServerOptions();
  } catch {
    httpsOptions = true;
  }

  return {
    base: './',
    server: {
      host: '127.0.0.1',
      port: 3000,
      https: httpsOptions
    },
    build: {
      outDir: 'dist-addin',
      emptyOutDir: true,
      rollupOptions: {
        input: {
          taskpane: 'addin/taskpane.html'
        }
      }
    }
  };
});
