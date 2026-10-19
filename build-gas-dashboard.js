import fs from 'node:fs/promises';
import path from 'node:path';
import vm from 'node:vm';
import { fileURLToPath } from 'node:url';
import { build } from 'vite';

const root = path.dirname(fileURLToPath(import.meta.url));
const temporaryOutput = path.join(root, '.gas-dashboard-build');

await build({
  configFile: false,
  logLevel: 'warn',
  build: {
    outDir: temporaryOutput,
    emptyOutDir: true,
    lib: {
      entry: path.join(root, 'app.js'),
      name: 'AgendaSuratDashboard',
      formats: ['iife'],
      fileName: () => 'dashboard.js',
    },
    cssCodeSplit: false,
    assetsInlineLimit: 100_000_000,
    rollupOptions: { output: { inlineDynamicImports: true } },
  },
});

const htmlPath = path.join(root, 'index.html');
let html = await fs.readFile(htmlPath, 'utf8');
const css = await fs.readFile(path.join(root, 'app.css'), 'utf8');
const javascript = await fs.readFile(path.join(temporaryOutput, 'dashboard.js'), 'utf8');
// HtmlService may serialize the whole HTML document through document.write.
// Keep the large bundle out of HTML raw-text parsing entirely: even a safely
// escaped closing script token can become fragile when another layer rewrites
// the document. Base64 is ASCII and cannot contain HTML markup delimiters.
new vm.Script(javascript, { filename: 'dashboard.js' });
const encodedJavascript = Buffer.from(javascript, 'utf8').toString('base64');
const cityLogo = (await fs.readFile(path.join(root, 'public/assets/logo-kota.jpeg'))).toString('base64');
const schoolLogo = (await fs.readFile(path.join(root, 'public/assets/logo-smpn13.png'))).toString('base64');

if (!html.includes('<link rel="stylesheet" href="app.css">')
  || !html.includes('<script type="module" src="app.js"></script>')) {
  throw new Error('index.html tidak memiliki link CSS dan entry script yang diharapkan.');
}
html = html.replace('<link rel="stylesheet" href="app.css">', () => `<style>\n${css}\n  </style>`);
html = html.replace('<script type="module" src="app.js"></script>', () => `
  <script>
    window.GAS_ASSETS = {
      cityLogo: 'data:image/jpeg;base64,${cityLogo}',
      schoolLogo: 'data:image/png;base64,${schoolLogo}'
    };
  </script>
`);

if (!html.includes('</body>')) {
  throw new Error('index.html tidak memiliki tag penutup </body>.');
}
html = html.replace('</body>', () => `  <script>
    const dashboardBytes = Uint8Array.from(atob('${encodedJavascript}'), character => character.charCodeAt(0));
    const dashboardSource = new TextDecoder().decode(dashboardBytes);
    const dashboardScript = document.createElement('script');
    dashboardScript.textContent = dashboardSource;
    document.body.appendChild(dashboardScript);
  </script>
</body>`);

await fs.writeFile(path.join(root, 'Dashboard.html'), html);
await fs.rm(temporaryOutput, { recursive: true, force: true });
console.log('Dashboard.html dibuat sebagai satu berkas HTML untuk Apps Script HtmlService.');
