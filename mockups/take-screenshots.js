'use strict';

const puppeteer = require('puppeteer-core');
const http      = require('http');
const fs        = require('fs');
const path      = require('path');

const CHROME    = 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe';
const PORT      = 8998;
const BASE_URL  = `http://localhost:${PORT}`;
const ROOT_DIR  = path.resolve(__dirname, '..');
const SHOTS_DIR = path.resolve(ROOT_DIR, 'screenshots');

// ── HTTP server serving the project root ─────────────────────────────────────

function startServer() {
  return new Promise(resolve => {
    const MIME = {
      '.html': 'text/html', '.js': 'application/javascript',
      '.css': 'text/css',   '.png': 'image/png', '.csv': 'text/plain',
    };
    const server = http.createServer((req, res) => {
      let filePath = path.join(ROOT_DIR, req.url.split('?')[0]);
      if (fs.existsSync(filePath) && fs.statSync(filePath).isDirectory()) {
        filePath = path.join(filePath, 'index.html');
      }
      if (!fs.existsSync(filePath)) { res.writeHead(404); res.end('Not found'); return; }
      res.writeHead(200, { 'Content-Type': MIME[path.extname(filePath)] || 'text/plain' });
      fs.createReadStream(filePath).pipe(res);
    });
    server.listen(PORT, () => { console.log(`  Server → ${BASE_URL}`); resolve(server); });
  });
}

function wait(ms) { return new Promise(r => setTimeout(r, ms)); }

// ── Filename for each .screenshot-frame element in DOM order ─────────────────
// Keep this in sync with the frame order in mockups/screenshot-guide.html.

const FRAME_FILENAMES = [
  'datasource-empty.png',           // Screen 1A
  'datasource-file-loaded.png',     // Screen 1B
  'hero-chart.png',                 // Screen 1C
  'preview-mode.png',               // Screen 1D
  'chart-bar.png',                  // Chart 1 — Bar (Vertical)
  'chart-hbar.png',                 // Chart 2 — Bar (Horizontal)
  'chart-line.png',                 // Chart 3 — Line
  'chart-area.png',                 // Chart 4 — Area
  'chart-scatter.png',              // Chart 5 — Scatter
  'chart-bubble.png',               // Chart 6 — Bubble
  'chart-pie.png',                  // Chart 7 — Pie
  'chart-doughnut.png',             // Chart 8 — Doughnut
  'chart-radar.png',                // Chart 9 — Radar
  'feature-stacked.png',            // Feature — Stacked Bar
  'feature-data-labels.png',        // Feature — Data Labels
  'feature-palettes.png',           // Feature — Color Palettes
  'feature-data-table.png',         // Feature — Data Table
  'datasource-sharepoint-list.png', // Feature — SharePoint List
];

// ── Main ─────────────────────────────────────────────────────────────────────

(async () => {
  fs.mkdirSync(SHOTS_DIR, { recursive: true });

  const server  = await startServer();
  const browser = await puppeteer.launch({
    executablePath: CHROME,
    headless:       true,
    args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage', '--disable-gpu', '--no-first-run'],
    defaultViewport: { width: 960, height: 900 },
  });

  const page = await browser.newPage();
  page.on('console', m => { if (m.type() === 'error') console.error('  Page error:', m.text()); });

  try {
    console.log('\n  Loading mockup page…');
    await page.goto(`${BASE_URL}/mockups/screenshot-guide.html`, { waitUntil: 'networkidle0' });

    // Wait for Chart.js CDN scripts to load and all charts to render
    await page.waitForFunction(() => typeof window.Chart !== 'undefined', { timeout: 20000 });
    // chart animation is disabled so 1.5s is enough for layout to settle
    await wait(1500);

    // Hide page chrome (sticky nav, TOC) — they're mockup scaffolding, not web part UI.
    // Without this, the sticky header lands in the middle of tall element screenshots.
    await page.addStyleTag({
      content: '.page-header, .toc { display: none !important; }',
    });
    await wait(100);

    // Use ElementHandle.screenshot() — Puppeteer scrolls the element into view
    // automatically and clips exactly to its bounding box, avoiding manual
    // scroll-then-measure coordinate arithmetic entirely.
    const frames = await page.$$('.screenshot-frame');
    console.log(`\n  Found ${frames.length} .screenshot-frame elements`);
    console.log('  Taking screenshots…\n');

    for (let i = 0; i < FRAME_FILENAMES.length; i++) {
      const filename = FRAME_FILENAMES[i];
      if (i >= frames.length) {
        console.warn(`  ⚠  Frame ${i} missing in HTML — skipping ${filename}`);
        continue;
      }
      await frames[i].screenshot({ path: path.join(SHOTS_DIR, filename), type: 'png' });
      console.log(`  ✓  ${filename}`);
    }

    // chart-gallery.png — the 2-column .chart-grid containing Charts 1–8
    const galleryHandle = await page.$('.chart-grid');
    if (galleryHandle) {
      await galleryHandle.screenshot({ path: path.join(SHOTS_DIR, 'chart-gallery.png'), type: 'png' });
      console.log('  ✓  chart-gallery.png');
    } else {
      console.warn('  ⚠  .chart-grid not found — chart-gallery.png skipped');
    }

    console.log(`\n  All screenshots saved to ${SHOTS_DIR}\n`);

  } catch (err) {
    console.error('\n  Error:', err.message);
    await page.screenshot({ path: path.join(SHOTS_DIR, '_debug.png') }).catch(() => {});
    process.exitCode = 1;
  } finally {
    await browser.close();
    server.close();
  }
})();
