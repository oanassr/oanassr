// Rasterise SVG files to PNG with the preinstalled Chromium (used by illustrations.py).
const { chromium } = require('playwright');
const fs = require('fs');
(async () => {
  const jobs = JSON.parse(process.argv[2]);
  const browser = await chromium.launch();
  const page = await browser.newPage();
  for (const j of jobs) {
    await page.setViewportSize({ width: j.w, height: j.h });
    await page.setContent(`<html><body style="margin:0">${fs.readFileSync(j.svg, 'utf8')}</body></html>`);
    await page.screenshot({ path: j.png, clip: { x: 0, y: 0, width: j.w, height: j.h } });
  }
  await browser.close();
})();
