// Render the book cover (cover.html) to img/book_cover.jpg via Chromium.
const { chromium } = require('playwright');
const path = require('path');
(async () => {
  const b = await chromium.launch();
  const p = await b.newPage({ viewport: { width: 2480, height: 3508 } });
  await p.goto('file://' + path.join(__dirname, 'cover.html'));
  await p.waitForFunction(() => document.body.dataset.ready === '1', null, { timeout: 15000 });
  await p.waitForTimeout(300);
  await p.screenshot({ path: process.argv[2], type: 'jpeg', quality: 90 });
  await b.close();
})();
