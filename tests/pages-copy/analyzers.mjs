// Page-copy analyzers — one check module per web part under ./analyzers/,
// each run against its live-captured fixture (fixtures/live-shapes.json).

const MODULES = [
  './analyzers/header.mjs',
  './analyzers/text.mjs',
  './analyzers/quick-links.mjs',
  './analyzers/news.mjs',
  './analyzers/list-library.mjs',
];

export async function run({ browser, check, WB_URL }) {
  const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
  await page.goto(WB_URL);
  await page.waitForSelector('.wb-home-cards', { timeout: 60000 });
  for (const path of MODULES) {
    const mod = await import(path);
    await mod.checks({ page, check });
  }
  await page.close();
}
