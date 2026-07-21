import { test, expect } from '@playwright/test';
import { existsSync } from 'node:fs';
import { EditorPage } from '../helpers/editor-page';

/**
 * Typing-latency budget (docs/INCREMENTAL-LAYOUT.md).
 *
 * Perceived keystroke latency is a hard product requirement: <50ms is the
 * acceptable standard, 16ms (RAIL frame budget) the goal. This spec types
 * into the heaviest corpus file and asserts budgets on the MEDIAN (a single
 * full-pipeline pass — an ineligible edit or cold cache — must not fail the
 * run, but a regression of the typical case must).
 *
 * Budgets include Playwright's per-keystroke harness overhead (~10-20ms),
 * so the editor's real budget is tighter than the asserted numbers.
 *
 * Skips when the corpus file is absent (it is customer data, not committed).
 */

const DOCX =
  process.env.PERF_DOCX ??
  '/Users/ryanrudd/Source/xyz-ai-docx-demo/src/scripts/docx-spike/samples/drc-qualification-v2.docx';

/** Median budget per keystroke, harness overhead included. */
const MEDIAN_BUDGET_MS = 50;
/** p90 budget — allows occasional full-pipeline passes without flakiness. */
const P90_BUDGET_MS = 160;
const KEYSTROKES = 12;

function median(xs: number[]): number {
  const s = [...xs].sort((a, b) => a - b);
  return s[Math.floor(s.length / 2)];
}
function p90(xs: number[]): number {
  const s = [...xs].sort((a, b) => a - b);
  return s[Math.min(s.length - 1, Math.ceil(s.length * 0.9) - 1)];
}

async function typeAndMeasure(
  page: import('@playwright/test').Page,
  text: string
): Promise<number[]> {
  const times: number[] = [];
  for (const ch of text.slice(0, KEYSTROKES)) {
    const t0 = Date.now();
    await page.keyboard.type(ch);
    await page.evaluate(() => 1 + 1); // wait until the main thread is responsive again
    times.push(Date.now() - t0);
  }
  return times;
}

test('keystroke latency stays within budget (body, wrap, table)', async ({ page }) => {
  test.skip(!existsSync(DOCX), 'perf corpus file not present');
  test.setTimeout(300_000);
  const editor = new EditorPage(page);
  await editor.goto();
  await editor.waitForReady();
  await editor.loadDocxFile(DOCX);
  await page.waitForTimeout(15_000);
  const mod = process.platform === 'darwin' ? 'Meta' : 'Control';
  const undoAll = async (n: number) => {
    for (let i = 0; i < n; i++) await page.keyboard.press(`${mod}+z`);
    await page.waitForTimeout(400);
  };

  // --- body paragraph, steady text (fast path M1) ---
  await page.evaluate(() => {
    document.querySelectorAll('.layout-page')[1]?.scrollIntoView({ block: 'center' });
  });
  await page.waitForTimeout(1500);
  await page
    .locator('.layout-page')
    .nth(1)
    .locator('.layout-paragraph .layout-run')
    .first()
    .click({ force: true, timeout: 30000 });
  await page.waitForTimeout(400);
  const body = await typeAndMeasure(page, 'xxxxxxxxxxxx');
  console.log(`body: median=${median(body)}ms p90=${p90(body)}ms [${body.join(',')}]`);

  // --- same paragraph, wrap-inducing text (fast path M2b) ---
  const wrap = await typeAndMeasure(page, 'wwwwwwwwwwww wwwwwwwwwwww wwwwwwwwwwww ');
  console.log(`wrap: median=${median(wrap)}ms p90=${p90(wrap)}ms [${wrap.join(',')}]`);
  await undoAll(30);

  // --- table cell (fast path M2) ---
  await page.evaluate(() => {
    document.querySelectorAll('.layout-page')[26]?.scrollIntoView({ block: 'center' });
  });
  await page.waitForTimeout(1500);
  await page
    .locator('.layout-page')
    .nth(26)
    .locator('.layout-table .layout-run')
    .first()
    .click({ force: true, timeout: 30000 });
  await page.waitForTimeout(400);
  const table = await typeAndMeasure(page, 'xxxxxxxxxxxx');
  console.log(`table: median=${median(table)}ms p90=${p90(table)}ms [${table.join(',')}]`);
  await undoAll(14);

  expect(median(body)).toBeLessThanOrEqual(MEDIAN_BUDGET_MS);
  expect(median(wrap)).toBeLessThanOrEqual(MEDIAN_BUDGET_MS);
  expect(median(table)).toBeLessThanOrEqual(MEDIAN_BUDGET_MS);
  expect(p90(body)).toBeLessThanOrEqual(P90_BUDGET_MS);
  expect(p90(wrap)).toBeLessThanOrEqual(P90_BUDGET_MS);
  expect(p90(table)).toBeLessThanOrEqual(P90_BUDGET_MS);
});
