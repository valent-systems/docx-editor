import { test } from '@playwright/test';
import { existsSync } from 'node:fs';
import { EditorPage } from '../helpers/editor-page';
const DOCX =
  '/Users/ryanrudd/Source/xyz-ai-docx-demo/src/scripts/docx-spike/samples/drc-qualification-v2.docx';
const OUT =
  '/private/tmp/claude-501/-Users-ryanrudd-Source-xyz-ai/0adb217f-bd48-46b5-b9df-55128ce83b22/scratchpad';
test('new DRC header banner render', async ({ page }) => {
  test.skip(!existsSync(DOCX), 'no file');
  test.setTimeout(300_000);
  const editor = new EditorPage(page);
  await editor.goto();
  await editor.waitForReady();
  await editor.loadDocxFile(DOCX);
  await page.waitForTimeout(10_000);
  const total = await page.evaluate(() => document.querySelectorAll('.layout-page').length);
  console.log('pages:', total);
  for (const idx of [1, 2]) {
    await page.evaluate((i) => {
      document.querySelectorAll('.layout-page')[i]?.scrollIntoView({ block: 'start' });
    }, idx);
    await page.waitForTimeout(500);
    await page
      .locator('.layout-page')
      .nth(idx)
      .screenshot({ path: `${OUT}/drc2-page-${idx + 1}.png` });
  }
});
