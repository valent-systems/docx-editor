import { test } from '@playwright/test';
import { existsSync } from 'node:fs';
import { EditorPage } from '../helpers/editor-page';

const DOCX =
  '/Users/ryanrudd/Source/xyz-ai-docx-demo/src/scripts/docx-spike/samples/tpx-proposal-template.docx';
const OUT =
  '/private/tmp/claude-501/-Users-ryanrudd-Source-xyz-ai/0adb217f-bd48-46b5-b9df-55128ce83b22/scratchpad';

test('screenshot pages 10-11', async ({ page }) => {
  test.skip(!existsSync(DOCX), 'corpus file not present');
  test.setTimeout(240_000);
  const editor = new EditorPage(page);
  await editor.goto();
  await editor.waitForReady();
  await editor.loadDocxFile(DOCX);
  await page.waitForTimeout(8_000);

  for (const idx of [9, 10]) {
    await page.evaluate((i) => {
      document.querySelectorAll('.layout-page')[i]?.scrollIntoView({ block: 'start' });
    }, idx);
    await page.waitForTimeout(500);
    await page
      .locator('.layout-page')
      .nth(idx)
      .screenshot({ path: `${OUT}/fixed-page-${idx + 1}.png` });
  }
});
