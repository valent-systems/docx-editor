import { test } from '@playwright/test';
import { existsSync } from 'node:fs';
import { EditorPage } from '../helpers/editor-page';

const DOCX =
  '/Users/ryanrudd/Source/xyz-ai-docx-demo/src/scripts/docx-spike/samples/drc-qualification-template.docx';
const OUT =
  '/private/tmp/claude-501/-Users-ryanrudd-Source-xyz-ai/0adb217f-bd48-46b5-b9df-55128ce83b22/scratchpad';

test('census DRC pages 12-28 + shots', async ({ page }) => {
  test.skip(!existsSync(DOCX), 'corpus file not present');
  test.setTimeout(300_000);
  const editor = new EditorPage(page);
  await editor.goto();
  await editor.waitForReady();
  await editor.loadDocxFile(DOCX);
  await page.waitForTimeout(10_000);

  for (let i = 11; i < 28; i++) {
    await page.evaluate((idx) => {
      document.querySelectorAll('.layout-page')[idx]?.scrollIntoView({ block: 'center' });
    }, i);
    await page.waitForTimeout(200);
    const info = await page.evaluate((idx) => {
      const p = document.querySelectorAll('.layout-page')[idx]!;
      const content = p.querySelector('.layout-page-content') as HTMLElement | null;
      const text = (content?.innerText ?? '').replace(/\s+/g, ' ').trim();
      const frags = p.querySelectorAll('.layout-page-content > *').length;
      const imgs = Array.from(p.querySelectorAll('.layout-page-content img')).map(
        (el) => `${(el as HTMLElement).offsetWidth}x${(el as HTMLElement).offsetHeight}`
      );
      const tables = p.querySelectorAll(
        '.layout-page-content .layout-table, .layout-page-content table'
      ).length;
      return { textLen: text.length, head: text.slice(0, 45), frags, imgs, tables };
    }, i);
    console.log(
      `p${i + 1}: textLen=${info.textLen} frags=${info.frags} tables=${info.tables} imgs=[${info.imgs.join(',')}] "${info.head}"`
    );
  }
  for (const idx of [12, 16, 20]) {
    await page.evaluate((i) => {
      document.querySelectorAll('.layout-page')[i]?.scrollIntoView({ block: 'start' });
    }, idx);
    await page.waitForTimeout(400);
    await page
      .locator('.layout-page')
      .nth(idx)
      .screenshot({ path: `${OUT}/drc-page-${idx + 1}.png` });
  }
});
