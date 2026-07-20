import { test } from '@playwright/test';
import { existsSync } from 'node:fs';
import { EditorPage } from '../helpers/editor-page';

const DOCX =
  '/Users/ryanrudd/Source/xyz-ai-docx-demo/src/scripts/docx-spike/samples/tpx-proposal-template.docx';

test('census: blank pages + image render state', async ({ page }) => {
  test.skip(!existsSync(DOCX), 'corpus file not present');
  test.setTimeout(300_000);
  const editor = new EditorPage(page);
  await editor.goto();
  await editor.waitForReady();
  await editor.loadDocxFile(DOCX);
  await page.waitForTimeout(6_000);

  const count = await page.evaluate(() => document.querySelectorAll('.layout-page').length);
  for (let i = 0; i < count; i++) {
    await page.evaluate((idx) => {
      document.querySelectorAll('.layout-page')[idx]?.scrollIntoView({ block: 'center' });
    }, i);
    await page.waitForTimeout(150);
    const info = await page.evaluate((idx) => {
      const p = document.querySelectorAll('.layout-page')[idx]!;
      const content = p.querySelector('.layout-page-content') as HTMLElement | null;
      const text = (content?.innerText ?? '').replace(/\s+/g, ' ').trim();
      const imgs = Array.from(p.querySelectorAll('.layout-page-content img')).map((el) => {
        const im = el as HTMLImageElement;
        return {
          loaded: im.complete && im.naturalWidth > 0,
          w: im.offsetWidth,
          h: im.offsetHeight,
          src: im.src.startsWith('data:')
            ? `data:${im.src.slice(5, 20)}…(${im.src.length})`
            : im.src.slice(0, 40),
        };
      });
      return { text: text.slice(0, 30), imgs };
    }, i);
    if (!info.text || info.imgs.some((m) => !m.loaded)) {
      console.log(
        `ours p${i + 1}: text=${JSON.stringify(info.text)} imgs=${JSON.stringify(info.imgs)}`
      );
    }
  }
});
