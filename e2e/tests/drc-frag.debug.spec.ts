import { test } from '@playwright/test';
import { existsSync } from 'node:fs';
import { EditorPage } from '../helpers/editor-page';
const DOCX =
  '/Users/ryanrudd/Source/xyz-ai-docx-demo/src/scripts/docx-spike/samples/drc-qualification-template.docx';
test('fragment geometry pages 4-6', async ({ page }) => {
  test.skip(!existsSync(DOCX), 'no corpus');
  test.setTimeout(300_000);
  const editor = new EditorPage(page);
  await editor.goto();
  await editor.waitForReady();
  await editor.loadDocxFile(DOCX);
  await page.waitForTimeout(10_000);
  for (const idx of [3, 4, 5]) {
    await page.evaluate((i) => {
      document.querySelectorAll('.layout-page')[i]?.scrollIntoView({ block: 'center' });
    }, idx);
    await page.waitForTimeout(300);
    const info = await page.evaluate((i) => {
      const p = document.querySelectorAll('.layout-page')[i]!;
      const content = p.querySelector('.layout-page-content') as HTMLElement;
      const cs = getComputedStyle(content);
      const frags = Array.from(content.children)
        .slice(0, 6)
        .map((f) => {
          const el = f as HTMLElement;
          return `${el.className.split(' ')[0]}@${el.style.top} h=${Math.round(el.getBoundingClientRect().height)} "${(el.innerText ?? '').replace(/\s+/g, ' ').slice(0, 25)}"`;
        });
      return { top: cs.top, bottom: cs.bottom, frags };
    }, idx);
    console.log(`page ${idx + 1}: contentTop=${info.top} bottom=${info.bottom}`);
    info.frags.forEach((f) => console.log('   ', f));
  }
});
