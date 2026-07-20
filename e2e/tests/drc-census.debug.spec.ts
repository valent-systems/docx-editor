import { test } from '@playwright/test';
import { existsSync } from 'node:fs';
import { EditorPage } from '../helpers/editor-page';

const DOCX =
  '/Users/ryanrudd/Source/xyz-ai-docx-demo/src/scripts/docx-spike/samples/drc-qualification-template.docx';
const OUT =
  '/private/tmp/claude-501/-Users-ryanrudd-Source-xyz-ai/0adb217f-bd48-46b5-b9df-55128ce83b22/scratchpad';

test('census DRC pages 1-8 + shots', async ({ page }) => {
  test.skip(!existsSync(DOCX), 'corpus file not present');
  test.setTimeout(300_000);
  const editor = new EditorPage(page);
  await editor.goto();
  await editor.waitForReady();
  await editor.loadDocxFile(DOCX);
  await page.waitForTimeout(10_000);

  for (let i = 0; i < 8; i++) {
    await page.evaluate((idx) => {
      document.querySelectorAll('.layout-page')[idx]?.scrollIntoView({ block: 'center' });
    }, i);
    await page.waitForTimeout(300);
    const info = await page.evaluate((idx) => {
      const p = document.querySelectorAll('.layout-page')[idx]!;
      const content = p.querySelector('.layout-page-content') as HTMLElement | null;
      const text = (content?.innerText ?? '').replace(/\s+/g, ' ').trim();
      const imgs = Array.from(p.querySelectorAll('.layout-page-content img')).map(
        (el) => `${(el as HTMLElement).offsetWidth}x${(el as HTMLElement).offsetHeight}`
      );
      return { textLen: text.length, head: text.slice(0, 60), imgs };
    }, i);
    console.log(`p${i + 1}: textLen=${info.textLen} imgs=[${info.imgs.join(',')}] "${info.head}"`);
    await page
      .locator('.layout-page')
      .nth(i)
      .screenshot({ path: `${OUT}/drc-early-${i + 1}.png` });
  }
});
