import { test } from '@playwright/test';
import { existsSync } from 'node:fs';
import { EditorPage } from '../helpers/editor-page';

const DOCX =
  '/Users/ryanrudd/Source/xyz-ai-docx-demo/src/scripts/docx-spike/samples/tpx-proposal-template.docx';
const OUT =
  '/private/tmp/claude-501/-Users-ryanrudd-Source-xyz-ai/0adb217f-bd48-46b5-b9df-55128ce83b22/scratchpad';

test('capture page seams as the user sees them', async ({ page }) => {
  test.skip(!existsSync(DOCX), 'corpus file not present');
  test.setTimeout(240_000);
  const editor = new EditorPage(page);
  await editor.goto();
  await editor.waitForReady();
  await editor.loadDocxFile(DOCX);
  await page.waitForTimeout(8_000);

  // Center each seam (bottom of page N + top of page N+1) in the viewport
  // and screenshot the viewport — exactly what the user sees.
  for (const n of [1, 2, 3]) {
    await page.evaluate((idx) => {
      const pages = document.querySelectorAll('.layout-page');
      const a = pages[idx] as HTMLElement;
      // Scroll so page N's bottom edge sits mid-viewport.
      a?.scrollIntoView({ block: 'end' });
      const scroller = document.querySelector('.docx-editor__scroll-container');
      if (scroller) scroller.scrollTop += 200;
    }, n - 1);
    await page.waitForTimeout(800);
    await page.screenshot({ path: `${OUT}/seam-${n}-${n + 1}.png` });
  }
});
