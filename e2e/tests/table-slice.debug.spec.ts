import { test } from '@playwright/test';
import { existsSync } from 'node:fs';
import { EditorPage } from '../helpers/editor-page';

const DOCX =
  process.env.SLICE_DOCX ??
  '/Users/ryanrudd/Source/xyz-ai-docx-demo/src/scripts/docx-spike/samples/drc-qualification-v2.docx';
const OUT =
  '/private/tmp/claude-501/-Users-ryanrudd-Source-xyz-ai/0adb217f-bd48-46b5-b9df-55128ce83b22/scratchpad';

/**
 * Hunt for table fragments whose bottomClip lands mid-glyph: for every page
 * that shows a table fragment with a clipped last row, measure the PAINTED
 * line rects near the clip edge and report how far the clip is from the
 * nearest painted line bottom.
 */
test('find mid-glyph table slices', async ({ page }) => {
  test.skip(!existsSync(DOCX), 'corpus file not present');
  test.setTimeout(600_000);
  const editor = new EditorPage(page);
  await editor.goto();
  await editor.waitForReady();
  await editor.loadDocxFile(DOCX);
  await page.waitForTimeout(15_000);

  const pageCount = await page.evaluate(() => document.querySelectorAll('.layout-page').length);
  console.log(`pages: ${pageCount}`);

  const offenders: number[] = [];
  for (let i = 0; i < pageCount; i++) {
    await page.evaluate((idx) => {
      document.querySelectorAll('.layout-page')[idx]?.scrollIntoView({ block: 'center' });
    }, i);
    await page.waitForTimeout(150);
    const info = await page.evaluate((idx) => {
      const p = document.querySelectorAll('.layout-page')[idx]!;
      const tables = Array.from(p.querySelectorAll('.layout-table')) as HTMLElement[];
      const results: string[] = [];
      for (const t of tables) {
        const tRect = t.getBoundingClientRect();
        // A clipped table hides overflow; find text spans whose rect crosses
        // the table's bottom edge (mid-glyph slice) or top edge.
        const spans = Array.from(t.querySelectorAll('span, div')).filter(
          (el) => el.childNodes.length === 1 && el.childNodes[0].nodeType === Node.TEXT_NODE
        ) as HTMLElement[];
        for (const s of spans) {
          const r = s.getBoundingClientRect();
          if (r.height === 0) continue;
          const cutBottom = r.top < tRect.bottom - 1 && r.bottom > tRect.bottom + 1;
          const cutTop = r.top < tRect.top - 1 && r.bottom > tRect.top + 1;
          if (cutBottom || cutTop) {
            results.push(
              `${cutBottom ? 'BOTTOM' : 'TOP'}-cut "${(s.textContent ?? '').slice(0, 40)}" spanY=[${Math.round(r.top - tRect.top)},${Math.round(r.bottom - tRect.top)}] tableH=${Math.round(tRect.height)}`
            );
          }
        }
      }
      return results;
    }, i);
    if (info.length > 0) {
      offenders.push(i + 1);
      console.log(`page ${i + 1}:`);
      info.slice(0, 6).forEach((l) => console.log('   ', l));
      await page
        .locator('.layout-page')
        .nth(i)
        .screenshot({ path: `${OUT}/slice-p${i + 1}.png` });
    }
  }
  console.log(`offending pages: [${offenders.join(', ')}]`);
});
