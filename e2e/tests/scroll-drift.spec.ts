import { test, expect } from '@playwright/test';
import { existsSync } from 'node:fs';
import { EditorPage } from '../helpers/editor-page';

/**
 * Regression: deep-scroll drift. Anchored text boxes walked down the page a
 * little more on every virtualization repaint — the painter resolved their
 * position FROM fragment.y and then wrote the resolved position back INTO
 * fragment.y, so each repaint re-added the anchor offset (fragments persist
 * across paints; paint runs on every scroll-back-into-view). Fixed by
 * resolving from an immutable anchorX/anchorY recorded on first paint.
 *
 * Method: load the 46-page TPX doc, snapshot page-1 geometry, scroll through
 * the whole doc (forcing every virtualized page to render), scroll back,
 * re-snapshot, twice. Geometry must be bit-identical after both round trips.
 */

const DOCX =
  '/Users/ryanrudd/Source/xyz-ai-docx-demo/src/scripts/docx-spike/samples/tpx-proposal-template.docx';

test('scroll drift repro', async ({ page }) => {
  test.skip(!existsSync(DOCX), 'corpus file not present');
  test.setTimeout(300_000);

  const editor = new EditorPage(page);
  await editor.goto();
  await editor.waitForReady();
  await editor.loadDocxFile(DOCX);
  await page.waitForTimeout(6_000);

  const snapshot = () =>
    page.evaluate(() => {
      const p = document.querySelectorAll('.layout-page')[0]!;
      return Array.from(p.querySelectorAll('.layout-page-content > *')).map((f) => {
        const el = f as HTMLElement;
        return {
          cls: el.className.split(' ').slice(0, 2).join('.'),
          top: el.style.top || 'flow',
          pm: el.dataset.pmStart ?? el.getAttribute('data-pm-start') ?? '-',
          text: el.innerText?.replace(/\s+/g, ' ').slice(0, 30) ?? '',
          img: el.querySelector('img') ? 'img' : '',
          tb: el.className.includes('textbox') || el.querySelector('.layout-textbox') ? 'TB' : '',
        };
      });
    });

  const scrollFullRoundTrip = async () => {
    // Scroll down through every page so virtualization renders each.
    const count = await page.evaluate(() => document.querySelectorAll('.layout-page').length);
    for (let i = 0; i < count; i += 3) {
      await page.evaluate((idx) => {
        document.querySelectorAll('.layout-page')[idx]?.scrollIntoView({ block: 'start' });
      }, i);
      await page.waitForTimeout(120);
    }
    // Back to top.
    await page.evaluate(() => {
      const sc = document.querySelector('.docx-editor__scroll-container');
      if (sc) sc.scrollTop = 0;
    });
    await page.waitForTimeout(1_500);
  };

  const before = await snapshot();
  console.log('BEFORE :', JSON.stringify(before));

  await scrollFullRoundTrip();
  const after1 = await snapshot();
  console.log('AFTER 1:', JSON.stringify(after1));

  await scrollFullRoundTrip();
  const after2 = await snapshot();
  console.log('AFTER 2:', JSON.stringify(after2));

  // Geometry must be bit-identical after any number of scroll round trips.
  expect(JSON.stringify(after1)).toBe(JSON.stringify(before));
  expect(JSON.stringify(after2)).toBe(JSON.stringify(before));
});
