import { test } from '@playwright/test';
import { existsSync } from 'node:fs';
import { EditorPage } from '../helpers/editor-page';

const DOCX =
  '/Users/ryanrudd/Source/xyz-ai-docx-demo/src/scripts/docx-spike/samples/tpx-proposal-template.docx';

test('probe: icons-table cell vertical anatomy', async ({ page }) => {
  test.skip(!existsSync(DOCX), 'corpus file not present');
  test.setTimeout(120_000);
  const editor = new EditorPage(page);
  await editor.goto();
  await editor.waitForReady();
  await editor.loadDocxFile(DOCX);
  await page.waitForTimeout(8_000);

  // Page 2 (index 1) holds the icons grid.
  await page.evaluate(() => {
    document.querySelectorAll('.layout-page')[1]?.scrollIntoView({ block: 'start' });
  });
  await page.waitForTimeout(600);

  const report = await page.evaluate(() => {
    const page2 = document.querySelectorAll('.layout-page')[1]!;
    const table = page2.querySelector(
      '.layout-page-content table, .layout-page-content .layout-table'
    );
    if (!table) return 'no table found on page 2';
    const rows: string[] = [];
    const tr = table.getBoundingClientRect();
    rows.push(
      `TABLE total: ${Math.round(tr.height)}px tall, top(page-rel)=${Math.round(tr.top - page2.getBoundingClientRect().top)}`
    );
    // First cell anatomy
    const cell = table.querySelector('td, .layout-table-cell');
    if (!cell) return rows.concat('no cell').join('\n');
    const cr = cell.getBoundingClientRect();
    rows.push(`CELL[0,0]: ${Math.round(cr.height)}px  padding=${getComputedStyle(cell).padding}`);
    for (const child of Array.from(
      cell.querySelectorAll(':scope .layout-paragraph, :scope p, :scope div')
    )) {
      const r = (child as HTMLElement).getBoundingClientRect();
      if (r.height === 0) continue;
      const cs = getComputedStyle(child as HTMLElement);
      const text = (child as HTMLElement).innerText.replace(/\s+/g, ' ').slice(0, 24);
      rows.push(
        `  <${(child as HTMLElement).className.split(' ')[0] || child.tagName}> h=${Math.round(r.height)} mt=${cs.marginTop} mb=${cs.marginBottom} lh=${cs.lineHeight} fs=${cs.fontSize} "${text}"`
      );
    }
    return rows.join('\n');
  });
  console.log(report);
});
