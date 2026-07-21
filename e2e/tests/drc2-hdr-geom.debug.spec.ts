import { test } from '@playwright/test';
import { existsSync } from 'node:fs';
import { EditorPage } from '../helpers/editor-page';
const DOCX =
  '/Users/ryanrudd/Source/xyz-ai-docx-demo/src/scripts/docx-spike/samples/drc-qualification-v2.docx';
test('header paint geometry', async ({ page }) => {
  test.skip(!existsSync(DOCX), 'no file');
  test.setTimeout(300_000);
  const editor = new EditorPage(page);
  await editor.goto();
  await editor.waitForReady();
  await editor.loadDocxFile(DOCX);
  await page.waitForTimeout(10_000);
  await page.evaluate(() => {
    document.querySelectorAll('.layout-page')[1]?.scrollIntoView({ block: 'start' });
  });
  await page.waitForTimeout(500);
  const info = await page.evaluate(() => {
    const p = document.querySelectorAll('.layout-page')[1]!;
    const pr = p.getBoundingClientRect();
    const hdr = p.querySelector('.layout-page-header') as HTMLElement;
    const hr = hdr.getBoundingClientRect();
    const out: string[] = [
      `header box: left=${Math.round(hr.left - pr.left)} top=${Math.round(hr.top - pr.top)} w=${Math.round(hr.width)} h=${Math.round(hr.height)}`,
    ];
    for (const el of Array.from(hdr.querySelectorAll('*')).slice(0, 30)) {
      const e = el as HTMLElement;
      const r = e.getBoundingClientRect();
      if (r.width < 5 || r.height < 5) continue;
      const cls =
        e.className && typeof e.className === 'string' ? e.className.split(' ')[0] : e.tagName;
      const txt = (e.innerText ?? '').replace(/\s+/g, ' ').slice(0, 18);
      if (
        ['layout-textbox', 'layout-paragraph', 'layout-line', 'IMG'].includes(cls) ||
        e.tagName === 'IMG'
      )
        out.push(
          `  <${cls}> page-rel left=${Math.round(r.left - pr.left)} top=${Math.round(r.top - pr.top)} w=${Math.round(r.width)} h=${Math.round(r.height)} align=${getComputedStyle(e).textAlign} "${txt}"`
        );
    }
    return out.join('\n');
  });
  console.log(info);
});
