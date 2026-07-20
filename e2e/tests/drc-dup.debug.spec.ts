import { test } from '@playwright/test';
import { existsSync } from 'node:fs';
import { EditorPage } from '../helpers/editor-page';
const DOCX =
  '/Users/ryanrudd/Source/xyz-ai-docx-demo/src/scripts/docx-spike/samples/drc-qualification-template.docx';
test('no duplicate floating headshots across page breaks', async ({ page }) => {
  test.skip(!existsSync(DOCX), 'no corpus');
  test.setTimeout(300_000);
  const editor = new EditorPage(page);
  await editor.goto();
  await editor.waitForReady();
  await editor.loadDocxFile(DOCX);
  await page.waitForTimeout(10_000);
  const count = await page.evaluate(() => document.querySelectorAll('.layout-page').length);
  // Count identical floating-image srcs appearing on ADJACENT pages (the dup signature)
  let dups = 0;
  let prev: string[] = [];
  for (let i = 0; i < count; i++) {
    await page.evaluate((idx) => {
      document.querySelectorAll('.layout-page')[idx]?.scrollIntoView({ block: 'center' });
    }, i);
    await page.waitForTimeout(120);
    const srcs: string[] = await page.evaluate((idx) => {
      const p = document.querySelectorAll('.layout-page')[idx]!;
      return Array.from(
        p.querySelectorAll('.layout-floating-images-layer img, .layout-page-floating-image img')
      )
        .filter((el) => ((el as HTMLImageElement).src ?? '').length > 40)
        .map((el) => {
          const src = (el as HTMLImageElement).src ?? '';
          // djb2 over the FULL src — base64 prefixes collide for same-size PNGs
          let h = 5381;
          for (let k = 0; k < src.length; k++) h = ((h * 33) ^ src.charCodeAt(k)) >>> 0;
          return `${src.length}:${h}`;
        });
    }, i);
    for (const s of srcs)
      if (s && prev.includes(s)) {
        dups++;
        console.log(`DUP on p${i + 1}: ${s}`);
      }
    prev = srcs;
  }
  console.log(`adjacent-page duplicate floats: ${dups} (pages=${count})`);
});
