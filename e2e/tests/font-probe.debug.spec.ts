import { test } from '@playwright/test';
import { existsSync } from 'node:fs';
import { EditorPage } from '../helpers/editor-page';

const DOCX =
  '/Users/ryanrudd/Source/xyz-ai-docx-demo/src/scripts/docx-spike/samples/tpx-proposal-template.docx';

test('probe: which fonts actually load', async ({ page }) => {
  test.skip(!existsSync(DOCX), 'corpus file not present');
  test.setTimeout(120_000);
  const requests: string[] = [];
  page.on('request', (r) => {
    if (r.url().includes('fonts.googleapis') || r.url().includes('fonts.gstatic'))
      requests.push(`${r.method()} ${r.url().slice(0, 110)}`);
  });
  const responses: string[] = [];
  page.on('response', (r) => {
    if (r.url().includes('fonts.googleapis'))
      responses.push(`${r.status()} ${r.url().slice(0, 110)}`);
  });

  const editor = new EditorPage(page);
  await editor.goto();
  await editor.waitForReady();
  await editor.loadDocxFile(DOCX);
  await page.waitForTimeout(8_000);

  const probe = await page.evaluate(() => ({
    onest400: document.fonts.check('400 16px "Onest"'),
    onest600: document.fonts.check('600 16px "Onest"'),
    roboto300: document.fonts.check('300 16px "Roboto"'),
    carlito: document.fonts.check('400 16px "Carlito"'),
    pages: document.querySelectorAll('.layout-page').length,
  }));
  console.log('css responses:', JSON.stringify(responses, null, 1));
  console.log('font checks:', JSON.stringify(probe));
});
