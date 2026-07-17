import { test, expect } from '@playwright/test';
import { execSync } from 'node:child_process';
import { existsSync } from 'node:fs';
import { EditorPage } from '../helpers/editor-page';

/**
 * Pagination parity against Word.
 *
 * A Word-authored .docx embeds Word's own page count in `docProps/app.xml`
 * (`<Pages>`), giving an automatic ground-truth oracle: parse + paginate the
 * file and compare our page count against the count the authoring app saved.
 *
 * The corpus below points at real customer-shaped documents kept OUTSIDE the
 * repo (they are large and private). Files that are missing are skipped, so CI
 * without the corpus still passes; locally this is the measurement harness for
 * pagination work.
 */

const CORPUS: Array<{ name: string; path: string; tolerance: number }> = [
  {
    name: 'tpx-proposal-template',
    path: '/Users/ryanrudd/Source/xyz-ai-docx-demo/src/scripts/docx-spike/samples/tpx-proposal-template.docx',
    tolerance: 1,
  },
];

function wordPageCount(docxPath: string): number | null {
  try {
    const xml = execSync(`unzip -p ${JSON.stringify(docxPath)} docProps/app.xml`, {
      encoding: 'utf8',
    });
    const m = xml.match(/<Pages>(\d+)<\/Pages>/);
    return m ? parseInt(m[1], 10) : null;
  } catch {
    return null;
  }
}

test.describe('pagination parity (docProps oracle)', () => {
  for (const doc of CORPUS) {
    test(`${doc.name}: page count within ±${doc.tolerance} of Word`, async ({ page }) => {
      test.skip(!existsSync(doc.path), `corpus file not present: ${doc.path}`);
      const expected = wordPageCount(doc.path);
      test.skip(expected == null, 'no <Pages> in docProps/app.xml');

      const editor = new EditorPage(page);
      await editor.goto();
      await editor.waitForReady();
      test.setTimeout(180_000); // 11MB image-heavy docs parse slowly
      await editor.loadDocxFile(doc.path);
      // Let pagination settle (image measure passes re-run layout).
      await page.waitForTimeout(5_000);

      const rendered = await page.evaluate(() => document.querySelectorAll('.layout-page').length);

      console.log(`[pagination-parity] ${doc.name}: ours=${rendered} word=${expected}`);
      expect(Math.abs(rendered - (expected as number))).toBeLessThanOrEqual(doc.tolerance);
    });
  }
});
