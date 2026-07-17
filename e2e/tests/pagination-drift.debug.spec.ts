import { test } from '@playwright/test';
import { execSync } from 'node:child_process';
import { existsSync } from 'node:fs';
import { EditorPage } from '../helpers/editor-page';

/**
 * DEBUG (not for CI): map Word's rendered page breaks
 * (`<w:lastRenderedPageBreak/>`) onto OUR pages to locate pagination drift.
 * For each Word page-start snippet, find the page we render it on and print
 * word-page → our-page. The first place the delta grows is the divergence.
 */

const DOCX =
  '/Users/ryanrudd/Source/xyz-ai-docx-demo/src/scripts/docx-spike/samples/tpx-proposal-template.docx';

function norm(s: string): string {
  return s.replace(/[’‘]/g, "'").replace(/[”“]/g, '"').replace(/\s+/g, ' ').trim().toLowerCase();
}

function wordBreakSnippets(): string[] {
  const xml = execSync(`unzip -p ${JSON.stringify(DOCX)} word/document.xml`, {
    encoding: 'utf8',
    maxBuffer: 64 * 1024 * 1024,
  });
  const snippets: string[] = [];
  const re = /<w:lastRenderedPageBreak\/>/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(xml))) {
    const after = xml.slice(m.index, m.index + 6000);
    const texts = [...after.matchAll(/<w:t[^>]*>([^<]+)/g)].map((t) => t[1]).join(' ');
    snippets.push(norm(texts).slice(0, 60));
  }
  return snippets;
}

test('map Word page breaks onto our pages', async ({ page }) => {
  test.skip(!existsSync(DOCX), `corpus file not present: ${DOCX}`);
  test.setTimeout(240_000);
  const snippets = wordBreakSnippets();

  const editor = new EditorPage(page);
  await editor.goto();
  await editor.waitForReady();
  await editor.loadDocxFile(DOCX);
  await page.waitForTimeout(6_000);

  // Virtualization renders only pages near the viewport — scroll through the
  // doc so every shell paints, capturing each page's text as it renders.
  const pageCount = await page.evaluate(() => document.querySelectorAll('.layout-page').length);
  const pages: string[] = [];
  for (let i = 0; i < pageCount; i++) {
    await page.evaluate((idx) => {
      document.querySelectorAll('.layout-page')[idx]?.scrollIntoView({ block: 'center' });
    }, i);
    await page.waitForTimeout(250);
    pages.push(
      await page.evaluate((idx) => {
        const p = document.querySelectorAll('.layout-page')[idx];
        return (p?.querySelector('.layout-page-content') as HTMLElement)?.innerText ?? '';
      }, i)
    );
  }
  const normPages = pages.map((t) =>
    t.replace(/[’‘]/g, "'").replace(/[”“]/g, '"').replace(/\s+/g, ' ').trim().toLowerCase()
  );

  console.log(`our pages: ${pages.length}; word breaks: ${snippets.length}`);
  let prevOurs = 1;
  for (let i = 0; i < snippets.length; i++) {
    const wordPage = i + 2; // marker i starts Word page i+2
    // Try progressively shorter prefixes — table cell order can differ
    // slightly between the XML run order and DOM innerText.
    let ours = -1;
    let probe = '';
    for (const len of [40, 28, 18, 12]) {
      probe = snippets[i].slice(0, len);
      if (!probe) continue;
      for (let p = 0; p < normPages.length; p++) {
        if (normPages[p].includes(probe)) {
          ours = p + 1;
          break;
        }
      }
      if (ours > 0) break;
    }
    const drift = ours > 0 ? ours - wordPage : NaN;
    const jump = ours > 0 && ours < prevOurs ? ' (!)' : '';
    if (ours > 0) prevOurs = ours;
    console.log(
      `word p${String(wordPage).padStart(2)} → ours p${String(ours).padStart(2)}  drift=${String(drift).padStart(3)}${jump}  "${probe.slice(0, 38)}"`
    );
  }
});
