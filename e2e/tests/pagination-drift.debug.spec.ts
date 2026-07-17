import { test } from '@playwright/test';
import { execSync } from 'node:child_process';
import { existsSync } from 'node:fs';
import { EditorPage } from '../helpers/editor-page';

/**
 * DEBUG (not for CI): interval-accurate pagination drift map.
 *
 * Reconstructs Word's TRUE page-start list from document.xml — every
 * `<w:lastRenderedPageBreak/>` is a rendered page start; an explicit
 * `w:br type="page"` or nextPage sectPr immediately followed by a marker
 * (no text between) is the SAME start; a lone one is its own. Pages with no
 * text between consecutive starts are textless (full-page images / blanks).
 *
 * Then maps each TEXTFUL page start onto our rendered pages SEQUENTIALLY
 * (search forward from the previous match — repeated headings can't alias),
 * and reports word-span vs our-span per interval. Intervals where the spans
 * differ are the exact local divergences.
 */

const DOCX =
  '/Users/ryanrudd/Source/xyz-ai-docx-demo/src/scripts/docx-spike/samples/tpx-proposal-template.docx';

const norm = (s: string) =>
  s.replace(/[’‘]/g, "'").replace(/[”“]/g, '"').replace(/\s+/g, ' ').trim().toLowerCase();

interface Anchor {
  pos: number;
  kinds: string[];
  /** normalized text of the page this anchor starts (up to the next anchor) */
  pageText: string;
}

function buildAnchors(): Anchor[] {
  const xml = execSync(`unzip -p ${JSON.stringify(DOCX)} word/document.xml`, {
    encoding: 'utf8',
    maxBuffer: 64 * 1024 * 1024,
  });
  const raw: Array<{ pos: number; kind: string }> = [];
  for (const m of xml.matchAll(/<w:lastRenderedPageBreak\s*\/>/g))
    raw.push({ pos: m.index!, kind: 'marker' });
  for (const m of xml.matchAll(/<w:br w:type="page"\s*\/>/g))
    raw.push({ pos: m.index!, kind: 'br' });
  const sects = [...xml.matchAll(/<w:sectPr[ >][\s\S]*?<\/w:sectPr>/g)];
  sects.slice(0, -1).forEach((m) => {
    if (!/<w:type w:val="continuous"\s*\/>/.test(m[0]))
      raw.push({ pos: m.index! + m[0].length, kind: 'sect' });
  });
  raw.sort((a, b) => a.pos - b.pos);

  const textBetween = (a: number, b: number) =>
    norm([...xml.slice(a, b).matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((t) => t[1]).join(' '));

  // Merge a br/sect with an immediately-following marker (same page start).
  const merged: Array<{ pos: number; kinds: string[] }> = [];
  for (const { pos, kind } of raw) {
    const prev = merged[merged.length - 1];
    if (
      prev &&
      kind === 'marker' &&
      !prev.kinds.includes('marker') &&
      !textBetween(prev.pos, pos)
    ) {
      prev.kinds.push(kind);
    } else {
      merged.push({ pos, kinds: [kind] });
    }
  }

  return merged.map((a, i) => ({
    pos: a.pos,
    kinds: a.kinds,
    pageText: textBetween(a.pos, merged[i + 1]?.pos ?? a.pos + 30000),
  }));
}

test('interval drift map: Word page spans vs ours', async ({ page }) => {
  test.skip(!existsSync(DOCX), `corpus file not present: ${DOCX}`);
  test.setTimeout(300_000);

  const anchors = buildAnchors();
  console.log(`Word page starts: ${anchors.length} → Word renders ${anchors.length + 1} pages`);

  const editor = new EditorPage(page);
  await editor.goto();
  await editor.waitForReady();
  await editor.loadDocxFile(DOCX);
  await page.waitForTimeout(6_000);

  // Scroll through so virtualization paints every page; capture each page's text.
  const pageCount = await page.evaluate(() => document.querySelectorAll('.layout-page').length);
  const pages: string[] = [];
  for (let i = 0; i < pageCount; i++) {
    await page.evaluate((idx) => {
      document.querySelectorAll('.layout-page')[idx]?.scrollIntoView({ block: 'center' });
    }, i);
    await page.waitForTimeout(200);
    pages.push(
      await page.evaluate((idx) => {
        const p = document.querySelectorAll('.layout-page')[idx];
        return (p?.querySelector('.layout-page-content') as HTMLElement)?.innerText ?? '';
      }, i)
    );
  }
  const normPages = pages.map(norm);
  console.log(`our pages: ${pages.length}`);

  // Per-page census: first chars + image count + tallest image (localizes the
  // full-page diagram pages Word renders textless).
  const census: string[] = await page.evaluate(() =>
    Array.from(document.querySelectorAll('.layout-page')).map((p) => {
      const imgs = Array.from(p.querySelectorAll('.layout-page-content img'));
      const maxH = imgs.reduce((m, el) => Math.max(m, (el as HTMLElement).offsetHeight), 0);
      return `imgs=${imgs.length}${maxH ? ` maxH=${maxH}` : ''}`;
    })
  );
  console.log('\n=== OUR pages ===');
  pages.forEach((t, i) => {
    const head = norm(t).slice(0, 60) || '(no text)';
    console.log(`ours p${String(i + 1).padStart(2)} [${census[i]}] "${head}"`);
  });
  console.log('\n=== WORD pages ===');
  anchors.forEach((a, i) => {
    console.log(
      `word p${String(i + 2).padStart(2)} [${a.kinds.join('+')}] "${a.pageText.slice(0, 60) || '(textless)'}"`
    );
  });

  // Sequentially match textful anchors onto our pages. Space-insensitive:
  // Word stores headings like "M anaged Firewalls" as split runs whose DOM
  // innerText joins differently, so compare with all spaces stripped.
  const squash = (s: string) => s.replace(/ /g, '');
  const squashedPages = normPages.map(squash);
  const matches: Array<{ anchorIdx: number; ourPage: number }> = [];
  let cursor = 0;
  anchors.forEach((a, i) => {
    if (a.pageText.length < 8) return;
    const squashed = squash(a.pageText);
    for (const len of [34, 24, 16, 10]) {
      const probe = squashed.slice(0, len);
      for (let p = cursor; p < squashedPages.length; p++) {
        if (squashedPages[p].includes(probe)) {
          matches.push({ anchorIdx: i, ourPage: p });
          cursor = p;
          return;
        }
      }
    }
    console.log(`  (unmatched) word p${i + 2} "${a.pageText.slice(0, 40)}"`);
  });

  // Interval report: word span vs our span between consecutive matches.
  console.log('\n=== intervals (≠ marks divergence) ===');
  let sumWord = 0;
  let sumOurs = 0;
  for (let k = 1; k < matches.length; k++) {
    const a = matches[k - 1];
    const b = matches[k];
    const wordSpan = b.anchorIdx - a.anchorIdx;
    const ourSpan = b.ourPage - a.ourPage;
    sumWord += wordSpan;
    sumOurs += ourSpan;
    const textless = anchors
      .slice(a.anchorIdx + 1, b.anchorIdx)
      .filter((x) => x.pageText.length < 8).length;
    const flag = wordSpan === ourSpan ? '  ' : wordSpan > ourSpan ? 'MERGE' : 'SPLIT';
    if (wordSpan !== ourSpan || textless > 0) {
      console.log(
        `${flag} word p${a.anchorIdx + 2}→p${b.anchorIdx + 2} (${wordSpan} pg, ${textless} textless) = ours p${a.ourPage + 1}→p${b.ourPage + 1} (${ourSpan} pg)  "${anchors[a.anchorIdx].pageText.slice(0, 34)}"`
      );
    }
  }
  console.log(`\nTOTAL matched intervals: word ${sumWord} pages vs ours ${sumOurs} pages`);
});
