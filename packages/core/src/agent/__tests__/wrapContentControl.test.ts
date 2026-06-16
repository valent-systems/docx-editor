/**
 * Creating inline content controls around occurrence-precise placeholder spans
 * (`wrapInlineContentControl`) — run-splitting, occurrence targeting, table
 * cells, and the guard for spans crossing un-splittable inline content.
 */

import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, test } from 'bun:test';

import { wrapInlineContentControl } from '../wrapContentControl';
import {
  findContentControls,
  findContentControl,
  setContentControlContent,
  fillContentControl,
} from '../contentControls';
import { getParagraphText } from '../text-utils';
import { parseDocx } from '../../docx/parser';
import { createDocx } from '../../docx/rezip';
import type {
  Document,
  BlockContent,
  Paragraph,
  Run,
  Table,
  InlineSdt,
} from '../../types/document';

const COMPREHENSIVE_FIXTURE = join(
  import.meta.dir,
  '../../../../../e2e/fixtures/block-sdt-comprehensive.docx'
);
async function loadComprehensive(): Promise<Document> {
  const buf = readFileSync(COMPREHENSIVE_FIXTURE);
  return parseDocx(buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength));
}

// ── builders ────────────────────────────────────────────────────────────────

function run(text: string): Run {
  return { type: 'run', content: [{ type: 'text', text }] };
}
function para(runs: Run[], paraId?: string): Paragraph {
  return { type: 'paragraph', content: runs, ...(paraId ? { paraId } : {}) };
}
function cell(blocks: BlockContent[]) {
  return { type: 'tableCell' as const, content: blocks };
}
function table(rows: Paragraph[][][]): Table {
  return {
    type: 'table',
    rows: rows.map((cells) => ({ type: 'tableRow' as const, cells: cells.map((b) => cell(b)) })),
  } as unknown as Table;
}
function doc(content: BlockContent[]): Document {
  return { package: { document: { content } } } as unknown as Document;
}

/** Plain text of the document body in reading order (paragraphs + cells). */
function bodyText(d: Document): string {
  const parts: string[] = [];
  const walk = (blocks: BlockContent[]): void => {
    for (const b of blocks) {
      if (b.type === 'paragraph') parts.push(getParagraphText(b));
      else if (b.type === 'table') for (const r of b.rows) for (const c of r.cells) walk(c.content);
      else if (b.type === 'blockSdt') walk(b.content);
    }
  };
  walk(d.package.document.content);
  return parts.join('|');
}

// ── single-paragraph wrapping ────────────────────────────────────────────────

describe('wrapInlineContentControl — run splitting', () => {
  test('wraps a whole run as its own inline control', () => {
    const d = doc([para([run('Owner: '), run('[INSERT]'), run(' (end)')])]);
    const res = wrapInlineContentControl(d, { text: '[INSERT]' }, { tag: 'f1' });
    expect(res.status).toBe('wrapped');
    if (res.status !== 'wrapped') return;

    const found = findContentControl(res.doc, { tag: 'f1' });
    expect(found?.kind).toBe('inline');
    expect(found?.text).toBe('[INSERT]');
    // The surrounding text is unchanged and still readable.
    expect(getParagraphText(res.doc.package.document.content[0] as Paragraph)).toBe(
      'Owner: [INSERT] (end)'
    );
  });

  test('splits a run mid-text to wrap a partial span (placeholder + trailing period)', () => {
    const d = doc([para([run('see [volume].')])]);
    const res = wrapInlineContentControl(d, { text: '[volume]' }, { tag: 'v' });
    expect(res.status).toBe('wrapped');
    if (res.status !== 'wrapped') return;

    const p = res.doc.package.document.content[0] as Paragraph;
    // before-run "see ", the control, after-run "."
    const sdt = p.content.find((i) => i.type === 'inlineSdt') as InlineSdt;
    expect(sdt).toBeDefined();
    expect(
      sdt.content.map((r) => (r.type === 'run' ? (r.content[0] as any).text : '')).join('')
    ).toBe('[volume]');
    expect(getParagraphText(p)).toBe('see [volume].'); // text identical
  });

  test('wraps a span that crosses several runs', () => {
    const d = doc([para([run('[Eng'), run('land and '), run('Wales]')])]);
    const res = wrapInlineContentControl(d, { text: '[England and Wales]' }, { tag: 'c' });
    expect(res.status).toBe('wrapped');
    if (res.status !== 'wrapped') return;
    expect(findContentControl(res.doc, { tag: 'c' })?.text).toBe('[England and Wales]');
  });
});

// ── occurrence-precision ─────────────────────────────────────────────────────

describe('wrapInlineContentControl — occurrence precision', () => {
  test('two identical spans in different paragraphs get independent controls', () => {
    let d = doc([
      para([run('governed by '), run('[LIST]')], 'AAAA'),
      para([run('disputes by '), run('[LIST]')], 'BBBB'),
    ]);
    const r0 = wrapInlineContentControl(d, { text: '[LIST]', occurrence: 0 }, { tag: 'gov' });
    expect(r0.status).toBe('wrapped');
    if (r0.status !== 'wrapped') return;
    const r1 = wrapInlineContentControl(r0.doc, { text: '[LIST]', occurrence: 1 }, { tag: 'jur' });
    expect(r1.status).toBe('wrapped');
    if (r1.status !== 'wrapped') return;
    d = r1.doc;

    // Two distinct controls, each text '[LIST]'.
    expect(
      findContentControls(d)
        .map((c) => c.tag)
        .sort()
    ).toEqual(['gov', 'jur']);

    // Filling each by tag is independent — no cross-contamination.
    let filled = setContentControlContent(d, { tag: 'gov' }, 'France');
    filled = setContentControlContent(filled, { tag: 'jur' }, 'Germany');
    expect(bodyText(filled)).toBe('governed by France|disputes by Germany');
  });

  test('occurrence index selects the right match within one paragraph', () => {
    const d = doc([para([run('from [X] to [X]')])]);
    const res = wrapInlineContentControl(d, { text: '[X]', occurrence: 1 }, { tag: 'second' });
    expect(res.status).toBe('wrapped');
    if (res.status !== 'wrapped') return;
    // The first [X] is untouched (no control), the second is wrapped.
    const controls = findContentControls(res.doc);
    expect(controls.length).toBe(1);
    const filled = setContentControlContent(res.doc, { tag: 'second' }, 'END');
    expect(getParagraphText(filled.package.document.content[0] as Paragraph)).toBe(
      'from [X] to END'
    );
  });

  test('paraId scopes the occurrence to one paragraph', () => {
    const d = doc([
      para([run('30/40/60/90EOM+5')], 'HEAD'),
      para([run('following 30/40/60/90 days')], 'SENT'),
    ]);
    // paraId-scoped occurrence 0 targets the sentence, never the heading.
    const res = wrapInlineContentControl(
      d,
      { text: '30/40/60/90', paraId: 'SENT', occurrence: 0 },
      { tag: 'pay' }
    );
    expect(res.status).toBe('wrapped');
    if (res.status !== 'wrapped') return;
    const filled = setContentControlContent(res.doc, { tag: 'pay' }, '60');
    expect(bodyText(filled)).toBe('30/40/60/90EOM+5|following 60 days'); // heading untouched
  });
});

// ── table cells ──────────────────────────────────────────────────────────────

describe('wrapInlineContentControl — table cells', () => {
  test('wraps a placeholder inside a table cell and discovers it there', () => {
    const d = doc([
      para([run('Agreement number')]),
      table([[[para([run('Agreement number')])], [para([run('[INSERT]')])]]]),
    ]);
    const res = wrapInlineContentControl(d, { text: '[INSERT]' }, { tag: 'num' });
    expect(res.status).toBe('wrapped');
    if (res.status !== 'wrapped') return;

    const found = findContentControl(res.doc, { tag: 'num' });
    expect(found?.kind).toBe('inline');
    expect(found?.container).toBe('body');
    expect(found?.text).toBe('[INSERT]');

    const filled = setContentControlContent(res.doc, { tag: 'num' }, 'AGR-001');
    expect(findContentControl(filled, { tag: 'num' })?.text).toBe('AGR-001');
  });
});

// ── failure outcomes ─────────────────────────────────────────────────────────

describe('wrapInlineContentControl — outcomes', () => {
  test('not-found when the text is absent', () => {
    const d = doc([para([run('nothing here')])]);
    expect(wrapInlineContentControl(d, { text: '[X]' }, { tag: 't' }).status).toBe('not-found');
  });

  test('occurrence-out-of-range reports the available count', () => {
    const d = doc([para([run('[X] and [X]')])]);
    const res = wrapInlineContentControl(d, { text: '[X]', occurrence: 5 }, { tag: 't' });
    expect(res.status).toBe('occurrence-out-of-range');
    if (res.status === 'occurrence-out-of-range') expect(res.matches).toBe(2);
  });

  test('refuses a span that crosses un-splittable inline content (a hyperlink)', () => {
    const link = {
      type: 'hyperlink' as const,
      children: [run('mid')],
    };
    // Logical text is "[a" + "mid" + "b]" = "[amidb]"; the span crosses the
    // hyperlink, which the wrapper cannot split.
    const p: Paragraph = { type: 'paragraph', content: [run('[a'), link as any, run('b]')] };
    const d = doc([p]);
    const res = wrapInlineContentControl(d, { text: '[amidb]' }, { tag: 't' });
    expect(res.status).toBe('crosses-inline-boundary');
  });

  test('source document is not mutated (pure)', () => {
    const d = doc([para([run('[X]')])]);
    const before = JSON.stringify(d);
    wrapInlineContentControl(d, { text: '[X]' }, { tag: 't' });
    expect(JSON.stringify(d)).toBe(before);
  });
});

// ── real .docx table-cell round-trip ─────────────────────────────────────────

describe('wrapInlineContentControl — table-cell round-trip (real .docx)', () => {
  // `m-B1` is the text of a cell inside the `multi` block control's table in the
  // comprehensive fixture — i.e. a span inside a table cell that is itself inside
  // a content control. Planting an inline control there, saving, and re-opening
  // exercises the full cell round-trip the CLM pipeline depends on.
  test('plants a control in a table cell, survives save → re-open, and fills', async () => {
    const doc0 = await loadComprehensive();
    const before = findContentControls(doc0).length;

    const wrapped = wrapInlineContentControl(doc0, { text: 'm-B1' }, { tag: 'cell-field' });
    expect(wrapped.status).toBe('wrapped');
    if (wrapped.status !== 'wrapped') return;

    // Save → re-open: the planted inline control survives inside the cell.
    const reopened = await parseDocx(await createDocx(wrapped.doc));
    const found = findContentControl(reopened, { tag: 'cell-field' });
    expect(found).toBeDefined();
    expect(found?.kind).toBe('inline');
    expect(found?.text).toBe('m-B1');
    // Every pre-existing control is still present (table not corrupted).
    expect(findContentControls(reopened).length).toBe(before + 1);

    // Fill by tag → save → re-open: the value persists in the cell.
    const filled = fillContentControl(reopened, { tag: 'cell-field' }, 'FILLED');
    expect(filled.status).toBe('filled');
    const reopened2 = await parseDocx(await createDocx(filled.doc!));
    expect(findContentControl(reopened2, { tag: 'cell-field' })?.text).toBe('FILLED');
  });
});
