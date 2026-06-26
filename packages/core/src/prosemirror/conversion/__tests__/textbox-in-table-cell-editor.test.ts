/**
 * Regression — a text box anchored from a run inside a TABLE CELL must survive
 * the editor (ProseMirror) round-trip, not just the headless file round-trip.
 *
 * The body parser promotes a paragraph's anchored text boxes to sibling `textBox`
 * PM nodes; the cell path didn't, so a cell-anchored box was dropped by
 * toProseDoc — it never rendered in the editor and was lost on an in-editor save.
 * Now cell paragraphs go through the same promotion (and the schema permits a
 * `textBox` inside a `tableCell`), so the box surfaces as a node in the cell and
 * round-trips back into the cell's content on save.
 *
 * Fixture: e2e/fixtures/textbox-in-table-cell.docx (synthetic; a one-cell table
 * whose cell anchors a text box reading "Internal use only").
 */
import { describe, test, expect } from 'bun:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import JSZip from 'jszip';
import { parseDocx } from '../../../docx/parser';
import { repackDocx } from '../../../docx/rezip';
import { toProseDoc } from '../toProseDoc';
import { fromProseDoc } from '../fromProseDoc';

const FIXTURE = join(import.meta.dir, '../../../../../../e2e/fixtures/textbox-in-table-cell.docx');
const BOX_TEXT = 'Internal use only';

function load() {
  const b = readFileSync(FIXTURE);
  return parseDocx(b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength));
}

describe('text box in a table cell — editor (PM) round-trip', () => {
  test('toProseDoc surfaces the cell box as a textBox node inside the cell', async () => {
    const pm = toProseDoc(await load());
    let total = 0;
    let inCell = 0;
    pm.descendants((n) => {
      if (n.type.name === 'textBox') total++;
    });
    pm.descendants((n) => {
      if (n.type.name === 'tableCell') {
        n.descendants((c) => {
          if (c.type.name === 'textBox') inCell++;
        });
      }
    });
    expect(total).toBe(1);
    expect(inCell).toBe(1);
  });

  test('an in-editor open → export keeps the cell text box', async () => {
    const doc = await load();
    const out = await repackDocx(fromProseDoc(toProseDoc(doc), doc));
    const xml = await (await JSZip.loadAsync(out)).file('word/document.xml')!.async('text');
    expect(xml).toContain('<w:txbxContent>');
    expect(xml).toContain(BOX_TEXT);
    // and the box is still inside the table, not hoisted out of it.
    expect(/<w:tbl>[\s\S]*Internal use only[\s\S]*<\/w:tbl>/.test(xml)).toBe(true);
  });
});
