/**
 * Regression — a text box anchored from a run inside a table cell must RENDER in
 * the editor (it previously vanished: toProseDoc dropped it, so the painter never
 * saw it). It is rendered in-flow inside the cell (Option A).
 *
 * Fixture: e2e/fixtures/textbox-in-table-cell.docx.
 */
import { afterAll, beforeAll, describe, expect, test } from 'bun:test';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { parseDocx } from '../../docx/parser';
import { toProseDoc } from '../../prosemirror/conversion/toProseDoc';
import { toFlowBlocks } from '../../layout-bridge/toFlowBlocks';
import { measureParagraph } from '../../layout-bridge/measuring';
import { measureTableBlock } from '../../layout-bridge/measureTable';
import { layoutDocument } from '../../layout-engine';
import {
  DEFAULT_TEXTBOX_MARGINS,
  DEFAULT_TEXTBOX_WIDTH,
  type FlowBlock,
  type Measure,
  type TextBoxBlock,
  type TableBlock,
} from '../../layout-engine/types';
import { renderPage } from '../renderPage';

const FIXTURE = resolve(process.cwd(), 'e2e/fixtures/textbox-in-table-cell.docx');
const CONTENT_WIDTH = 624;

let originalGetContext: typeof HTMLCanvasElement.prototype.getContext | undefined;

beforeAll(() => {
  GlobalRegistrator.register({
    settings: {
      disableCSSFileLoading: true,
      disableJavaScriptFileLoading: true,
      handleDisabledFileLoadingAsSuccess: true,
    },
  });
  originalGetContext = HTMLCanvasElement.prototype.getContext;
  HTMLCanvasElement.prototype.getContext = function getContext(type: string) {
    if (type !== '2d') return null;
    return {
      font: '',
      measureText: (text: string) => ({
        width: text.length * 9,
        actualBoundingBoxAscent: 12,
        actualBoundingBoxDescent: 4,
      }),
    } as unknown as CanvasRenderingContext2D;
  } as typeof HTMLCanvasElement.prototype.getContext;
});

afterAll(() => {
  if (originalGetContext) HTMLCanvasElement.prototype.getContext = originalGetContext;
  GlobalRegistrator.unregister();
});

function measureBlock(block: FlowBlock, contentWidth: number): Measure {
  if (block.kind === 'paragraph') return measureParagraph(block, contentWidth);
  if (block.kind === 'table')
    return measureTableBlock(block as TableBlock, contentWidth, measureBlock);
  if (block.kind === 'textBox') {
    const tb = block as TextBoxBlock;
    const m = tb.margins ?? DEFAULT_TEXTBOX_MARGINS;
    const innerWidth = (tb.width ?? DEFAULT_TEXTBOX_WIDTH) - m.left - m.right;
    const innerMeasures = tb.content.map((p) => measureParagraph(p, innerWidth));
    return {
      kind: 'textBox',
      width: tb.width ?? DEFAULT_TEXTBOX_WIDTH,
      height: tb.height ?? innerMeasures.reduce((s, x) => s + x.totalHeight, 0) + m.top + m.bottom,
      innerMeasures,
    };
  }
  if (block.kind === 'pageBreak' || block.kind === 'columnBreak' || block.kind === 'sectionBreak') {
    return { kind: block.kind };
  }
  throw new Error(`Unexpected fixture block kind: ${block.kind}`);
}

describe('text box anchored in a table cell — editor render', () => {
  test('paints the cell text box, in-flow inside the table cell', async () => {
    const buffer = readFileSync(FIXTURE);
    const parsed = await parseDocx(
      buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength)
    );
    const pmDoc = toProseDoc(parsed, { styles: parsed.package.styles });
    const blocks = toFlowBlocks(pmDoc, { theme: parsed.package.theme });
    const measures = blocks.map((b) => measureBlock(b, CONTENT_WIDTH));
    const layout = layoutDocument(blocks, measures, {
      pageSize: { w: 816, h: 1056 },
      margins: { top: 96, right: 96, bottom: 96, left: 96, header: 48, footer: 48 },
    });
    const blockLookup = new Map(
      blocks.map((block, index) => [String(block.id), { block, measure: measures[index] }])
    );

    const pageEl = renderPage(
      layout.pages[0],
      { pageNumber: 1, totalPages: 1, contentWidth: CONTENT_WIDTH, section: 'body' },
      { document, blockLookup }
    );

    const box = pageEl.querySelector<HTMLElement>('.layout-textbox');
    expect(box).not.toBeNull();
    expect(box?.textContent).toContain('Internal use only');
    // It renders INSIDE a table cell (not hoisted to the page).
    expect(box?.closest('.layout-table-cell')).not.toBeNull();
    // In-flow (Option A): relative, not page-absolute.
    expect(box?.style.position).toBe('relative');
  });
});
