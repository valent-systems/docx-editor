import { describe, expect, test } from 'bun:test';
import type {
  FlowBlock,
  Measure,
  ParagraphBlock,
  ParagraphMeasure,
} from '@valent/docx-editor-core/layout-engine';
import { measureTableCellBlockVisualHeight } from '@valent/docx-editor-core/layout-bridge';

function paragraphBlock(
  runs: ParagraphBlock['runs'],
  attrs?: ParagraphBlock['attrs']
): ParagraphBlock {
  return {
    kind: 'paragraph',
    id: 'p1',
    runs,
    attrs,
  };
}

function paragraphMeasure(totalHeight: number, lineHeight: number): ParagraphMeasure {
  return {
    kind: 'paragraph',
    lines: [
      {
        fromRun: 0,
        toRun: 0,
        fromChar: 0,
        toChar: 0,
        width: lineHeight,
        ascent: lineHeight,
        descent: 0,
        lineHeight,
      },
    ],
    totalHeight,
  };
}

describe('measureTableCellBlockVisualHeight', () => {
  test('keeps normal paragraph total height for text content', () => {
    const block = paragraphBlock([{ kind: 'text', text: 'hello' }]);
    const measure = paragraphMeasure(17.9, 17.9);

    expect(measureTableCellBlockVisualHeight(block, measure)).toBe(17.9);
  });

  test('image-only paragraphs use the measured height, exactly as painted', () => {
    // Previously special-cased to the image's bare intrinsic height (29), but
    // the painter renders the line at the paragraph measure's height (image +
    // baseline descent) — the mismatch clipped icon-grid cells mid-glyph and
    // cut icons in half at row breaks (TPX). Dual-renderer rule: the cell box
    // is sized by the same measure the painter draws.
    const block = paragraphBlock([{ kind: 'image', src: 'logo.png', width: 186, height: 29 }], {
      spacing: { before: 0, after: 0 },
    });
    const measure = paragraphMeasure(34.859375, 34.859375);

    expect(measureTableCellBlockVisualHeight(block, measure)).toBe(34.859375);
  });

  test('falls back to totalHeight for non-paragraph measures', () => {
    const block: FlowBlock = {
      kind: 'image',
      id: 'img1',
      src: 'logo.png',
      width: 186,
      height: 29,
    };
    const measure: Measure = {
      kind: 'image',
      width: 186,
      height: 29,
    };

    expect(measureTableCellBlockVisualHeight(block, measure)).toBe(29);
  });
});
