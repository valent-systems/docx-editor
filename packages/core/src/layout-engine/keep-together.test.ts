import { describe, test, expect } from 'bun:test';
import { computeKeepNextChains, calculateChainHeight } from './keep-together';
import type { FlowBlock, ParagraphBlock, ParagraphMeasure, TextRun } from './types';

const textRun = (text: string): TextRun =>
  ({ kind: 'text', text, formatting: {} }) as unknown as TextRun;

let id = 0;
const para = (text: string, attrs?: ParagraphBlock['attrs']): ParagraphBlock =>
  ({
    kind: 'paragraph',
    id: `p${++id}`,
    attrs: attrs ?? {},
    runs: text ? [textRun(text)] : [],
  }) as unknown as ParagraphBlock;

describe('computeKeepNextChains', () => {
  test('a keepNext heading anchors on the next paragraph', () => {
    const blocks: FlowBlock[] = [para('Heading', { keepNext: true }), para('Body text')];
    const chain = computeKeepNextChains(blocks).get(0)!;
    expect(chain.memberIndices).toEqual([0]);
    expect(chain.anchorIndex).toBe(1);
  });

  test('blank spacers ride the chain; the anchor is the first CONTENT paragraph', () => {
    // Heading1 + blank line + intro — the TPX pattern. Break policy (product
    // call 2026-07-17): match the Google Docs presentation — the heading is
    // kept with its real content, not with an empty line, so it moves to the
    // next page instead of stranding at a page bottom. (Desktop Word's cached
    // pagination leaves "heading + blank" at the bottom; deliberately not
    // matched.)
    const blocks: FlowBlock[] = [
      para('UCx with Webex', { keepNext: true, keepLines: true }),
      para(''), // empty spacer — rides the chain
      para('   '), // whitespace-only — also a spacer
      para('UCx with Webex Hosted Unified Communications Service ...'),
    ];
    const chain = computeKeepNextChains(blocks).get(0)!;
    expect(chain.memberIndices).toEqual([0, 1, 2]);
    expect(chain.anchorIndex).toBe(3);
  });

  test('consecutive keepNext paragraphs still chain as before', () => {
    const blocks: FlowBlock[] = [
      para('H1', { keepNext: true }),
      para('H2', { keepNext: true }),
      para('Body'),
    ];
    const chain = computeKeepNextChains(blocks).get(0)!;
    expect(chain.memberIndices).toEqual([0, 1]);
    expect(chain.anchorIndex).toBe(2);
  });

  test('a paragraph with content and no keepNext is never pulled into the chain', () => {
    const blocks: FlowBlock[] = [
      para('Heading', { keepNext: true }),
      para('Body A'),
      para('Body B'),
    ];
    const chain = computeKeepNextChains(blocks).get(0)!;
    expect(chain.memberIndices).toEqual([0]);
    expect(chain.anchorIndex).toBe(1);
  });
});

describe('calculateChainHeight anchor modes', () => {
  const measure = (lineHeights: number[]): ParagraphMeasure =>
    ({
      kind: 'paragraph',
      totalHeight: lineHeights.reduce((a, b) => a + b, 0),
      lines: lineHeights.map((h) => ({ lineHeight: h })),
    }) as unknown as ParagraphMeasure;

  const blocks: FlowBlock[] = [para('Heading', { keepNext: true }), para('Intro paragraph')];
  const chain = computeKeepNextChains(blocks).get(0)!;

  test('a SHORT anchor is kept whole (heading + intro travel together)', () => {
    const measures = [measure([20]), measure([17, 17, 17, 17])]; // anchor 68px
    // pageContentHeight 900 → cap 300 → anchor (68) kept whole.
    expect(calculateChainHeight(chain, blocks, measures, 900)).toBe(20 + 68);
  });

  test('a LONG anchor reserves only its first line (fragments under the heading)', () => {
    const long = Array.from({ length: 20 }, () => 35); // 700px
    const measures = [measure([20]), measure(long)];
    // 700 > 900/3 → first-line mode: no page-sized gaps.
    expect(calculateChainHeight(chain, blocks, measures, 900)).toBe(20 + 35);
  });

  test('without pageContentHeight the first-line rule applies (legacy callers)', () => {
    const measures = [measure([20]), measure([17, 17])];
    expect(calculateChainHeight(chain, blocks, measures)).toBe(20 + 17);
  });
});
