import { describe, test, expect } from 'bun:test';
import { computeKeepNextChains } from './keep-together';
import type { FlowBlock, ParagraphBlock, TextRun } from './types';

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

  test('an EMPTY next paragraph is still the anchor (Word-accurate)', () => {
    // Heading1 + blank line + intro — the TPX pattern. Word keeps a keepNext
    // heading with the next paragraph mark even when it is blank (verified
    // against TPX's lastRenderedPageBreak record: Word leaves "heading +
    // blank" at a page bottom). Chaining through blanks would push headings
    // Word doesn't push and diverge pagination — deliberately NOT done.
    const blocks: FlowBlock[] = [
      para('UCx with Webex', { keepNext: true, keepLines: true }),
      para(''), // empty spacer — the anchor, per Word
      para('UCx with Webex Hosted Unified Communications Service ...'),
    ];
    const chain = computeKeepNextChains(blocks).get(0)!;
    expect(chain.memberIndices).toEqual([0]);
    expect(chain.anchorIndex).toBe(1);
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
