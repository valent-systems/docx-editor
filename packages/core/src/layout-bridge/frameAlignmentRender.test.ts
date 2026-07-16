import { describe, test, expect } from 'bun:test';
import { headerFooterToProseDoc, toProseDoc } from '../prosemirror/conversion/toProseDoc';
import { fromProseDoc } from '../prosemirror/conversion/fromProseDoc';
import { toFlowBlocks } from './toFlowBlocks';
import type { BlockContent, Paragraph } from '../types/content';
import type { Document } from '../types/document';

// A framed paragraph (w:framePr) — e.g. a centered page number in a footer —
// positions via its frame, which the layout engine does not implement. On the
// header/footer render path we approximate the frame's xAlign as paragraph
// alignment. It must be render-only: the saved model keeps its framePr and
// gains no spurious w:jc.

const framedParagraph = (xAlign: string): BlockContent =>
  ({
    type: 'paragraph',
    content: [{ type: 'run', content: [{ type: 'text', text: '1' }] }],
    formatting: { frame: { xAlign, hAnchor: 'margin', wrap: 'none' } },
  }) as BlockContent;

const paragraphBlocks = (content: BlockContent[], frameAlignment: boolean) => {
  const pm = headerFooterToProseDoc(content);
  return toFlowBlocks(pm, { frameAlignment }).filter((b) => b.kind === 'paragraph');
};

describe('frame alignment on the header/footer render path', () => {
  test('a centered framed paragraph renders centered when frameAlignment is on', () => {
    const [block] = paragraphBlocks([framedParagraph('center')], true);
    expect(block?.attrs?.alignment).toBe('center');
  });

  test('right / outside map to right; inside maps to left', () => {
    const alignOf = (xAlign: string) =>
      paragraphBlocks([framedParagraph(xAlign)], true)[0]?.attrs?.alignment;
    expect(alignOf('right')).toBe('right');
    expect(alignOf('outside')).toBe('right');
    expect(alignOf('inside')).toBe('left');
  });

  test('body path (frameAlignment off) leaves a framed paragraph unaligned', () => {
    const [block] = paragraphBlocks([framedParagraph('center')], false);
    expect(block?.attrs?.alignment).toBeUndefined();
  });

  test('an explicit alignment still wins over the frame', () => {
    const p = framedParagraph('center') as Paragraph;
    p.formatting!.alignment = 'right';
    const [block] = paragraphBlocks([p], true);
    expect(block?.attrs?.alignment).toBe('right');
  });
});

describe('frame save fidelity (round-trip)', () => {
  test('framePr survives toProseDoc → fromProseDoc with no spurious alignment', () => {
    const doc = {
      package: {
        headers: new Map(),
        footers: new Map(),
        document: {
          content: [framedParagraph('center')],
          sections: [],
        },
      },
    } as unknown as Document;

    const pm = toProseDoc(doc);
    const back = fromProseDoc(pm, doc);
    const para = back.package.document.content[0] as Paragraph;

    // The frame is preserved (via _originalFormatting) and NO w:jc was injected.
    expect(para.formatting?.frame?.xAlign).toBe('center');
    expect(para.formatting?.alignment).toBeUndefined();
  });
});
