import { describe, test, expect } from 'bun:test';
import { assignPageSectionIndices } from './assignPageSectionIndices';
import type { FlowBlock, Page } from './types';

// Stamps each page with the section it belongs to, so the painter can select
// that section's header/footer. Section index of a page = number of section
// breaks preceding its first content fragment. This ordering must match
// `collectSectionHeaderFooters` (both walk document order) so page.sectionIndex
// indexes straight into that array.

const block = (id: string): FlowBlock => ({ id }) as unknown as FlowBlock;
const page = (blockIds: string[]): Page =>
  ({ fragments: blockIds.map((blockId) => ({ blockId })) }) as unknown as Page;

describe('assignPageSectionIndices', () => {
  test('assigns each page the section of its first fragment', () => {
    // b0 b1 | [break@2] | b2 b3 | [break@5] | b4
    const blocks = [
      block('b0'),
      block('b1'),
      block('sb0'),
      block('b2'),
      block('b3'),
      block('sb1'),
      block('b4'),
    ];
    const breakIndices = [2, 5];
    const pages = [page(['b0', 'b1']), page(['b2', 'b3']), page(['b4'])];

    assignPageSectionIndices(pages, blocks, breakIndices);

    expect(pages[0].sectionIndex).toBe(0);
    expect(pages[1].sectionIndex).toBe(1);
    expect(pages[2].sectionIndex).toBe(2);
    // Each of these pages starts a new section.
    expect(pages.map((p) => p.isSectionStart)).toEqual([true, true, true]);
  });

  test('isSectionStart is true only on the first page of each section', () => {
    // Section 0 spans two pages, then section 1.
    const blocks = [block('a'), block('b'), block('sb0'), block('c')];
    const breakIndices = [2];
    const pages = [page(['a']), page(['b']), page(['c'])];

    assignPageSectionIndices(pages, blocks, breakIndices);

    expect(pages.map((p) => p.sectionIndex)).toEqual([0, 0, 1]);
    expect(pages.map((p) => p.isSectionStart)).toEqual([true, false, true]);
  });

  test('a section-ending paragraph stays in the section it ends (break block follows it)', () => {
    // The cover paragraph carries the sectPr; its break block sits AFTER it, so
    // the cover page must resolve to section 0 — not the next section.
    const blocks = [block('cover'), block('coverBreak'), block('body')];
    const breakIndices = [1];
    const pages = [page(['cover']), page(['body'])];

    assignPageSectionIndices(pages, blocks, breakIndices);

    expect(pages[0].sectionIndex).toBe(0); // cover → its own section
    expect(pages[1].sectionIndex).toBe(1);
  });

  test('an empty page inherits the previous page section', () => {
    const blocks = [block('b0'), block('sb0'), block('b1')];
    const breakIndices = [1];
    const pages = [page(['b0']), page([]), page(['b1'])];

    assignPageSectionIndices(pages, blocks, breakIndices);

    expect(pages[0].sectionIndex).toBe(0);
    expect(pages[1].sectionIndex).toBe(0); // inherited
    expect(pages[2].sectionIndex).toBe(1);
  });

  test('single-section document → every page is section 0', () => {
    const blocks = [block('b0'), block('b1')];
    const pages = [page(['b0']), page(['b1'])];

    assignPageSectionIndices(pages, blocks, []);

    expect(pages[0].sectionIndex).toBe(0);
    expect(pages[1].sectionIndex).toBe(0);
  });
});
