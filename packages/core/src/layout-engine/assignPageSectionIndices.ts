/**
 * Assign each laid-out page the index of the section it belongs to.
 *
 * The painter uses `page.sectionIndex` to pick that section's header/footer from
 * the per-section array produced by `collectSectionHeaderFooters`. Both this and
 * that array are ordered by document position, so the index lines up directly.
 *
 * A page's section is that of its first content fragment: the number of section
 * breaks whose flow position precedes that fragment. A section-ending paragraph
 * carries the `w:sectPr`, and `toFlowBlocks` emits the section-break block AFTER
 * it, so the paragraph stays in the section it ends (e.g. a cover page resolves
 * to its own section, not the next). Pages with no fragment (rare) inherit the
 * previous page's section.
 *
 * Limitation: a `continuous` section break changes sections mid-page. This uses
 * the section in effect at the page's top; Word draws the footer from the
 * section in effect at the page bottom. The overwhelmingly common `nextPage`
 * break — where each section starts on a fresh page — is exact.
 */

import type { BlockId, FlowBlock, Page } from './types';

export function assignPageSectionIndices(
  pages: Page[],
  blocks: FlowBlock[],
  breakIndices: number[]
): void {
  const flowIndexById = new Map<BlockId, number>();
  blocks.forEach((b, i) => flowIndexById.set(b.id, i));

  // breakIndices is ascending; count how many breaks precede a flow position.
  const sectionOf = (flowIndex: number): number => {
    let section = 0;
    for (const breakIndex of breakIndices) {
      if (breakIndex < flowIndex) section++;
      else break;
    }
    return section;
  };

  let previous: number | undefined;
  for (const page of pages) {
    const firstFragment = page.fragments[0];
    const flowIndex = firstFragment ? flowIndexById.get(firstFragment.blockId) : undefined;
    const section = flowIndex != null ? sectionOf(flowIndex) : (previous ?? 0);
    page.sectionIndex = section;
    // The first page whose section differs from the prior page starts a section
    // (page 1 always does). Drives the per-section title-page variant.
    page.isSectionStart = section !== previous;
    previous = section;
  }
}
