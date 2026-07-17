/**
 * Keep Together Logic - Handle keepNext and keepLines paragraph properties
 *
 * DOCX paragraphs can have keepNext (keep with next paragraph) and keepLines
 * (keep all lines together) properties that affect pagination.
 */

import type { FlowBlock, ParagraphBlock, Measure, ParagraphMeasure } from './types';

/**
 * A chain of consecutive keepNext paragraphs.
 */
export type KeepNextChain = {
  /** Index of the first paragraph in the chain. */
  startIndex: number;
  /** Index of the last paragraph in the chain. */
  endIndex: number;
  /** All paragraph indices in the chain. */
  memberIndices: number[];
  /** Index of the anchor paragraph (first non-keepNext after chain), or -1 if none. */
  anchorIndex: number;
};

/**
 * True for a paragraph with no visible content (no runs, or only whitespace
 * text runs) — a blank spacer line.
 */
function isEmptyParagraph(para: ParagraphBlock): boolean {
  return (para.runs ?? []).every(
    (run) => run.kind === 'text' && !(run as { text?: string }).text?.trim()
  );
}

/**
 * Pre-scan blocks to find all keepNext chains.
 *
 * A keepNext chain is a sequence of consecutive paragraphs with keepNext=true,
 * followed by an anchor paragraph (the first non-keepNext paragraph).
 * The entire chain must stay on the same page as the anchor's first line.
 *
 * Returns a map from chain start index to chain info.
 */
export function computeKeepNextChains(blocks: FlowBlock[]): Map<number, KeepNextChain> {
  const chains = new Map<number, KeepNextChain>();
  const processed = new Set<number>();

  for (let i = 0; i < blocks.length; i++) {
    // Skip already-processed blocks (mid-chain members)
    if (processed.has(i)) continue;

    const block = blocks[i];
    // Only paragraphs can have keepNext
    if (block.kind !== 'paragraph') continue;

    const para = block as ParagraphBlock;
    // Skip paragraphs without keepNext
    if (!para.attrs?.keepNext) continue;

    // Found a keepNext paragraph - scan forward to find full chain
    const memberIndices: number[] = [i];
    let endIndex = i;

    for (let j = i + 1; j < blocks.length; j++) {
      const nextBlock = blocks[j];

      // Breaks terminate the chain
      if (
        nextBlock.kind === 'sectionBreak' ||
        nextBlock.kind === 'pageBreak' ||
        nextBlock.kind === 'columnBreak'
      ) {
        break;
      }

      // Non-paragraphs terminate the chain
      if (nextBlock.kind !== 'paragraph') {
        break;
      }

      const nextPara = nextBlock as ParagraphBlock;
      if (nextPara.attrs?.keepNext) {
        // Continue the chain
        memberIndices.push(j);
        endIndex = j;
        processed.add(j);
      } else if (isEmptyParagraph(nextPara)) {
        // Blank spacer paragraphs ride the chain — the heading is kept with
        // its real CONTENT, not with an empty line. (Break policy: desktop
        // Word's cached pagination leaves "heading + blank" at a page
        // bottom; Google Docs — the presentation our customers reference —
        // pushes the heading to the next page. Product call 2026-07-17:
        // match the Google Docs behavior.)
        memberIndices.push(j);
        endIndex = j;
        processed.add(j);
      } else {
        // Found the anchor - stop here
        break;
      }
    }

    // Find the anchor (first paragraph after the chain)
    const potentialAnchor = endIndex + 1;
    let anchorIndex = -1;

    if (potentialAnchor < blocks.length) {
      const anchorBlock = blocks[potentialAnchor];
      // Anchor must not be a break
      if (
        anchorBlock.kind !== 'sectionBreak' &&
        anchorBlock.kind !== 'pageBreak' &&
        anchorBlock.kind !== 'columnBreak'
      ) {
        anchorIndex = potentialAnchor;
      }
    }

    // Record the chain
    chains.set(i, {
      startIndex: i,
      endIndex,
      memberIndices,
      anchorIndex,
    });
  }

  return chains;
}

/**
 * Fraction of the page content height below which a chain's anchor paragraph
 * is kept WHOLE with its heading (pushed together to the next page); above
 * it, only the anchor's first line is reserved and the paragraph fragments
 * under the heading. Break policy (product call 2026-07-17): matches the
 * Google Docs presentation — a heading is not stranded above a short section
 * intro, while a long paragraph still fills the heading's page instead of
 * leaving a page-sized gap (the anchor-fragment-trailing-spacing regression).
 */
const KEEP_WHOLE_ANCHOR_MAX_FRACTION = 1 / 3;

/**
 * Calculate the total height needed to keep a chain together.
 *
 * Includes all chain members plus the anchor paragraph — whole when it is
 * short (≤ KEEP_WHOLE_ANCHOR_MAX_FRACTION of the page content height), first
 * line only when it is long. Pass `pageContentHeight` to enable the
 * whole-anchor mode; without it the first-line rule applies.
 */
export function calculateChainHeight(
  chain: KeepNextChain,
  blocks: FlowBlock[],
  measures: Measure[],
  pageContentHeight?: number
): number {
  let totalHeight = 0;

  // Sum heights of all chain members
  for (const memberIndex of chain.memberIndices) {
    const block = blocks[memberIndex];
    const measure = measures[memberIndex];

    if (block.kind !== 'paragraph' || measure.kind !== 'paragraph') continue;

    const para = block as ParagraphBlock;
    const paraMeasure = measure as ParagraphMeasure;

    // Add spacing before (simplified - could be more sophisticated with collapse)
    const spacingBefore = para.attrs?.spacing?.before ?? 0;
    totalHeight += spacingBefore;

    // Add paragraph height
    totalHeight += paraMeasure.totalHeight;

    // Add spacing after
    const spacingAfter = para.attrs?.spacing?.after ?? 0;
    totalHeight += spacingAfter;
  }

  // Add the anchor paragraph: whole when short (heading + intro travel
  // together), first line only when long (paragraph fragments under the
  // heading). See KEEP_WHOLE_ANCHOR_MAX_FRACTION.
  if (chain.anchorIndex !== -1) {
    const anchorMeasure = measures[chain.anchorIndex];
    if (anchorMeasure?.kind === 'paragraph') {
      const anchorPara = anchorMeasure as ParagraphMeasure;
      const keepWholeCap =
        pageContentHeight !== undefined
          ? pageContentHeight * KEEP_WHOLE_ANCHOR_MAX_FRACTION
          : -Infinity;
      if (anchorPara.totalHeight <= keepWholeCap) {
        totalHeight += anchorPara.totalHeight;
      } else if (anchorPara.lines.length > 0) {
        totalHeight += anchorPara.lines[0].lineHeight;
      }
    }
  }

  return totalHeight;
}

/**
 * Get the set of indices that are mid-chain (not chain starters).
 * These should skip the keepNext check since their chain starter already decided.
 */
export function getMidChainIndices(chains: Map<number, KeepNextChain>): Set<number> {
  const midChain = new Set<number>();

  for (const chain of chains.values()) {
    // All members except the first are mid-chain
    for (let i = 1; i < chain.memberIndices.length; i++) {
      midChain.add(chain.memberIndices[i]);
    }
  }

  return midChain;
}

/**
 * Check if a paragraph has keepLines property (all lines must stay together).
 */
export function hasKeepLines(block: FlowBlock): boolean {
  if (block.kind !== 'paragraph') return false;
  const para = block as ParagraphBlock;
  return para.attrs?.keepLines === true;
}

/**
 * Check if a paragraph should start on a new page (pageBreakBefore).
 */
export function hasPageBreakBefore(block: FlowBlock): boolean {
  if (block.kind !== 'paragraph') return false;
  const para = block as ParagraphBlock;
  return para.attrs?.pageBreakBefore === true;
}

/**
 * True when the paragraph's break comes from a leading `w:br type="page"` —
 * an explicit action that always breaks (even onto an empty page), unlike
 * pPr `pageBreakBefore`, which is conditional. See paginator.forcePageBreak.
 */
export function hasExplicitPageBreakBefore(block: FlowBlock): boolean {
  if (block.kind !== 'paragraph') return false;
  return (block as ParagraphBlock).attrs?.explicitPageBreakBefore === true;
}
