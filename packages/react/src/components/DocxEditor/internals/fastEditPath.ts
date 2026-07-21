/**
 * Typing fast path (docs/INCREMENTAL-LAYOUT.md, milestone M1).
 *
 * For the common keystroke — a plain-text edit inside one top-level,
 * non-list paragraph whose re-measured geometry (line count + total height)
 * is unchanged — skip the full layout pipeline: patch the block + measure
 * in place, repaint only the page(s) showing that paragraph, and let the
 * deferred settle pass (the ordinary full pipeline, scheduled at
 * typing-idle by the caller) reconcile everything else. The fast path only
 * has to be visually correct until settle runs; any ineligible or
 * ambiguous edit returns false and the caller runs the full pipeline
 * immediately, exactly as before.
 *
 * Deliberately conservative eligibility (widened in M2):
 * - every step's change range sits inside ONE top-level `paragraph` node
 * - the paragraph is not a list item (no `numId`) and not framed (framePr)
 * - old and new runs are text/tab/lineBreak only (no images, no shapes)
 * - re-measured at the fragment's width with no floating zones, the
 *   result matches the old measure's line count and total height — a
 *   mismatch implies float context or reflow, both full-pipeline work
 */

import type { EditorState, Transaction } from 'prosemirror-state';
import type {
  FlowBlock,
  Measure,
  Layout,
  ParagraphBlock,
  ParagraphMeasure,
} from '@valent/docx-editor-core/layout-engine';
import { convertSingleParagraph } from '@valent/docx-editor-core/layout-bridge/toFlowBlocks';
import { measureParagraph } from '@valent/docx-editor-core/layout-bridge';
import { repaintPage } from '@valent/docx-editor-core/layout-painter';
import type { Theme } from '@valent/docx-editor-core/types';

const FAST_RUN_KINDS = new Set(['text', 'tab', 'lineBreak']);
const HEIGHT_EPSILON = 0.5;

export interface FastEditContext {
  /** Live block/measure arrays from the last full pass — patched in place. */
  blocks: FlowBlock[];
  measures: Measure[];
  layout: Layout | null;
  pagesContainer: HTMLElement | null;
  theme: Theme | null | undefined;
}

/** Union of changed ranges in NEW-doc coordinates, or null if none. */
function changedRange(tr: Transaction): { from: number; to: number } | null {
  let from = Infinity;
  let to = -Infinity;
  for (let i = 0; i < tr.mapping.maps.length; i++) {
    tr.mapping.maps[i].forEach((_oldStart, _oldEnd, newStart, newEnd) => {
      // Map this step's new-range through the REMAINING steps to final coords.
      const rest = tr.mapping.slice(i + 1);
      from = Math.min(from, rest.map(newStart, -1));
      to = Math.max(to, rest.map(newEnd, 1));
    });
  }
  return from <= to ? { from, to } : null;
}

function paragraphIsEligible(node: {
  type: { name: string };
  attrs: Record<string, unknown>;
}): boolean {
  if (node.type.name !== 'paragraph') return false;
  if (node.attrs.numId != null) return false;
  const orig = node.attrs._originalFormatting as { frame?: unknown } | undefined;
  if (orig?.frame) return false;
  return true;
}

function runsAreFast(block: ParagraphBlock): boolean {
  for (const run of block.runs) {
    if (!FAST_RUN_KINDS.has((run as { kind: string }).kind)) return false;
  }
  return true;
}

/**
 * Attempt the fast path for a docChanged transaction. Returns true when the
 * edit was fully absorbed (blocks/measures patched, dirty pages repainted) —
 * the caller must then schedule the deferred settle pass. Returns false when
 * the edit needs the full pipeline now.
 */
export function attemptFastEdit(
  tr: Transaction,
  newState: EditorState,
  ctx: FastEditContext
): boolean {
  const { blocks, measures, layout, pagesContainer } = ctx;
  if (!layout || !pagesContainer || blocks.length === 0) return false;

  const range = changedRange(tr);
  if (!range) return false;

  // The whole change must live inside one top-level node.
  const doc = newState.doc;
  const clampedFrom = Math.min(range.from, doc.content.size);
  const $from = doc.resolve(clampedFrom);
  if ($from.depth < 1) return false;
  const nodeStart = $from.before(1);
  const node = doc.child($from.index(0));
  const nodeEnd = nodeStart + node.nodeSize;
  if (range.to > nodeEnd) return false;
  if (!paragraphIsEligible(node as unknown as Parameters<typeof paragraphIsEligible>[0])) {
    return false;
  }

  // Locate the old block: the paragraph's start position is before the edit,
  // so it is unchanged by this transaction — pmStart matches directly.
  const blockIndex = blocks.findIndex(
    (b) => b.kind === 'paragraph' && (b as ParagraphBlock).pmStart === nodeStart
  );
  if (blockIndex < 0) return false;
  const oldBlock = blocks[blockIndex] as ParagraphBlock;
  const oldMeasure = measures[blockIndex];
  if (oldMeasure?.kind !== 'paragraph') return false;
  if (!runsAreFast(oldBlock)) return false;

  // Find the fragment(s) painting this block and their pages.
  const fragmentPages: number[] = [];
  let fragmentWidth: number | null = null;
  for (let p = 0; p < layout.pages.length; p++) {
    for (const frag of layout.pages[p].fragments) {
      if (frag.kind === 'paragraph' && frag.blockId === oldBlock.id) {
        fragmentPages.push(p);
        fragmentWidth = frag.width;
      }
    }
  }
  if (fragmentPages.length === 0 || fragmentWidth == null) return false;

  // Re-convert and re-measure just this paragraph.
  const defaultTabStopTwips = doc.attrs?.defaultTabStopTwips as number | undefined;
  const newBlock = convertSingleParagraph(node, nodeStart, {
    theme: ctx.theme ?? undefined,
    defaultTabStopTwips,
  });
  if (!runsAreFast(newBlock)) return false;
  const newMeasure = measureParagraph(newBlock, fragmentWidth);

  const om = oldMeasure as ParagraphMeasure;
  if (newMeasure.lines.length !== om.lines.length) return false;
  if (Math.abs(newMeasure.totalHeight - om.totalHeight) > HEIGHT_EPSILON) return false;

  // Patch in place. Keep the old identity so fragments and the painter's
  // lookup still resolve; bump pmEnd by the transaction's size delta.
  const delta = tr.doc.content.size - tr.before.content.size;
  newBlock.id = oldBlock.id;
  blocks[blockIndex] = newBlock;
  measures[blockIndex] = newMeasure;
  for (const p of fragmentPages) {
    for (const frag of layout.pages[p].fragments) {
      if (frag.kind === 'paragraph' && frag.blockId === oldBlock.id) {
        if (frag.pmEnd !== undefined) frag.pmEnd += delta;
      }
    }
  }

  let painted = true;
  for (let i = 0; i < fragmentPages.length; i++) {
    painted =
      repaintPage(
        pagesContainer,
        fragmentPages[i],
        i === 0 ? { blockId: oldBlock.id, block: newBlock, measure: newMeasure } : undefined
      ) && painted;
  }
  return painted;
}
