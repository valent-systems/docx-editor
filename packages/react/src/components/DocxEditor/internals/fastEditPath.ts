/**
 * Typing fast path (docs/INCREMENTAL-LAYOUT.md, milestones M1+M2).
 *
 * For the common keystroke — a plain-text edit inside one paragraph whose
 * re-measured geometry (line count + total height) is unchanged — skip the
 * full layout pipeline: patch the block + measure in place, repaint only
 * the page(s) showing it, and let the deferred settle pass (the ordinary
 * full pipeline, scheduled at typing-idle by the caller) reconcile
 * everything else. The fast path only has to be visually correct until
 * settle runs; any ineligible or ambiguous edit returns false and the
 * caller runs the full pipeline immediately, exactly as before.
 *
 * Two eligible shapes:
 * - a top-level paragraph (M1)
 * - a paragraph directly inside a non-nested table cell (M2) — re-measured
 *   at the cell's content width; the row only keeps its geometry when the
 *   paragraph's does, which the measure comparison guarantees
 *
 * Shared eligibility: text/tab/lineBreak runs only (no images/shapes), not
 * framed (framePr). List paragraphs ARE eligible — the pre-computed
 * `listMarker` text is carried over from the old block, since a text edit
 * cannot renumber anything (structure edits fail the one-paragraph range
 * check and take the full pipeline).
 */

import type { EditorState, Transaction } from 'prosemirror-state';
import type { Node as PMNode } from 'prosemirror-model';
import type {
  FlowBlock,
  Measure,
  Layout,
  ParagraphBlock,
  ParagraphMeasure,
  TableBlock,
  TableMeasure,
} from '@valent/docx-editor-core/layout-engine';
import { convertSingleParagraph } from '@valent/docx-editor-core/layout-bridge/toFlowBlocks';
import { measureParagraph } from '@valent/docx-editor-core/layout-bridge';
import { repaintPage } from '@valent/docx-editor-core/layout-painter';
import type { Theme } from '@valent/docx-editor-core/types';

const FAST_RUN_KINDS = new Set(['text', 'tab', 'lineBreak']);
const HEIGHT_EPSILON = 0.5;
/** Word's TableNormal default cell side padding (see measureTable.ts). */
const DEFAULT_CELL_PADDING_X = 7;

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

function runsAreFast(block: ParagraphBlock): boolean {
  for (const run of block.runs) {
    if (!FAST_RUN_KINDS.has((run as { kind: string }).kind)) return false;
  }
  return true;
}

/**
 * Convert + re-measure one paragraph (eligibility-checked). Returns the new
 * pair (identity taken from `oldBlock`) or null when ineligible. Geometry
 * comparison is the caller's job — M1 requires it unchanged; M2b absorbs a
 * height change within the page's slack.
 */
function remeasureEligibleParagraph(
  parNode: PMNode,
  parStart: number,
  width: number,
  oldBlock: ParagraphBlock,
  ctx: FastEditContext,
  defaultTabStopTwips: number | undefined
): { newBlock: ParagraphBlock; newMeasure: ParagraphMeasure } | null {
  const orig = parNode.attrs._originalFormatting as { frame?: unknown } | undefined;
  if (orig?.frame) return null;
  if (!runsAreFast(oldBlock)) return null;

  const newBlock = convertSingleParagraph(parNode, parStart, {
    theme: ctx.theme ?? undefined,
    defaultTabStopTwips,
  });
  if (!runsAreFast(newBlock)) return null;

  // Numbered paragraphs: the marker text needs document-wide counter state
  // that convertSingleParagraph doesn't have. A text edit can't renumber, so
  // carry the old pre-computed marker over.
  const oldAttrs = oldBlock.attrs;
  if (parNode.attrs.numId != null && oldAttrs && newBlock.attrs) {
    newBlock.attrs.listMarker = oldAttrs.listMarker;
    newBlock.attrs.listMarkerHidden = oldAttrs.listMarkerHidden;
    newBlock.attrs.listMarkerFontFamily = oldAttrs.listMarkerFontFamily;
    newBlock.attrs.listMarkerFontSize = oldAttrs.listMarkerFontSize;
  }

  const newMeasure = measureParagraph(newBlock, width);
  newBlock.id = oldBlock.id;
  return { newBlock, newMeasure };
}

function geometryUnchanged(a: ParagraphMeasure, b: ParagraphMeasure): boolean {
  return (
    a.lines.length === b.lines.length && Math.abs(a.totalHeight - b.totalHeight) <= HEIGHT_EPSILON
  );
}

/**
 * M2b: absorb a height/line-count change within one page. Eligible when the
 * paragraph paints as a single whole fragment on a single-column page, no
 * floating content (text boxes / images) sits below it on that page, and
 * growth still fits above the page's content bottom. Fragments below shift
 * by the height delta; the settle pass repaginates properly.
 *
 * Returns the pageIndex to repaint, or null when not absorbable.
 */
function tryAbsorbReflowInPage(
  layoutPages: Layout['pages'],
  pages: number[],
  blockId: FlowBlock['id'],
  oldMeasure: ParagraphMeasure,
  newMeasure: ParagraphMeasure
): number | null {
  if (pages.length !== 1) return null;
  const pageIndex = pages[0];
  const page = layoutPages[pageIndex];
  if ((page.columns?.count ?? 1) > 1) return null;

  const frags = page.fragments;
  const parFrag = frags.find((f) => f.blockId === blockId);
  if (!parFrag || parFrag.kind !== 'paragraph') return null;
  if (frags.filter((f) => f.blockId === blockId).length !== 1) return null;
  // Whole paragraph on this page (not a split fragment).
  if (parFrag.fromLine !== 0 || parFrag.toLine !== oldMeasure.lines.length) return null;

  const deltaH = newMeasure.totalHeight - oldMeasure.totalHeight;
  const parBottomOld = parFrag.y + parFrag.height;
  const contentBottom =
    page.size.h - page.margins.top - page.margins.bottom - (page.footnoteReservedHeight ?? 0);

  let maxBottom = parFrag.y + newMeasure.totalHeight;
  for (const f of frags) {
    if (f === parFrag) continue;
    const below = f.y >= parBottomOld - HEIGHT_EPSILON;
    if (below && (f.kind === 'textBox' || f.kind === 'image')) return null; // anchored content would misalign
    if (below) maxBottom = Math.max(maxBottom, f.y + f.height + deltaH);
  }
  if (deltaH > 0 && maxBottom > contentBottom + HEIGHT_EPSILON) return null;

  // Commit: resize the paragraph fragment, shift everything below.
  parFrag.height = newMeasure.totalHeight;
  parFrag.toLine = newMeasure.lines.length;
  for (const f of frags) {
    if (f === parFrag) continue;
    if (f.y >= parBottomOld - HEIGHT_EPSILON) f.y += deltaH;
  }
  return pageIndex;
}

/** Pages showing fragments of `blockId`, plus the paragraph-fragment width. */
function findFragmentPages(
  layout: Layout,
  blockId: FlowBlock['id']
): { pages: number[]; paragraphWidth: number | null } {
  const pages: number[] = [];
  let paragraphWidth: number | null = null;
  for (let p = 0; p < layout.pages.length; p++) {
    for (const frag of layout.pages[p].fragments) {
      if (frag.blockId !== blockId) continue;
      if (pages[pages.length - 1] !== p) pages.push(p);
      if (frag.kind === 'paragraph') paragraphWidth = frag.width;
    }
  }
  return { pages, paragraphWidth };
}

function repaintPages(
  container: HTMLElement,
  pages: number[],
  patch?: { blockId: string | number; block: unknown; measure: unknown }
): boolean {
  let ok = true;
  for (let i = 0; i < pages.length; i++) {
    ok = repaintPage(container, pages[i], i === 0 ? patch : undefined) && ok;
  }
  return ok;
}

/** Fast path for an edit inside a top-level paragraph (M1). */
function attemptParagraphFastEdit(
  node: PMNode,
  nodeStart: number,
  delta: number,
  ctx: FastEditContext,
  defaultTabStopTwips: number | undefined
): boolean {
  const { blocks, measures, layout, pagesContainer } = ctx;
  const blockIndex = blocks.findIndex(
    (b) => b.kind === 'paragraph' && (b as ParagraphBlock).pmStart === nodeStart
  );
  if (blockIndex < 0) return false;
  const oldBlock = blocks[blockIndex] as ParagraphBlock;
  const oldMeasure = measures[blockIndex];
  if (oldMeasure?.kind !== 'paragraph') return false;

  const { pages, paragraphWidth } = findFragmentPages(layout!, oldBlock.id);
  if (pages.length === 0 || paragraphWidth == null) return false;

  const patched = remeasureEligibleParagraph(
    node,
    nodeStart,
    paragraphWidth,
    oldBlock,
    ctx,
    defaultTabStopTwips
  );
  if (!patched) return false;

  const om = oldMeasure as ParagraphMeasure;
  if (!geometryUnchanged(patched.newMeasure, om)) {
    // M2b: a line wrapped (or unwrapped) — absorb the height change within
    // the page's slack instead of bailing to the full pipeline.
    const absorbed = tryAbsorbReflowInPage(
      layout!.pages,
      pages,
      oldBlock.id,
      om,
      patched.newMeasure
    );
    if (absorbed == null) return false;
  }

  blocks[blockIndex] = patched.newBlock;
  measures[blockIndex] = patched.newMeasure;
  for (const p of pages) {
    for (const frag of layout!.pages[p].fragments) {
      if (frag.blockId === oldBlock.id && frag.pmEnd !== undefined) frag.pmEnd += delta;
    }
  }
  return repaintPages(pagesContainer!, pages, {
    blockId: oldBlock.id,
    block: patched.newBlock,
    measure: patched.newMeasure,
  });
}

/** Fast path for an edit inside a paragraph in a non-nested table cell (M2). */
function attemptTableCellFastEdit(
  tableStart: number,
  $from: ReturnType<EditorState['doc']['resolve']>,
  range: { from: number; to: number },
  delta: number,
  ctx: FastEditContext,
  defaultTabStopTwips: number | undefined
): boolean {
  const { blocks, measures, layout, pagesContainer } = ctx;

  // The innermost textblock must be a paragraph whose ancestry contains
  // exactly one table (no nested tables) and the whole change is inside it.
  if ($from.parent.type.name !== 'paragraph') return false;
  const parDepth = $from.depth;
  const parStart = $from.before(parDepth);
  const parEnd = $from.after(parDepth);
  if (range.from < parStart || range.to > parEnd) return false;
  let tableCount = 0;
  for (let d = 1; d <= parDepth; d++) {
    if ($from.node(d).type.name === 'table') tableCount++;
  }
  if (tableCount !== 1) return false;

  const blockIndex = blocks.findIndex(
    (b) => b.kind === 'table' && (b as TableBlock).pmStart === tableStart
  );
  if (blockIndex < 0) return false;
  const tableBlock = blocks[blockIndex] as TableBlock;
  const tableMeasure = measures[blockIndex];
  if (tableMeasure?.kind !== 'table') return false;

  // Locate the edited paragraph inside the table's nested structure.
  for (let r = 0; r < tableBlock.rows.length; r++) {
    const row = tableBlock.rows[r];
    for (let c = 0; c < row.cells.length; c++) {
      const cell = row.cells[c];
      for (let i = 0; i < cell.blocks.length; i++) {
        const b = cell.blocks[i];
        if (b.kind !== 'paragraph' || (b as ParagraphBlock).pmStart !== parStart) continue;

        const oldPara = b as ParagraphBlock;
        const cellMeasure = (tableMeasure as TableMeasure).rows[r]?.cells?.[c];
        const oldParaMeasure = cellMeasure?.blocks?.[i];
        if (!cellMeasure || oldParaMeasure?.kind !== 'paragraph') return false;

        const padLeft = cell.padding?.left ?? DEFAULT_CELL_PADDING_X;
        const padRight = cell.padding?.right ?? DEFAULT_CELL_PADDING_X;
        const cellContentWidth = Math.max(1, cellMeasure.width - padLeft - padRight);

        const patched = remeasureEligibleParagraph(
          $from.parent,
          parStart,
          cellContentWidth,
          oldPara,
          ctx,
          defaultTabStopTwips
        );
        // Table cells stay geometry-strict: a height change moves the row,
        // which moves the whole table — full-pipeline work.
        if (
          !patched ||
          !geometryUnchanged(patched.newMeasure, oldParaMeasure as ParagraphMeasure)
        ) {
          return false;
        }

        // Patch the nested structures in place. The painter's blockLookup
        // entry points at these same table objects, so no lookup patch is
        // needed — only the repaint.
        cell.blocks[i] = patched.newBlock;
        cellMeasure.blocks[i] = patched.newMeasure;
        if (tableBlock.pmEnd !== undefined) tableBlock.pmEnd += delta;

        const { pages } = findFragmentPages(layout!, tableBlock.id);
        if (pages.length === 0) return false;
        for (const p of pages) {
          for (const frag of layout!.pages[p].fragments) {
            if (frag.blockId === tableBlock.id && frag.pmEnd !== undefined) frag.pmEnd += delta;
          }
        }
        return repaintPages(pagesContainer!, pages);
      }
    }
  }
  return false;
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
  const { blocks, layout, pagesContainer } = ctx;
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

  const delta = tr.doc.content.size - tr.before.content.size;
  const defaultTabStopTwips = doc.attrs?.defaultTabStopTwips as number | undefined;

  if (node.type.name === 'paragraph') {
    // Top-level paragraph: the change must be inside it (already guaranteed
    // by the nodeEnd check above).
    return attemptParagraphFastEdit(node, nodeStart, delta, ctx, defaultTabStopTwips);
  }
  if (node.type.name === 'table') {
    return attemptTableCellFastEdit(nodeStart, $from, range, delta, ctx, defaultTabStopTwips);
  }
  return false;
}
