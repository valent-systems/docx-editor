/**
 * Table row-break geometry.
 *
 * Word lets a table row break across a page boundary ("allow row to break
 * across pages", on by default). When a row is taller than the space left on
 * the page, the portion that fits stays and the rest continues on the next
 * page — broken between whole text lines, never through a glyph.
 *
 * This module computes, per row, the set of safe break offsets (the y of every
 * line bottom across the row's content, including vertically-merged cells that
 * span into the row) so the paginator can snap a break to the deepest whole
 * line that still fits.
 */

import type { TableBlock, TableMeasure } from './types';
import { resolveCellGrid } from '../layout-bridge/tableWidthUtils';
import { layoutCellContent } from '../layout-bridge/cellBlockLayout';

/**
 * Precomputed break geometry for a table.
 */
export interface TableRowBreakInfo {
  /** Cumulative y of the top of each row; rowTops[rows.length] is the table height. */
  rowTops: number[];
  /**
   * Per-row sorted, de-duplicated line-bottom offsets (relative to the row top)
   * at which a break is clean. Always includes the row's full height as the
   * final boundary.
   */
  breakOffsets: number[][];
}

/**
 * Build break geometry for a table from its block + measure.
 */
export function buildTableRowBreakInfo(
  block: TableBlock,
  measure: TableMeasure
): TableRowBreakInfo {
  const rowCount = measure.rows.length;
  // True (unrounded) cumulative row offsets — the paginator splits against
  // exact measured heights. The painter has a sibling `buildRowYPositions`
  // that rounds to whole pixels for crisp borders; keep the two SEPARATE
  // (don't "dedupe") or you break either break-offset alignment or crispness.
  const rowTops: number[] = [];
  let acc = 0;
  for (let r = 0; r < rowCount; r++) {
    rowTops.push(acc);
    acc += measure.rows[r]?.height ?? 0;
  }
  rowTops.push(acc);

  // Use the shared grid resolution so "which cells cover row r" matches the
  // measurer and painter. A cell starting in row `sr` with rowSpan covers
  // rows [sr, sr + rowSpan); a merged cell spills its line bottoms into the
  // rows below its restart row.
  const resolved = resolveCellGrid(block);
  const breakOffsets: number[][] = [];
  for (let r = 0; r < rowCount; r++) {
    const rowHeight = measure.rows[r]?.height ?? 0;
    const offsets = new Set<number>();
    offsets.add(rowHeight); // a row boundary is always a clean break
    // Painted line spans of ALL covering cells: a break offset is only clean
    // when NO column has a line straddling it (Word never cuts through a
    // glyph — a line bottom in one column can land mid-line in another, e.g.
    // a 3-line cell next to a vertically-centered 1-line cell).
    const lineSpans: Array<[number, number]> = [];

    for (const g of resolved) {
      if (g.rowIndex > r || g.rowIndex + g.rowSpan - 1 < r) continue;
      const sourceCell = block.rows[g.rowIndex]?.cells?.[g.cellIndex];
      const measuredCell = measure.rows[g.rowIndex]?.cells?.[g.cellIndex];
      if (!sourceCell || !measuredCell) continue;
      // Break offsets must live in PAINTED space, so mirror renderTableCell's
      // defaults: painter cell padding defaults to 1px top/bottom.
      const padTop = sourceCell.padding?.top ?? 1;
      const padBottom = sourceCell.padding?.bottom ?? 1;
      const { flatBottoms, flatTops, contentHeight } = layoutCellContent(
        sourceCell.blocks,
        measuredCell.blocks,
        padTop
      );
      // Mirror renderTableCell's vertical alignment: when content doesn't fill
      // the cell box, vAlign center/bottom flex-shifts the painted content down
      // by the leftover slack. The line bottoms must shift identically or the
      // paginator clips straight through painted glyphs (DRC certifications
      // table: style-cascaded vAlign=center put lines ~6px below the measured
      // offsets, slicing the boundary row's text on both pages).
      const boxH = rowTops[g.rowIndex + g.rowSpan] - rowTops[g.rowIndex];
      const contentFillsBox = (measuredCell.height ?? 0) >= boxH - 0.5;
      let vShift = 0;
      if (
        !contentFillsBox &&
        (sourceCell.verticalAlign === 'center' || sourceCell.verticalAlign === 'bottom')
      ) {
        const slack = Math.max(0, boxH - padTop - padBottom - contentHeight);
        vShift = sourceCell.verticalAlign === 'center' ? slack / 2 : slack;
      }
      // Map cell-content y (relative to the cell/region top at rowTops[startRow])
      // into this row's coordinate space (relative to rowTops[r]).
      const shift = rowTops[r] - rowTops[g.rowIndex];
      for (let i = 0; i < flatBottoms.length; i++) {
        const off = flatBottoms[i] + vShift - shift;
        if (off > 0 && off < rowHeight) offsets.add(off);
        lineSpans.push([flatTops[i] + vShift - shift, flatBottoms[i] + vShift - shift]);
      }
    }
    // Drop offsets that would cut through another column's line. 0.5px of
    // tolerance absorbs float jitter between near-identical line grids.
    const EPS = 0.5;
    const clean = [...offsets].filter(
      (off) => off >= rowHeight || !lineSpans.some(([t, b]) => off > t + EPS && off < b - EPS)
    );
    breakOffsets.push(clean.sort((a, b) => a - b));
  }

  return { rowTops, breakOffsets };
}

/**
 * Given a row and how much of it has already been placed (`fromOffset`),
 * return how many more px can be placed ending on a whole line, without
 * exceeding `maxSlice`. Returns 0 when not even the first line fits.
 */
export function snapRowBreak(
  info: TableRowBreakInfo,
  rowIndex: number,
  fromOffset: number,
  maxSlice: number
): number {
  const offsets = info.breakOffsets[rowIndex];
  if (!offsets || offsets.length === 0) return 0;
  const limit = fromOffset + maxSlice;
  let best = 0;
  for (const off of offsets) {
    if (off <= fromOffset) continue;
    if (off <= limit) best = off - fromOffset;
    else break;
  }
  return best;
}
