/**
 * Shared table measurement helper.
 *
 * Both React's PagedEditor and Vue's useDocxEditor measure tables the
 * same way: resolve column widths through `tableWidthUtils`, track
 * column occupancy across vertically merged cells, then call back into
 * the adapter's `measureBlock` for each cell's contained blocks.
 *
 * This module lives in core (not in either adapter) so a fix on one
 * side automatically applies to both. The recursive cell-content
 * measurement is delegated to the caller via `measureBlock` because
 * each adapter's block coverage differs (React supports floating
 * zones, footnotes, textBoxes; Vue is a subset).
 */

import type {
  FlowBlock,
  Measure,
  TableBlock,
  TableCell,
  TableMeasure,
} from '../layout-engine/types';
import {
  countTableColumns,
  normalizeTableColumnWidths,
  resolveTableWidthPx,
} from './tableWidthUtils';

/** Word's TableNormal default — 108 twips ≈ 7px. */
const DEFAULT_CELL_PADDING_X = 7;
/** OOXML/TableNormal default for top/bottom cell padding. */
const DEFAULT_CELL_PADDING_Y = 0;

/**
 * Visual height of a single block inside a table cell.
 *
 * A one-line paragraph that contains only image runs is laid out at the
 * image's intrinsic height (plus the paragraph's explicit spacing), not
 * a full text line — matching Word's per-cell layout. Everything else
 * uses the measured `totalHeight` / `height`.
 */
export function measureTableCellBlockVisualHeight(block: FlowBlock, blockMeasure: Measure): number {
  if (block.kind !== 'paragraph' || blockMeasure.kind !== 'paragraph') {
    if ('totalHeight' in blockMeasure) return blockMeasure.totalHeight;
    if ('height' in blockMeasure) return blockMeasure.height;
    return 0;
  }

  const nonEmptyRuns = block.runs.filter((run) => run.kind !== 'text' || run.text.length > 0);
  const imageOnlySingleLine =
    blockMeasure.lines.length === 1 &&
    nonEmptyRuns.length > 0 &&
    nonEmptyRuns.every((run) => run.kind === 'image');

  if (!imageOnlySingleLine) {
    return blockMeasure.totalHeight;
  }

  const maxImageHeight = nonEmptyRuns.reduce(
    (h, run) => (run.kind === 'image' ? Math.max(h, run.height) : h),
    0
  );
  const spacingBefore = block.attrs?.spacing?.before ?? 0;
  const spacingAfter = block.attrs?.spacing?.after ?? 0;
  return spacingBefore + maxImageHeight + spacingAfter;
}

/** Combined top + bottom border width of a cell, in pixels. */
function getTableCellVerticalBorderHeight(cell: TableCell | undefined): number {
  const top = cell?.borders?.top?.width ?? 0;
  const bottom = cell?.borders?.bottom?.width ?? 0;
  return top + bottom;
}

/**
 * Measure a `TableBlock` against a content-width budget.
 *
 * `measureBlock` is the per-cell-content measurement callback the
 * adapter uses for everything inside a cell. The adapter passes its
 * own `measureBlock` so block coverage stays per-renderer.
 */
export function measureTableBlock(
  tableBlock: TableBlock,
  contentWidth: number,
  measureBlock: (block: FlowBlock, contentWidth: number) => Measure
): TableMeasure {
  let columnWidths = tableBlock.columnWidths ?? [];
  const explicitWidthPx = resolveTableWidthPx(tableBlock.width, tableBlock.widthType, contentWidth);
  const colCount = countTableColumns(tableBlock);
  const targetWidth = explicitWidthPx ?? contentWidth;

  if (tableBlock.rows.length > 0) {
    columnWidths = normalizeTableColumnWidths(columnWidths, colCount, targetWidth);
  }

  if (columnWidths.length > 0 && explicitWidthPx) {
    const totalWidth = columnWidths.reduce((sum, w) => sum + w, 0);
    if (totalWidth > 0 && Math.abs(totalWidth - explicitWidthPx) > 1) {
      const scale = explicitWidthPx / totalWidth;
      columnWidths = columnWidths.map((w) => w * scale);
    }
  }

  // Track columns occupied by spanning cells from previous rows so vertically
  // merged cells in later rows don't shift their column-width assignment.
  const occupiedColumnsPerRow = new Map<number, Set<number>>();
  for (let rowIdx = 0; rowIdx < tableBlock.rows.length; rowIdx++) {
    const row = tableBlock.rows[rowIdx];
    if (!row) continue;
    let colIdx = 0;
    const occupied = occupiedColumnsPerRow.get(rowIdx) ?? new Set<number>();
    while (occupied.has(colIdx)) colIdx++;

    for (const cell of row.cells) {
      const colSpan = cell.colSpan ?? 1;
      const rowSpan = cell.rowSpan ?? 1;
      if (rowSpan > 1) {
        for (let r = rowIdx + 1; r < rowIdx + rowSpan; r++) {
          if (!occupiedColumnsPerRow.has(r)) occupiedColumnsPerRow.set(r, new Set());
          const occSet = occupiedColumnsPerRow.get(r)!;
          for (let c = 0; c < colSpan; c++) occSet.add(colIdx + c);
        }
      }
      colIdx += colSpan;
      while (occupied.has(colIdx)) colIdx++;
    }
  }

  const rows = tableBlock.rows.map((row, rowIdx) => {
    let columnIndex = 0;
    const occupied = occupiedColumnsPerRow.get(rowIdx) ?? new Set<number>();
    while (occupied.has(columnIndex)) columnIndex++;

    return {
      cells: row.cells.map((cell) => {
        const colSpan = cell.colSpan ?? 1;
        let cellWidth = 0;
        for (let c = 0; c < colSpan && columnIndex + c < columnWidths.length; c++) {
          cellWidth += columnWidths[columnIndex + c] ?? 0;
        }
        if (cellWidth === 0) {
          cellWidth =
            (cell.width && cell.width > 0
              ? cell.width
              : resolveTableWidthPx(cell.widthValue, cell.widthType, targetWidth)) ?? 100;
        }
        columnIndex += colSpan;
        while (occupied.has(columnIndex)) columnIndex++;

        const padLeft = cell.padding?.left ?? DEFAULT_CELL_PADDING_X;
        const padRight = cell.padding?.right ?? DEFAULT_CELL_PADDING_X;
        const cellContentWidth = Math.max(1, cellWidth - padLeft - padRight);

        return {
          blocks: cell.blocks.map((b) => measureBlock(b, cellContentWidth)),
          width: cellWidth,
          height: 0,
          colSpan: cell.colSpan,
          rowSpan: cell.rowSpan,
        };
      }),
      height: 0,
    };
  });

  for (let rowIdx = 0; rowIdx < rows.length; rowIdx++) {
    const row = rows[rowIdx];
    const sourceRowCells = tableBlock.rows[rowIdx]?.cells;
    let maxHeight = 0;
    let maxVerticalBorderHeight = 0;
    for (let cellIdx = 0; cellIdx < row.cells.length; cellIdx++) {
      const cell = row.cells[cellIdx];
      const sourceCell = sourceRowCells?.[cellIdx];
      // `paragraphMeasure.totalHeight` already includes spacing.before /
      // spacing.after; just sum the block heights. Adjacent-paragraph
      // collapse rules don't apply across the cell-content boundary, so this
      // matches Word's per-cell layout.
      let contentHeight = 0;
      for (let blockIdx = 0; blockIdx < cell.blocks.length; blockIdx++) {
        const sourceBlock = sourceCell?.blocks[blockIdx];
        const blockMeasure = cell.blocks[blockIdx];
        if (!sourceBlock || !blockMeasure) continue;
        contentHeight += measureTableCellBlockVisualHeight(sourceBlock, blockMeasure);
      }
      cell.height = contentHeight;
      const padTop = sourceCell?.padding?.top ?? DEFAULT_CELL_PADDING_Y;
      const padBottom = sourceCell?.padding?.bottom ?? DEFAULT_CELL_PADDING_Y;
      cell.height += padTop + padBottom;
      maxHeight = Math.max(maxHeight, cell.height);
      maxVerticalBorderHeight = Math.max(
        maxVerticalBorderHeight,
        getTableCellVerticalBorderHeight(sourceCell)
      );
    }

    const sourceRow = tableBlock.rows[rowIdx];
    const explicitHeight = sourceRow?.height;
    const heightRule = sourceRow?.heightRule;
    if (explicitHeight && heightRule === 'exact') {
      row.height = explicitHeight;
    } else if (explicitHeight) {
      // ECMA-376 §17.4.81: when hRule is absent or "auto", val is the minimum row height.
      row.height = Math.max(maxHeight + maxVerticalBorderHeight, explicitHeight);
    } else {
      row.height = maxHeight + maxVerticalBorderHeight;
    }
  }

  const totalHeight = rows.reduce((h, r) => h + r.height, 0);
  const totalWidth = columnWidths.reduce((w, cw) => w + cw, 0) || explicitWidthPx || contentWidth;

  return { kind: 'table', rows, columnWidths, totalWidth, totalHeight } as TableMeasure;
}
