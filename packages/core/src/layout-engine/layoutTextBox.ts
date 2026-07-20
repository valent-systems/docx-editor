/**
 * Text-box block layout: floating boxes attach to the current page at the
 * anchor's flow position (the painter resolves the final anchored offset);
 * inline boxes participate in the normal flow via addFragment.
 */

import type { TextBoxBlock, TextBoxMeasure, TextBoxFragment } from './types';
import type { Paginator } from './paginator';
import { isFloatingTextBoxBlock } from './textBoxFlow';
import {
  pageGeometryFromPage,
  resolveAnchoredObjectVerticalTop,
} from '../layout-painter/anchoredObjectPosition';

export function layoutTextBox(
  block: TextBoxBlock,
  measure: TextBoxMeasure,
  paginator: Paginator
): void {
  if (measure.kind !== 'textBox') {
    throw new Error(`layoutTextBox: expected textBox measure`);
  }

  if (isFloatingTextBoxBlock(block)) {
    let state = paginator.getCurrentState();
    // Word keeps a paragraph-anchored floating box on its anchor's page: when
    // the RESOLVED box would overflow past the page content bottom and a
    // fresh page could hold it, the anchor advances to the next page (TPX
    // 'Teams-certified Devices': a 733px box anchored near a page bottom
    // starts the next page in Word — we painted it bleeding over the
    // footer). Page/margin-anchored boxes are absolute and never pushed.
    // Resolution uses the painter's own resolver (dual-renderer rule).
    const relV = block.position?.vertical?.relativeTo;
    if (relV === undefined || relV === 'paragraph' || relV === 'line') {
      const geometry = pageGeometryFromPage(state.page);
      const resolveInput = {
        width: measure.width,
        height: measure.height,
        position: block.position,
        cssFloat: block.cssFloat,
      };
      const anchorContentY = state.cursorY - state.topMargin;
      const top = resolveAnchoredObjectVerticalTop(resolveInput, anchorContentY, geometry);
      const freshTop = resolveAnchoredObjectVerticalTop(resolveInput, 0, geometry);
      if (
        top + measure.height > geometry.contentHeight &&
        freshTop + measure.height <= geometry.contentHeight &&
        state.page.fragments.length > 0
      ) {
        state = paginator.forcePageBreak();
      }
    }
    const fragment: TextBoxFragment = {
      kind: 'textBox',
      blockId: block.id,
      x: paginator.getColumnX(state.columnIndex),
      y: state.cursorY,
      width: measure.width,
      height: measure.height,
      pmStart: block.pmStart,
      pmEnd: block.pmEnd,
      isFloating: true,
      zIndex: block.wrapType === 'behind' ? -1 : 1,
    };
    state.page.fragments.push(fragment);
    return;
  }

  const state = paginator.ensureFits(measure.height);

  const fragment: TextBoxFragment = {
    kind: 'textBox',
    blockId: block.id,
    x: paginator.getColumnX(state.columnIndex),
    y: 0,
    width: measure.width,
    height: measure.height,
    pmStart: block.pmStart,
    pmEnd: block.pmEnd,
  };

  const result = paginator.addFragment(fragment, measure.height, 0, 0);
  fragment.y = result.y;
}
