/**
 * React wiring for the typing fast path (docs/INCREMENTAL-LAYOUT.md M1).
 *
 * Owns the live context refs (layout/blocks/measures are read at call time,
 * not closed over) and the settle timer: after a successful fast-path
 * keystroke the full layout pipeline runs at typing-idle to reconcile
 * pagination, positions, and overlays. `tryFastEdit` returns true when the
 * edit was absorbed (patch + repaint done, settle scheduled); the caller
 * runs the full pipeline itself on false — `cancelSettle` is invoked
 * internally in that case so a stale settle never races a live full pass.
 */

import { useCallback, useEffect, useRef } from 'react';
import type { EditorState, Transaction } from 'prosemirror-state';
import type { FlowBlock, Layout, Measure } from '@valent/docx-editor-core/layout-engine';
import type { Theme } from '@valent/docx-editor-core/types';
import { attemptFastEdit } from '../internals/fastEditPath';

/**
 * Idle window after the last fast-path keystroke before the full layout
 * pipeline reconciles. Long enough to stay out of a typing burst, short
 * enough that page counts and downstream pages never feel stale.
 */
const FAST_EDIT_SETTLE_MS = 250;

export function useFastEditPath(inputs: {
  layout: Layout | null;
  blocks: FlowBlock[];
  measures: Measure[];
  theme: Theme | null | undefined;
  pagesContainerRef: React.RefObject<HTMLElement | null>;
  scheduleLayout: (state: EditorState) => void;
}): { tryFastEdit: (tr: Transaction, newState: EditorState) => boolean } {
  const { layout, blocks, measures, theme, pagesContainerRef, scheduleLayout } = inputs;

  const ctxRef = useRef({ layout, blocks, measures, theme });
  ctxRef.current = { layout, blocks, measures, theme };

  const settleTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const cancelSettle = useCallback(() => {
    if (settleTimerRef.current != null) {
      clearTimeout(settleTimerRef.current);
      settleTimerRef.current = null;
    }
  }, []);
  const scheduleSettle = useCallback(
    (state: EditorState) => {
      cancelSettle();
      settleTimerRef.current = setTimeout(() => {
        settleTimerRef.current = null;
        scheduleLayout(state);
      }, FAST_EDIT_SETTLE_MS);
    },
    [cancelSettle, scheduleLayout]
  );
  useEffect(() => cancelSettle, [cancelSettle]);

  const tryFastEdit = useCallback(
    (tr: Transaction, newState: EditorState): boolean => {
      const container = pagesContainerRef.current;
      const ctx = ctxRef.current;
      const ok =
        !!container &&
        attemptFastEdit(tr, newState, {
          blocks: ctx.blocks,
          measures: ctx.measures,
          layout: ctx.layout,
          pagesContainer: container,
          theme: ctx.theme,
        });
      if (ok) {
        scheduleSettle(newState);
      } else {
        cancelSettle();
      }
      return ok;
    },
    [pagesContainerRef, scheduleSettle, cancelSettle]
  );

  return { tryFastEdit };
}
