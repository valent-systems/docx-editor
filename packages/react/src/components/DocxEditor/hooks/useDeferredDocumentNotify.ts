/**
 * Deferred document-change notification.
 *
 * `getDocument()` runs a full PM → model conversion (fromProseDoc over the
 * whole document) and the notification fans out to history.push + onChange +
 * the bridge — 30-60ms of main-thread work per keystroke on large documents,
 * none of which is needed mid-typing-burst. Coalesce it to a short idle
 * window. Undo/redo are unaffected (they run on the PM history plugin, not
 * the notified model) and explicit save reads `getDocument()` directly;
 * anything else that reads the notified document synchronously must call
 * `flush()` first. Flushes on unmount so the last edits aren't dropped.
 */

import { useCallback, useEffect, useRef } from 'react';
import type { Document } from '@valent/docx-editor-core/types';

/** Idle window before a docChanged transaction is converted and fanned out. */
const DOC_NOTIFY_IDLE_MS = 300;

export function useDeferredDocumentNotify(
  getDocument: () => Document | null | undefined,
  notify: (doc: Document) => void
): { schedule: () => void; flush: () => void } {
  const pendingRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const getDocumentRef = useRef(getDocument);
  getDocumentRef.current = getDocument;
  const notifyRef = useRef(notify);
  notifyRef.current = notify;

  const flush = useCallback(() => {
    if (pendingRef.current == null) return;
    clearTimeout(pendingRef.current);
    pendingRef.current = null;
    const doc = getDocumentRef.current();
    if (doc) notifyRef.current(doc);
  }, []);

  const schedule = useCallback(() => {
    if (pendingRef.current != null) {
      clearTimeout(pendingRef.current);
    }
    pendingRef.current = setTimeout(() => {
      pendingRef.current = null;
      const doc = getDocumentRef.current();
      if (doc) notifyRef.current(doc);
    }, DOC_NOTIFY_IDLE_MS);
  }, []);

  useEffect(() => flush, [flush]);

  return { schedule, flush };
}
