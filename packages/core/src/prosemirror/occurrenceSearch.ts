/**
 * Occurrence-precise text search over a live ProseMirror document.
 *
 * The high-level `findInDocument`/`findTextInPmParagraph` helpers reject text
 * that occurs more than once in a paragraph and take no occurrence index, so
 * they cannot address a specific match when a placeholder repeats. This module
 * enumerates EVERY match of a literal in reading order with its true PM range,
 * and resolves a `(text, occurrence, paraId?)` locator to exactly one range —
 * the addressing the content-control wrap API needs to plant a marker over the
 * right occurrence (e.g. the two byte-identical CCEP country-list clauses).
 *
 * Correctness note (the inline-SDT gotcha): an inline content control (`sdt`)
 * is inline + NON-leaf — it contributes two PM positions (open/close) but zero
 * characters. A flat `pos + 1 + index` therefore drifts past it. We build a
 * per-character `index → PM position` table by walking inline content, emitting
 * one object-replacement char per inline leaf so string offsets stay 1:1 with
 * PM positions, and recursing through inline non-leaf nodes.
 */

import type { Node as PMNode } from 'prosemirror-model';

/** One concrete location of a literal in the live document. */
export interface OccurrenceMatch {
  /** PM position of the first matched char (inclusive). */
  from: number;
  /** PM position just past the last matched char (exclusive). */
  to: number;
  /** Host paragraph's `w14:paraId`, or null if it has none. */
  paraId: string | null;
  /** Character index of the match within the paragraph's plain text. */
  indexInPara: number;
  /** 0-based occurrence index within the host paragraph. */
  occurrenceInPara: number;
}

// One object-replacement char per inline leaf keeps string offsets 1:1 with PM
// positions across images/breaks (mirrors `node.textBetween(…, '￼')`).
const LEAF = '￼';

/**
 * Every match of `needle` in `doc`, in reading order, with its true PM range.
 * Descends into containers (tables, etc.) and through inline non-leaf nodes.
 */
export function enumerateMatches(doc: PMNode, needle: string): OccurrenceMatch[] {
  const out: OccurrenceMatch[] = [];
  if (!needle) return out;

  doc.descendants((node, pos) => {
    if (!node.isTextblock) return true; // descend into containers

    let paraText = '';
    const posMap: number[] = []; // posMap[k] = PM position of paraText[k]
    node.descendants((child, childPos) => {
      if (child.isText) {
        const base = pos + 1 + childPos;
        const t = child.text ?? '';
        for (let i = 0; i < t.length; i++) {
          paraText += t[i];
          posMap.push(base + i);
        }
        return false;
      }
      if (child.isLeaf) {
        paraText += LEAF;
        posMap.push(pos + 1 + childPos);
        return false;
      }
      return true; // inline non-leaf (e.g. sdt) — recurse into its children
    });

    const paraId = (node.attrs as { paraId?: string } | undefined)?.paraId ?? null;
    let search = 0;
    let occ = 0;
    for (;;) {
      const idx = paraText.indexOf(needle, search);
      if (idx < 0) break;
      out.push({
        from: posMap[idx],
        to: posMap[idx + needle.length - 1] + 1,
        paraId,
        indexInPara: idx,
        occurrenceInPara: occ,
      });
      search = idx + Math.max(1, needle.length);
      occ += 1;
    }
    return false; // inline content already walked
  });

  return out;
}

/** Locates which occurrence of a literal to address. */
export interface OccurrenceLocator {
  /** Exact literal to find, verbatim. */
  text: string;
  /** 0-based index within the resolution scope. Defaults to 0. */
  occurrence?: number;
  /**
   * Optional `w14:paraId`. When it matches a paragraph in the doc, the scope
   * (and the occurrence index) is restricted to that paragraph; otherwise the
   * scope is the whole document in reading order.
   */
  paraId?: string;
}

/**
 * Resolve a locator to exactly one live PM range, or `null` if the text is
 * absent or the occurrence index is out of range. Never silently retargets to
 * occurrence 0 — an out-of-range index returns `null` so the caller skips
 * rather than edits the wrong span.
 */
export function resolveOccurrence(doc: PMNode, locator: OccurrenceLocator): OccurrenceMatch | null {
  const all = enumerateMatches(doc, locator.text);
  if (all.length === 0) return null;

  let scope = all;
  if (locator.paraId) {
    const inPara = all.filter((m) => m.paraId === locator.paraId);
    if (inPara.length > 0) scope = inPara;
  }
  return scope[locator.occurrence ?? 0] ?? null;
}
