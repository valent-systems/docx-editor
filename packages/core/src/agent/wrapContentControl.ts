/**
 * Create (plant) inline content controls around occurrence-precise text spans.
 *
 * The read/fill side of `agent/contentControls` can *find and edit* controls
 * that already exist; this module is the missing *create* side. It wraps a
 * located placeholder span in a new inline content control (`w:sdt`) carrying a
 * caller-supplied stable `tag` — the durable identifier that later resolves the
 * control for filling.
 *
 * Why inline controls: CLM placeholders are short text spans that occur
 * mid-sentence and (overwhelmingly) inside table cells. An inline `w:sdt` lives
 * in a paragraph's run stream, so it wraps a mid-paragraph span and works
 * unchanged inside a table cell (a cell holds paragraphs, paragraphs hold inline
 * content) — no table-cell model change required. A freshly wrapped control has
 * no `rawPropertiesXml`, so the serializer synthesizes its `w:sdtPr` from the
 * modeled `tag`/`id`/`sdtType` (see {@link synthesizeSdtPr}); it round-trips
 * through `.docx` and Microsoft Word.
 *
 * Occurrence-precision: when the same placeholder text occurs more than once,
 * each occurrence is addressed by its 0-based index in document reading order
 * (optionally scoped to a single paragraph by `paraId`). Each occurrence is
 * wrapped as its own control with its own `tag`, so two identical clauses
 * (e.g. governing-law vs jurisdiction) receive independent values and are never
 * confused. The wrap is a pure transform — the source `doc` is not mutated.
 */

import type {
  Document,
  DocumentBody,
  BlockContent,
  Paragraph,
  ParagraphContent,
  Run,
  RunContent,
  SdtProperties,
  SdtType,
} from '../types/document';
import { getParagraphText, getRunText, getHyperlinkText, getInlineSdtText } from './text-utils';

/** Locates which occurrence of a placeholder string to wrap. */
export interface WrapLocator {
  /** Exact placeholder literal to wrap, verbatim. */
  text: string;
  /**
   * 0-based index among identical matches within the scope. Defaults to `0`.
   * Scope is the `paraId` paragraph if set, else the whole body in reading
   * order.
   */
  occurrence?: number;
  /**
   * Optional `w14:paraId` of the host paragraph. When set, the scope (and the
   * `occurrence` index) is restricted to that paragraph; identical text in
   * other paragraphs is ignored.
   */
  paraId?: string;
}

/** Properties for the control to plant. `tag` is the durable identifier. */
export interface WrapProps {
  /** Stable unique identifier (`w:tag`) — the durable fill-time anchor. */
  tag: string;
  /** Friendly name (`w:alias`). */
  alias?: string;
  /** Numeric id (`w:id`). Emitted for Word-friendliness; addressing uses `tag`. */
  id?: number;
  /** Control type. Only `richText`/`plainText` are free-text fillable. */
  sdtType?: Extract<SdtType, 'richText' | 'plainText'>;
}

/**
 * Outcome of a wrap attempt. `wrapped` carries the new document and the tag
 * assigned; the other variants describe why nothing was planted — surfaced to
 * an instrument UI rather than thrown, since "not found" is an expected result
 * when a template drifts.
 */
export type WrapResult =
  | { status: 'wrapped'; doc: Document; tag: string }
  | { status: 'not-found' }
  | { status: 'occurrence-out-of-range'; matches: number }
  | { status: 'crosses-inline-boundary' };

/** A located match: the paragraph that holds it plus its char span in that paragraph. */
interface ParaMatch {
  para: Paragraph;
  /** Inclusive start / exclusive end char offset within the paragraph's text. */
  start: number;
  end: number;
}

/** Find every char-offset range of `needle` in `haystack` (non-overlapping). */
function findRanges(haystack: string, needle: string): { start: number; end: number }[] {
  if (!needle) return [];
  const out: { start: number; end: number }[] = [];
  let from = 0;
  for (;;) {
    const idx = haystack.indexOf(needle, from);
    if (idx === -1) break;
    out.push({ start: idx, end: idx + needle.length });
    from = idx + needle.length;
  }
  return out;
}

/**
 * Logical text length an inline item contributes to its paragraph's text —
 * matching {@link getParagraphText} (runs, hyperlinks, and inline controls
 * carry text; bookmarks, fields, and tracked-change markers do not). Offsets
 * are computed over this logical text so occurrence indices stay stable even
 * after earlier spans have been wrapped into inline controls.
 */
function itemTextLength(item: ParagraphContent): number {
  if (item.type === 'run') return getRunText(item).length;
  if (item.type === 'hyperlink') return getHyperlinkText(item).length;
  if (item.type === 'inlineSdt') return getInlineSdtText(item).length;
  return 0;
}

/**
 * Collect matches of `locator.text` across the body in reading order,
 * descending into tables (cells) and block content controls, restricted to the
 * `paraId` scope if set.
 */
function collectMatches(body: DocumentBody, locator: WrapLocator): ParaMatch[] {
  const matches: ParaMatch[] = [];
  const visitParagraph = (para: Paragraph): void => {
    if (locator.paraId !== undefined && para.paraId !== locator.paraId) return;
    const text = getParagraphText(para);
    for (const r of findRanges(text, locator.text)) {
      matches.push({ para, start: r.start, end: r.end });
    }
  };
  const visitBlocks = (blocks: BlockContent[]): void => {
    for (const block of blocks) {
      if (block.type === 'paragraph') {
        visitParagraph(block);
      } else if (block.type === 'table') {
        for (const row of block.rows) {
          for (const cell of row.cells) visitBlocks(cell.content);
        }
      } else if (block.type === 'blockSdt') {
        visitBlocks(block.content);
      }
    }
  };
  visitBlocks(body.content);
  return matches;
}

/**
 * Split a run's content at text offset `k` (over the run's plain text).
 * Zero-length content (tabs/breaks/drawings) at the boundary stays with the
 * left piece if it precedes the split, the right piece otherwise.
 */
function splitRunContentAt(content: RunContent[], k: number): [RunContent[], RunContent[]] {
  const left: RunContent[] = [];
  const right: RunContent[] = [];
  let acc = 0;
  for (const item of content) {
    if (item.type === 'text') {
      const len = item.text.length;
      if (acc + len <= k) {
        left.push(item);
      } else if (acc >= k) {
        right.push(item);
      } else {
        const cut = k - acc;
        left.push({ ...item, text: item.text.slice(0, cut) });
        right.push({ ...item, text: item.text.slice(cut) });
      }
      acc += len;
    } else {
      // Zero text length: assign by position relative to the split.
      (acc < k ? left : right).push(item);
    }
  }
  return [left, right];
}

/** A run carrying just `content`, preserving formatting and property changes. */
function runWith(run: Run, content: RunContent[]): Run {
  const next: Run = { type: 'run', content };
  if (run.formatting) next.formatting = run.formatting;
  if (run.propertyChanges) next.propertyChanges = run.propertyChanges;
  return next;
}

function inlineSdtProps(props: WrapProps): SdtProperties {
  return {
    sdtType: props.sdtType ?? 'richText',
    tag: props.tag,
    ...(props.alias !== undefined ? { alias: props.alias } : {}),
    ...(props.id !== undefined ? { id: props.id } : {}),
    // No rawPropertiesXml → the serializer synthesizes a fresh w:sdtPr.
  };
}

/**
 * Rebuild `para` with the run text in `[start, end)` wrapped in a new inline
 * content control. Returns `null` if the span crosses inline content the
 * wrapper does not split — a hyperlink, field, or existing inline control
 * between the first and last covered run — which the caller maps to
 * `crosses-inline-boundary`. (Runs spanning the boundary are split; whole runs
 * inside the span move into the control.)
 */
function wrapSpanInParagraph(
  para: Paragraph,
  start: number,
  end: number,
  props: WrapProps
): Paragraph | null {
  // Logical char span of each inline item (over getParagraphText coordinates).
  const itemStart: number[] = [];
  const itemEnd: number[] = [];
  let cursor = 0;
  para.content.forEach((item, idx) => {
    itemStart[idx] = cursor;
    cursor += itemTextLength(item);
    itemEnd[idx] = cursor;
  });

  // The runs in which the span starts and ends (both endpoints must land in a
  // run, and everything between must be a run — else the span crosses inline
  // content the wrapper cannot split, and the caller surfaces a refusal).
  let firstIdx = -1;
  let lastIdx = -1;
  for (let i = 0; i < para.content.length; i++) {
    if (itemEnd[i] === itemStart[i]) continue; // zero-length: not a span endpoint
    if (start >= itemStart[i] && start < itemEnd[i]) firstIdx = i;
    if (end - 1 >= itemStart[i] && end - 1 < itemEnd[i]) lastIdx = i;
  }
  if (firstIdx === -1 || lastIdx === -1) return null;
  for (let i = firstIdx; i <= lastIdx; i++) {
    if (para.content[i].type !== 'run') return null;
  }

  const out: ParagraphContent[] = [];
  const inner: Run[] = [];
  for (let idx = 0; idx < para.content.length; idx++) {
    const item = para.content[idx];
    if (idx < firstIdx || idx > lastIdx) {
      out.push(item);
      continue;
    }
    const run = item as Run; // validated above
    const localStart = Math.max(start, itemStart[idx]) - itemStart[idx];
    const localEnd = Math.min(end, itemEnd[idx]) - itemStart[idx];
    const [beforeContent, restContent] = splitRunContentAt(run.content, localStart);
    const [insideContent, afterContent] = splitRunContentAt(restContent, localEnd - localStart);

    if (idx === firstIdx && beforeContent.length > 0) out.push(runWith(run, beforeContent));
    if (insideContent.length > 0) inner.push(runWith(run, insideContent));
    if (idx === lastIdx) {
      if (inner.length === 0) return null; // nothing captured
      out.push({ type: 'inlineSdt', properties: inlineSdtProps(props), content: inner });
      if (afterContent.length > 0) out.push(runWith(run, afterContent));
    }
  }

  return { ...para, content: out };
}

/**
 * Rebuild a block-content tree, replacing only the single target paragraph (by
 * identity) with its wrapped version. Generic over the array element type so it
 * preserves a table cell's narrower `(Paragraph | Table)[]` content type while
 * also serving the body's `BlockContent[]`. Wrapping a paragraph keeps it a
 * paragraph, so the element type never widens.
 */
function rebuildWithWrappedParagraph<T extends BlockContent>(
  blocks: T[],
  target: Paragraph,
  wrappedPara: Paragraph
): T[] {
  return blocks.map((block): T => {
    if ((block as BlockContent) === target) return wrappedPara as unknown as T;
    if (block.type === 'table') {
      return {
        ...block,
        rows: block.rows.map((row) => ({
          ...row,
          cells: row.cells.map((cell) => ({
            ...cell,
            content: rebuildWithWrappedParagraph(cell.content, target, wrappedPara),
          })),
        })),
      } as T;
    }
    if (block.type === 'blockSdt') {
      return {
        ...block,
        content: rebuildWithWrappedParagraph(block.content, target, wrappedPara),
      } as T;
    }
    return block;
  });
}

/**
 * Wrap an occurrence-precise placeholder span in a new inline content control
 * carrying `props.tag`. Pure: returns a new {@link Document} on success and
 * leaves `doc` untouched. The match is located in body reading order
 * (descending into table cells and block controls), scoped to `locator.paraId`
 * when given.
 *
 * Returns a {@link WrapResult}: `wrapped` with the new doc, or `not-found` /
 * `occurrence-out-of-range` / `crosses-inline-boundary` describing why nothing
 * was planted.
 */
export function wrapInlineContentControl(
  doc: Document,
  locator: WrapLocator,
  props: WrapProps
): WrapResult {
  const body = doc.package.document;
  const matches = collectMatches(body, locator);
  if (matches.length === 0) return { status: 'not-found' };

  const occurrence = locator.occurrence ?? 0;
  if (occurrence < 0 || occurrence >= matches.length) {
    return { status: 'occurrence-out-of-range', matches: matches.length };
  }

  const target = matches[occurrence];
  const wrappedPara = wrapSpanInParagraph(target.para, target.start, target.end, props);
  if (!wrappedPara) return { status: 'crosses-inline-boundary' };

  const content = rebuildWithWrappedParagraph(body.content, target.para, wrappedPara);
  const next: Document = {
    ...doc,
    package: { ...doc.package, document: { ...body, content } },
  };
  return { status: 'wrapped', doc: next, tag: props.tag };
}
