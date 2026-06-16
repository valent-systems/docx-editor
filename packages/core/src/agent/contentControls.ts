/**
 * Content-control (SDT) addressing for the document model.
 *
 * Block-level content controls (`w:sdt`) are the natural anchor for template
 * logic and agent edits: they survive the round trip (see the parser +
 * serializer) and carry a stable `tag`/`alias`/`id`. This module is the
 * read side of that contract — discover controls and read their content
 * without a DOM or an editor instance, so server-side pipelines and AI
 * agents can find an anchor by tag and act on it.
 *
 * Walks the body recursively, descending into nested controls so a control
 * inside another control is still found. (The model places block content
 * controls at body level or nested in other controls, not inside table
 * cells.) The returned `path` (block indices from the body root) addresses
 * the control unambiguously for a follow-up edit.
 */

import type {
  Document,
  DocumentBody,
  BlockContent,
  BlockSdt,
  InlineSdt,
  ParagraphContent,
  Run,
  HeaderFooter,
  SdtType,
  SdtProperties,
  SdtDataBinding,
} from '../types/document';
import { getParagraphText, getRunText, getTableText } from './text-utils';

/** Filter for {@link findContentControls}. All provided fields must match (AND). */
export interface ContentControlFilter {
  /** Developer identifier (`w:tag`), exact match. */
  tag?: string;
  /** Friendly name (`w:alias`), exact match. */
  alias?: string;
  /** Numeric id (`w:id`), exact match. */
  id?: number;
  /** Control type projection (`richText`, `dropDownList`, …). */
  type?: SdtType;
}

/** A discovered content control plus enough context to address and edit it. */
export interface ContentControlInfo {
  /** Developer identifier (`w:tag`). */
  tag?: string;
  /** Friendly name (`w:alias`). */
  alias?: string;
  /** Numeric id (`w:id`). */
  id?: number;
  /** Control type projection. */
  sdtType: SdtType;
  /** Lock setting, if any. A locked control should refuse content edits. */
  lock?: SdtProperties['lock'];
  /** Dropdown/combobox list items, if modeled. */
  listItems?: { displayText: string; value: string }[];
  /** Placeholder docPart reference, if any. */
  placeholder?: string;
  /** Whether the control is currently showing placeholder text (`w:showingPlcHdr`). */
  showingPlaceholder?: boolean;
  /** Checkbox state, for checkbox controls. */
  checked?: boolean;
  /** Date format string, for date controls. */
  dateFormat?: string;
  /** XML data binding (`w:dataBinding`), if the control is bound. */
  dataBinding?: SdtDataBinding;
  /** Plain text of the control's content (paragraphs/tables/nested controls flattened). */
  text: string;
  /**
   * Whether this is a block-level (`w:sdt` wrapping paragraphs/tables) or an
   * inline (`w:sdt` mid-paragraph) content control. Inline controls are the
   * ones planted by {@link wrapInlineContentControl} around placeholder spans.
   */
  kind: 'block' | 'inline';
  /**
   * Which content tree the control lives in. `header`/`footer` controls live
   * outside the document body and are addressed by `tag`/`id`, not by `path`.
   */
  container: 'body' | 'header' | 'footer';
  /**
   * Index path from the container root to this control. Block controls are
   * addressed `[i]` (or `[i, j]` when nested in another control); a control
   * inside a table cell is `[blockIdx, rowIdx, cellIdx, …]`; the last element of
   * an inline control's path is its index within the host paragraph's inline
   * content. Informational — the durable address is `tag`/`id`.
   */
  path: number[];
  /** Path nesting depth (`path.length - 1`). */
  depth: number;
}

/** Narrow a {@link Document} or {@link DocumentBody} to its block list. */
function bodyOf(input: Document | DocumentBody): DocumentBody {
  return 'package' in input ? input.package.document : input;
}

/** Plain text of a block control's content, descending into tables and nested SDTs. */
export function getContentControlText(control: BlockSdt): string {
  return blocksText(control.content);
}

function blocksText(blocks: BlockContent[]): string {
  const parts: string[] = [];
  for (const block of blocks) {
    if (block.type === 'paragraph') parts.push(getParagraphText(block));
    else if (block.type === 'table') parts.push(getTableText(block));
    else if (block.type === 'blockSdt') parts.push(blocksText(block.content));
  }
  return parts.join('\n');
}

/** Plain text of an inline control's content (runs + nested inline controls). */
function inlineText(content: InlineSdt['content']): string {
  let text = '';
  for (const item of content) {
    if (item.type === 'run') text += getRunText(item);
    else if (item.type === 'inlineSdt') text += inlineText(item.content);
  }
  return text;
}

function matches(props: SdtProperties, filter: ContentControlFilter): boolean {
  if (filter.tag !== undefined && props.tag !== filter.tag) return false;
  if (filter.alias !== undefined && props.alias !== filter.alias) return false;
  if (filter.id !== undefined && props.id !== filter.id) return false;
  if (filter.type !== undefined && props.sdtType !== filter.type) return false;
  return true;
}

function infoOf(
  props: SdtProperties,
  text: string,
  kind: 'block' | 'inline',
  container: ContentControlInfo['container'],
  path: number[]
): ContentControlInfo {
  return {
    tag: props.tag,
    alias: props.alias,
    id: props.id,
    sdtType: props.sdtType,
    lock: props.lock,
    listItems: props.listItems,
    placeholder: props.placeholder,
    showingPlaceholder: props.showingPlaceholder,
    checked: props.checked,
    dateFormat: props.dateFormat,
    dataBinding: props.dataBinding,
    text,
    kind,
    container,
    path,
    depth: path.length - 1,
  };
}

/** Discover inline controls within a paragraph's inline content (recursing nested). */
function walkInline(
  content: ParagraphContent[],
  parentPath: number[],
  container: ContentControlInfo['container'],
  filter: ContentControlFilter,
  out: ContentControlInfo[]
): void {
  for (let k = 0; k < content.length; k++) {
    const item = content[k];
    if (item.type === 'inlineSdt') {
      const path = [...parentPath, k];
      if (matches(item.properties, filter)) {
        out.push(infoOf(item.properties, inlineText(item.content), 'inline', container, path));
      }
      walkInline(item.content as ParagraphContent[], path, container, filter, out); // nested
    }
  }
}

/**
 * Discover block + inline controls within a block-content tree, descending into
 * block controls, table cells, and paragraph inline content.
 */
function walkBlocks(
  blocks: BlockContent[],
  parentPath: number[],
  container: ContentControlInfo['container'],
  filter: ContentControlFilter,
  out: ContentControlInfo[]
): void {
  for (let i = 0; i < blocks.length; i++) {
    const block = blocks[i];
    const path = [...parentPath, i];
    if (block.type === 'blockSdt') {
      if (matches(block.properties, filter)) {
        out.push(infoOf(block.properties, getContentControlText(block), 'block', container, path));
      }
      walkBlocks(block.content, path, container, filter, out); // nested block controls
    } else if (block.type === 'paragraph') {
      walkInline(block.content, path, container, filter, out); // inline controls
    } else if (block.type === 'table') {
      for (let r = 0; r < block.rows.length; r++) {
        const row = block.rows[r];
        for (let c = 0; c < row.cells.length; c++) {
          walkBlocks(row.cells[c].content, [...path, r, c], container, filter, out);
        }
      }
    }
  }
}

/**
 * Find every content control in the document, optionally filtered by
 * tag/alias/id/type. Results are in document order; nested controls follow
 * their parent. Discovery reaches **block** controls (body, nested, header/
 * footer) and **inline** controls wherever they live — including inside table
 * cells and mid-paragraph (the spans planted by {@link wrapInlineContentControl}).
 * Each result carries `kind` (`block`/`inline`) and `container`
 * (`body`/`header`/`footer`). When given a {@link DocumentBody} only the body is
 * searched; pass the {@link Document} to also reach headers/footers.
 */
export function findContentControls(
  input: Document | DocumentBody,
  filter: ContentControlFilter = {}
): ContentControlInfo[] {
  const out: ContentControlInfo[] = [];
  walkBlocks(bodyOf(input).content, [], 'body', filter, out);
  if ('package' in input) {
    for (const hf of input.package.headers?.values() ?? []) {
      walkBlocks(hf.content, [], 'header', filter, out);
    }
    for (const hf of input.package.footers?.values() ?? []) {
      walkBlocks(hf.content, [], 'footer', filter, out);
    }
  }
  return out;
}

/** Convenience: the first control matching `filter`, or `undefined`. */
export function findContentControl(
  input: Document | DocumentBody,
  filter: ContentControlFilter
): ContentControlInfo | undefined {
  return findContentControls(input, filter)[0];
}

// ============================================================================
// MUTATION (edit a control by tag)
// ============================================================================

/** No control matched the filter. */
export class ContentControlNotFoundError extends Error {
  constructor(filter: ContentControlFilter) {
    super(`No content control matched ${JSON.stringify(filter)}`);
    this.name = 'ContentControlNotFoundError';
  }
}

/** The matched control's lock forbids the attempted edit (pass `force` to override). */
export class ContentControlLockedError extends Error {
  constructor(lock: SdtProperties['lock'], op: 'edit' | 'remove') {
    super(`Content control is ${lock}; cannot ${op} it without { force: true }`);
    this.name = 'ContentControlLockedError';
  }
}

/**
 * The control's type doesn't support free text/block replacement (e.g. a
 * dropdown, date, checkbox, or picture control), so writing arbitrary content
 * would desync the type marker from its value. Use a type-specific setter, or
 * pass `{ force: true }` to override.
 */
export class ContentControlTypeError extends Error {
  constructor(sdtType: SdtType) {
    super(
      `Content control is a '${sdtType}' control; replacing its content with free text ` +
        `would desync it. Use a type-specific value setter or pass { force: true }.`
    );
    this.name = 'ContentControlTypeError';
  }
}

/**
 * The control is bound to a Custom XML data store (`w:dataBinding`). Writing its
 * content won't stick — Word re-renders the control from the bound XML node — so
 * the write is refused. Update the data store instead, or pass `{ force: true }`.
 */
export class ContentControlBoundError extends Error {
  constructor() {
    super(
      'Content control is data-bound (w:dataBinding); its content is driven by the ' +
        'Custom XML store and a direct write will not persist. Update the store, or pass { force: true }.'
    );
    this.name = 'ContentControlBoundError';
  }
}

/**
 * Control types whose content is free-form and safe to replace with text/blocks.
 * Typed controls (dropdown, date, checkbox, picture) carry structured state that
 * arbitrary content would contradict, and `group` exists to lock/contain nested
 * structure — all gated unless forced.
 */
const TEXT_REPLACEABLE_TYPES = new Set<SdtType>(['richText', 'plainText', 'unknown']);

/** True if free text/block content can safely replace this control type's content. */
export function isTextReplaceable(sdtType: SdtType): boolean {
  return TEXT_REPLACEABLE_TYPES.has(sdtType);
}

/** `w:lock` values that forbid editing the control's content. */
export function isContentLocked(lock: SdtProperties['lock']): boolean {
  return lock === 'contentLocked' || lock === 'sdtContentLocked';
}

/** `w:lock` values that forbid deleting the control. */
export function isDeletionLocked(lock: SdtProperties['lock']): boolean {
  return lock === 'sdtLocked' || lock === 'sdtContentLocked';
}

/**
 * True if the raw `w:sdtPr` carries a (w15) repeating-section structure. Matches
 * the element name (`<w15:repeatingSection>` / `<w15:repeatingSectionItem>`) so
 * a tag/alias value that merely contains the word doesn't false-match.
 */
export function hasRepeatingSection(props: SdtProperties): boolean {
  return /<w15:repeatingSection(Item)?[\s/>]/.test(props.rawPropertiesXml ?? '');
}

/** True if the control is bound to a Custom XML data store (`w:dataBinding`). */
export function isDataBound(props: SdtProperties): boolean {
  return props.dataBinding != null;
}

/**
 * Strip `<w:showingPlcHdr/>` from a raw `w:sdtPr` string. When real content is
 * written into a control that was showing its placeholder, the flag must go or
 * Word keeps rendering the (now-stale) placeholder styling over real content.
 */
export function clearShowingPlaceholderXml(raw: string | undefined): string | undefined {
  if (!raw) return raw;
  return raw
    .replace(/<w:showingPlcHdr\b[^>]*\/>/g, '')
    .replace(/<w:showingPlcHdr\b[^>]*>[\s\S]*?<\/w:showingPlcHdr>/g, '');
}

/** Properties for a control after real content is written: placeholder flag cleared. */
function propsAfterContentWrite(props: SdtProperties): SdtProperties {
  if (!props.showingPlaceholder && !/showingPlcHdr/.test(props.rawPropertiesXml ?? '')) {
    return props;
  }
  const next: SdtProperties = { ...props, showingPlaceholder: false };
  const cleaned = clearShowingPlaceholderXml(props.rawPropertiesXml);
  if (cleaned !== undefined) next.rawPropertiesXml = cleaned;
  return next;
}

function paragraph(text: string): BlockContent {
  return {
    type: 'paragraph',
    content: text ? [{ type: 'run', content: [{ type: 'text', text }] }] : [],
  };
}

/**
 * Turn a string into paragraphs (one per newline), or deep-clone block input.
 * A `plainText` control is single-paragraph in OOXML, so its string content is
 * collapsed to one paragraph rather than split — multiple paragraphs would make
 * Word repair the control on open.
 */
function toBlocks(
  replacement: string | BlockContent[],
  opts: { singleParagraph?: boolean } = {}
): BlockContent[] {
  if (typeof replacement !== 'string') {
    // Clone so the caller can't later mutate content shared with the result.
    return structuredClone(replacement);
  }
  if (opts.singleParagraph) return [paragraph(replacement)];
  return replacement.split('\n').map(paragraph);
}

export type ControlOp = (control: BlockSdt) => BlockContent[];

/**
 * Rebuild `blocks`, applying `op` to the first control matching `filter`. The
 * op's result (0, 1, or many blocks) is spliced in at the control's own level
 * — including when the control is nested inside another control — so a
 * remove/unwrap never leaves a placeholder behind. `state.done` stops the
 * walk after the first match.
 */
export function applyToFirst(
  blocks: BlockContent[],
  filter: ContentControlFilter,
  op: ControlOp,
  state: { done: boolean }
): BlockContent[] {
  const out: BlockContent[] = [];
  for (const block of blocks) {
    if (state.done) {
      out.push(block);
      continue;
    }
    // Controls are searched at body level and inside other controls. Table
    // cells are not searched: the current model types a cell as
    // (Paragraph | Table)[], and the table parser does not yet surface a
    // cell-level w:sdt (which OOXML's CT_Tc does permit) — see CONTENT-CONTROLS.md.
    if (block.type === 'blockSdt') {
      if (matches(block.properties, filter)) {
        out.push(...op(block));
        state.done = true;
        continue;
      }
      out.push({ ...block, content: applyToFirst(block.content, filter, op, state) });
    } else {
      out.push(block);
    }
  }
  return out;
}

export function rebuild(doc: Document, content: BlockContent[]): Document {
  return {
    ...doc,
    package: {
      ...doc.package,
      document: { ...doc.package.document, content },
    },
  };
}

// ============================================================================
// UNIVERSAL EDIT (first matching control anywhere — block OR inline)
// ============================================================================
//
// `applyToFirst` (above) edits block controls in the body and is shared by the
// repeating-section / typed-value paths. The mutators below need to reach
// *inline* controls (the spans planted by `wrapInlineContentControl`) wherever
// they live — including inside table cells and headers/footers — so they use
// this broader traversal, applying `blockOp` to a matching block control and
// `inlineOp` to a matching inline control. The first match (document order)
// wins; untouched subtrees keep their identity (returned arrays are reused when
// nothing changed) so the transform stays pure and cheap.

type BlockControlOp = (control: BlockSdt) => BlockContent[];
type InlineControlOp = (control: InlineSdt) => InlineSdt['content'];

interface EditState {
  done: boolean;
}

function editInline(
  content: ParagraphContent[],
  filter: ContentControlFilter,
  inlineOp: InlineControlOp,
  state: EditState
): ParagraphContent[] {
  let changed = false;
  const out: ParagraphContent[] = [];
  for (const item of content) {
    if (state.done || item.type !== 'inlineSdt') {
      out.push(item);
      continue;
    }
    if (matches(item.properties, filter)) {
      out.push(...inlineOp(item));
      state.done = true;
      changed = true;
      continue;
    }
    const nested = editInline(item.content as ParagraphContent[], filter, inlineOp, state);
    if (nested !== (item.content as ParagraphContent[])) {
      out.push({ ...item, content: nested as InlineSdt['content'] });
      changed = true;
    } else {
      out.push(item);
    }
  }
  return changed ? out : content;
}

function editBlocks(
  blocks: BlockContent[],
  filter: ContentControlFilter,
  blockOp: BlockControlOp,
  inlineOp: InlineControlOp,
  state: EditState
): BlockContent[] {
  let changed = false;
  const out: BlockContent[] = [];
  for (const block of blocks) {
    if (state.done) {
      out.push(block);
      continue;
    }
    if (block.type === 'blockSdt') {
      if (matches(block.properties, filter)) {
        out.push(...blockOp(block));
        state.done = true;
        changed = true;
        continue;
      }
      const nested = editBlocks(block.content, filter, blockOp, inlineOp, state);
      if (nested !== block.content) {
        out.push({ ...block, content: nested });
        changed = true;
      } else {
        out.push(block);
      }
    } else if (block.type === 'paragraph') {
      const content = editInline(block.content, filter, inlineOp, state);
      if (content !== block.content) {
        out.push({ ...block, content });
        changed = true;
      } else {
        out.push(block);
      }
    } else if (block.type === 'table') {
      let tableChanged = false;
      const rows = block.rows.map((row) => {
        if (state.done) return row;
        let rowChanged = false;
        const cells = row.cells.map((cell) => {
          if (state.done) return cell;
          const cellContent = editBlocks(cell.content, filter, blockOp, inlineOp, state);
          if (cellContent !== cell.content) {
            rowChanged = true;
            // Fill/remove keeps a paragraph a paragraph and never adds a
            // blockSdt to a cell, so the cell's narrower content type holds.
            return { ...cell, content: cellContent as typeof cell.content };
          }
          return cell;
        });
        if (rowChanged) {
          tableChanged = true;
          return { ...row, cells };
        }
        return row;
      });
      if (tableChanged) {
        out.push({ ...block, rows });
        changed = true;
      } else {
        out.push(block);
      }
    } else {
      out.push(block);
    }
  }
  return changed ? out : blocks;
}

function editHeaderFooterMap(
  map: Map<string, HeaderFooter>,
  filter: ContentControlFilter,
  blockOp: BlockControlOp,
  inlineOp: InlineControlOp,
  state: EditState
): Map<string, HeaderFooter> {
  let changed = false;
  const next = new Map<string, HeaderFooter>();
  for (const [key, hf] of map) {
    if (state.done) {
      next.set(key, hf);
      continue;
    }
    const content = editBlocks(hf.content, filter, blockOp, inlineOp, state);
    if (content !== hf.content) {
      next.set(key, { ...hf, content });
      changed = true;
    } else {
      next.set(key, hf);
    }
  }
  return changed ? next : map;
}

/**
 * Apply the first matching control's op across the document body, table cells,
 * and headers/footers. Returns the new document and whether a match was found.
 * Pure — `doc` is not mutated.
 */
function applyToFirstControl(
  doc: Document,
  filter: ContentControlFilter,
  blockOp: BlockControlOp,
  inlineOp: InlineControlOp
): { doc: Document; found: boolean } {
  const state: EditState = { done: false };
  const content = editBlocks(doc.package.document.content, filter, blockOp, inlineOp, state);
  let headers = doc.package.headers;
  let footers = doc.package.footers;
  if (!state.done && headers) {
    headers = editHeaderFooterMap(headers, filter, blockOp, inlineOp, state);
  }
  if (!state.done && footers) {
    footers = editHeaderFooterMap(footers, filter, blockOp, inlineOp, state);
  }
  if (!state.done) return { doc, found: false };
  return {
    doc: {
      ...doc,
      package: {
        ...doc.package,
        document: { ...doc.package.document, content },
        headers,
        footers,
      },
    },
    found: true,
  };
}

/** A run carrying a single text node, used to fill an inline control. */
function inlineRun(text: string): Run {
  return {
    type: 'run',
    content: text ? [{ type: 'text', text, preserveSpace: true }] : [],
  };
}

/**
 * Replace the content of the first control matching `filter` — a **block** or
 * an **inline** control, wherever it lives (body, table cell, header/footer).
 * `replacement` may be a string (split into paragraphs on newlines for a block
 * control; collapsed to one run for an inline control) or block content (block
 * controls only — flattened to text for an inline control). The control's
 * properties, tag/alias, and lossless raw `w:sdtPr` are preserved — only the
 * contained content changes, so the result still round-trips.
 *
 * When the control was showing its placeholder (`w:showingPlcHdr`), that flag
 * is cleared so Word doesn't render the new content as placeholder text.
 *
 * Throws {@link ContentControlNotFoundError} if nothing matches,
 * {@link ContentControlLockedError} if the control's lock forbids editing, and
 * {@link ContentControlTypeError} if the control is a typed (dropdown/date/…)
 * control whose value shouldn't be set as free text. Pass `{ force: true }` to
 * override the lock/type guards.
 */
export function setContentControlContent(
  doc: Document,
  filter: ContentControlFilter,
  replacement: string | BlockContent[],
  options: { force?: boolean } = {}
): Document {
  const guard = (props: SdtProperties): void => {
    if (!options.force && isContentLocked(props.lock)) {
      throw new ContentControlLockedError(props.lock, 'edit');
    }
    if (!options.force && !isTextReplaceable(props.sdtType)) {
      throw new ContentControlTypeError(props.sdtType);
    }
    if (!options.force && isDataBound(props)) {
      throw new ContentControlBoundError();
    }
  };
  const blockOp: BlockControlOp = (control) => {
    guard(control.properties);
    return [
      {
        ...control,
        properties: propsAfterContentWrite(control.properties),
        content: toBlocks(replacement, {
          singleParagraph: control.properties.sdtType === 'plainText',
        }),
      },
    ];
  };
  const inlineOp: InlineControlOp = (control) => {
    guard(control.properties);
    const text = typeof replacement === 'string' ? replacement : blocksText(replacement);
    return [
      {
        ...control,
        properties: propsAfterContentWrite(control.properties),
        content: text ? [inlineRun(text)] : [],
      },
    ];
  };
  const { doc: next, found } = applyToFirstControl(doc, filter, blockOp, inlineOp);
  if (!found) throw new ContentControlNotFoundError(filter);
  return next;
}

/**
 * Remove the first control matching `filter` from the document. With
 * `keepContent: true` the control's blocks are unwrapped in place (the box
 * goes away, the content stays) — useful for "resolve this conditional
 * section into plain content". Otherwise the whole region is deleted.
 *
 * Unwrapping a repeating-section (item) is refused unless `force`, since
 * lifting its blocks out would orphan the (w15) repeating structure.
 *
 * Throws {@link ContentControlNotFoundError} / {@link ContentControlLockedError}
 * as {@link setContentControlContent} does.
 */
export function removeContentControl(
  doc: Document,
  filter: ContentControlFilter,
  options: { force?: boolean; keepContent?: boolean } = {}
): Document {
  const guardRemoval = (props: SdtProperties): void => {
    if (!options.force && isDeletionLocked(props.lock)) {
      throw new ContentControlLockedError(props.lock, 'remove');
    }
    if (options.keepContent && !options.force && hasRepeatingSection(props)) {
      throw new ContentControlLockedError(props.lock, 'remove');
    }
  };
  const blockOp: BlockControlOp = (control) => {
    guardRemoval(control.properties);
    return options.keepContent ? control.content : [];
  };
  const inlineOp: InlineControlOp = (control) => {
    guardRemoval(control.properties);
    return options.keepContent ? control.content : [];
  };
  const { doc: next, found } = applyToFirstControl(doc, filter, blockOp, inlineOp);
  if (!found) throw new ContentControlNotFoundError(filter);
  // Never leave a structurally empty body (matches the live-editor path, which
  // auto-fills a paragraph). An empty <w:body> is invalid for Word consumers.
  if (next.package.document.content.length === 0) {
    return rebuild(next, [{ type: 'paragraph', content: [] }]);
  }
  return next;
}

// ============================================================================
// RESULT-RETURNING FILL (success/blocked outcome, no throw)
// ============================================================================

/** Why a fill did or didn't happen — surfaced to a fill UI per field. */
export type FillStatus = 'filled' | 'not-found' | 'locked' | 'typed' | 'data-bound';

/** Outcome of {@link fillContentControl}; `doc` is the new document on `filled`. */
export interface FillResult {
  status: FillStatus;
  doc?: Document;
}

/**
 * Fill the first control matching `filter` with `value`, returning a
 * discriminated {@link FillResult} instead of throwing. This is the
 * result-shaped wrapper over {@link setContentControlContent} for a fill UI
 * that must show a per-field outcome (filled / not found / locked / typed /
 * data-bound) rather than fail silently. Handles block and inline controls
 * everywhere (body, table cell, header/footer).
 */
export function fillContentControl(
  doc: Document,
  filter: ContentControlFilter,
  value: string,
  options: { force?: boolean } = {}
): FillResult {
  try {
    return { status: 'filled', doc: setContentControlContent(doc, filter, value, options) };
  } catch (err) {
    if (err instanceof ContentControlNotFoundError) return { status: 'not-found' };
    if (err instanceof ContentControlLockedError) return { status: 'locked' };
    if (err instanceof ContentControlTypeError) return { status: 'typed' };
    if (err instanceof ContentControlBoundError) return { status: 'data-bound' };
    throw err;
  }
}
