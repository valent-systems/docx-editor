/**
 * Section-properties → page geometry helpers.
 *
 * Both adapters need the same translation from `SectionProperties` (twips
 * + DOCX field names) into the layout engine's pixel-shaped `PageSize`
 * and `PageMargins`. Living in core eliminates the subtle React/Vue
 * drift that grows when each adapter ships its own twipsToPixels math.
 */

import type { PageMargins } from '../layout-engine/types';
import type { Document, SectionProperties } from '../types/document';
import type { HeaderFooter } from '../types/content';

/** US Letter at 96 DPI — Word's default page size. */
export const DEFAULT_PAGE_WIDTH_PX = 816;
/** US Letter at 96 DPI. */
export const DEFAULT_PAGE_HEIGHT_PX = 1056;
/** 1 inch at 96 DPI — Word's default body margins. */
export const DEFAULT_BODY_MARGIN_PX = 96;
/** Word's default `headerDistance` / `footerDistance` (0.5in = 48px). */
export const DEFAULT_HF_DISTANCE_PX = 48;

function twipsToPixels(twips: number): number {
  return Math.round((twips / 1440) * 96);
}

/** Convert SectionProperties page size (twips) → pixel `{ w, h }`. */
export function getPageSize(sp: SectionProperties | null | undefined): {
  w: number;
  h: number;
} {
  return {
    w: sp?.pageWidth ? twipsToPixels(sp.pageWidth) : DEFAULT_PAGE_WIDTH_PX,
    h: sp?.pageHeight ? twipsToPixels(sp.pageHeight) : DEFAULT_PAGE_HEIGHT_PX,
  };
}

/**
 * Convert SectionProperties margins (twips) → pixel `PageMargins`.
 *
 * `header` / `footer` default to 48px (Word's 0.5-inch default) so the
 * HF margin-extension math doesn't have to special-case undefined.
 */
export function getMargins(sp: SectionProperties | null | undefined): PageMargins {
  return {
    top: sp?.marginTop ? twipsToPixels(sp.marginTop) : DEFAULT_BODY_MARGIN_PX,
    right: sp?.marginRight ? twipsToPixels(sp.marginRight) : DEFAULT_BODY_MARGIN_PX,
    bottom: sp?.marginBottom ? twipsToPixels(sp.marginBottom) : DEFAULT_BODY_MARGIN_PX,
    left: sp?.marginLeft ? twipsToPixels(sp.marginLeft) : DEFAULT_BODY_MARGIN_PX,
    header: sp?.headerDistance ? twipsToPixels(sp.headerDistance) : DEFAULT_HF_DISTANCE_PX,
    footer: sp?.footerDistance ? twipsToPixels(sp.footerDistance) : DEFAULT_HF_DISTANCE_PX,
  };
}

/**
 * Resolve the HeaderFooter pair (default + first-page) for a section.
 *
 * Mirrors React's lookup in DocxEditor.tsx: read `pkg.headers`/`footers`
 * (Maps keyed by rId), resolve through `sp.headerReferences` /
 * `footerReferences`. When `titlePg` is unset and only `first` HFs
 * exist, they serve as the default — same Word fallback both adapters
 * have shipped.
 */
export function resolveHeaderFooter(
  doc: Document | null,
  sp: SectionProperties | null | undefined
): {
  header: HeaderFooter | null;
  footer: HeaderFooter | null;
  firstHeader: HeaderFooter | null;
  firstFooter: HeaderFooter | null;
} {
  const empty = { header: null, footer: null, firstHeader: null, firstFooter: null };
  if (!doc?.package) return empty;
  const headers = doc.package.headers;
  const footers = doc.package.footers;

  let header: HeaderFooter | null = null;
  let footer: HeaderFooter | null = null;
  let firstHeader: HeaderFooter | null = null;
  let firstFooter: HeaderFooter | null = null;

  if (headers && sp?.headerReferences) {
    const def = sp.headerReferences.find((r) => r.type === 'default');
    if (def?.rId) header = headers.get(def.rId) ?? null;
    const first = sp.headerReferences.find((r) => r.type === 'first');
    if (first?.rId) firstHeader = headers.get(first.rId) ?? null;
  }

  if (footers && sp?.footerReferences) {
    const def = sp.footerReferences.find((r) => r.type === 'default');
    if (def?.rId) footer = footers.get(def.rId) ?? null;
    const first = sp.footerReferences.find((r) => r.type === 'first');
    if (first?.rId) firstFooter = footers.get(first.rId) ?? null;
  }

  if (!sp?.titlePg) {
    if (!header && firstHeader) header = firstHeader;
    if (!footer && firstFooter) footer = firstFooter;
  }

  return { header, footer, firstHeader, firstFooter };
}
