/**
 * Per-section header/footer resolution.
 *
 * The render pipeline historically resolved a SINGLE header/footer set for the
 * whole document (preferring `finalSectionProperties` — the body section), so a
 * multi-section document painted the body's footer on every page. A cover page
 * whose own `w:sectPr` sits inside a content control (`w:sdt`) therefore showed
 * the body's (often out-of-date) footer instead of its own.
 *
 * This module resolves one header/footer set PER SECTION, in document order, by:
 *  1. Walking the document body in order and collecting every section-ending
 *     paragraph's `sectionProperties` — recursing into block SDTs, where Word
 *     legitimately nests section breaks (`buildSections` only scans top-level
 *     blocks, which is why `body.sections` collapses for these documents).
 *  2. Appending `finalSectionProperties` (the body-level `w:sectPr`) as the last
 *     section.
 *  3. Applying ECMA-376 §17.6 section inheritance (reusing the parser's own
 *     `applySectionInheritance`) so a section that omits a header/footer ref
 *     inherits it from the previous section — exactly as Word renders it.
 *  4. Resolving each section's `headerReferences`/`footerReferences` against the
 *     package's header/footer maps via `resolveHeaderFooter`.
 *
 * The result is ordered to match the section-break sequence the layout engine
 * produces from the same `_sectionProperties`, so a page tagged with section
 * index `i` selects `result[i]`.
 */

import { applySectionInheritance, getDefaultSectionProperties } from '../docx/sectionParser';
import { resolveHeaderFooter, twipsToPixels } from './sectionGeometry';
import type { Document, Section, SectionProperties } from '../types/document';
import type { BlockContent } from '../types/content';
import type { HeaderFooter } from '../types/content';

/** A resolved header/footer set for one section (default + first-page variants). */
export interface SectionHeaderFooter {
  header: HeaderFooter | null;
  footer: HeaderFooter | null;
  firstHeader: HeaderFooter | null;
  firstFooter: HeaderFooter | null;
  /** When true, the section's first page uses the first-page (title) variant. */
  titlePg: boolean;
  /**
   * This section's `w:pgMar w:header` distance (px) — where the header band
   * starts from the page top. Sections legitimately differ (e.g. a cover at
   * 1008 twips vs body sections at 432), and a paragraph-anchored header logo
   * with a negative offset is positioned off this value, so using one global
   * distance pushes it off the page top on sections with a larger distance.
   * Undefined when the section doesn't specify one.
   */
  headerDistancePx?: number;
  /** This section's `w:pgMar w:footer` distance (px). See headerDistancePx. */
  footerDistancePx?: number;
}

/**
 * Ordered section properties in document order — one per section-ending
 * paragraph (recursing into block SDTs), with `finalSectionProperties` last.
 * Inheritance is NOT applied here; see {@link collectSectionHeaderFooters}.
 */
export function collectOrderedSectionProperties(
  document: Document | null | undefined
): SectionProperties[] {
  const body = document?.package?.document;
  if (!body) return [];

  const ordered: SectionProperties[] = [];
  const walk = (blocks: BlockContent[] | undefined): void => {
    for (const block of blocks ?? []) {
      if (block.type === 'paragraph') {
        if (block.sectionProperties) ordered.push(block.sectionProperties);
      } else if (block.type === 'blockSdt') {
        // Section breaks are legitimately nested inside content controls.
        walk(block.content);
      }
      // Tables carry no section break (OOXML forbids `w:sectPr` in a cell).
    }
  };
  walk(body.content);

  ordered.push(body.finalSectionProperties ?? getDefaultSectionProperties());
  return ordered;
}

/**
 * Resolve one {@link SectionHeaderFooter} per section, in document order.
 * A page laid out under section index `i` selects `result[i]`.
 */
export function collectSectionHeaderFooters(
  document: Document | null | undefined
): SectionHeaderFooter[] {
  const ordered = collectOrderedSectionProperties(document);
  // Reuse the parser's tested §17.6 inheritance so omitted refs fall back to the
  // previous section, matching Word — then resolve each against the HF maps.
  const sections: Section[] = ordered.map((properties) => ({ properties, content: [] }));
  const inherited = applySectionInheritance(sections);
  return inherited.map((s) => ({
    ...resolveHeaderFooter(document ?? null, s.properties),
    titlePg: s.properties.titlePg === true,
    headerDistancePx:
      s.properties.headerDistance != null ? twipsToPixels(s.properties.headerDistance) : undefined,
    footerDistancePx:
      s.properties.footerDistance != null ? twipsToPixels(s.properties.footerDistance) : undefined,
  }));
}
