/**
 * Per-page header/footer selection.
 *
 * Historically the painter rendered one document-global header/footer (swapping
 * to a first-page variant only on the title page), so multi-section documents
 * showed a single section's footer everywhere. This selects the correct set for
 * each page:
 *  - When `sectionHeaderFootersForRender` is present, a page uses its own
 *    section's set (`page.sectionIndex`), and its section's first-page variant
 *    on the section's first page (`page.isSectionStart` + that section's
 *    `titlePg`).
 *  - When it is absent, behavior is byte-identical to the legacy single-section
 *    path: the global header/footer, with the first-page variant on page 1 when
 *    `titlePg` is set.
 */

import type { Page } from '../../layout-engine/types';
import type { HeaderFooterContent } from './headerFooter';

/**
 * A section's header/footer content in render form (default + first-page
 * variants), indexed by `page.sectionIndex`.
 */
export interface SectionHeaderFooterContent {
  header?: HeaderFooterContent;
  footer?: HeaderFooterContent;
  firstHeader?: HeaderFooterContent;
  firstFooter?: HeaderFooterContent;
  /** When true, this section's first page uses the first-page variant. */
  titlePg: boolean;
  /**
   * This section's `w:pgMar w:header` / `w:footer` distances (px). Sections
   * legitimately differ; the page painter must use ITS section's distance or a
   * paragraph-anchored header image is positioned off the wrong band origin
   * (e.g. a cover logo pushed off the page top). Undefined = fall back to the
   * document-global distance.
   */
  headerDistancePx?: number;
  footerDistancePx?: number;
}

export interface PageHeaderFooterInputs {
  headerContent?: HeaderFooterContent;
  footerContent?: HeaderFooterContent;
  firstPageHeaderContent?: HeaderFooterContent;
  firstPageFooterContent?: HeaderFooterContent;
  titlePg?: boolean;
  sectionHeaderFootersForRender?: SectionHeaderFooterContent[];
}

export interface PageHeaderFooterSelection {
  headerContent?: HeaderFooterContent;
  footerContent?: HeaderFooterContent;
  /** The page's section-specific band distances (px); undefined = global. */
  headerDistancePx?: number;
  footerDistancePx?: number;
}

export function selectPageHeaderFooter(
  page: Pick<Page, 'number' | 'sectionIndex' | 'isSectionStart'>,
  opts: PageHeaderFooterInputs
): PageHeaderFooterSelection {
  // Per-section path: a page uses its own section's header/footer.
  if (opts.sectionHeaderFootersForRender && page.sectionIndex != null) {
    const section = opts.sectionHeaderFootersForRender[page.sectionIndex];
    if (section) {
      const useFirst = section.titlePg && page.isSectionStart === true;
      return {
        headerContent: (useFirst ? section.firstHeader : section.header) ?? opts.headerContent,
        footerContent: (useFirst ? section.firstFooter : section.footer) ?? opts.footerContent,
        headerDistancePx: section.headerDistancePx,
        footerDistancePx: section.footerDistancePx,
      };
    }
  }

  // Legacy single-section path: first-page variant on page 1 when titlePg is set.
  if (opts.titlePg && page.number === 1) {
    return {
      headerContent: opts.firstPageHeaderContent,
      footerContent: opts.firstPageFooterContent,
    };
  }

  return { headerContent: opts.headerContent, footerContent: opts.footerContent };
}
