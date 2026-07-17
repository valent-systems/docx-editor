import { describe, test, expect } from 'bun:test';
import { selectPageHeaderFooter, type SectionHeaderFooterContent } from './selectPageHeaderFooter';
import type { HeaderFooterContent } from './headerFooter';
import type { Page } from '../../layout-engine/types';

// Identify each header/footer content by object identity.
const hf = (tag: string): HeaderFooterContent => ({ tag }) as unknown as HeaderFooterContent;
const pg = (p: Partial<Page>): Pick<Page, 'number' | 'sectionIndex' | 'isSectionStart'> =>
  ({ number: 1, ...p }) as Page;

describe('selectPageHeaderFooter', () => {
  describe('legacy single-section path (no per-section array)', () => {
    const globalFooter = hf('GLOBAL');
    const firstFooter = hf('FIRST');

    test('page 1 with titlePg uses the first-page footer', () => {
      const sel = selectPageHeaderFooter(pg({ number: 1 }), {
        footerContent: globalFooter,
        firstPageFooterContent: firstFooter,
        titlePg: true,
      });
      expect(sel.footerContent).toBe(firstFooter);
    });

    test('page 2 uses the global footer even with titlePg', () => {
      const sel = selectPageHeaderFooter(pg({ number: 2 }), {
        footerContent: globalFooter,
        firstPageFooterContent: firstFooter,
        titlePg: true,
      });
      expect(sel.footerContent).toBe(globalFooter);
    });

    test('without titlePg, every page uses the global footer', () => {
      const sel = selectPageHeaderFooter(pg({ number: 1 }), { footerContent: globalFooter });
      expect(sel.footerContent).toBe(globalFooter);
    });
  });

  describe('per-section path', () => {
    const coverFooter = hf('COVER');
    const bodyFooter = hf('BODY');
    const coverFirstFooter = hf('COVER_FIRST');
    const globalFooter = hf('GLOBAL');

    const sections: SectionHeaderFooterContent[] = [
      { footer: coverFooter, firstFooter: coverFirstFooter, titlePg: false },
      { footer: bodyFooter, titlePg: false },
    ];

    test('a cover page selects the cover footer, not the body/global footer', () => {
      const sel = selectPageHeaderFooter(pg({ number: 1, sectionIndex: 0, isSectionStart: true }), {
        footerContent: globalFooter,
        sectionHeaderFootersForRender: sections,
      });
      expect(sel.footerContent).toBe(coverFooter);
    });

    test('a body page selects the body footer', () => {
      const sel = selectPageHeaderFooter(pg({ number: 3, sectionIndex: 1, isSectionStart: true }), {
        footerContent: globalFooter,
        sectionHeaderFootersForRender: sections,
      });
      expect(sel.footerContent).toBe(bodyFooter);
    });

    test('a section with titlePg uses its first-page footer on its first page', () => {
      const withTitle: SectionHeaderFooterContent[] = [
        { footer: coverFooter, firstFooter: coverFirstFooter, titlePg: true },
        { footer: bodyFooter, titlePg: false },
      ];
      const first = selectPageHeaderFooter(
        pg({ number: 1, sectionIndex: 0, isSectionStart: true }),
        { sectionHeaderFootersForRender: withTitle }
      );
      expect(first.footerContent).toBe(coverFirstFooter);

      // A later page of the same section falls back to the section default.
      const later = selectPageHeaderFooter(
        pg({ number: 2, sectionIndex: 0, isSectionStart: false }),
        { sectionHeaderFootersForRender: withTitle }
      );
      expect(later.footerContent).toBe(coverFooter);
    });

    test('falls back to the global footer when a section has no footer of its own', () => {
      const partial: SectionHeaderFooterContent[] = [{ titlePg: false }];
      const sel = selectPageHeaderFooter(pg({ number: 1, sectionIndex: 0, isSectionStart: true }), {
        footerContent: globalFooter,
        sectionHeaderFootersForRender: partial,
      });
      expect(sel.footerContent).toBe(globalFooter);
    });

    test("a page carries its own section's header/footer band distances", () => {
      // Sections legitimately differ (e.g. TPX: 67px cover vs 29px body). A
      // paragraph-anchored header logo positions off this value, so using one
      // global distance shoves it off the page top on the larger-distance section.
      const withDistances: SectionHeaderFooterContent[] = [
        { footer: coverFooter, titlePg: false, headerDistancePx: 67, footerDistancePx: 15 },
        { footer: bodyFooter, titlePg: false, headerDistancePx: 29, footerDistancePx: 19 },
      ];
      const cover = selectPageHeaderFooter(
        pg({ number: 1, sectionIndex: 0, isSectionStart: true }),
        { sectionHeaderFootersForRender: withDistances }
      );
      expect(cover.headerDistancePx).toBe(67);
      expect(cover.footerDistancePx).toBe(15);

      const body = selectPageHeaderFooter(
        pg({ number: 2, sectionIndex: 1, isSectionStart: true }),
        { sectionHeaderFootersForRender: withDistances }
      );
      expect(body.headerDistancePx).toBe(29);
      expect(body.footerDistancePx).toBe(19);
    });

    test('distances stay undefined when the section omits them (global fallback)', () => {
      const partial: SectionHeaderFooterContent[] = [{ titlePg: false }];
      const sel = selectPageHeaderFooter(pg({ number: 1, sectionIndex: 0, isSectionStart: true }), {
        sectionHeaderFootersForRender: partial,
      });
      expect(sel.headerDistancePx).toBeUndefined();
      expect(sel.footerDistancePx).toBeUndefined();
    });
  });
});
