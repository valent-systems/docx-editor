import { describe, test, expect } from 'bun:test';
import { collectSectionHeaderFooters } from './collectSectionHeaderFooters';
import type { Document, SectionProperties } from '../types/document';
import type { HeaderFooter } from '../types/content';

// Per-section header/footer resolution. The render pipeline historically
// resolved ONE footer for the whole document (preferring finalSectionProperties
// — the body), so a cover page whose own `w:sectPr` sits inside a content
// control (`w:sdt`) showed the body's footer. These tests pin the corrected
// behavior: one resolved header/footer set per section, in document order,
// discovered even when the section break is nested inside a block SDT.

/** A footer part we can identify by object identity once resolved. */
function footer(tag: string): HeaderFooter {
  return {
    type: 'footer',
    hdrFtrType: 'default',
    content: [
      { type: 'paragraph', content: [{ type: 'run', content: [{ type: 'text', text: tag }] }] },
    ],
  } as HeaderFooter;
}

function paragraph(sectionProperties?: Partial<SectionProperties>) {
  return {
    type: 'paragraph' as const,
    content: [],
    ...(sectionProperties ? { sectionProperties: sectionProperties as SectionProperties } : {}),
  };
}

/** Build a minimal Document whose package carries the given footer parts. */
function docWith(
  content: unknown[],
  finalSectionProperties: Partial<SectionProperties> | undefined,
  footers: Map<string, HeaderFooter>
): Document {
  return {
    package: {
      headers: new Map(),
      footers,
      document: {
        content,
        finalSectionProperties: finalSectionProperties as SectionProperties | undefined,
        sections: [],
      },
    },
  } as unknown as Document;
}

describe('collectSectionHeaderFooters', () => {
  test('resolves a distinct footer per section in document order', () => {
    const coverFooter = footer('COVER');
    const bodyFooter = footer('BODY');
    const footers = new Map([
      ['rIdCover', coverFooter],
      ['rIdBody', bodyFooter],
    ]);

    const doc = docWith(
      [
        // Cover section ends here (top-level paragraph carrying the sectPr).
        paragraph({ footerReferences: [{ type: 'default', rId: 'rIdCover' }] }),
        paragraph(), // body text
      ],
      { footerReferences: [{ type: 'default', rId: 'rIdBody' }] },
      footers
    );

    const result = collectSectionHeaderFooters(doc);
    expect(result).toHaveLength(2);
    expect(result[0].footer).toBe(coverFooter);
    expect(result[1].footer).toBe(bodyFooter);
  });

  test('discovers a section break nested inside a block SDT (the real bug)', () => {
    const coverFooter = footer('COVER');
    const bodyFooter = footer('BODY');
    const footers = new Map([
      ['rIdCover', coverFooter],
      ['rIdBody', bodyFooter],
    ]);

    const doc = docWith(
      [
        // Cover section break lives INSIDE a content control, not at top level.
        {
          type: 'blockSdt',
          properties: {},
          content: [paragraph({ footerReferences: [{ type: 'default', rId: 'rIdCover' }] })],
        },
        paragraph(), // body text
      ],
      { footerReferences: [{ type: 'default', rId: 'rIdBody' }] },
      footers
    );

    const result = collectSectionHeaderFooters(doc);
    expect(result).toHaveLength(2);
    expect(result[0].footer).toBe(coverFooter); // cover footer, not body's
    expect(result[1].footer).toBe(bodyFooter);
  });

  test('handles multiple breaks inside a single SDT (TPX shape) in order', () => {
    const f0 = footer('S0');
    const f1 = footer('S1');
    const fBody = footer('BODY');
    const footers = new Map([
      ['rId0', f0],
      ['rId1', f1],
      ['rIdBody', fBody],
    ]);

    const doc = docWith(
      [
        {
          type: 'blockSdt',
          properties: {},
          content: [
            paragraph({ footerReferences: [{ type: 'default', rId: 'rId0' }] }),
            paragraph({ footerReferences: [{ type: 'default', rId: 'rId1' }] }),
          ],
        },
        paragraph(),
      ],
      { footerReferences: [{ type: 'default', rId: 'rIdBody' }] },
      footers
    );

    const result = collectSectionHeaderFooters(doc);
    expect(result).toHaveLength(3);
    expect(result[0].footer).toBe(f0);
    expect(result[1].footer).toBe(f1);
    expect(result[2].footer).toBe(fBody);
  });

  test('single-section document resolves exactly one footer (regression guard)', () => {
    const bodyFooter = footer('BODY');
    const footers = new Map([['rIdBody', bodyFooter]]);

    const doc = docWith(
      [paragraph(), paragraph()],
      { footerReferences: [{ type: 'default', rId: 'rIdBody' }] },
      footers
    );

    const result = collectSectionHeaderFooters(doc);
    expect(result).toHaveLength(1);
    expect(result[0].footer).toBe(bodyFooter);
  });
});
