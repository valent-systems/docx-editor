/**
 * Paragraph Serializer - Serialize paragraphs to OOXML XML
 *
 * Converts Paragraph objects back to <w:p> XML format for DOCX files.
 * Handles all paragraph properties and child content (runs, hyperlinks, fields, bookmarks).
 *
 * pPr property serializers (borders/shading/tabs/spacing/indentation/
 * numbering/frame) live in `paragraphSerializer/properties.ts`; child
 * content serializers (hyperlinks/fields/SDT/tracked-change wrappers)
 * live in `paragraphSerializer/content.ts`. This file orchestrates
 * paragraph-level serialization and re-exports the public API consumed
 * by sibling serializers.
 *
 * OOXML Reference:
 * - Paragraph: w:p
 * - Paragraph properties: w:pPr
 * - Runs, hyperlinks, bookmarks, fields as child elements
 */

import type {
  Paragraph,
  ParagraphFormatting,
  ParagraphPropertyChange,
  TextFormatting,
} from '../../types/document';

import { serializeTextFormatting } from './runSerializer';
import { escapeXml } from './xmlUtils';
import {
  serializeFrameProperties,
  serializeIndentation,
  serializeNumbering,
  serializeParagraphBorders,
  serializeShading,
  serializeSpacing,
  serializeTabStops,
} from './paragraphSerializer/properties';
import { serializeParagraphContent } from './paragraphSerializer/content';

/**
 * Serialize paragraph formatting properties to w:pPr XML
 */
export function serializeParagraphFormatting(
  formatting: ParagraphFormatting | undefined,
  propertyChanges?: ParagraphPropertyChange[]
): string {
  const parts: string[] = [];

  if (formatting) {
    // Style reference (must be first)
    if (formatting.styleId) {
      parts.push(`<w:pStyle w:val="${escapeXml(formatting.styleId)}"/>`);
    }

    // Keep next/lines/widow
    if (formatting.keepNext) {
      parts.push('<w:keepNext/>');
    }

    if (formatting.keepLines) {
      parts.push('<w:keepLines/>');
    }

    if (formatting.contextualSpacing) {
      parts.push('<w:contextualSpacing/>');
    }

    if (formatting.pageBreakBefore) {
      parts.push('<w:pageBreakBefore/>');
    }

    // Frame properties
    const frameXml = serializeFrameProperties(formatting.frame);
    if (frameXml) {
      parts.push(frameXml);
    }

    // Widow control
    if (formatting.widowControl === false) {
      parts.push('<w:widowControl w:val="0"/>');
    } else if (formatting.widowControl === true) {
      parts.push('<w:widowControl/>');
    }

    // Numbering
    const numPrXml = serializeNumbering(formatting.numPr);
    if (numPrXml) {
      parts.push(numPrXml);
    }

    // Paragraph borders
    const bordersXml = serializeParagraphBorders(formatting.borders);
    if (bordersXml) {
      parts.push(bordersXml);
    }

    // Shading
    const shadingXml = serializeShading(formatting.shading);
    if (shadingXml) {
      parts.push(shadingXml);
    }

    // Tabs
    const tabsXml = serializeTabStops(formatting.tabs);
    if (tabsXml) {
      parts.push(tabsXml);
    }

    // Suppress line numbers
    if (formatting.suppressLineNumbers) {
      parts.push('<w:suppressLineNumbers/>');
    }

    // Suppress auto hyphens
    if (formatting.suppressAutoHyphens) {
      parts.push('<w:suppressAutoHyphens/>');
    }

    // Spacing
    const spacingXml = serializeSpacing(formatting);
    if (spacingXml) {
      parts.push(spacingXml);
    }

    // Indentation
    const indXml = serializeIndentation(formatting);
    if (indXml) {
      parts.push(indXml);
    }

    // Text direction (bidi)
    if (formatting.bidi) {
      parts.push('<w:bidi/>');
    }

    // Justification
    if (formatting.alignment) {
      parts.push(`<w:jc w:val="${formatting.alignment}"/>`);
    }

    // Outline level
    if (formatting.outlineLevel !== undefined) {
      parts.push(`<w:outlineLvl w:val="${formatting.outlineLevel}"/>`);
    }

    // Run properties (default run formatting for paragraph)
    if (formatting.runProperties) {
      const rPrXml = serializeTextFormatting(formatting.runProperties);
      if (rPrXml) {
        parts.push(rPrXml);
      }
    }
  }

  if (propertyChanges && propertyChanges.length > 0) {
    parts.push(...propertyChanges.map((change) => serializeParagraphPropertyChange(change)));
  }

  if (parts.length === 0) return '';

  return `<w:pPr>${parts.join('')}</w:pPr>`;
}

function extractPPrInner(pPrXml: string): string {
  if (!pPrXml.startsWith('<w:pPr>') || !pPrXml.endsWith('</w:pPr>')) {
    return '';
  }
  return pPrXml.slice('<w:pPr>'.length, -'</w:pPr>'.length);
}

function serializeParagraphPropertyChange(change: ParagraphPropertyChange): string {
  const normalizedId = Number.isInteger(change.info.id) && change.info.id >= 0 ? change.info.id : 0;
  const authorCandidate = typeof change.info.author === 'string' ? change.info.author.trim() : '';
  const normalizedAuthor = authorCandidate.length > 0 ? authorCandidate : 'Unknown';
  const normalizedDate = typeof change.info.date === 'string' ? change.info.date.trim() : undefined;
  const normalizedRsid = typeof change.info.rsid === 'string' ? change.info.rsid.trim() : undefined;
  const attrs = [`w:id="${normalizedId}"`, `w:author="${escapeXml(normalizedAuthor)}"`];
  if (normalizedDate) {
    attrs.push(`w:date="${escapeXml(normalizedDate)}"`);
  }
  if (normalizedRsid) {
    attrs.push(`w:rsid="${escapeXml(normalizedRsid)}"`);
  }

  const previousPPrXml = serializeParagraphFormatting(change.previousFormatting) || '<w:pPr/>';
  const previousPPrInner = extractPPrInner(previousPPrXml);
  const normalizedPreviousPPr =
    previousPPrInner.length > 0 ? `<w:pPr>${previousPPrInner}</w:pPr>` : '<w:pPr/>';
  return `<w:pPrChange ${attrs.join(' ')}>${normalizedPreviousPPr}</w:pPrChange>`;
}

/**
 * Serialize a paragraph to OOXML XML (w:p)
 *
 * @param paragraph - The paragraph to serialize
 * @returns XML string for the paragraph
 */
export function serializeParagraph(paragraph: Paragraph): string {
  const parts: string[] = [];

  // Paragraph ID attributes
  const attrs: string[] = [];
  if (paragraph.paraId) {
    attrs.push(`w14:paraId="${paragraph.paraId}"`);
  }
  if (paragraph.textId) {
    attrs.push(`w14:textId="${paragraph.textId}"`);
  }
  const attrsStr = attrs.length > 0 ? ' ' + attrs.join(' ') : '';

  // Add paragraph properties if present
  const pPrXml = serializeParagraphFormatting(paragraph.formatting, paragraph.propertyChanges);
  if (pPrXml) {
    parts.push(pPrXml);
  }

  // Add paragraph content. Marker injection (when `renderedPageBreakBefore`
  // is set) is handled by `injectRenderedPageBreakIntoFirstRun` below.
  let pendingRenderedPageBreak = !!paragraph.renderedPageBreakBefore;
  for (const content of paragraph.content) {
    let contentXml = serializeParagraphContent(content);
    if (!contentXml) continue;
    if (pendingRenderedPageBreak) {
      const next = injectRenderedPageBreakIntoFirstRun(contentXml);
      if (next) {
        contentXml = next;
        pendingRenderedPageBreak = false;
      }
    }
    parts.push(contentXml);
  }

  return `<w:p${attrsStr}>${parts.join('')}</w:p>`;
}

/**
 * Insert `<w:lastRenderedPageBreak/>` after the first `<w:r ...>` opening
 * tag in `xml` (matches runs nested in hyperlink / sdt / ins / del /
 * moveFrom / moveTo / smartTag wrappers). Returns `null` when no `<w:r>`
 * is present so the caller can keep scanning later siblings. The lookahead
 * `(?=[\s>/])` skips `<w:rPr>` and any other prefix-collision tag.
 */
function injectRenderedPageBreakIntoFirstRun(xml: string): string | null {
  const re = /<w:r(?=[\s>/])[^>]*>/;
  if (!re.test(xml)) return null;
  return xml.replace(re, (match) => `${match}<w:lastRenderedPageBreak/>`);
}

/**
 * Serialize multiple paragraphs to OOXML XML
 *
 * @param paragraphs - The paragraphs to serialize
 * @returns XML string for all paragraphs
 */
export function serializeParagraphs(paragraphs: Paragraph[]): string {
  return paragraphs.map(serializeParagraph).join('');
}

/**
 * Check if a paragraph has any content
 */
export function hasParagraphContent(paragraph: Paragraph): boolean {
  return paragraph.content.length > 0;
}

/**
 * Check if a paragraph has formatting
 */
export function hasParagraphFormatting(paragraph: Paragraph): boolean {
  return paragraph.formatting !== undefined && Object.keys(paragraph.formatting).length > 0;
}

/**
 * Get plain text from a paragraph (for comparison/debugging)
 */
export function getParagraphPlainText(paragraph: Paragraph): string {
  const texts: string[] = [];

  for (const content of paragraph.content) {
    if (content.type === 'run') {
      for (const item of content.content) {
        if (item.type === 'text') {
          texts.push(item.text);
        } else if (item.type === 'tab') {
          texts.push('\t');
        } else if (item.type === 'break') {
          texts.push('\n');
        }
      }
    } else if (content.type === 'hyperlink') {
      for (const child of content.children) {
        if (child.type === 'run') {
          for (const item of child.content) {
            if (item.type === 'text') {
              texts.push(item.text);
            }
          }
        }
      }
    } else if (content.type === 'simpleField') {
      for (const item of content.content) {
        if (item.type === 'run') {
          for (const subItem of item.content) {
            if (subItem.type === 'text') {
              texts.push(subItem.text);
            }
          }
        }
      }
    } else if (content.type === 'complexField') {
      for (const run of content.fieldResult) {
        for (const item of run.content) {
          if (item.type === 'text') {
            texts.push(item.text);
          }
        }
      }
    } else if (content.type === 'inlineSdt') {
      for (const item of content.content) {
        if (item.type === 'run') {
          for (const subItem of item.content) {
            if (subItem.type === 'text') {
              texts.push(subItem.text);
            }
          }
        }
      }
    } else if (
      content.type === 'insertion' ||
      content.type === 'deletion' ||
      content.type === 'moveFrom' ||
      content.type === 'moveTo'
    ) {
      for (const item of content.content) {
        if (item.type === 'run') {
          for (const subItem of item.content) {
            if (subItem.type === 'text') {
              texts.push(subItem.text);
            }
          }
        }
      }
    }
  }

  return texts.join('');
}

/**
 * Create an empty paragraph
 */
export function createEmptyParagraph(formatting?: ParagraphFormatting): Paragraph {
  return {
    type: 'paragraph',
    formatting,
    content: [],
  };
}

/**
 * Create a paragraph with a single text run
 */
export function createTextParagraph(
  text: string,
  paragraphFormatting?: ParagraphFormatting,
  textFormatting?: TextFormatting
): Paragraph {
  return {
    type: 'paragraph',
    formatting: paragraphFormatting,
    content: [
      {
        type: 'run',
        formatting: textFormatting,
        content: [{ type: 'text', text }],
      },
    ],
  };
}

/**
 * Check if paragraph is a list item
 */
export function isListParagraph(paragraph: Paragraph): boolean {
  return paragraph.formatting?.numPr !== undefined;
}

/**
 * Get list level of a paragraph (0-8, or -1 if not a list)
 */
export function getListLevel(paragraph: Paragraph): number {
  return paragraph.formatting?.numPr?.ilvl ?? -1;
}

export default serializeParagraph;
