/**
 * Regression tests for the text-box save round-trip data loss (DRC template).
 *
 * Three failure modes, all fixed together:
 * 1. serializeShapeContent only wrote wps:txbx for shapeType 'textBox', but
 *    fromProseDoc rebuilds PM text boxes as shapeType 'rect' — every save
 *    silently deleted text-box text, leaving empty rectangles.
 * 2. parseDrawing let a textless wps-shape drawing fall through to
 *    parseInline/parseAnchor, producing an Image with no blip — rendered as
 *    a broken empty picture frame after reload.
 * 3. Decorative textless shapes were converted to BOTH a block-level PM
 *    textBox (extractTextBoxesFromParagraph) and an inline PM shape node
 *    (runs.ts), duplicating the shape on every save.
 */
import { describe, test, expect } from 'bun:test';
import { serializeShapeContent } from './serializer/runSerializer/drawing';
import { parseDrawing } from './imageParser';
import { parseXml } from './xmlParser';
import { convertPMTextBoxRun } from '../prosemirror/conversion/fromProseDoc/textbox';
import { isBlockExtractedShape } from '../prosemirror/conversion/toProseDoc/textbox';
import { schema } from '../prosemirror/schema';
import type { Shape, ShapeContent } from '../types/document';

function rectShape(overrides: Partial<Shape> = {}): Shape {
  return {
    type: 'shape',
    shapeType: 'rect',
    size: { width: 914400, height: 457200 },
    ...overrides,
  };
}

describe('text-box save round-trip', () => {
  test('rect shape with textBody serializes its paragraphs as wps:txbx', () => {
    const content: ShapeContent = {
      type: 'shape',
      shape: rectShape({
        textBody: {
          content: [
            {
              type: 'paragraph',
              content: [{ type: 'run', content: [{ type: 'text', text: 'Mr. Varnadoe' }] }],
            },
          ],
        },
      }),
    };
    const xml = serializeShapeContent(content);
    expect(xml).toContain('<wps:txbx><w:txbxContent>');
    expect(xml).toContain('Mr. Varnadoe');
  });

  test('textless decorative shape still serializes without dropping fill', () => {
    const content: ShapeContent = {
      type: 'shape',
      shape: rectShape({
        fill: { type: 'solid', color: { rgb: '162849' } },
        textBody: { content: [] },
      }),
    };
    const xml = serializeShapeContent(content);
    expect(xml).toContain('162849');
  });

  test('PM textBox → run → XML keeps the text (the DRC caption loss)', () => {
    const pmTextBox = schema.node('textBox', { width: 150, height: 40 }, [
      schema.node('paragraph', {}, [schema.text('VALERIE BLANTON')]),
    ]);
    const run = convertPMTextBoxRun(pmTextBox);
    const shapeContent = run.content[0] as ShapeContent;
    expect(shapeContent.type).toBe('shape');
    const xml = serializeShapeContent(shapeContent);
    expect(xml).toContain('<wps:txbx><w:txbxContent>');
    expect(xml).toContain('VALERIE BLANTON');
  });

  test('parseDrawing never turns a blipless wps shape into an Image', () => {
    const drawingXml =
      '<w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
      '<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
      '<wp:simplePos x="0" y="0"/><wp:extent cx="1076325" cy="323850"/><wp:docPr id="1" name="Shape 1"/>' +
      '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
      '<wps:wsp><wps:cNvSpPr/><wps:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
      '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></wps:spPr>' +
      '<wps:bodyPr rot="0" vert="horz"/></wps:wsp>' +
      '</a:graphicData></a:graphic></wp:anchor></w:drawing>';
    const el = parseXml(drawingXml);
    expect(parseDrawing(el, undefined, undefined)).toBeNull();
  });

  test('block-extraction predicate covers text-bearing and decorative shapes exactly once', () => {
    const withText = rectShape({
      textBody: {
        content: [
          { type: 'paragraph', content: [{ type: 'run', content: [{ type: 'text', text: 'x' }] }] },
        ],
      },
    });
    const decorative = rectShape({
      fill: { type: 'solid', color: { rgb: '162849' } },
      textBody: { content: [] },
    });
    const bare = rectShape(); // no textBody at all → stays inline
    expect(isBlockExtractedShape(withText)).toBe(true);
    expect(isBlockExtractedShape(decorative)).toBe(true);
    expect(isBlockExtractedShape(bare)).toBe(false);
  });
});
