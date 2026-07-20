/**
 * Regression — images INSIDE a text box must resolve their media (src) just
 * like body images. parseTextBoxContent accepted the media map but dropped it
 * when delegating to parseParagraph, so every text-box image parsed with an
 * empty src and painted as a broken image (TPX Teams-page icons).
 */

import { describe, expect, test } from 'bun:test';
import { parseDocumentBody } from '../documentParser';
import type { MediaFile, RelationshipMap } from '../../types/document';

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"';

const INLINE_IMAGE_RUN = `
  <w:r>
    <w:drawing>
      <wp:inline distT="0" distB="0" distL="0" distR="0">
        <wp:extent cx="914400" cy="914400"/>
        <wp:docPr id="7" name="Icon"/>
        <a:graphic>
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
            <pic:pic>
              <pic:nvPicPr>
                <pic:cNvPr id="7" name="icon.png"/>
                <pic:cNvPicPr/>
              </pic:nvPicPr>
              <pic:blipFill>
                <a:blip r:embed="rId9"/>
                <a:stretch><a:fillRect/></a:stretch>
              </pic:blipFill>
              <pic:spPr>
                <a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>
                <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
              </pic:spPr>
            </pic:pic>
          </a:graphicData>
        </a:graphic>
      </wp:inline>
    </w:drawing>
  </w:r>`;

function buildDocWithTextBoxImage(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:document ${NS}>
      <w:body>
        <w:p>
          <w:r>
            <mc:AlternateContent>
              <mc:Choice Requires="wps">
                <w:drawing>
                  <wp:anchor distT="0" distB="0" distL="114300" distR="114300"
                    simplePos="0" relativeHeight="251658240" behindDoc="0" locked="0"
                    layoutInCell="1" allowOverlap="1">
                    <wp:simplePos x="0" y="0"/>
                    <wp:positionH relativeFrom="margin"><wp:align>left</wp:align></wp:positionH>
                    <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
                    <wp:extent cx="1390650" cy="1390650"/>
                    <wp:effectExtent l="0" t="0" r="0" b="0"/>
                    <wp:wrapSquare wrapText="bothSides"/>
                    <wp:docPr id="1" name="Text Box 1"/>
                    <wp:cNvGraphicFramePr/>
                    <a:graphic>
                      <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                        <wps:wsp>
                          <wps:cNvSpPr txBox="1"/>
                          <wps:spPr>
                            <a:xfrm><a:off x="0" y="0"/><a:ext cx="1390650" cy="1390650"/></a:xfrm>
                            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                          </wps:spPr>
                          <wps:txbx>
                            <w:txbxContent>
                              <w:p>${INLINE_IMAGE_RUN}</w:p>
                            </w:txbxContent>
                          </wps:txbx>
                          <wps:bodyPr/>
                        </wps:wsp>
                      </a:graphicData>
                    </a:graphic>
                  </wp:anchor>
                </w:drawing>
              </mc:Choice>
              <mc:Fallback><w:pict/></mc:Fallback>
            </mc:AlternateContent>
          </w:r>
        </w:p>
      </w:body>
    </w:document>`;
}

// 1x1 transparent PNG
const PNG_DATA_URL =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==';

function makeRelsAndMedia(): { rels: RelationshipMap; media: Map<string, MediaFile> } {
  const rels: RelationshipMap = new Map([
    [
      'rId9',
      {
        id: 'rId9',
        type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
        target: 'media/icon.png',
      },
    ],
  ]) as unknown as RelationshipMap;
  const mediaFile = {
    filename: 'icon.png',
    mimeType: 'image/png',
    dataUrl: PNG_DATA_URL,
  } as unknown as MediaFile;
  const media = new Map<string, MediaFile>([
    ['word/media/icon.png', mediaFile],
    ['media/icon.png', mediaFile],
  ]);
  return { rels, media };
}

function findImagesDeep(n: unknown, out: Array<{ src?: string; rId?: string }>): void {
  if (!n || typeof n !== 'object') return;
  const obj = n as Record<string, unknown>;
  if (obj.type === 'image') out.push(obj as { src?: string; rId?: string });
  for (const v of Object.values(obj)) {
    if (Array.isArray(v)) v.forEach((c) => findImagesDeep(c, out));
    else if (v && typeof v === 'object') findImagesDeep(v, out);
  }
}

describe('text-box image media resolution', () => {
  test('an image inside a wps text box resolves its src from the media map', () => {
    const { rels, media } = makeRelsAndMedia();
    const body = parseDocumentBody(buildDocWithTextBoxImage(), null, null, null, rels, media);

    const images: Array<{ src?: string; rId?: string }> = [];
    findImagesDeep(body.content, images);

    expect(images.length).toBe(1);
    expect(images[0].rId).toBe('rId9');
    // The regression: this was undefined because parseTextBoxContent dropped
    // the media map on its way to parseParagraph.
    expect(images[0].src).toBeTruthy();
    expect(images[0].src!.startsWith('data:image/png')).toBe(true);
  });
});
