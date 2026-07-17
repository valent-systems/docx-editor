import { describe, test, expect } from 'bun:test';
import { floatingImageIsBehindDoc } from './floatingImageFlow';
import { parseWrapElement } from '../docx/drawingUtils';
import { serializeDrawingContent } from '../docx/serializer/runSerializer/drawing';
import type { Image } from '../types/content';

// A picture set "behind text" in Word keeps behindDoc=1 while wrapping tightly.
// Its z-order intent (behind body text and text boxes) must survive parse,
// the layout classification, and save — independent of the wrap type.

describe('behindDoc z-order (independent of wrap type)', () => {
  test('a tight image with behindDoc paints behind; a plain tight image does not', () => {
    expect(floatingImageIsBehindDoc({ wrapType: 'tight', behindDoc: true })).toBe(true);
    expect(floatingImageIsBehindDoc({ wrapType: 'tight' })).toBe(false);
    // wrapNone folds behindDoc into the wrap type — still behind.
    expect(floatingImageIsBehindDoc({ wrapType: 'behind' })).toBe(true);
    expect(floatingImageIsBehindDoc({ wrapType: 'inFront' })).toBe(false);
  });

  test('parseWrapElement preserves behindDoc on a wrapTight element', () => {
    const wrapEl = { name: 'wp:wrapTight', attributes: {}, children: [] } as never;
    const wrap = parseWrapElement(wrapEl, true);
    expect(wrap.type).toBe('tight');
    expect(wrap.behindDoc).toBe(true);

    const notBehind = parseWrapElement(wrapEl, false);
    expect(notBehind.behindDoc).toBeUndefined();
  });

  test('serializer emits behindDoc="1" for a tight+behindDoc image (round-trip fidelity)', () => {
    const image: Image = {
      type: 'image',
      rId: 'rId1',
      size: { width: 914400, height: 914400 },
      wrap: { type: 'tight', behindDoc: true },
      position: {
        horizontal: { relativeTo: 'column', posOffset: 0 },
        vertical: { relativeTo: 'paragraph', posOffset: 0 },
      },
    };
    const xml = serializeDrawingContent({ type: 'drawing', image });
    expect(xml).toContain('behindDoc="1"');
    expect(xml).toContain('<wp:wrapTight');
  });
});
