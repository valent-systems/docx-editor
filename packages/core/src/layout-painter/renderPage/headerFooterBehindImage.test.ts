import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { beforeAll, afterAll, describe, test, expect } from 'bun:test';
import {
  renderHeaderFooterContent,
  renderBehindHeaderFooterImagesLayer,
  type HeaderFooterContent,
  type HeaderFooterLayoutInfo,
  type BehindHeaderFooterImage,
} from './headerFooter';

beforeAll(() => GlobalRegistrator.register());
afterAll(() => GlobalRegistrator.unregister());

// A single header paragraph whose only run is a floating image with the given
// wrap type — mirrors how a full-bleed background (wrapNone + behindDoc) or a
// tight-wrapped logo reaches the painter after flow conversion.
function headerWithFloatImage(wrapType: string): HeaderFooterContent {
  const block = {
    kind: 'paragraph' as const,
    id: 'p1',
    attrs: {},
    runs: [
      {
        kind: 'image' as const,
        src: 'data:image/png;base64,AAAA',
        width: 800,
        height: 700,
        alt: 'bg',
        wrapType,
        // relativeTo column, posOffset -95px ; relativeTo page, +120px
        position: {
          horizontal: { relativeTo: 'column', posOffset: -904875 },
          vertical: { relativeTo: 'page', posOffset: 1143000 },
        },
      },
    ],
  };
  const measure = { kind: 'paragraph' as const, totalHeight: 20, lines: [{ height: 20 }] };
  return { blocks: [block], measures: [measure], height: 20 } as unknown as HeaderFooterContent;
}

const layout: HeaderFooterLayoutInfo = {
  flowTop: 67,
  flowLeft: 48,
  contentWidth: 720,
  pageWidth: 816,
  pageHeight: 1056,
  margins: { top: 96, right: 48, bottom: 96, left: 48 },
};

const ctx = { pageNumber: 2, totalPages: 3, section: 'header' as const, contentWidth: 720 };

describe('header/footer behind-image lifting', () => {
  test('a `behind` float is lifted out of the container into the out-array (page-absolute)', () => {
    const out: BehindHeaderFooterImage[] = [];
    const el = renderHeaderFooterContent(
      headerWithFloatImage('behind'),
      ctx as never,
      {} as never,
      layout,
      out
    );

    // Not painted inside the header container...
    expect(el.querySelectorAll('img').length).toBe(0);
    // ...but collected for the caller, in page-absolute coords.
    expect(out).toHaveLength(1);
    expect(out[0]!.width).toBe(800);
    // horizontal: flowLeft(48) + posOffset(-95) = -47 (full-bleed off left edge)
    expect(out[0]!.left).toBeCloseTo(48 + -904875 / 9525, 1);
    // vertical: page-anchored posOffset resolves to the absolute page offset (120)
    expect(out[0]!.top).toBeCloseTo(1143000 / 9525, 1);
  });

  test('a non-`behind` float still renders inside the header container', () => {
    const out: BehindHeaderFooterImage[] = [];
    const el = renderHeaderFooterContent(
      headerWithFloatImage('tight'),
      ctx as never,
      {} as never,
      layout,
      out
    );
    expect(out).toHaveLength(0);
    expect(el.querySelectorAll('img').length).toBe(1);
  });

  test('without an out-array, a `behind` float falls back to in-container paint (back-compat)', () => {
    const el = renderHeaderFooterContent(
      headerWithFloatImage('behind'),
      ctx as never,
      {} as never,
      layout
    );
    expect(el.querySelectorAll('img').length).toBe(1);
  });

  test('the behind layer paints each image at its page-absolute position', () => {
    const layer = renderBehindHeaderFooterImagesLayer(
      [
        {
          src: 'data:image/png;base64,AAAA',
          width: 800,
          height: 700,
          alt: 'bg',
          left: -47,
          top: 120,
        },
      ],
      document
    );
    const img = layer.querySelector('img')!;
    expect(layer.className).toBe('layout-hf-behind-layer');
    expect(layer.style.pointerEvents).toBe('none');
    expect(img.style.left).toBe('-47px');
    expect(img.style.top).toBe('120px');
  });
});
