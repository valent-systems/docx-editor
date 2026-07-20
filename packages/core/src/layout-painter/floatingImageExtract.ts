/**
 * Extraction of paragraph-anchored floating images for page painting.
 * Split from renderPage.ts (line-cap): owns the PageFloatingImage record and
 * the fragment-scoped extraction that keeps a split paragraph's anchored
 * images on the page holding their anchor line (Word semantics).
 */

import type { Measure, ParagraphBlock, ImageRun } from '../layout-engine/types';
import { isFloatingImageRun, imageWrapTextFromCssFloat } from './floatingImageFlow';
import { resolveAnchoredObjectPosition, type PageGeometry } from './anchoredObjectPosition';

/**
 * Page-level floating image that has been extracted from paragraphs.
 * These are positioned absolutely within the page's content area.
 */
export interface PageFloatingImage {
  src: string;
  width: number;
  height: number;
  alt?: string;
  transform?: string;
  /** Which side: 'left' for left margin, 'right' for right margin */
  side: 'left' | 'right';
  /** X position relative to content area (0 = left edge of content) */
  x: number;
  /** Y position relative to content area (0 = top of content) */
  y: number;
  /** Wrap distances */
  distTop: number;
  distBottom: number;
  distLeft: number;
  distRight: number;
  /** ProseMirror start position for click-to-select */
  pmStart?: number;
  /** ProseMirror end position */
  pmEnd?: number;
  /** OOXML wrapText: which side(s) TEXT flows on */
  wrapText?: 'bothSides' | 'left' | 'right' | 'largest';
  /** Wrap type (square, tight, through, topAndBottom) */
  wrapType?: string;
  /** `wp:anchor@behindDoc`: paint behind text regardless of wrap type. */
  behindDoc?: boolean;
  cropTop?: number;
  cropRight?: number;
  cropBottom?: number;
  cropLeft?: number;
  /** a:alphaModFix → opacity. */
  opacity?: number;
}

/**
 * Extract floating images from a paragraph block and determine their page-level positions.
 * Returns extracted images and info for the paragraph about space reserved.
 */
export function extractFloatingImagesFromParagraph(
  block: ParagraphBlock,
  fragmentY: number, // Y position of the paragraph fragment on the page (relative to content area)
  contentWidth: number, // Width of the content area
  geometry?: PageGeometry,
  // When the paragraph splits across pages, Word paints an anchored image
  // ONCE — on the page holding its anchor line. Pass the fragment's line
  // window + the paragraph measure to scope extraction; without them (legacy
  // callers) every run extracts, which duplicated headshots on both sides of
  // a page break and stamped the copy over the continuation's text.
  fragmentLines?: { fromLine: number; toLine: number },
  measure?: Measure
): PageFloatingImage[] {
  const floatingImages: PageFloatingImage[] = [];

  const lines = measure?.kind === 'paragraph' ? measure.lines : undefined;
  const runInFragment = (runIndex: number): boolean => {
    if (!fragmentLines || !lines) return true;
    const lineIdx = lines.findIndex((l) => runIndex >= l.fromRun && runIndex <= l.toRun);
    // Anchor line not found (e.g. floating run excluded from line boxes):
    // attribute it to the fragment containing the paragraph's first line.
    if (lineIdx === -1) return fragmentLines.fromLine === 0;
    return lineIdx >= fragmentLines.fromLine && lineIdx < fragmentLines.toLine;
  };

  for (const [runIndex, run] of block.runs.entries()) {
    if (run.kind !== 'image') continue;
    const imgRun = run as ImageRun;

    if (!isFloatingImageRun(imgRun)) continue;
    if (!runInFragment(runIndex)) continue;

    const distTop = imgRun.distTop ?? 0;
    const distBottom = imgRun.distBottom ?? 0;
    const distLeft = imgRun.distLeft ?? 12;
    const distRight = imgRun.distRight ?? 12;
    const { x, y, side } = resolveAnchoredObjectPosition(imgRun, fragmentY, contentWidth, geometry);

    floatingImages.push({
      src: imgRun.src,
      width: imgRun.width,
      height: imgRun.height,
      alt: imgRun.alt,
      transform: imgRun.transform,
      side,
      x,
      y,
      distTop,
      distBottom,
      distLeft,
      distRight,
      pmStart: imgRun.pmStart,
      pmEnd: imgRun.pmEnd,
      wrapText: imageWrapTextFromCssFloat(imgRun.cssFloat),
      wrapType: imgRun.wrapType,
      behindDoc: imgRun.behindDoc,
      cropTop: imgRun.cropTop,
      cropRight: imgRun.cropRight,
      cropBottom: imgRun.cropBottom,
      cropLeft: imgRun.cropLeft,
      opacity: imgRun.opacity,
    });
  }

  return floatingImages;
}
