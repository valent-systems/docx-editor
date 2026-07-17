import { useLayoutEffect, useRef, type RefObject } from 'react';

/**
 * Keep a horizontally-overflowing scroll container centered on its content.
 *
 * The editor enforces a minimum layout width (page + outline/ruler gutters +
 * comments-sidebar reservation); in a host pane narrower than that, the scroll
 * container h-scrolls — but would land at scrollLeft 0, showing the left gutter
 * with the page shoved off-view. Land (and keep) the horizontal scroll centered
 * so the page sits mid-view — Google Docs behavior — re-centering on pane /
 * content resize until the user h-scrolls deliberately. Vertical scroll is
 * never touched.
 */
export function useCenteredHorizontalScroll(
  scrollContainerRef: RefObject<HTMLDivElement | null>,
  deps: { minLayoutWidth: number; isLoading: boolean }
): void {
  const userHScrolledRef = useRef(false);
  const { minLayoutWidth, isLoading } = deps;

  useLayoutEffect(() => {
    const el = scrollContainerRef.current;
    if (!el) return;
    const center = () => {
      if (userHScrolledRef.current) return;
      const overflow = el.scrollWidth - el.clientWidth;
      if (overflow > 1 && Math.abs(el.scrollLeft - overflow / 2) > 1) {
        el.scrollLeft = overflow / 2;
      }
    };
    center();
    const ro = new ResizeObserver(center);
    ro.observe(el);
    for (const child of el.children) ro.observe(child);
    const onWheel = (e: WheelEvent) => {
      if (Math.abs(e.deltaX) > Math.abs(e.deltaY)) userHScrolledRef.current = true;
    };
    el.addEventListener('wheel', onWheel, { passive: true });
    return () => {
      ro.disconnect();
      el.removeEventListener('wheel', onWheel);
    };
  }, [scrollContainerRef, minLayoutWidth, isLoading]);
}
