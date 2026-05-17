/**
 * Framework-agnostic print helpers shared by the React and Vue
 * adapters. Lifted from packages/react/src/components/ui/PrintPreview.tsx
 * so both adapters use the same parsing / preview-window code path.
 *
 * The thin button component + the print-time CSS injection stay
 * adapter-local (they're framework-specific JSX/SFC bits); the data
 * helpers below are pure functions.
 */

export interface PrintOptions {
  includeHeaders?: boolean;
  includeFooters?: boolean;
  includePageNumbers?: boolean;
  pageRange?: { start: number; end: number } | null;
  scale?: number;
  printBackground?: boolean;
  margins?: 'default' | 'none' | 'minimum';
}

const DEFAULT_PRINT_OPTIONS: PrintOptions = {
  includeHeaders: true,
  includeFooters: true,
  includePageNumbers: true,
  pageRange: null,
  scale: 1.0,
  printBackground: true,
  margins: 'default',
};

export function getDefaultPrintOptions(): PrintOptions {
  return { ...DEFAULT_PRINT_OPTIONS };
}

/** Trigger browser print dialog for the current document. */
export function triggerPrint(): void {
  if (typeof window !== 'undefined') window.print();
}

/** Open a new window with print-optimised body content. */
export function openPrintWindow(title: string, content: string): Window | null {
  if (typeof window === 'undefined') return null;
  const w = window.open('', '_blank');
  if (!w) return null;
  w.document.write(
    `<!DOCTYPE html><html><head><title>${title}</title>` +
      `<style>@media print { body { margin: 0; padding: 0; } @page { margin: 0; } }</style>` +
      `</head><body>${content}</body></html>`
  );
  w.document.close();
  return w;
}

/** Parse "1", "1-5", etc. into a page range, or null on invalid. */
export function parsePageRange(
  input: string,
  maxPages: number
): { start: number; end: number } | null {
  if (!input || !input.trim()) return null;
  const t = input.trim();
  if (/^\d+$/.test(t)) {
    const p = parseInt(t, 10);
    return p >= 1 && p <= maxPages ? { start: p, end: p } : null;
  }
  const m = t.match(/^(\d+)-(\d+)$/);
  if (m) {
    const start = parseInt(m[1], 10);
    const end = parseInt(m[2], 10);
    if (start >= 1 && end <= maxPages && start <= end) return { start, end };
  }
  return null;
}

export function formatPageRange(
  range: { start: number; end: number } | null,
  totalPages: number
): string {
  if (!range) return `All (${totalPages} pages)`;
  if (range.start === range.end) return `Page ${range.start}`;
  return `Pages ${range.start}-${range.end}`;
}

export function isPrintSupported(): boolean {
  return typeof window !== 'undefined' && typeof window.print === 'function';
}
