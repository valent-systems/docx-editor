/**
 * Sidebar card chrome — shared between React and Vue. Numeric
 * pixel values rather than `'8px'` strings so both adapters'
 * CSSProperties shapes accept them. Lifted from
 * `packages/react/src/components/sidebar/cardStyles.ts` and the
 * Vue mirror so there's one canonical table.
 */
import type { CSSProperties } from './cssTypes';

export const CARD_STYLE_COLLAPSED: CSSProperties = {
  padding: '8px 10px',
  borderRadius: 8,
  backgroundColor: '#f8fbff',
  cursor: 'pointer',
  boxShadow: '0 1px 3px rgba(60,64,67,0.2), 0 2px 6px rgba(60,64,67,0.08)',
};

export const CARD_STYLE_EXPANDED: CSSProperties = {
  padding: '10px 12px',
  borderRadius: 8,
  backgroundColor: '#fff',
  cursor: 'pointer',
  boxShadow: '0 1px 3px rgba(60,64,67,0.3), 0 4px 8px 3px rgba(60,64,67,0.15)',
};
