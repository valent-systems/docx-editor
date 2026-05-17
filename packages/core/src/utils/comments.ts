/**
 * Framework-agnostic comment + tracked-change helpers shared by the
 * React and Vue adapters. The data shapes and string-formatting rules
 * here are part of the visible UI (avatar colors, date strings) so
 * keep this file as the single source of truth — both adapters import
 * from here. CSS-property factories live in adapter-specific files
 * (sidebar/cardUtils.ts in React, sidebar/sidebarUtils.ts in Vue).
 */
import type { Paragraph } from '../types/content';

/** Extract plain text from a Comment's paragraph content. */
export function getCommentText(paragraphs?: Paragraph[]): string {
  if (!paragraphs?.length) return '';
  return paragraphs
    .flatMap((p) =>
      p.content
        .filter((c) => c.type === 'run')
        .flatMap((r) => ('content' in r ? r.content : []))
        .filter((c) => c.type === 'text')
        .map((t) => ('text' in t ? t.text : ''))
    )
    .join('');
}

export function formatDate(dateStr?: string): string {
  if (!dateStr) return '';
  const d = new Date(dateStr);
  return d.toLocaleString(undefined, {
    hour: 'numeric',
    minute: '2-digit',
    month: 'short',
    day: 'numeric',
  });
}

export function getInitials(name: string): string {
  return name
    .split(/\s+/)
    .map((w) => w[0])
    .join('')
    .toUpperCase()
    .slice(0, 2);
}

// Kibana-style avatar palette — deterministic per author name.
const AVATAR_COLORS = [
  '#6DCCB1',
  '#79AAD9',
  '#EE789D',
  '#A987D1',
  '#E6A85F',
  '#F2CC8F',
  '#68B3A2',
  '#B07AA1',
  '#59A14F',
  '#FF9DA7',
  '#E15759',
  '#76B7B2',
];

export function getAvatarColor(name: string): string {
  let hash = 0;
  for (let i = 0; i < name.length; i++) {
    hash = name.charCodeAt(i) + ((hash << 5) - hash);
  }
  return AVATAR_COLORS[Math.abs(hash) % AVATAR_COLORS.length];
}

export function truncateText(text: string, maxLength = 50): string {
  return text.length > maxLength ? text.slice(0, maxLength) + '...' : text;
}

export interface TrackedChangeEntry {
  type: 'insertion' | 'deletion' | 'replacement';
  text: string;
  /** For replacements: the deleted text that was replaced */
  deletedText?: string;
  author: string;
  date?: string;
  from: number;
  to: number;
  revisionId: number;
  /** For replacements: the insertion part's revisionId (different from the deletion's) */
  insertionRevisionId?: number;
}
