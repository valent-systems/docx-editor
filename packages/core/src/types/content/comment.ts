/**
 * Comments (`w:comment` in `comments.xml`) and the inline range markers
 * (`w:commentRangeStart`/`End`) that anchor them inside paragraphs.
 */

import type { Paragraph } from './paragraph';

/**
 * A comment (w:comment) from comments.xml
 */
export interface Comment {
  /** Comment ID (matches commentRangeStart/End) */
  id: number;
  /** Author name */
  author: string;
  /** Author initials */
  initials?: string;
  /** Date */
  date?: string;
  /** Comment content (paragraphs) */
  content: Paragraph[];
  /** Parent comment ID (for replies) */
  parentId?: number;
  /** Whether the comment is resolved/done */
  done?: boolean;
}

/**
 * Comment range start marker in paragraph content
 */
export interface CommentRangeStart {
  type: 'commentRangeStart';
  id: number;
}

/**
 * Comment range end marker in paragraph content
 */
export interface CommentRangeEnd {
  type: 'commentRangeEnd';
  id: number;
}
