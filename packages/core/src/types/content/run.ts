/**
 * Run content (`w:r`) and the inline pieces that live inside a run —
 * text, tab, break, symbol, footnote/endnote references, field chars,
 * instruction text, soft/no-break hyphens, drawings, shapes.
 */

import type { TextFormatting } from '../formatting';
import type { Image } from './image';
import type { Shape } from './shape';
import type { RunPropertyChange } from './trackedChange';

/**
 * Plain text content
 */
export interface TextContent {
  type: 'text';
  /** The text string */
  text: string;
  /** Preserve whitespace (xml:space="preserve") */
  preserveSpace?: boolean;
}

/**
 * Tab character
 */
export interface TabContent {
  type: 'tab';
}

/**
 * Line break
 */
export interface BreakContent {
  type: 'break';
  /** Break type */
  breakType?: 'page' | 'column' | 'textWrapping';
  /** Clear type for text wrapping break */
  clear?: 'none' | 'left' | 'right' | 'all';
}

/**
 * Symbol character (special font character)
 */
export interface SymbolContent {
  type: 'symbol';
  /** Font name */
  font: string;
  /** Character code */
  char: string;
}

/**
 * Footnote or endnote reference
 */
export interface NoteReferenceContent {
  type: 'footnoteRef' | 'endnoteRef';
  /** Note ID */
  id: number;
}

/**
 * Field character (begin/separate/end)
 */
export interface FieldCharContent {
  type: 'fieldChar';
  /** Field character type */
  charType: 'begin' | 'separate' | 'end';
  /** Field is locked */
  fldLock?: boolean;
  /** Field is dirty (needs update) */
  dirty?: boolean;
}

/**
 * Field instruction text
 */
export interface InstrTextContent {
  type: 'instrText';
  /** Field instruction */
  text: string;
}

/**
 * Soft hyphen
 */
export interface SoftHyphenContent {
  type: 'softHyphen';
}

/**
 * Non-breaking hyphen
 */
export interface NoBreakHyphenContent {
  type: 'noBreakHyphen';
}

/**
 * Drawing/image reference
 */
export interface DrawingContent {
  type: 'drawing';
  /** Image data */
  image: Image;
}

/**
 * Shape reference
 */
export interface ShapeContent {
  type: 'shape';
  /** Shape data */
  shape: Shape;
}

/**
 * All possible run content types
 */
export type RunContent =
  | TextContent
  | TabContent
  | BreakContent
  | SymbolContent
  | NoteReferenceContent
  | FieldCharContent
  | InstrTextContent
  | SoftHyphenContent
  | NoBreakHyphenContent
  | DrawingContent
  | ShapeContent;

/**
 * A run is a contiguous region of text with the same formatting
 */
export interface Run {
  type: 'run';
  /** Text formatting properties */
  formatting?: TextFormatting;
  /** Run-level tracked property changes (w:rPrChange) */
  propertyChanges?: RunPropertyChange[];
  /** Run content (text, tabs, breaks, etc.) */
  content: RunContent[];
}
