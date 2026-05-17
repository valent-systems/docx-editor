/**
 * Structured Document Tags / content controls (`w:sdt`) — inline and
 * block variants, plus properties (alias, tag, lock, list items,
 * checkbox state) for the supported SDT types.
 */

import type { Run } from './run';
import type { Hyperlink, SimpleField, ComplexField } from './link';
import type { MathEquation } from './math';
import type { Paragraph } from './paragraph';
import type { Table } from './table';

/**
 * SDT type (content control type)
 */
export type SdtType =
  | 'richText'
  | 'plainText'
  | 'date'
  | 'dropdown'
  | 'comboBox'
  | 'checkbox'
  | 'picture'
  | 'buildingBlockGallery'
  | 'group'
  | 'unknown';

/**
 * SDT properties (w:sdtPr)
 */
export interface SdtProperties {
  /** SDT type */
  sdtType: SdtType;
  /** Alias (friendly name) */
  alias?: string;
  /** Tag (developer identifier) */
  tag?: string;
  /** Lock content editing */
  lock?: 'sdtLocked' | 'contentLocked' | 'sdtContentLocked' | 'unlocked';
  /** Placeholder text */
  placeholder?: string;
  /** Whether showing placeholder */
  showingPlaceholder?: boolean;
  /** Date format for date controls */
  dateFormat?: string;
  /** Dropdown/combobox list items */
  listItems?: { displayText: string; value: string }[];
  /** Checkbox checked state */
  checked?: boolean;
}

/**
 * Inline SDT (content control within a paragraph)
 */
export interface InlineSdt {
  type: 'inlineSdt';
  /** SDT properties */
  properties: SdtProperties;
  /**
   * Inline content held inside the control. OOXML allows runs,
   * hyperlinks, simple/complex fields, nested SDTs, and math at this
   * level; the renderer must descend into all of them so docProps-bound
   * fields and similar template content survive paged rendering.
   */
  content: (Run | Hyperlink | SimpleField | ComplexField | InlineSdt | MathEquation)[];
}

/**
 * Block-level SDT (content control wrapping paragraphs/tables)
 */
export interface BlockSdt {
  type: 'blockSdt';
  /** SDT properties */
  properties: SdtProperties;
  /** Block content inside the control */
  content: (Paragraph | Table)[];
}
