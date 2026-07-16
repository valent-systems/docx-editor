/**
 * Pure ref-API query helpers — read-only inspectors over the PM document
 * and the paginated layout. Used by DocxEditor's `defineExpose` ref API
 * (`findInDocument`, `getSelectionInfo`, `getPageContent`).
 *
 * Lifted to `@valent/docx-editor-core/prosemirror/queries` and shared
 * with the React adapter; re-exported here to keep existing import sites stable.
 */

export {
  findInDocument,
  getSelectionInfo,
  getPageContent,
} from '@valent/docx-editor-core/prosemirror/queries';
export type {
  FindInDocumentMatch,
  SelectionInfo,
  PageContent,
} from '@valent/docx-editor-core/prosemirror/queries';
