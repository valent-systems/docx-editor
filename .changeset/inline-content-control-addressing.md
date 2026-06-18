---
'@eigenpal/docx-editor-core': minor
---

Content-control addressing now covers inline controls. `findContentControls`, `findContentControl`, `setContentControlContent`, and `removeContentControl` discover and edit inline (`w:sdt`-in-paragraph) controls, including inside table cells — not just block-level controls in the document body. Results carry `kind`, `location`, and a structural `address`; pass `{ scope: 'all' }` to also reach headers and footers. The live-editor `DocxEditorRef` methods (React and Vue) gain the same inline support.
