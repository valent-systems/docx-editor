---
'@sqren/docx-editor-core': minor
'@sqren/docx-editor-react': minor
'@sqren/docx-editor-vue': minor
---

Create content controls and address them anywhere. New `wrapInlineContentControl` (headless) and `wrapContentControl` (editor ref) plant an inline content control with a stable tag around an occurrence-precise placeholder span — including inside table cells and mid-sentence. Discovery (`findContentControls`/`getContentControls`) and editing (`setContentControlContent`/`removeContentControl`) now reach inline controls and controls in table cells and headers/footers, not just block controls in the body; `ContentControlInfo` gains `kind` and `container`. Added `fillContentControl` returning a success/blocked result instead of throwing. Programmatically created controls serialize a synthesized `w:sdtPr` so they round-trip through `.docx` and Word.
