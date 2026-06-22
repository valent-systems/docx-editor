---
'@eigenpal/docx-editor-core': patch
---

Render and round-trip text boxes anchored inside table cells. A text box anchored from a run inside a table cell was dropped — it didn't appear in the editor and was lost on save. It now parses, renders in-flow inside the cell, and survives both a headless and an in-editor save.
