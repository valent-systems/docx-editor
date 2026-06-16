# Content controls (SDTs)

Word **content controls** — `w:sdt` (Structured Document Tags) — are labeled,
bounded regions of a document. Block-level controls wrap one or more paragraphs
or a table and carry a stable `tag`, `alias`, and `id`. They are the natural
anchor for templates, conditional sections, and document automation.

The editor parses block controls into the document model, keeps them editable,
renders their boundary, and round-trips them losslessly on save (including
unmodeled properties such as `w:dataBinding` and `w15:repeatingSection`, which
are preserved verbatim). This guide covers the APIs for finding and editing them.

## Headless (server-side / AI pipelines)

Operate directly on a parsed `Document` with no editor or DOM. Import from
`@sqren/docx-editor-core/headless` (or `/agent`).

```ts
import {
  parseDocx,
  wrapInlineContentControl,
  findContentControls,
  findContentControl,
  setContentControlContent,
  fillContentControl,
  removeContentControl,
  serializeDocx, // or createDocx for bytes
} from '@sqren/docx-editor-core/headless';

const doc = await parseDocx(buffer);

// Create: wrap an occurrence-precise placeholder span in a new inline control
// with a stable tag. `occurrence` (+ optional `paraId` scope) disambiguates
// identical text, so two identical clauses get independent controls.
const r = wrapInlineContentControl(doc, { text: '[volume]' }, { tag: 'volume' });
//   { status: 'wrapped'; doc; tag } | 'not-found' | 'occurrence-out-of-range' | 'crosses-inline-boundary'

// Discover controls (filter by tag / alias / id / type) — block AND inline,
// in the body, table cells, and headers/footers.
const controls = findContentControls(doc); // ContentControlInfo[]
const intro = findContentControl(doc, { tag: 'intro' });
//   { tag, alias, id, sdtType, lock, listItems, placeholder, text, kind, container, path, depth }

// Fill a control by tag (string → paragraphs for a block control, one run for
// an inline control). Throwing variant, plus a result-returning wrapper:
let next = setContentControlContent(doc, { tag: 'intro' }, 'Filled by template');
const filled = fillContentControl(doc, { tag: 'volume' }, '500,000 units');
//   { status: 'filled'; doc } | 'not-found' | 'locked' | 'typed' | 'data-bound'

// Conditional sections: drop a control, or unwrap it (keep its content)
next = removeContentControl(next, { tag: 'optionalClause' });
next = removeContentControl(next, { tag: 'wrapper' }, { keepContent: true });
```

All mutators are **pure** (they return a new `Document`) and preserve the
control's identity and raw `w:sdtPr`, so the result still round-trips. Block
input is deep-cloned, so you can safely reuse the array you pass in.

Safety rules (all overridable with `{ force: true }`):

- A **content-locked** control (`contentLocked` / `sdtContentLocked`) throws
  `ContentControlLockedError` on edit; a **deletion-locked** one
  (`sdtLocked` / `sdtContentLocked`) throws on remove.
- A **typed** control (dropdown, date, checkbox, picture, group) throws
  `ContentControlTypeError` on free-text replacement — its value shouldn't be
  set as arbitrary text. Only `richText` / `plainText` controls accept it.
- A **data-bound** control (`w:dataBinding`) throws `ContentControlBoundError`:
  its content is driven by the Custom XML store, so a direct write would not
  persist in Word. Update the store instead.
- **Unwrapping a repeating-section** control is refused (it would orphan the
  w15 structure).
- Filling a control that was **showing its placeholder** (`w:showingPlcHdr`)
  clears that flag, so Word doesn't render real content as placeholder text.
- A `plainText` control stays a single paragraph (newlines are not split into
  paragraphs, which Word would repair).

An unmatched filter throws `ContentControlNotFoundError`.

**Mutators affect the first match only.** When a tag repeats (common in
templates — e.g. a `date` control in the header and the body), enumerate with
`findContentControls` and disambiguate by `id`, then edit each.

## Editor ref (React / Vue)

The same operations are on `DocxEditorRef`, running against the live editor so
writes are normal undoable edits.

```ts
const ref = useRef<DocxEditorRef>(null);

ref.current?.getContentControls({ type: 'dropDownList' }); // PMContentControl[]
ref.current?.scrollToContentControl({ tag: 'intro' });
ref.current?.setContentControlContent({ tag: 'intro' }, 'Filled'); // → boolean
ref.current?.removeContentControl({ tag: 'optionalClause' });
```

These return `false` when no control matches; a locked or typed control throws
(same rules as headless) unless `{ force: true }` is passed. The editor ref's
`setContentControlContent` takes a **string** today (newlines become
paragraphs); the headless API takes string or `BlockContent[]`.

## Typed value setters (dropdown / checkbox / date)

For typed controls, use `setContentControlValue` instead of
`setContentControlContent` (which refuses them). It updates both the visible
content and the structured `w:sdtPr` state.

```ts
// headless
setContentControlValue(doc, { tag: 'status' }, { kind: 'dropdown', value: '2' });
setContentControlValue(doc, { tag: 'agree' }, { kind: 'checkbox', checked: true });
setContentControlValue(doc, { tag: 'effective' }, { kind: 'date', date: '2026-06-01' });

// editor ref (React / Vue)
ref.current?.setContentControlValue({ tag: 'status' }, { kind: 'dropdown', value: '2' });
```

A dropdown value must match one of the control's list items (by value or
display text); a date is ISO `yyyy-mm-dd` and is formatted with the control's
`w:dateFormat`. A data-bound control (`w:dataBinding`) is refused (its value is
driven by the Custom XML store); a content-locked one is refused unless forced.
**comboBox controls are treated as pick-only** — typing a value outside the
list is not supported yet.

### Interactive UI

In the editor, typed controls render a small trigger at the top-right of their
box (revealed on hover/focus). Clicking it toggles a checkbox, opens a menu of
the dropdown's items, or opens a date picker — each runs through
`setContentControlValue` as a normal undoable edit. Available in both the React
and Vue adapters; no wiring required. The trigger is a focusable button and the
dropdown menu is arrow-key navigable (Enter to choose, Escape to close). Data-
bound and content-locked controls don't render a trigger.

## Reading bound / typed state

Both `ContentControlInfo` (headless) and `PMContentControl` (editor) expose
`showingPlaceholder`, `checked`, `dateFormat`, `listItems`, and `dataBinding`
(`{ xpath, storeItemID, prefixMappings }`) so automation can inspect a
control's state before deciding what to write. `dataBinding` and
`w15:repeatingSection` round-trip verbatim but are **not** modeled as live
behavior (no bound-value resolution, no repeat expansion).

When `showingPlaceholder` is `true`, the control's `text` is the **placeholder
boilerplate** (e.g. "Click here to enter text"), not entered content — check
this flag before treating `text` as real data.

## Notes and limits

- **Addressing is by `tag`/`alias`/`id`** — content controls are not the same
  as the `{{mustache}}` template variables used by the docxtemplater plugin.
- Locks: `contentLocked`/`sdtContentLocked` block content edits;
  `sdtLocked`/`sdtContentLocked` block removal.
- **Discovery and edit reach both block and inline controls, wherever they
  live** — document body, **table cells**, **headers/footers**, and
  mid-paragraph (inline). `ContentControlInfo` carries `kind` (`block`/`inline`)
  and `container` (`body`/`header`/`footer`). `setContentControlContent` /
  `removeContentControl` / `fillContentControl` operate on the first match in
  document order across all of these.
- **Creating controls.** {@link wrapInlineContentControl} (headless) and the
  editor ref's `wrapContentControl` plant a new inline control around an
  occurrence-precise placeholder span. A created control has no captured
  `rawPropertiesXml`, so the serializer synthesizes a sequence-valid `w:sdtPr`
  from its `tag`/`id`/`sdtType` — it round-trips through `.docx` and Word.
- **Cell-level controls are inline.** A control inside a table cell is planted
  as an _inline_ `w:sdt` within the cell's paragraph (a cell holds paragraphs;
  paragraphs hold inline content). The model does **not** add a block `w:sdt`
  directly under `w:tc` — `TableCell.content` is unchanged.
- Typed-control value setters, live `dataBinding`, and `repeatingSection`
  expansion remain as documented above / on the roadmap.
