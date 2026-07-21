# Incremental layout: typing-latency architecture

Status: proposal (v0). Owner: editor engine.

## Problem

Keystroke-to-paint latency on large documents misses the interaction budget.
Measured on the 96-page DRC corpus file (`drc-qualification-v2.docx`, dev
build, M-series MacBook):

| stage                               | per keystroke |
| ----------------------------------- | ------------- |
| `toFlowBlocks` (PM doc → blocks)    | ~7ms          |
| `measureBlocks`                     | ~16ms         |
| `layoutDocument` (pagination)       | ~5ms          |
| paint (`renderPages`, incremental)  | ~11ms         |
| React state fan-out + overlays      | ~50ms         |
| **total (after the 2026-07 fixes)** | **60–130ms**  |

Two earlier per-keystroke costs are already fixed and excluded from the table:
the history deep-compare (`JSON.stringify` of the full model, ~200ms on
image-heavy files) and the synchronous `getDocument()` conversion (30–60ms,
now coalesced to a 300ms idle window by `useDeferredDocumentNotify`).

Targets, per Google's RAIL model for continuous input:

- **Acceptable standard: <50ms** worst case on the largest realistic document.
- **Goal: 16ms** (one 60fps frame) for the visible response to a keystroke.

## What is architecturally wrong

The engine is a **batch compiler run on every edit**. Each docChanged
transaction recomputes PM doc → blocks → measures → pages from scratch. The
caches that exist reduce constants, not asymptotics — every stage remains
O(document) per keystroke, so latency scales with document size and no amount
of stage-level tuning reaches a fixed frame budget.

Concretely:

1. **Identity churn.** `toFlowBlocks` allocates fresh block objects with fresh
   ids (`nextBlockId()` counter) on every pass. Nothing downstream can
   memoize by block identity because identity never survives a pass.
2. **Absolute positions baked into blocks.** Every block and run carries
   `pmStart`/`pmEnd`. A one-character edit shifts positions for the entire
   rest of the document, so even byte-identical trailing content differs
   from last pass's output.
3. **Discarded change information.** PM transactions describe exactly what
   changed (`tr.mapping`, step ranges). The pipeline ignores this and
   re-derives everything.
4. **Content-hash caching pays O(doc) anyway.** The paragraph measure cache
   is keyed by a string hash of runs + attrs; computing the keys means
   walking every run of every paragraph, every keystroke.
5. **Cache bypasses on float-heavy documents.** Paragraphs affected by
   floating-object zones skip the measure cache entirely (context-dependent
   measurement), and text-box inner paragraphs are never cached. The DRC file
   (63 anchored images, 30 text boxes) is exactly this shape.
6. **Pagination always restarts from page 1** and rebuilds table row-break
   geometry (`buildTableRowBreakInfo`, `layoutCellContent` per covering cell
   per row) for every table on every pass.
7. **React as the fan-out bus.** `setBlocks`/`setMeasures`/`setLayout` run per
   keystroke; overlays, sidebar anchors, and decoration layers re-render on
   every pass regardless of whether their inputs changed. This is the single
   largest line item (~50ms dev; smaller in prod but still O(doc)
   reconciliation).

What is already right and should be kept: rAF transaction coalescing
(`createLayoutScheduler`), the incremental painter (dirty-page DOM reuse),
the content-hash measure cache as a correctness baseline, and the deferred
document-change notification.

## Target architecture

Make edit cost proportional to what changed, then split remaining work across
the frame boundary. Five pillars:

### 1. Stable block identity (enabler for everything else)

Key converted blocks by PM node identity: PM documents are persistent trees,
so unchanged nodes keep reference identity across transactions.
`WeakMap<PMNode, ConvertedBlock>` at the top-level walk; a cache hit reuses
the block object (same id, same run objects). Sequential context that
conversion depends on (list counter state, theme, section geometry,
`pageContentHeight`) becomes an explicit, comparable input: snapshot the
small counter vector per block and reuse only when the incoming context
matches. Fallback on mismatch is plain reconversion — correctness never
depends on the cache.

Positions move out of blocks: a separate `positions: Int32Array` (or a
per-block `pmStart` map) is recomputed each pass by one O(n) integer walk
over node sizes — no allocation, no object churn. Downstream consumers that
need `pmStart` read it from the index. (Interim step, if run positions prove
hard to externalize: keep positions in blocks but patch them in place on
cached hits — mutation is safe because the cache owns the objects.)

### 2. Identity-keyed measurement

With stable block identity, the measure cache becomes
`WeakMap<Block, {width, zonesSig, measure}>` — a pointer compare instead of
string hashing. The content hash stays as the fallback tier (survives
identity misses after e.g. undo). Float-zone-affected paragraphs stop
bypassing the cache: key on `(block, zoneSignature, yBucket)` where
`zoneSignature` is derived from the float blocks (themselves
identity-stable). Text-box inner measures get the same treatment.

### 3. Resumable pagination

Keep the previous `Layout`. On edit, find the first dirty block (first
identity miss), start pagination from the page/column state that block
occupied last pass, and stop as soon as the emitted fragment stream
re-synchronizes with the previous layout (same block id opens a page at the
same y with the same paginator state → splice the remaining pages
unchanged). Typical edit touches 1–3 pages regardless of document length.
`buildTableRowBreakInfo` caches by `(block, measure)` identity — stable
under pillars 1–2.

### 4. Store-based fan-out

Layout results go into a subscription store (plain emitter, not React
state). Overlays, sidebar anchors, and decoration layers subscribe with a
dirty-page filter and update imperatively or via fine-grained selectors.
React state updates (`blocks`/`measures`/`layout` for components that truly
need them) coalesce to the same idle window as the document notification.
The painter already works this way; this pillar extends the pattern to the
rest of the view layer.

### 5. Frame-split scheduling (the 16ms step)

Even a perfectly incremental pipeline occasionally owes >16ms (edit inside a
490-row table). Split the work at the frame boundary:

- **Sync (this frame):** re-layout and repaint only the page containing the
  caret. The typed character appears on the next frame.
- **Deferred (idle/next frames):** repaginate forward from the dirty page,
  splice, repaint any pages that moved.

This is visibly what Google Docs does — the current line responds instantly;
downstream reflow settles behind it. It converts worst-case latency into
worst-case _settle time_, which RAIL permits.

## Correctness strategy

Incremental output must be bit-identical to batch output.

- **Dual-run verification:** dev flag + e2e mode that runs both pipelines per
  keystroke and asserts deep-equal layouts. The pagination-parity corpus
  (tpx / pws / drc) runs under this mode in CI.
- **Fallback:** any context mismatch, cache anomaly, or dual-run divergence
  drops the pass to the batch path (the batch code never leaves the tree).
- **Property tests:** randomized edit scripts (insert/delete/split/merge at
  random positions, including inside tables and text boxes) comparing
  incremental vs batch.

## Milestones

Each ships independently, gated on the dual-run harness staying green and the
typing-latency probe (keystroke timing on the DRC corpus file) improving.

| milestone                                   | expected effect (DRC file)                 |
| ------------------------------------------- | ------------------------------------------ |
| M1: stable identity + incremental convert   | convert ~7ms → <1ms; enables M2+           |
| M2: identity-keyed measure (+float/textbox) | measure ~16ms → ~1ms                       |
| M3: resumable pagination + table-info cache | layout ~5ms → ~1ms; stable at any doc size |
| M4: store fan-out, overlay dirty filtering  | React ~50ms → ~5-10ms                      |
| M5: frame-split scheduling                  | visible response ≤16ms always              |

M1–M3 land in core (shared with the Vue adapter for free). M4 is per
adapter. M5 is core scheduling + adapter paint hooks.

Latency probe: `e2e` keystroke timing (see the 2026-07 session's
`type-freeze` probe pattern — click into body text and a large table on the
DRC corpus file, measure keydown → responsive). Add it as a checked-in perf
spec with budget assertions once M1 lands, so regressions fail CI rather
than reaching users.

## Non-goals

- Web worker / OffscreenCanvas measurement: revisit only if M1–M5 miss the
  budget; workers add serialization costs that identity-based caching avoids.
- Virtualized painting (only mounting visible pages): the painter's
  incremental DOM reuse already bounds paint cost; full virtualization is a
  scroll-perf project, not a typing-latency one.
