/**
 * Block-level Structured Document Tag (`w:sdt`) serializer.
 *
 * Round-trips a {@link BlockSdt} by replaying the captured `w:sdtPr` /
 * `w:sdtEndPr` verbatim (preserving the `CT_SdtPr` sequence order and any
 * unmodeled features such as data binding, `w14:`/`w15:` extensions, and
 * `@lastValue`), then serializing the child blocks inside `w:sdtContent`.
 *
 * A control parsed from a `.docx` round-trips by replaying its captured
 * `rawPropertiesXml` verbatim — preserving `CT_SdtPr` sequence order and any
 * unmodeled features (data binding, `w14:`/`w15:` extensions, `@lastValue`).
 * A control *created* programmatically (the wrap-a-range API) has no captured
 * raw XML, so its `w:sdtPr` is synthesized from the modeled projection via the
 * shared {@link synthesizeSdtPr} (the same synthesizer the inline serializer
 * uses, so block and inline created controls cannot drift).
 *
 * Shared by the body serializer and the header/footer serializer so block
 * SDTs round-trip identically wherever block content can appear.
 */

import type { BlockContent, BlockSdt } from '../../types/document';
import { synthesizeSdtPr } from './sdtPrSynth';

/**
 * Serialize a {@link BlockSdt} to a `<w:sdt>` element.
 *
 * @param block - the block SDT to serialize
 * @param serializeChild - serializer for a single child block (lets the body
 *   and header/footer paths recurse with their own block dispatcher)
 */
export function serializeBlockSdt(
  block: BlockSdt,
  serializeChild: (child: BlockContent) => string
): string {
  const sdtContentXml = block.content.map(serializeChild).join('');
  // Replay the captured properties verbatim for a parsed control; synthesize a
  // sequence-valid `w:sdtPr` from the modeled fields for a created one (no raw).
  const sdtPrXml = block.properties.rawPropertiesXml ?? synthesizeSdtPr(block.properties);
  const sdtEndPrXml = block.properties.rawEndPropertiesXml ?? '';
  return `<w:sdt>${sdtPrXml}${sdtEndPrXml}<w:sdtContent>${sdtContentXml}</w:sdtContent></w:sdt>`;
}
