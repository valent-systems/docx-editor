/**
 * Synthesize a `<w:sdtPr>` from a modeled {@link SdtProperties} projection.
 *
 * Content controls parsed from a `.docx` round-trip by replaying their
 * captured `rawPropertiesXml` verbatim. Controls created *programmatically*
 * (e.g. by the wrap-a-range API) have no captured raw XML, so the serializer
 * must build a sequence-valid `w:sdtPr` from the modeled fields instead. This
 * is the single source of truth for that synthesis, shared by the block
 * ({@link serializeBlockSdt}) and inline ({@link serializeInlineSdt}) paths so
 * the two cannot drift.
 *
 * Children are emitted in `CT_SdtPr` (ECMA-376 §17.5.2.38) sequence order —
 * `alias, tag, id, lock, placeholder, showingPlcHdr, <type marker>` — so strict
 * OOXML validators (and Word's open-time repair) accept the result. A
 * `richText` control emits no type marker (it is the spec default); an
 * `unknown`/`group`/gallery/citation type also emits none (there is no faithful
 * marker to synthesize, and emitting the wrong one would mislabel the control).
 */

import type { SdtProperties } from '../../types/document';
import { escapeXml } from './xmlUtils';

/** The type-marker element for a synthesized control, or `''` for none. */
function synthesizeTypeMarker(props: SdtProperties): string {
  switch (props.sdtType) {
    case 'plainText':
      return '<w:text/>';
    case 'date':
      return props.dateFormat
        ? `<w:date w:fullDate="${escapeXml(props.dateFormat)}"/>`
        : '<w:date/>';
    case 'dropDownList': {
      const items = (props.listItems ?? [])
        .map(
          (i) =>
            `<w:listItem w:displayText="${escapeXml(i.displayText)}" w:value="${escapeXml(i.value)}"/>`
        )
        .join('');
      return `<w:dropDownList>${items}</w:dropDownList>`;
    }
    case 'comboBox': {
      const items = (props.listItems ?? [])
        .map(
          (i) =>
            `<w:listItem w:displayText="${escapeXml(i.displayText)}" w:value="${escapeXml(i.value)}"/>`
        )
        .join('');
      return `<w:comboBox>${items}</w:comboBox>`;
    }
    case 'checkbox':
      return `<w14:checkbox><w14:checked w14:val="${props.checked ? '1' : '0'}"/></w14:checkbox>`;
    case 'picture':
      return '<w:picture/>';
    // richText (default), unknown, group, buildingBlockGallery, equation,
    // citation, bibliography → no synthesizable marker.
    default:
      return '';
  }
}

/**
 * Build a sequence-valid `<w:sdtPr>` string from the modeled projection.
 * Returns the full element including the `<w:sdtPr>` wrapper.
 */
export function synthesizeSdtPr(props: SdtProperties): string {
  const parts: string[] = [];
  if (props.alias) parts.push(`<w:alias w:val="${escapeXml(props.alias)}"/>`);
  if (props.tag) parts.push(`<w:tag w:val="${escapeXml(props.tag)}"/>`);
  if (props.id != null) parts.push(`<w:id w:val="${props.id}"/>`);
  if (props.lock && props.lock !== 'unlocked') parts.push(`<w:lock w:val="${props.lock}"/>`);
  if (props.placeholder)
    parts.push(
      `<w:placeholder><w:docPart w:val="${escapeXml(props.placeholder)}"/></w:placeholder>`
    );
  if (props.showingPlaceholder) parts.push('<w:showingPlcHdr/>');
  parts.push(synthesizeTypeMarker(props));
  return `<w:sdtPr>${parts.join('')}</w:sdtPr>`;
}
