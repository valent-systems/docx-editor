import { describe, test, expect } from 'bun:test';
import type { Node as PMNode } from 'prosemirror-model';

import { parseDocumentBody } from '../documentParser';
import { serializeDocumentBody } from '../serializer/documentSerializer';
import { toProseDoc } from '../../prosemirror/conversion/toProseDoc';
import { fromProseDoc } from '../../prosemirror/conversion/fromProseDoc';
import { toFlowBlocks } from '../../layout-bridge/toFlowBlocks';
import type { BlockSdt, Document, DocumentBody, Table, TableCell } from '../../types/document';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const W15 = 'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"';

/**
 * A block-level content control (`w:sdt`) nested directly inside a table cell —
 * the CLM OptionSet shape: the option paragraphs of one cell are wrapped in a
 * single control so picking a value replaces the whole region. Before this work
 * `parseCellContent` dropped the `w:sdt` (and its content) entirely. These tests
 * pin that the control + its paragraphs now survive parse → model → serialize,
 * and flow through the model → PM → layout pipeline so the cell renders and the
 * control is addressable by tag for in-place fill.
 */
const CELL_SDT = `<w:document ${W} ${W15}><w:body>
  <w:tbl>
    <w:tr>
      <w:tc>
        <w:sdt>
          <w:sdtPr><w:alias w:val="Renewal mechanism"/><w:tag w:val="renewal"/><w:id w:val="1"/></w:sdtPr>
          <w:sdtContent>
            <w:p><w:r><w:t>Automatic renewal</w:t></w:r></w:p>
            <w:p><w:r><w:t>Automatic termination</w:t></w:r></w:p>
            <w:p><w:r><w:t>CCEP choice</w:t></w:r></w:p>
          </w:sdtContent>
        </w:sdt>
      </w:tc>
    </w:tr>
  </w:tbl>
</w:body></w:document>`;

function firstCell(body: DocumentBody): TableCell {
  const table = body.content[0] as Table;
  expect(table.type).toBe('table');
  const cell = table.rows[0]!.cells[0]!;
  expect(cell.type).toBe('tableCell');
  return cell;
}

function firstCellSdt(body: DocumentBody): BlockSdt {
  const cell = firstCell(body);
  const sdt = cell.content[0];
  if (sdt?.type !== 'blockSdt') {
    throw new Error(`expected cell content[0] to be blockSdt, got ${sdt?.type}`);
  }
  return sdt;
}

describe('block SDT inside a table cell — DOCX parse + serialize', () => {
  test('parses the cell `w:sdt` into a BlockSdt holding the option paragraphs (not dropped)', () => {
    const body = parseDocumentBody(CELL_SDT);
    const cell = firstCell(body);
    expect(cell.content).toHaveLength(1);
    const sdt = firstCellSdt(body);
    expect(sdt.properties.tag).toBe('renewal');
    expect(sdt.properties.alias).toBe('Renewal mechanism');
    expect(sdt.content).toHaveLength(3);
    expect(sdt.content.every((b) => b.type === 'paragraph')).toBe(true);
  });

  test('serializes the cell BlockSdt back to a `<w:sdt>` inside the `<w:tc>` losslessly', () => {
    const body = parseDocumentBody(CELL_SDT);
    const xml = serializeDocumentBody(body);
    expect(xml).toContain('<w:sdt>');
    expect(xml).toContain('<w:tag w:val="renewal"');
    expect(xml).toContain('Automatic renewal');
    expect(xml).toContain('CCEP choice');
    // The control is inside the cell, and re-parses to the same shape.
    const sdt = firstCellSdt(
      parseDocumentBody(`<w:document ${W} ${W15}><w:body>${xml}</w:body></w:document>`)
    );
    expect(sdt.properties.tag).toBe('renewal');
    expect(sdt.content).toHaveLength(3);
  });
});

describe('block SDT inside a table cell — model → PM → model + layout', () => {
  const docFrom = (body: DocumentBody): Document => ({
    package: { document: { content: body.content } },
  });

  test('model → PM emits a `blockSdt` PM node inside the `tableCell` (so it renders + is fillable)', () => {
    const pm = toProseDoc(docFrom(parseDocumentBody(CELL_SDT)));
    let cellSdt: PMNode | null = null;
    pm.descendants((node) => {
      if (node.type.name === 'tableCell') {
        node.forEach((child) => {
          if (child.type.name === 'blockSdt') cellSdt = child;
        });
      }
      return true;
    });
    expect(cellSdt).not.toBeNull();
    expect(cellSdt!.attrs.tag).toBe('renewal');
    expect(cellSdt!.childCount).toBe(3);
    expect(cellSdt!.child(0).textContent).toBe('Automatic renewal');
  });

  test('PM → model preserves the cell BlockSdt (full round-trip)', () => {
    const back = fromProseDoc(toProseDoc(docFrom(parseDocumentBody(CELL_SDT))));
    const sdt = firstCellSdt({ content: back.package.document.content } as DocumentBody);
    expect(sdt.properties.tag).toBe('renewal');
    expect(sdt.content).toHaveLength(3);
  });

  test('layout flattens the cell control into 3 cell flow blocks, each tagged with the SDT group', () => {
    const pm = toProseDoc(docFrom(parseDocumentBody(CELL_SDT)));
    const flow = toFlowBlocks(pm);
    const table = flow.find((b) => 'kind' in b && b.kind === 'table') as
      | {
          kind: 'table';
          rows: Array<{ cells: Array<{ blocks: Array<{ sdtGroups?: Array<{ tag?: string }> }> }> }>;
        }
      | undefined;
    expect(table).toBeDefined();
    const cellBlocks = table!.rows[0]!.cells[0]!.blocks;
    expect(cellBlocks).toHaveLength(3);
    // Every option paragraph renders as an ordinary cell block AND carries the
    // control's group membership (so the in-cell boundary box can find it).
    expect(cellBlocks.every((b) => (b.sdtGroups ?? []).some((g) => g.tag === 'renewal'))).toBe(
      true
    );
  });
});
