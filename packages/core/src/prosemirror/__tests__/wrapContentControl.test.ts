/**
 * Live-editor (PM) content-control creation: occurrence-precise text search,
 * wrapping a range in a new inline `sdt`, and filling it — the editor-side
 * mirror of the headless wrap/fill APIs.
 */

import { describe, expect, test } from 'bun:test';
import { EditorState } from 'prosemirror-state';

import { singletonManager } from '../schema';
import {
  findContentControlsInPM,
  setContentControlContentTr,
  wrapRangeInContentControlTr,
  wrapContentControlByTextTr,
} from '../contentControls';
import { enumerateMatches, resolveOccurrence } from '../occurrenceSearch';

const schema = singletonManager.getSchema();

function para(text: string, paraId?: string) {
  return schema.nodes.paragraph.create(paraId ? { paraId } : null, text ? schema.text(text) : null);
}
function stateOf(...paragraphs: ReturnType<typeof para>[]) {
  return EditorState.create({ schema, doc: schema.nodes.doc.create(null, paragraphs) });
}

describe('occurrence search', () => {
  test('enumerates every match in reading order with PM ranges', () => {
    const state = stateOf(para('from [X] to [X]', 'P1'), para('also [X]', 'P2'));
    const matches = enumerateMatches(state.doc, '[X]');
    expect(matches.length).toBe(3);
    expect(matches.map((m) => m.paraId)).toEqual(['P1', 'P1', 'P2']);
    expect(matches[0].occurrenceInPara).toBe(0);
    expect(matches[1].occurrenceInPara).toBe(1);
    // Each range renders back to the needle.
    for (const m of matches) expect(state.doc.textBetween(m.from, m.to, '\n', '￼')).toBe('[X]');
  });

  test('resolveOccurrence scopes by paraId and never retargets out-of-range', () => {
    const state = stateOf(para('a [X] b [X]', 'P1'), para('c [X]', 'P2'));
    expect(resolveOccurrence(state.doc, { text: '[X]', paraId: 'P2', occurrence: 0 })?.paraId).toBe(
      'P2'
    );
    expect(resolveOccurrence(state.doc, { text: '[X]', occurrence: 1 })?.paraId).toBe('P1');
    expect(resolveOccurrence(state.doc, { text: '[X]', paraId: 'P2', occurrence: 3 })).toBeNull();
    expect(resolveOccurrence(state.doc, { text: '[missing]' })).toBeNull();
  });
});

describe('wrapContentControlByTextTr', () => {
  test('wraps the chosen occurrence in an inline control and fills it independently', () => {
    let state = stateOf(para('from [X] to [X]', 'P1'));
    // Wrap occurrence 1 (the second [X]).
    const tr = wrapContentControlByTextTr(
      state,
      { text: '[X]', occurrence: 1 },
      {
        sdtType: 'richText',
        tag: 'second',
      }
    );
    expect(tr).not.toBeNull();
    state = state.apply(tr!);

    const controls = findContentControlsInPM(state.doc, { tag: 'second' });
    expect(controls.length).toBe(1);
    expect(controls[0].text).toBe('[X]');

    // Filling by tag edits only the wrapped occurrence.
    state = state.apply(setContentControlContentTr(state, { tag: 'second' }, 'END'));
    expect(state.doc.textContent).toBe('from [X] to END');
  });

  test('two identical spans in different paragraphs wrap to independent controls', () => {
    let state = stateOf(para('gov [LIST]', 'G'), para('jur [LIST]', 'J'));
    state = state.apply(
      wrapContentControlByTextTr(
        state,
        { text: '[LIST]', occurrence: 0 },
        {
          sdtType: 'richText',
          tag: 'gov',
        }
      )!
    );
    state = state.apply(
      wrapContentControlByTextTr(
        state,
        { text: '[LIST]', occurrence: 1 },
        {
          sdtType: 'richText',
          tag: 'jur',
        }
      )!
    );
    expect(
      findContentControlsInPM(state.doc)
        .map((c) => c.tag)
        .sort()
    ).toEqual(['gov', 'jur']);

    state = state.apply(setContentControlContentTr(state, { tag: 'gov' }, 'France'));
    state = state.apply(setContentControlContentTr(state, { tag: 'jur' }, 'Germany'));
    expect(state.doc.textContent).toBe('gov Francejur Germany');
  });

  test('returns null for an unresolved occurrence', () => {
    const state = stateOf(para('only [X]', 'P1'));
    expect(
      wrapContentControlByTextTr(
        state,
        { text: '[X]', occurrence: 9 },
        {
          sdtType: 'richText',
          tag: 't',
        }
      )
    ).toBeNull();
  });
});

describe('wrapRangeInContentControlTr', () => {
  test('refuses a range that crosses a paragraph boundary', () => {
    const state = stateOf(para('one', 'P1'), para('two', 'P2'));
    // A range spanning into the second paragraph cannot be wrapped inline.
    const tr = wrapRangeInContentControlTr(
      state,
      { from: 2, to: state.doc.content.size - 1 },
      { sdtType: 'richText', tag: 't' }
    );
    expect(tr).toBeNull();
  });
});
