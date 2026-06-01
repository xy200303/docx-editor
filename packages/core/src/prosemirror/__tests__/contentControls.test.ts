/**
 * PM-level content-control addressing (live-editor path): discover by tag,
 * fill content, and remove, against an EditorState built from the singleton
 * schema.
 */

import { describe, expect, test } from 'bun:test';
import { EditorState } from 'prosemirror-state';

import { singletonManager } from '../schema';
import {
  findContentControlsInPM,
  findContentControlPos,
  setContentControlContentTr,
  removeContentControlTr,
} from '../contentControls';
import {
  ContentControlNotFoundError,
  ContentControlLockedError,
  ContentControlTypeError,
} from '../../agent/contentControls';

const schema = singletonManager.getSchema();

function blockSdt(attrs: Record<string, unknown>, text: string) {
  return schema.nodes.blockSdt.create(
    attrs,
    schema.nodes.paragraph.create(null, schema.text(text))
  );
}

function makeState() {
  const doc = schema.nodes.doc.create(null, [
    schema.nodes.paragraph.create(null, schema.text('before')),
    blockSdt({ sdtType: 'richText', tag: 'intro', alias: 'Intro' }, 'CONTROL #1'),
    blockSdt({ sdtType: 'richText', tag: 'locked', lock: 'sdtContentLocked' }, 'do not edit'),
    schema.nodes.paragraph.create(null, schema.text('after')),
  ]);
  return EditorState.create({ schema, doc });
}

describe('PM content-control addressing', () => {
  test('finds controls by tag and reads text', () => {
    const state = makeState();
    const all = findContentControlsInPM(state.doc);
    expect(all.map((c) => c.tag)).toEqual(['intro', 'locked']);
    expect(findContentControlsInPM(state.doc, { tag: 'intro' })[0].text).toBe('CONTROL #1');
  });

  test('findContentControlPos returns the node position', () => {
    const state = makeState();
    const pos = findContentControlPos(state.doc, { tag: 'intro' });
    expect(pos).not.toBeNull();
    expect(state.doc.nodeAt(pos!)?.type.name).toBe('blockSdt');
  });

  test('setContentControlContentTr replaces content, keeping the control', () => {
    const state = makeState();
    const tr = setContentControlContentTr(state, { tag: 'intro' }, 'Filled');
    const next = state.apply(tr);
    const ctrl = findContentControlsInPM(next.doc, { tag: 'intro' })[0];
    expect(ctrl.text).toBe('Filled');
    expect(ctrl.alias).toBe('Intro');
    expect(findContentControlsInPM(next.doc).length).toBe(2); // none lost
  });

  test('locked control is refused unless forced', () => {
    const state = makeState();
    expect(() => setContentControlContentTr(state, { tag: 'locked' }, 'x')).toThrow(
      ContentControlLockedError
    );
    const tr = setContentControlContentTr(state, { tag: 'locked' }, 'forced', { force: true });
    expect(findContentControlsInPM(state.apply(tr).doc, { tag: 'locked' })[0].text).toBe('forced');
  });

  test('missing tag throws', () => {
    const state = makeState();
    expect(() => setContentControlContentTr(state, { tag: 'nope' }, 'x')).toThrow(
      ContentControlNotFoundError
    );
  });

  test('removeContentControlTr deletes; keepContent unwraps', () => {
    const state = makeState();
    const removed = state.apply(removeContentControlTr(state, { tag: 'intro' }));
    expect(findContentControlsInPM(removed.doc, { tag: 'intro' })).toHaveLength(0);

    const unwrapped = state.apply(
      removeContentControlTr(state, { tag: 'intro' }, { keepContent: true })
    );
    expect(findContentControlsInPM(unwrapped.doc, { tag: 'intro' })).toHaveLength(0);
    expect(unwrapped.doc.textContent).toContain('CONTROL #1'); // content kept
  });

  test('refuses free-text replacement of a typed (dropdown) control unless forced', () => {
    const doc = schema.nodes.doc.create(null, [
      blockSdt({ sdtType: 'dropDownList', tag: 'choice' }, 'Active'),
    ]);
    const state = EditorState.create({ schema, doc });
    expect(() => setContentControlContentTr(state, { tag: 'choice' }, 'x')).toThrow(
      ContentControlTypeError
    );
    const tr = setContentControlContentTr(state, { tag: 'choice' }, 'Archived', { force: true });
    expect(findContentControlsInPM(state.apply(tr).doc, { tag: 'choice' })[0].text).toBe(
      'Archived'
    );
  });

  test('clears the showingPlaceholder attr when filling a placeholder control', () => {
    const doc = schema.nodes.doc.create(null, [
      blockSdt(
        {
          sdtType: 'richText',
          tag: 'ph',
          showingPlaceholder: true,
          rawPropertiesXml: '<w:sdtPr><w:showingPlcHdr/></w:sdtPr>',
        },
        'Click to enter'
      ),
    ]);
    const state = EditorState.create({ schema, doc });
    const next = state.apply(setContentControlContentTr(state, { tag: 'ph' }, 'Real'));
    expect(next.doc.firstChild!.attrs.showingPlaceholder).toBe(false);
    expect(String(next.doc.firstChild!.attrs.rawPropertiesXml)).not.toContain('showingPlcHdr');
  });
});
