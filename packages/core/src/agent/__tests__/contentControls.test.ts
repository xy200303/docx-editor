/**
 * Read/discovery API for block-level content controls, exercised against the
 * comprehensive fixture (10 SDT scenarios incl. nesting, table-wrapping,
 * dropdown, lock).
 */

import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, test } from 'bun:test';

import { parseDocx } from '../../docx/parser';
import { createDocx } from '../../docx/rezip';
import {
  findContentControls,
  findContentControl,
  setContentControlContent,
  removeContentControl,
  clearShowingPlaceholderXml,
  ContentControlNotFoundError,
  ContentControlLockedError,
  ContentControlTypeError,
  ContentControlBoundError,
} from '../contentControls';
import type { BlockContent, Document } from '../../types/document';

// Relative to this file (cwd-independent — CI runs tests from the repo root).
const FIXTURE = join(import.meta.dir, '../../../../../e2e/fixtures/block-sdt-comprehensive.docx');

async function loadFixture(): Promise<Document> {
  const buf = readFileSync(FIXTURE);
  return parseDocx(buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength));
}

describe('findContentControls', () => {
  test('discovers every control with its modeled properties', async () => {
    const doc = await loadFixture();
    const controls = findContentControls(doc);

    // 10 declared scenarios; one (#4) nests another, so 11 controls total.
    expect(controls.length).toBe(11);
    const tags = controls.map((c) => c.tag);
    expect(tags).toContain('intro');
    expect(tags).toContain('multi');
    expect(tags).toContain('outer');
    expect(tags).toContain('inner');
  });

  test('filters by tag and reads the control text', async () => {
    const doc = await loadFixture();
    const intro = findContentControl(doc, { tag: 'intro' });
    expect(intro).toBeDefined();
    expect(intro?.alias).toBe('Intro');
    expect(intro?.sdtType).toBe('richText');
    expect(intro?.text).toContain('CONTROL #1');
  });

  test('surfaces lock, dropdown list items, and table-wrapping text', async () => {
    const doc = await loadFixture();
    expect(findContentControl(doc, { tag: 'locked' })?.lock).toBe('sdtContentLocked');
    expect(findContentControl(doc, { tag: 'choice' })?.listItems?.length).toBe(3);
    // #2 wraps a table; its text flattens the cells.
    expect(findContentControl(doc, { tag: 'grid' })?.text).toContain('B2');
  });

  test('reports nesting via path/depth and finds the inner control', async () => {
    const doc = await loadFixture();
    const inner = findContentControl(doc, { tag: 'inner' });
    expect(inner).toBeDefined();
    expect(inner!.depth).toBeGreaterThan(0); // nested below the body root
    expect(inner!.path.length).toBe(inner!.depth + 1);
  });

  test('filter by type narrows to matching controls', async () => {
    const doc = await loadFixture();
    const dropdowns = findContentControls(doc, { type: 'dropDownList' });
    expect(dropdowns.length).toBe(1);
    expect(dropdowns[0].tag).toBe('choice');
  });
});

describe('setContentControlContent', () => {
  test('replaces a control’s text while keeping tag/alias and raw props', async () => {
    const doc = await loadFixture();
    const before = findContentControl(doc, { tag: 'intro' })!;
    const next = setContentControlContent(doc, { tag: 'intro' }, 'Filled by template');

    const after = findContentControl(next, { tag: 'intro' })!;
    expect(after.text).toBe('Filled by template');
    expect(after.alias).toBe('Intro'); // identity preserved
    expect(after.id).toBe(before.id);
    // original untouched (pure function)
    expect(findContentControl(doc, { tag: 'intro' })!.text).toBe(before.text);
  });

  test('refuses a locked control unless forced', async () => {
    const doc = await loadFixture();
    expect(() => setContentControlContent(doc, { tag: 'locked' }, 'x')).toThrow(
      ContentControlLockedError
    );
    const forced = setContentControlContent(doc, { tag: 'locked' }, 'overridden', { force: true });
    expect(findContentControl(forced, { tag: 'locked' })!.text).toBe('overridden');
  });

  test('throws when nothing matches', async () => {
    const doc = await loadFixture();
    expect(() => setContentControlContent(doc, { tag: 'nope' }, 'x')).toThrow(
      ContentControlNotFoundError
    );
  });

  test('edits survive a full save → reparse round-trip', async () => {
    const doc = await loadFixture();
    const edited = setContentControlContent(doc, { tag: 'multi' }, 'New multi body');
    const bytes = await createDocx(edited);
    const reparsed = await parseDocx(bytes);

    const ctrl = findContentControl(reparsed, { tag: 'multi' })!;
    expect(ctrl.text).toBe('New multi body');
    expect(ctrl.alias).toBe('Multi');
    // other controls still present after the round-trip
    expect(findContentControls(reparsed).length).toBeGreaterThanOrEqual(10);
  });
});

describe('removeContentControl', () => {
  test('deletes the control region', async () => {
    const doc = await loadFixture();
    const next = removeContentControl(doc, { tag: 'intro' });
    expect(findContentControl(next, { tag: 'intro' })).toBeUndefined();
    expect(findContentControls(next).length).toBe(findContentControls(doc).length - 1);
  });

  test('keepContent unwraps the control but leaves its blocks', async () => {
    const doc = await loadFixture();
    const introText = findContentControl(doc, { tag: 'intro' })!.text;
    const next = removeContentControl(doc, { tag: 'intro' }, { keepContent: true });
    expect(findContentControl(next, { tag: 'intro' })).toBeUndefined();
    // the text is still in the body (now as a plain paragraph)
    const stillThere = next.package.document.content.some(
      (b) => b.type === 'paragraph' && b.content.some((r) => r.type === 'run')
    );
    expect(stillThere).toBe(true);
    expect(introText.length).toBeGreaterThan(0);
  });

  test('refuses to unwrap a repeating-section control unless forced', async () => {
    const doc = await loadFixture();
    expect(() => removeContentControl(doc, { tag: 'repeat' }, { keepContent: true })).toThrow(
      ContentControlLockedError
    );
    // Plain delete (not unwrap) is allowed.
    expect(
      findContentControl(removeContentControl(doc, { tag: 'repeat' }), { tag: 'repeat' })
    ).toBeUndefined();
  });
});

describe('typed controls + data binding + purity', () => {
  test('refuses free-text replacement of a dropdown control unless forced', async () => {
    const doc = await loadFixture();
    expect(() => setContentControlContent(doc, { tag: 'choice' }, 'Whatever')).toThrow(
      ContentControlTypeError
    );
    const forced = setContentControlContent(doc, { tag: 'choice' }, 'Archived', { force: true });
    expect(findContentControl(forced, { tag: 'choice' })!.text).toBe('Archived');
  });

  test('surfaces dataBinding on a bound control', async () => {
    const doc = await loadFixture();
    const bound = findContentControl(doc, { tag: 'bound' })!;
    expect(bound.dataBinding).toBeDefined();
    expect(bound.dataBinding!.storeItemID).toContain('1B2C3D4E');
  });

  test('block-content input is cloned (caller mutation does not leak)', async () => {
    const doc = await loadFixture();
    const blocks: BlockContent[] = [
      { type: 'paragraph', content: [{ type: 'run', content: [{ type: 'text', text: 'orig' }] }] },
    ];
    const next = setContentControlContent(doc, { tag: 'intro' }, blocks);
    // mutate the caller's array AFTER the call
    (blocks[0] as { content: { content: { text: string }[] }[] }).content[0].content[0].text =
      'mutated';
    expect(findContentControl(next, { tag: 'intro' })!.text).toBe('orig');
  });
});

describe('clearShowingPlaceholderXml', () => {
  test('strips the showingPlcHdr element, leaves other props', () => {
    const raw = '<w:sdtPr><w:tag w:val="x"/><w:showingPlcHdr/><w:id w:val="1"/></w:sdtPr>';
    const out = clearShowingPlaceholderXml(raw)!;
    expect(out).not.toContain('showingPlcHdr');
    expect(out).toContain('w:tag');
    expect(out).toContain('w:id');
  });

  test('setContentControlContent clears the placeholder flag on a placeholder control', () => {
    const doc: Document = {
      package: {
        document: {
          content: [
            {
              type: 'blockSdt',
              properties: {
                sdtType: 'richText',
                tag: 'ph',
                showingPlaceholder: true,
                rawPropertiesXml: '<w:sdtPr><w:tag w:val="ph"/><w:showingPlcHdr/></w:sdtPr>',
              },
              content: [{ type: 'paragraph', content: [] }],
            },
          ],
        },
      },
    } as unknown as Document;

    const next = setContentControlContent(doc, { tag: 'ph' }, 'Real value');
    const ctrl = next.package.document.content[0];
    expect(ctrl.type).toBe('blockSdt');
    if (ctrl.type === 'blockSdt') {
      expect(ctrl.properties.showingPlaceholder).toBe(false);
      expect(ctrl.properties.rawPropertiesXml).not.toContain('showingPlcHdr');
    }
  });
});

// Build a one-control document for guard tests.
function docWith(props: Record<string, unknown>): Document {
  return {
    package: {
      document: {
        content: [
          {
            type: 'blockSdt',
            properties: props,
            content: [{ type: 'paragraph', content: [] }],
          },
        ],
      },
    },
  } as unknown as Document;
}

function firstControlText(doc: Document): string {
  const c = doc.package.document.content[0];
  if (c.type !== 'blockSdt') return '';
  return c.content
    .filter((b): b is Extract<BlockContent, { type: 'paragraph' }> => b.type === 'paragraph')
    .map((p) =>
      p.content
        .map((r) =>
          r.type === 'run' ? r.content.map((t) => ('text' in t ? t.text : '')).join('') : ''
        )
        .join('')
    )
    .join('\n');
}

describe('round-2 write guards', () => {
  test('refuses writing to a data-bound control unless forced', () => {
    const doc = docWith({
      sdtType: 'richText',
      tag: 'b',
      dataBinding: { xpath: '/root/x', storeItemID: '{X}' },
    });
    expect(() => setContentControlContent(doc, { tag: 'b' }, 'v')).toThrow(
      ContentControlBoundError
    );
    expect(
      firstControlText(setContentControlContent(doc, { tag: 'b' }, 'v', { force: true }))
    ).toBe('v');
  });

  test('plainText fill collapses to a single paragraph', () => {
    const doc = docWith({ sdtType: 'plainText', tag: 'p' });
    const next = setContentControlContent(doc, { tag: 'p' }, 'line1\nline2');
    const c = next.package.document.content[0];
    expect(c.type).toBe('blockSdt');
    if (c.type === 'blockSdt') expect(c.content.length).toBe(1); // not split into 2 paragraphs
  });

  test('refuses free-text replacement of a group control unless forced', () => {
    const doc = docWith({ sdtType: 'group', tag: 'g' });
    expect(() => setContentControlContent(doc, { tag: 'g' }, 'x')).toThrow(ContentControlTypeError);
  });
});
