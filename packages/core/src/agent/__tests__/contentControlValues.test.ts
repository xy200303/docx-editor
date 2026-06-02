/**
 * Typed value setters for content controls: dropdown selection, checkbox
 * toggle, and date — each updates the visible content and patches the raw
 * w:sdtPr state in place.
 */

import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, test } from 'bun:test';

import {
  setContentControlValue,
  formatSdtDate,
  ContentControlValueError,
} from '../contentControlValues';
import {
  findContentControl,
  ContentControlNotFoundError,
  ContentControlBoundError,
} from '../contentControls';
import { parseDocx } from '../../docx/parser';
import { createDocx } from '../../docx/rezip';
import type { Document, SdtProperties } from '../../types/document';

const WIDGET_FIXTURE = join(import.meta.dir, '../../../../../e2e/fixtures/block-sdt-widgets.docx');

function docWith(props: SdtProperties): Document {
  return {
    package: {
      document: {
        content: [
          { type: 'blockSdt', properties: props, content: [{ type: 'paragraph', content: [] }] },
        ],
      },
    },
  } as unknown as Document;
}

describe('setContentControlValue — dropdown', () => {
  const props: SdtProperties = {
    sdtType: 'dropDownList',
    tag: 'status',
    listItems: [
      { displayText: 'Draft', value: '1' },
      { displayText: 'Final', value: '2' },
    ],
    rawPropertiesXml: '<w:sdtPr><w:tag w:val="status"/><w:dropDownList w:lastValue="1"/></w:sdtPr>',
  };

  test('selects an item by value: sets display text and patches lastValue', () => {
    const next = setContentControlValue(
      docWith(props),
      { tag: 'status' },
      {
        kind: 'dropdown',
        value: '2',
      }
    );
    const ctrl = findContentControl(next, { tag: 'status' })!;
    expect(ctrl.text).toBe('Final');
    const c = next.package.document.content[0];
    if (c.type === 'blockSdt') {
      expect(c.properties.rawPropertiesXml).toContain('w:lastValue="2"');
    }
  });

  test('accepts matching by displayText', () => {
    const next = setContentControlValue(
      docWith(props),
      { tag: 'status' },
      {
        kind: 'dropdown',
        value: 'Draft',
      }
    );
    expect(findContentControl(next, { tag: 'status' })!.text).toBe('Draft');
  });

  test('rejects an unknown value', () => {
    expect(() =>
      setContentControlValue(docWith(props), { tag: 'status' }, { kind: 'dropdown', value: 'X' })
    ).toThrow(ContentControlValueError);
  });

  test('rejects a kind/type mismatch', () => {
    expect(() =>
      setContentControlValue(docWith(props), { tag: 'status' }, { kind: 'checkbox', checked: true })
    ).toThrow(ContentControlValueError);
  });
});

describe('setContentControlValue — checkbox', () => {
  const props: SdtProperties = {
    sdtType: 'checkbox',
    tag: 'agree',
    checked: false,
    rawPropertiesXml:
      '<w:sdtPr><w:tag w:val="agree"/><w14:checkbox><w14:checked w14:val="0"/>' +
      '<w14:checkedState w14:val="2612" w14:font="MS Gothic"/>' +
      '<w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/></w14:checkbox></w:sdtPr>',
  };

  test('checking sets checked + the checked glyph + raw val', () => {
    const next = setContentControlValue(
      docWith(props),
      { tag: 'agree' },
      {
        kind: 'checkbox',
        checked: true,
      }
    );
    const c = next.package.document.content[0];
    expect(c.type).toBe('blockSdt');
    if (c.type === 'blockSdt') {
      expect(c.properties.checked).toBe(true);
      expect(c.properties.rawPropertiesXml).toContain('w14:checked w14:val="1"');
      expect(findContentControl(next, { tag: 'agree' })!.text).toBe(String.fromCodePoint(0x2612));
    }
  });

  test('unchecking sets the unchecked glyph', () => {
    const checked = setContentControlValue(
      docWith(props),
      { tag: 'agree' },
      {
        kind: 'checkbox',
        checked: true,
      }
    );
    const next = setContentControlValue(
      checked,
      { tag: 'agree' },
      {
        kind: 'checkbox',
        checked: false,
      }
    );
    expect(findContentControl(next, { tag: 'agree' })!.text).toBe(String.fromCodePoint(0x2610));
    const c = next.package.document.content[0];
    if (c.type === 'blockSdt') expect(c.properties.rawPropertiesXml).toContain('w14:val="0"');
  });
});

describe('setContentControlValue — date', () => {
  const props: SdtProperties = {
    sdtType: 'date',
    tag: 'effective',
    rawPropertiesXml:
      '<w:sdtPr><w:tag w:val="effective"/><w:date w:fullDate="2020-01-01T00:00:00Z">' +
      '<w:dateFormat w:val="MMMM d, yyyy"/></w:date></w:sdtPr>',
  };

  test('sets fullDate and formats the display text', () => {
    const next = setContentControlValue(
      docWith(props),
      { tag: 'effective' },
      {
        kind: 'date',
        date: '2026-06-01',
      }
    );
    const c = next.package.document.content[0];
    expect(c.type).toBe('blockSdt');
    if (c.type === 'blockSdt') {
      expect(c.properties.rawPropertiesXml).toContain('w:fullDate="2026-06-01T00:00:00"');
    }
    expect(findContentControl(next, { tag: 'effective' })!.text).toBe('June 1, 2026');
  });

  test('rejects a malformed date', () => {
    expect(() =>
      setContentControlValue(
        docWith(props),
        { tag: 'effective' },
        {
          kind: 'date',
          date: 'not-a-date',
        }
      )
    ).toThrow(ContentControlValueError);
  });
});

describe('setContentControlValue — general', () => {
  test('throws when nothing matches', () => {
    const doc = docWith({ sdtType: 'checkbox', tag: 'a' });
    expect(() =>
      setContentControlValue(doc, { tag: 'zzz' }, { kind: 'checkbox', checked: true })
    ).toThrow(ContentControlNotFoundError);
  });
});

describe('setContentControlValue — fidelity', () => {
  test('checkbox glyph run carries the symbol font from w14:checkedState', () => {
    const props: SdtProperties = {
      sdtType: 'checkbox',
      tag: 'agree',
      rawPropertiesXml:
        '<w:sdtPr><w14:checkbox><w14:checked w14:val="0"/>' +
        '<w14:checkedState w14:val="2612" w14:font="MS Gothic"/>' +
        '<w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/></w14:checkbox></w:sdtPr>',
    };
    const next = setContentControlValue(
      docWith(props),
      { tag: 'agree' },
      {
        kind: 'checkbox',
        checked: true,
      }
    );
    const c = next.package.document.content[0];
    if (c.type === 'blockSdt') {
      const run = c.content[0].type === 'paragraph' ? c.content[0].content[0] : undefined;
      expect(run?.type).toBe('run');
      if (run?.type === 'run') expect(run.formatting?.fontFamily?.ascii).toBe('MS Gothic');
    }
  });

  test('clears the placeholder flag when setting a value', () => {
    const props: SdtProperties = {
      sdtType: 'dropDownList',
      tag: 's',
      showingPlaceholder: true,
      listItems: [{ displayText: 'A', value: '1' }],
      rawPropertiesXml: '<w:sdtPr><w:showingPlcHdr/><w:dropDownList w:lastValue=""/></w:sdtPr>',
    };
    const next = setContentControlValue(
      docWith(props),
      { tag: 's' },
      { kind: 'dropdown', value: '1' }
    );
    const c = next.package.document.content[0];
    if (c.type === 'blockSdt') {
      expect(c.properties.showingPlaceholder).toBe(false);
      expect(c.properties.rawPropertiesXml).not.toContain('showingPlcHdr');
    }
  });

  test('escapes special characters in a dropdown value (valid XML)', () => {
    const props: SdtProperties = {
      sdtType: 'dropDownList',
      tag: 's',
      listItems: [{ displayText: 'A & B', value: 'a&b' }],
      rawPropertiesXml: '<w:sdtPr><w:dropDownList w:lastValue="x"/></w:sdtPr>',
    };
    const next = setContentControlValue(
      docWith(props),
      { tag: 's' },
      { kind: 'dropdown', value: 'a&b' }
    );
    const c = next.package.document.content[0];
    if (c.type === 'blockSdt') {
      expect(c.properties.rawPropertiesXml).toContain('w:lastValue="a&amp;b"');
      expect(c.properties.rawPropertiesXml).not.toContain('w:lastValue="a&b"');
    }
  });
});

describe('formatSdtDate', () => {
  test('handles common patterns', () => {
    expect(formatSdtDate('2026-06-01', 'M/d/yyyy')).toBe('6/1/2026');
    expect(formatSdtDate('2026-06-01', 'MM/dd/yyyy')).toBe('06/01/2026');
    expect(formatSdtDate('2026-06-01', 'yyyy-MM-dd')).toBe('2026-06-01');
    expect(formatSdtDate('2026-06-01', 'MMM d, yyyy')).toBe('Jun 1, 2026');
    expect(formatSdtDate('2026-06-01')).toBe('6/1/2026'); // default
  });

  test('does not corrupt month names containing "M" (single-pass)', () => {
    expect(formatSdtDate('2026-03-05', 'MMMM d, yyyy')).toBe('March 5, 2026');
    expect(formatSdtDate('2026-05-09', 'MMMM d, yyyy')).toBe('May 9, 2026');
    expect(formatSdtDate('2026-03-05', 'MMM d')).toBe('Mar 5');
  });
});

describe('setContentControlValue — guards', () => {
  test('refuses a data-bound control unless forced', () => {
    const doc = docWith({
      sdtType: 'checkbox',
      tag: 'b',
      dataBinding: { xpath: '/x', storeItemID: '{X}' },
      rawPropertiesXml:
        '<w:sdtPr><w14:checkbox><w14:checked w14:val="0"/></w14:checkbox></w:sdtPr>',
    });
    expect(() =>
      setContentControlValue(doc, { tag: 'b' }, { kind: 'checkbox', checked: true })
    ).toThrow(ContentControlBoundError);
  });

  test('checkbox with no w14:checked state is rejected', () => {
    const doc = docWith({ sdtType: 'checkbox', tag: 'c', rawPropertiesXml: '<w:sdtPr/>' });
    expect(() =>
      setContentControlValue(doc, { tag: 'c' }, { kind: 'checkbox', checked: true })
    ).toThrow(ContentControlValueError);
  });
});

describe('setContentControlValue — full save → reparse round-trip', () => {
  async function load(): Promise<Document> {
    const buf = readFileSync(WIDGET_FIXTURE);
    return parseDocx(buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength));
  }

  test('dropdown, checkbox, and date values survive serialize→reparse', async () => {
    let doc = await load();
    doc = setContentControlValue(doc, { tag: 'status' }, { kind: 'dropdown', value: '2' });
    doc = setContentControlValue(doc, { tag: 'agree' }, { kind: 'checkbox', checked: true });
    doc = setContentControlValue(doc, { tag: 'effective' }, { kind: 'date', date: '2026-03-05' });

    const reparsed = await parseDocx(await createDocx(doc));
    expect(findContentControl(reparsed, { tag: 'status' })!.text).toBe('Final');
    expect(findContentControl(reparsed, { tag: 'agree' })!.text).toBe(String.fromCodePoint(0x2612));
    expect(findContentControl(reparsed, { tag: 'agree' })!.checked).toBe(true);
    expect(findContentControl(reparsed, { tag: 'effective' })!.text).toBe('March 5, 2026');
  });
});
