/**
 * Unit tests for contributesToHeaderFooterFlowHeight (issue #705).
 *
 * Only IN-FLOW header/footer content grows the band that pushes the body
 * margin. Anchored / floating objects (page or margin-anchored shapes, e.g. a
 * letterhead) are positioned on the page and must NOT contribute.
 */

import { describe, test, expect } from 'bun:test';
import { contributesToHeaderFooterFlowHeight } from '../headerFooterLayout';
import type {
  FlowBlock,
  ImageBlock,
  ParagraphBlock,
  TextBoxBlock,
} from '../../layout-engine/types';

const paragraph: ParagraphBlock = {
  kind: 'paragraph',
  id: 'p',
  runs: [{ kind: 'text', text: 'hi' }],
};

const table: FlowBlock = { kind: 'table', id: 't', rows: [] };

function image(anchored: boolean): ImageBlock {
  return {
    kind: 'image',
    id: 'img',
    src: 'logo.png',
    width: 100,
    height: 50,
    anchor: anchored ? { isAnchored: true } : undefined,
  };
}

function textBox(displayMode: TextBoxBlock['displayMode']): TextBoxBlock {
  return { kind: 'textBox', id: 'tb', width: 200, height: 400, content: [], displayMode };
}

describe('contributesToHeaderFooterFlowHeight', () => {
  test('paragraphs and tables always count', () => {
    expect(contributesToHeaderFooterFlowHeight(paragraph)).toBe(true);
    expect(contributesToHeaderFooterFlowHeight(table)).toBe(true);
  });

  test('inline images count; anchored images do not', () => {
    expect(contributesToHeaderFooterFlowHeight(image(false))).toBe(true);
    expect(contributesToHeaderFooterFlowHeight(image(true))).toBe(false);
  });

  test('inline text boxes count; floating / block text boxes do not', () => {
    expect(contributesToHeaderFooterFlowHeight(textBox('inline'))).toBe(true);
    expect(contributesToHeaderFooterFlowHeight(textBox(undefined))).toBe(true);
    // A page-anchored letterhead text box (the #705 trigger).
    expect(contributesToHeaderFooterFlowHeight(textBox('float'))).toBe(false);
    expect(contributesToHeaderFooterFlowHeight(textBox('block'))).toBe(false);
  });

  test('structural breaks never count', () => {
    expect(contributesToHeaderFooterFlowHeight({ kind: 'pageBreak', id: 'pb' } as FlowBlock)).toBe(
      false
    );
    expect(contributesToHeaderFooterFlowHeight({ kind: 'sectionBreak' } as FlowBlock)).toBe(false);
  });
});
