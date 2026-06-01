/**
 * ProseMirror-level content-control (SDT) addressing for the live editor.
 *
 * The headless equivalents in `agent/contentControls` operate on the parsed
 * Document model; these operate on the editor's PM state so the editor adapters
 * (React/Vue) can discover and edit a control by tag without a full reload and
 * with normal undo. Shared by both adapters to keep them in lockstep.
 */

import type { EditorState, Transaction } from 'prosemirror-state';
import type { Node as PMNode } from 'prosemirror-model';

import {
  ContentControlNotFoundError,
  ContentControlLockedError,
  ContentControlTypeError,
  ContentControlBoundError,
  isTextReplaceable,
  isContentLocked,
  isDeletionLocked,
  clearShowingPlaceholderXml,
  type ContentControlFilter,
} from '../agent/contentControls';
import type { SdtType, SdtProperties, SdtDataBinding } from '../types/document';

/** A control discovered in the PM doc, with its PM position for scroll/edit. */
export interface PMContentControl {
  tag?: string;
  alias?: string;
  id?: number;
  sdtType: SdtType;
  lock?: SdtProperties['lock'];
  /** Whether the control is currently showing placeholder text. */
  showingPlaceholder?: boolean;
  /** Checkbox state, for checkbox controls. */
  checked?: boolean;
  /** Date format, for date controls. */
  dateFormat?: string;
  /** Dropdown/combobox list items, if modeled. */
  listItems?: { displayText: string; value: string }[];
  /** XML data binding (`w:dataBinding`), if the control is bound. */
  dataBinding?: SdtDataBinding;
  /** Plain text of the control's content. */
  text: string;
  /** PM position of the `blockSdt` node (its `before` position). */
  pos: number;
  /** Nesting depth among content controls (0 = not inside another control). */
  depth: number;
}

/** Safe JSON.parse for a cached attr string; returns undefined on miss/throw. */
function parseAttrJson<T>(value: unknown): T | undefined {
  if (typeof value !== 'string' || !value) return undefined;
  try {
    return JSON.parse(value) as T;
  } catch {
    return undefined;
  }
}

function attrsMatch(attrs: Record<string, unknown>, filter: ContentControlFilter): boolean {
  if (filter.tag !== undefined && attrs.tag !== filter.tag) return false;
  if (filter.alias !== undefined && attrs.alias !== filter.alias) return false;
  if (filter.id !== undefined && attrs.id !== filter.id) return false;
  if (filter.type !== undefined && (attrs.sdtType ?? 'richText') !== filter.type) return false;
  return true;
}

function controlInfo(node: PMNode, pos: number, depth: number): PMContentControl {
  const a = node.attrs as Record<string, unknown>;
  return {
    tag: a.tag != null ? String(a.tag) : undefined,
    alias: a.alias != null ? String(a.alias) : undefined,
    id: typeof a.id === 'number' ? a.id : undefined,
    sdtType: (a.sdtType as SdtType) ?? 'richText',
    lock: (a.lock as SdtProperties['lock']) ?? undefined,
    showingPlaceholder: a.showingPlaceholder === true ? true : undefined,
    checked: typeof a.checked === 'boolean' ? a.checked : undefined,
    dateFormat: a.dateFormat != null ? String(a.dateFormat) : undefined,
    listItems: parseAttrJson<{ displayText: string; value: string }[]>(a.listItems),
    dataBinding: parseAttrJson<SdtDataBinding>(a.dataBinding),
    text: node.textBetween(0, node.content.size, '\n'),
    pos,
    depth,
  };
}

/** All block content controls in the PM doc (document order), optionally filtered. */
export function findContentControlsInPM(
  doc: PMNode,
  filter: ContentControlFilter = {}
): PMContentControl[] {
  const out: PMContentControl[] = [];
  // `base` is the position of `node`'s first child; a child at `offset` has
  // `before` position `base + offset`, and its own children start one past it.
  // `depth` counts only enclosing content controls.
  const walk = (node: PMNode, base: number, depth: number): void => {
    node.forEach((child, offset) => {
      const childPos = base + offset;
      const isSdt = child.type.name === 'blockSdt';
      if (isSdt && attrsMatch(child.attrs as Record<string, unknown>, filter)) {
        out.push(controlInfo(child, childPos, depth));
      }
      walk(child, childPos + 1, isSdt ? depth + 1 : depth);
    });
  };
  walk(doc, 0, 0);
  return out;
}

/** PM position of the first control matching `filter`, or `null`. */
export function findContentControlPos(doc: PMNode, filter: ContentControlFilter): number | null {
  return findContentControlsInPM(doc, filter)[0]?.pos ?? null;
}

/** Locate the first matching blockSdt node + its position in the PM doc. */
function locate(doc: PMNode, filter: ContentControlFilter): { node: PMNode; pos: number } | null {
  let found: { node: PMNode; pos: number } | null = null;
  doc.descendants((node, pos) => {
    if (found) return false;
    if (
      node.type.name === 'blockSdt' &&
      attrsMatch(node.attrs as Record<string, unknown>, filter)
    ) {
      found = { node, pos };
      return false;
    }
    return true;
  });
  return found;
}

/**
 * Build a transaction that replaces the first matching control's content with
 * `text` (newlines become paragraphs; a `plainText` control stays one
 * paragraph). Throws if nothing matches, the control is content-locked, a typed
 * (dropdown/date/…) control, or data-bound (unless `force`). The control's
 * identity/raw props are kept; a `w:showingPlcHdr` placeholder flag is cleared
 * so the new content isn't rendered as placeholder.
 */
export function setContentControlContentTr(
  state: EditorState,
  filter: ContentControlFilter,
  text: string,
  options: { force?: boolean } = {}
): Transaction {
  const target = locate(state.doc, filter);
  if (!target) throw new ContentControlNotFoundError(filter);
  const attrs = target.node.attrs as Record<string, unknown>;
  if (!options.force && isContentLocked(attrs.lock as SdtProperties['lock'])) {
    throw new ContentControlLockedError(attrs.lock as SdtProperties['lock'], 'edit');
  }
  const sdtType = (attrs.sdtType as SdtType) ?? 'richText';
  if (!options.force && !isTextReplaceable(sdtType)) {
    throw new ContentControlTypeError(sdtType);
  }
  if (!options.force && attrs.dataBinding != null) {
    throw new ContentControlBoundError();
  }
  const { schema } = state;
  const lines = sdtType === 'plainText' ? [text] : text.split('\n');
  const paragraphs = lines.map((line) =>
    schema.nodes.paragraph.create(null, line ? schema.text(line) : null)
  );
  const from = target.pos + 1;
  const to = target.pos + 1 + target.node.content.size;
  const tr = state.tr.replaceWith(from, to, paragraphs);
  // Clear placeholder state so the written content isn't styled as placeholder.
  // The node's own start position is unaffected by the inner content replace.
  if (attrs.showingPlaceholder || /showingPlcHdr/.test(String(attrs.rawPropertiesXml ?? ''))) {
    tr.setNodeMarkup(target.pos, undefined, {
      ...attrs,
      showingPlaceholder: false,
      rawPropertiesXml: clearShowingPlaceholderXml(
        attrs.rawPropertiesXml == null ? undefined : String(attrs.rawPropertiesXml)
      ),
    });
  }
  return tr;
}

/**
 * Build a transaction that removes the first matching control. With
 * `keepContent` the inner blocks are unwrapped in place; otherwise the whole
 * region is deleted. Throws if nothing matches or the control is
 * deletion-locked (unless `force`).
 */
export function removeContentControlTr(
  state: EditorState,
  filter: ContentControlFilter,
  options: { force?: boolean; keepContent?: boolean } = {}
): Transaction {
  const target = locate(state.doc, filter);
  if (!target) throw new ContentControlNotFoundError(filter);
  const lock = target.node.attrs.lock as SdtProperties['lock'];
  if (!options.force && isDeletionLocked(lock)) {
    throw new ContentControlLockedError(lock, 'remove');
  }
  // Unwrapping a repeating-section item would orphan its (w15) structure.
  if (
    options.keepContent &&
    !options.force &&
    /<w15:repeatingSection(Item)?[\s/>]/.test(String(target.node.attrs.rawPropertiesXml ?? ''))
  ) {
    throw new ContentControlLockedError(lock, 'remove');
  }
  const start = target.pos;
  const end = target.pos + target.node.nodeSize;
  return options.keepContent
    ? state.tr.replaceWith(start, end, target.node.content)
    : state.tr.delete(start, end);
}
