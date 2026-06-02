/**
 * Content-control (SDT) addressing for the document model.
 *
 * Block-level content controls (`w:sdt`) are the natural anchor for template
 * logic and agent edits: they survive the round trip (see the parser +
 * serializer) and carry a stable `tag`/`alias`/`id`. This module is the
 * read side of that contract — discover controls and read their content
 * without a DOM or an editor instance, so server-side pipelines and AI
 * agents can find an anchor by tag and act on it.
 *
 * Walks the body recursively, descending into nested controls so a control
 * inside another control is still found. (The model places block content
 * controls at body level or nested in other controls, not inside table
 * cells.) The returned `path` (block indices from the body root) addresses
 * the control unambiguously for a follow-up edit.
 */

import type {
  Document,
  DocumentBody,
  BlockContent,
  BlockSdt,
  SdtType,
  SdtProperties,
  SdtDataBinding,
} from '../types/document';
import { getParagraphText, getTableText } from './text-utils';

/** Filter for {@link findContentControls}. All provided fields must match (AND). */
export interface ContentControlFilter {
  /** Developer identifier (`w:tag`), exact match. */
  tag?: string;
  /** Friendly name (`w:alias`), exact match. */
  alias?: string;
  /** Numeric id (`w:id`), exact match. */
  id?: number;
  /** Control type projection (`richText`, `dropDownList`, …). */
  type?: SdtType;
}

/** A discovered content control plus enough context to address and edit it. */
export interface ContentControlInfo {
  /** Developer identifier (`w:tag`). */
  tag?: string;
  /** Friendly name (`w:alias`). */
  alias?: string;
  /** Numeric id (`w:id`). */
  id?: number;
  /** Control type projection. */
  sdtType: SdtType;
  /** Lock setting, if any. A locked control should refuse content edits. */
  lock?: SdtProperties['lock'];
  /** Dropdown/combobox list items, if modeled. */
  listItems?: { displayText: string; value: string }[];
  /** Placeholder docPart reference, if any. */
  placeholder?: string;
  /** Whether the control is currently showing placeholder text (`w:showingPlcHdr`). */
  showingPlaceholder?: boolean;
  /** Checkbox state, for checkbox controls. */
  checked?: boolean;
  /** Date format string, for date controls. */
  dateFormat?: string;
  /** XML data binding (`w:dataBinding`), if the control is bound. */
  dataBinding?: SdtDataBinding;
  /** Plain text of the control's content (paragraphs/tables/nested controls flattened). */
  text: string;
  /**
   * Block-index path from the document body to this control. Top-level
   * controls are `[i]`; a control nested inside the i-th body block's content
   * is `[i, j]`, and so on. Stable address for a follow-up edit.
   */
  path: number[];
  /** Nesting depth (0 = direct child of the body). */
  depth: number;
}

/** Narrow a {@link Document} or {@link DocumentBody} to its block list. */
function bodyOf(input: Document | DocumentBody): DocumentBody {
  return 'package' in input ? input.package.document : input;
}

/** Plain text of a control's content, descending into tables and nested SDTs. */
export function getContentControlText(control: BlockSdt): string {
  return blocksText(control.content);
}

function blocksText(blocks: BlockContent[]): string {
  const parts: string[] = [];
  for (const block of blocks) {
    if (block.type === 'paragraph') parts.push(getParagraphText(block));
    else if (block.type === 'table') parts.push(getTableText(block));
    else if (block.type === 'blockSdt') parts.push(blocksText(block.content));
  }
  return parts.join('\n');
}

function matches(props: SdtProperties, filter: ContentControlFilter): boolean {
  if (filter.tag !== undefined && props.tag !== filter.tag) return false;
  if (filter.alias !== undefined && props.alias !== filter.alias) return false;
  if (filter.id !== undefined && props.id !== filter.id) return false;
  if (filter.type !== undefined && props.sdtType !== filter.type) return false;
  return true;
}

function infoOf(control: BlockSdt, path: number[]): ContentControlInfo {
  const p = control.properties;
  return {
    tag: p.tag,
    alias: p.alias,
    id: p.id,
    sdtType: p.sdtType,
    lock: p.lock,
    listItems: p.listItems,
    placeholder: p.placeholder,
    showingPlaceholder: p.showingPlaceholder,
    checked: p.checked,
    dateFormat: p.dateFormat,
    dataBinding: p.dataBinding,
    text: getContentControlText(control),
    path,
    depth: path.length - 1,
  };
}

/**
 * Find every block-level content control in the document, optionally filtered
 * by tag/alias/id/type. Results are in document order; nested controls follow
 * their parent. Searches the body and controls nested inside controls. Table
 * cells are not searched (the current model/parser does not surface cell-level
 * controls), and headers/footers live in a separate content tree.
 */
export function findContentControls(
  input: Document | DocumentBody,
  filter: ContentControlFilter = {}
): ContentControlInfo[] {
  const out: ContentControlInfo[] = [];

  // Controls live at body level or nested inside other controls; a table cell
  // cannot hold one in this model, so we only recurse through blockSdt content.
  const walk = (blocks: BlockContent[], parentPath: number[]): void => {
    for (let i = 0; i < blocks.length; i++) {
      const block = blocks[i];
      if (block.type === 'blockSdt') {
        const path = [...parentPath, i];
        if (matches(block.properties, filter)) out.push(infoOf(block, path));
        walk(block.content, path); // nested controls
      }
    }
  };

  walk(bodyOf(input).content, []);
  return out;
}

/** Convenience: the first control matching `filter`, or `undefined`. */
export function findContentControl(
  input: Document | DocumentBody,
  filter: ContentControlFilter
): ContentControlInfo | undefined {
  return findContentControls(input, filter)[0];
}

// ============================================================================
// MUTATION (edit a control by tag)
// ============================================================================

/** No control matched the filter. */
export class ContentControlNotFoundError extends Error {
  constructor(filter: ContentControlFilter) {
    super(`No content control matched ${JSON.stringify(filter)}`);
    this.name = 'ContentControlNotFoundError';
  }
}

/** The matched control's lock forbids the attempted edit (pass `force` to override). */
export class ContentControlLockedError extends Error {
  constructor(lock: SdtProperties['lock'], op: 'edit' | 'remove') {
    super(`Content control is ${lock}; cannot ${op} it without { force: true }`);
    this.name = 'ContentControlLockedError';
  }
}

/**
 * The control's type doesn't support free text/block replacement (e.g. a
 * dropdown, date, checkbox, or picture control), so writing arbitrary content
 * would desync the type marker from its value. Use a type-specific setter, or
 * pass `{ force: true }` to override.
 */
export class ContentControlTypeError extends Error {
  constructor(sdtType: SdtType) {
    super(
      `Content control is a '${sdtType}' control; replacing its content with free text ` +
        `would desync it. Use a type-specific value setter or pass { force: true }.`
    );
    this.name = 'ContentControlTypeError';
  }
}

/**
 * The control is bound to a Custom XML data store (`w:dataBinding`). Writing its
 * content won't stick — Word re-renders the control from the bound XML node — so
 * the write is refused. Update the data store instead, or pass `{ force: true }`.
 */
export class ContentControlBoundError extends Error {
  constructor() {
    super(
      'Content control is data-bound (w:dataBinding); its content is driven by the ' +
        'Custom XML store and a direct write will not persist. Update the store, or pass { force: true }.'
    );
    this.name = 'ContentControlBoundError';
  }
}

/**
 * Control types whose content is free-form and safe to replace with text/blocks.
 * Typed controls (dropdown, date, checkbox, picture) carry structured state that
 * arbitrary content would contradict, and `group` exists to lock/contain nested
 * structure — all gated unless forced.
 */
const TEXT_REPLACEABLE_TYPES = new Set<SdtType>(['richText', 'plainText', 'unknown']);

/** True if free text/block content can safely replace this control type's content. */
export function isTextReplaceable(sdtType: SdtType): boolean {
  return TEXT_REPLACEABLE_TYPES.has(sdtType);
}

/** `w:lock` values that forbid editing the control's content. */
export function isContentLocked(lock: SdtProperties['lock']): boolean {
  return lock === 'contentLocked' || lock === 'sdtContentLocked';
}

/** `w:lock` values that forbid deleting the control. */
export function isDeletionLocked(lock: SdtProperties['lock']): boolean {
  return lock === 'sdtLocked' || lock === 'sdtContentLocked';
}

/**
 * True if the raw `w:sdtPr` carries a (w15) repeating-section structure. Matches
 * the element name (`<w15:repeatingSection>` / `<w15:repeatingSectionItem>`) so
 * a tag/alias value that merely contains the word doesn't false-match.
 */
export function hasRepeatingSection(props: SdtProperties): boolean {
  return /<w15:repeatingSection(Item)?[\s/>]/.test(props.rawPropertiesXml ?? '');
}

/** True if the control is bound to a Custom XML data store (`w:dataBinding`). */
export function isDataBound(props: SdtProperties): boolean {
  return props.dataBinding != null;
}

/**
 * Strip `<w:showingPlcHdr/>` from a raw `w:sdtPr` string. When real content is
 * written into a control that was showing its placeholder, the flag must go or
 * Word keeps rendering the (now-stale) placeholder styling over real content.
 */
export function clearShowingPlaceholderXml(raw: string | undefined): string | undefined {
  if (!raw) return raw;
  return raw
    .replace(/<w:showingPlcHdr\b[^>]*\/>/g, '')
    .replace(/<w:showingPlcHdr\b[^>]*>[\s\S]*?<\/w:showingPlcHdr>/g, '');
}

/** Properties for a control after real content is written: placeholder flag cleared. */
function propsAfterContentWrite(props: SdtProperties): SdtProperties {
  if (!props.showingPlaceholder && !/showingPlcHdr/.test(props.rawPropertiesXml ?? '')) {
    return props;
  }
  const next: SdtProperties = { ...props, showingPlaceholder: false };
  const cleaned = clearShowingPlaceholderXml(props.rawPropertiesXml);
  if (cleaned !== undefined) next.rawPropertiesXml = cleaned;
  return next;
}

function paragraph(text: string): BlockContent {
  return {
    type: 'paragraph',
    content: text ? [{ type: 'run', content: [{ type: 'text', text }] }] : [],
  };
}

/**
 * Turn a string into paragraphs (one per newline), or deep-clone block input.
 * A `plainText` control is single-paragraph in OOXML, so its string content is
 * collapsed to one paragraph rather than split — multiple paragraphs would make
 * Word repair the control on open.
 */
function toBlocks(
  replacement: string | BlockContent[],
  opts: { singleParagraph?: boolean } = {}
): BlockContent[] {
  if (typeof replacement !== 'string') {
    // Clone so the caller can't later mutate content shared with the result.
    return structuredClone(replacement);
  }
  if (opts.singleParagraph) return [paragraph(replacement)];
  return replacement.split('\n').map(paragraph);
}

export type ControlOp = (control: BlockSdt) => BlockContent[];

/**
 * Rebuild `blocks`, applying `op` to the first control matching `filter`. The
 * op's result (0, 1, or many blocks) is spliced in at the control's own level
 * — including when the control is nested inside another control — so a
 * remove/unwrap never leaves a placeholder behind. `state.done` stops the
 * walk after the first match.
 */
export function applyToFirst(
  blocks: BlockContent[],
  filter: ContentControlFilter,
  op: ControlOp,
  state: { done: boolean }
): BlockContent[] {
  const out: BlockContent[] = [];
  for (const block of blocks) {
    if (state.done) {
      out.push(block);
      continue;
    }
    // Controls are searched at body level and inside other controls. Table
    // cells are not searched: the current model types a cell as
    // (Paragraph | Table)[], and the table parser does not yet surface a
    // cell-level w:sdt (which OOXML's CT_Tc does permit) — see CONTENT-CONTROLS.md.
    if (block.type === 'blockSdt') {
      if (matches(block.properties, filter)) {
        out.push(...op(block));
        state.done = true;
        continue;
      }
      out.push({ ...block, content: applyToFirst(block.content, filter, op, state) });
    } else {
      out.push(block);
    }
  }
  return out;
}

export function rebuild(doc: Document, content: BlockContent[]): Document {
  return {
    ...doc,
    package: {
      ...doc.package,
      document: { ...doc.package.document, content },
    },
  };
}

/**
 * Replace the content of the first control matching `filter`. `replacement`
 * may be a string (split into paragraphs on newlines) or block content. The
 * control's properties, tag/alias, and lossless raw `w:sdtPr` are preserved —
 * only the contained blocks change, so the result still round-trips.
 *
 * When the control was showing its placeholder (`w:showingPlcHdr`), that flag
 * is cleared so Word doesn't render the new content as placeholder text.
 *
 * Throws {@link ContentControlNotFoundError} if nothing matches,
 * {@link ContentControlLockedError} if the control's lock forbids editing, and
 * {@link ContentControlTypeError} if the control is a typed (dropdown/date/…)
 * control whose value shouldn't be set as free text. Pass `{ force: true }` to
 * override the lock/type guards.
 */
export function setContentControlContent(
  doc: Document,
  filter: ContentControlFilter,
  replacement: string | BlockContent[],
  options: { force?: boolean } = {}
): Document {
  const state = { done: false };
  const op: ControlOp = (control) => {
    const props = control.properties;
    if (!options.force && isContentLocked(props.lock)) {
      throw new ContentControlLockedError(props.lock, 'edit');
    }
    if (!options.force && !isTextReplaceable(props.sdtType)) {
      throw new ContentControlTypeError(props.sdtType);
    }
    if (!options.force && isDataBound(props)) {
      throw new ContentControlBoundError();
    }
    return [
      {
        ...control,
        properties: propsAfterContentWrite(props),
        content: toBlocks(replacement, { singleParagraph: props.sdtType === 'plainText' }),
      },
    ];
  };
  const content = applyToFirst(doc.package.document.content, filter, op, state);
  if (!state.done) throw new ContentControlNotFoundError(filter);
  return rebuild(doc, content);
}

/**
 * Remove the first control matching `filter` from the document. With
 * `keepContent: true` the control's blocks are unwrapped in place (the box
 * goes away, the content stays) — useful for "resolve this conditional
 * section into plain content". Otherwise the whole region is deleted.
 *
 * Unwrapping a repeating-section (item) is refused unless `force`, since
 * lifting its blocks out would orphan the (w15) repeating structure.
 *
 * Throws {@link ContentControlNotFoundError} / {@link ContentControlLockedError}
 * as {@link setContentControlContent} does.
 */
export function removeContentControl(
  doc: Document,
  filter: ContentControlFilter,
  options: { force?: boolean; keepContent?: boolean } = {}
): Document {
  const state = { done: false };
  const op: ControlOp = (control) => {
    if (!options.force && isDeletionLocked(control.properties.lock)) {
      throw new ContentControlLockedError(control.properties.lock, 'remove');
    }
    if (options.keepContent && !options.force && hasRepeatingSection(control.properties)) {
      throw new ContentControlLockedError(control.properties.lock, 'remove');
    }
    return options.keepContent ? control.content : [];
  };
  const content = applyToFirst(doc.package.document.content, filter, op, state);
  if (!state.done) throw new ContentControlNotFoundError(filter);
  // Never leave a structurally empty body (matches the live-editor path, which
  // auto-fills a paragraph). An empty <w:body> is invalid for Word consumers.
  const safe = content.length > 0 ? content : [{ type: 'paragraph' as const, content: [] }];
  return rebuild(doc, safe);
}
