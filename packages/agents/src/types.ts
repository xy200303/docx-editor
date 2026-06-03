// ============================================================================
// CONTENT BLOCKS — what getContent() returns
// ============================================================================

export interface HeadingBlock {
  type: 'heading';
  index: number;
  /** Stable Word `w14:paraId`. Use this as the anchor for live-editor operations. */
  paraId?: string;
  level: number;
  text: string;
}

export interface ParagraphBlock {
  type: 'paragraph';
  index: number;
  /** Stable Word `w14:paraId`. Use this as the anchor for live-editor operations. */
  paraId?: string;
  text: string;
}

export interface TableCellBlock {
  type: 'table-cell';
  index: number;
  /** Stable Word `w14:paraId` of the cell's first paragraph. */
  paraId?: string;
  row: number;
  col: number;
  text: string;
}

export interface TableBlock {
  type: 'table';
  index: number;
  rows: string[][];
  /** Per-cell paraIds, parallel to `rows`. cells[r][c] is the first paraId in that cell. */
  cellParaIds?: (string | undefined)[][];
}

export interface ListItemBlock {
  type: 'list-item';
  index: number;
  /** Stable Word `w14:paraId`. Use this as the anchor for live-editor operations. */
  paraId?: string;
  text: string;
  listLevel: number;
  listType: 'bullet' | 'number';
}

export type ContentBlock = HeadingBlock | ParagraphBlock | TableBlock | ListItemBlock;

export interface GetContentOptions {
  fromIndex?: number;
  toIndex?: number;
  /** Annotate tracked changes inline. Default: true */
  includeTrackedChanges?: boolean;
  /** Annotate comments inline. Default: true */
  includeCommentAnchors?: boolean;
}

// ============================================================================
// DISCOVERY — what getChanges() / getComments() return
// ============================================================================

export interface ReviewChange {
  id: number;
  type: 'insertion' | 'deletion' | 'moveFrom' | 'moveTo';
  author: string;
  date: string | null;
  text: string;
  context: string;
  /**
   * Index of the containing paragraph. For body changes this is the
   * document-wide paragraph index; for note changes it is the paragraph index
   * *within that note* (note bodies have their own numbering), so pair it with
   * `noteId` / `noteType` rather than reading it as a body index.
   */
  paragraphIndex: number;
  /**
   * Set when the change lives inside a footnote or endnote. Such changes are
   * surfaced for discovery only — accept/reject operate on the body, so an id
   * that resolves *only* to a note change throws `NoteChangeNotEditableError`
   * (an id also present on a body change resolves to the body change). The
   * raw `id` is not namespaced across parts, so pair it with `noteId` /
   * `noteType` to identify the change.
   */
  noteId?: number;
  /** Which note store the change came from. Absent for body changes. */
  noteType?: 'footnote' | 'endnote';
}

export interface ReviewCommentReply {
  id: number;
  author: string;
  date: string | null;
  text: string;
}

export interface ReviewComment {
  id: number;
  author: string;
  date: string | null;
  text: string;
  anchoredText: string;
  paragraphIndex: number;
  replies: ReviewCommentReply[];
  done: boolean;
}

export interface ChangeFilter {
  author?: string;
  type?: 'insertion' | 'deletion' | 'moveFrom' | 'moveTo';
  /** Also report tracked changes inside footnote bodies. Default: false. */
  includeFootnotes?: boolean;
  /** Also report tracked changes inside endnote bodies. Default: false. */
  includeEndnotes?: boolean;
}

export interface CommentFilter {
  author?: string;
  done?: boolean;
}

// ============================================================================
// ACTION OPTIONS — author is optional (falls back to reviewer default)
// ============================================================================

/**
 * Live-editor (bridge) action options — anchored by Word `w14:paraId`.
 * Stable across edits; the agent gets paraIds from `read_document` / `find_text`.
 */
export interface AddCommentByParaIdOptions {
  paraId: string;
  text: string;
  author?: string;
  /** Optional: anchor to a specific phrase within the paragraph (must be unique). */
  search?: string;
}

/**
 * Headless / DocxReviewer action options — anchored by ordinal paragraph index.
 * Used by the static-document review pipeline.
 */
export interface AddCommentOptions {
  paragraphIndex: number;
  text: string;
  author?: string;
  /** Optional: anchor to specific text. Omit to anchor whole paragraph. */
  search?: string;
}

export interface ReplyOptions {
  text: string;
  author?: string;
}

/**
 * Live-editor change. Pass `replaceWith: ''` to delete; pass `search: ''` to
 * insert at end of paragraph; pass both non-empty for a replacement.
 */
export interface ProposeChangeOptions {
  paraId: string;
  search: string;
  replaceWith: string;
  author?: string;
}

export interface ProposeReplacementOptions {
  paragraphIndex: number;
  search: string;
  replaceWith: string;
  author?: string;
}

export interface ProposeInsertionOptions {
  paragraphIndex: number;
  insertText: string;
  author?: string;
  position?: 'before' | 'after';
  search?: string;
}

export interface ProposeDeletionOptions {
  paragraphIndex: number;
  search: string;
  author?: string;
}

/** Per-match handle returned by `findText` — pass `paraId` + `match` back as `search`. */
export interface FoundMatch {
  paraId: string;
  match: string;
  before: string;
  after: string;
}

/** @public */
export interface SelectionInfo {
  paraId: string | null;
  selectedText: string;
  paragraphText: string;
  before: string;
  after: string;
}

/**
 * Character formatting marks the agent can apply.
 *
 * Mirrors Word JS API `Range.font.*`. A `false` value clears that mark in the
 * range; a missing key leaves it untouched. `color.themeColor` follows ECMA-376
 * theme color values (e.g. `'accent1'`, `'text1'`) and resolves at render time.
 */
export interface CharacterFormatting {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean | { style?: string };
  strike?: boolean;
  color?: { rgb?: string; themeColor?: string };
  highlight?: string;
  fontSize?: number;
  fontFamily?: { ascii?: string; hAnsi?: string };
}

/**
 * Apply character formatting to a range. `paraId` is required; if `search`
 * is provided, the formatting only applies to that phrase within the
 * paragraph (must match exactly once). Otherwise it applies to the whole
 * paragraph's text.
 */
export interface ApplyFormattingOptions {
  paraId: string;
  search?: string;
  marks: CharacterFormatting;
}

/**
 * Apply a paragraph style by `styleId` (e.g. `'Heading1'`, `'Title'`,
 * `'Quote'`). The styleId must exist in the document's style definitions
 * — unknown ids are no-ops.
 */
export interface SetParagraphStyleOptions {
  paraId: string;
  styleId: string;
}

export type InsertTextPosition =
  | 'cursor'
  | 'paragraph_start'
  | 'paragraph_end'
  | 'before_paragraph'
  | 'after_paragraph';

export type InsertTextPlacement = 'before' | 'after' | 'replace';

/**
 * Direct text insertion. This edits the document immediately; it does not add
 * comments or create tracked-change suggestions.
 */
export interface InsertTextOptions {
  text: string;
  /** Optional paragraph anchor. Omit to insert at the current cursor/selection. */
  paraId?: string;
  /**
   * Paragraph-relative placement when `paraId` is supplied and `search` is not.
   * Defaults to `paragraph_end` with a paraId, otherwise `cursor`.
   */
  position?: InsertTextPosition;
  /** Optional unique phrase inside `paraId` to insert around or replace. */
  search?: string;
  /** Search-relative placement. Defaults to `after` when `search` is supplied. */
  placement?: InsertTextPlacement;
}

/**
 * Direct text replacement/deletion in a paragraph. This edits the document
 * immediately; it does not add comments or create tracked-change suggestions.
 */
export interface ReplaceTextOptions {
  paraId: string;
  search: string;
  replaceWith: string;
}

/** Insert a table at the current cursor, or after a paragraph when `paraId` is supplied. */
export interface InsertTableOptions {
  rows: number;
  columns: number;
  data?: string[][];
  hasHeader?: boolean;
  paraId?: string;
}

/** Insert an inline image at the current cursor, or at the end of `paraId` when supplied. */
export interface InsertImageOptions {
  /** Data URL, e.g. `data:image/png;base64,...`, so DOCX export can embed it. */
  src: string;
  alt?: string;
  width?: number;
  height?: number;
  paraId?: string;
}

export type ContentControlType =
  | 'richText'
  | 'plainText'
  | 'date'
  | 'dropDownList'
  | 'comboBox'
  | 'checkbox'
  | 'picture'
  | 'buildingBlockGallery'
  | 'group'
  | 'equation'
  | 'citation'
  | 'bibliography'
  | 'unknown';

/** Filter used to address Word content controls / SDTs by stable metadata. */
export interface ContentControlFilter {
  tag?: string;
  alias?: string;
  id?: number;
  type?: ContentControlType;
}

export interface ContentControlInfo extends ContentControlFilter {
  sdtType: ContentControlType;
  lock?: 'sdtLocked' | 'contentLocked' | 'sdtContentLocked' | 'unlocked';
  showingPlaceholder?: boolean;
  checked?: boolean;
  dateFormat?: string;
  listItems?: { displayText: string; value: string }[];
  dataBinding?: {
    xpath?: string;
    storeItemID?: string;
    prefixMappings?: string;
  };
  text: string;
  pos?: number;
  depth?: number;
}

/** Typed value to apply to a Word content control / SDT. */
export type ContentControlValue =
  | { kind: 'dropdown'; value: string }
  | { kind: 'checkbox'; checked: boolean }
  | { kind: 'date'; date: string };

export interface SetContentControlOptions extends ContentControlFilter {
  text: string;
  force?: boolean;
}

export interface SetContentControlValueOptions extends ContentControlFilter {
  value: ContentControlValue;
  force?: boolean;
}

export interface RemoveContentControlOptions extends ContentControlFilter {
  force?: boolean;
  keepContent?: boolean;
}

/** A single paragraph anchored on a page (returned by `getPage` / `getPages`). */
export interface PageParagraph {
  paraId: string;
  text: string;
  /** True for headings, list items, and other styled blocks. */
  styleId?: string;
}

/** What the agent sees when reading one or more pages. */
export interface PageContent {
  /** 1-indexed page number. */
  pageNumber: number;
  /** Plain text of the page, formatted as `[paraId] text` lines. */
  text: string;
  /** Paragraphs on the page, in document order. */
  paragraphs: PageParagraph[];
}

/**
 * Snapshot of what the user is looking at — pass this to your agent's system
 * prompt so it knows the current selection / page without an extra
 * `read_selection` round-trip.
 *
 * @public
 */
export interface AgentContextSnapshot {
  /** User's current selection or cursor (null if editor isn't focused). */
  selection: SelectionInfo | null;
  /** 1-indexed page the cursor / selection is on. 0 if unknown. */
  currentPage: number;
  /** Total number of pages currently rendered. */
  totalPages: number;
}

// ============================================================================
// BATCH — the main LLM-facing interface
// ============================================================================

export interface BatchReviewOptions {
  accept?: number[];
  reject?: number[];
  comments?: AddCommentOptions[];
  replies?: (ReplyOptions & { commentId: number })[];
  proposals?: ProposeReplacementOptions[];
}

export interface BatchError {
  operation: string;
  id?: number;
  search?: string;
  error: string;
}

export interface BatchResult {
  accepted: number;
  rejected: number;
  commentsAdded: number;
  repliesAdded: number;
  proposalsAdded: number;
  errors: BatchError[];
}
