/**
 * Formal Word JS API parity contract.
 *
 * This file declares what subset of Microsoft Word's Office.js JavaScript API
 * (https://learn.microsoft.com/en-us/javascript/api/word) we mirror, **at the
 * type level**. It is checked at compile time: `EditorBridge` must satisfy
 * `WordCompatBridge` (verified via the static assertion at the bottom).
 *
 * The assertion is the source of truth. If you change the bridge surface and
 * forget to update parity, typecheck breaks.
 *
 * ## What we mirror (✓)
 *
 * | Word JS API                         | Our equivalent                        |
 * | ----------------------------------- | ------------------------------------- |
 * | `Range` (stable handle)             | `paraId: string`                      |
 * | `body.search(text)` → Ranges        | `findText(query, opts) → FoundMatch[]`|
 * | `range.insertComment(text)`         | `addComment({paraId, text, search?})` |
 * | `comment.reply(text)`               | `replyTo(commentId, {text})`          |
 * | `comment.resolved = true`           | `resolveComment(commentId)`           |
 * | `range.insertText(text, location)`  | `proposeChange({paraId, search, replaceWith})` (3 modes via empty-string semantics) |
 * | `document.getSelection()`           | `getSelection() → SelectionInfo|null` |
 * | `range.scrollIntoView()`            | `scrollTo(paraId)`                    |
 * | `commentCollection.getItems()`      | `getComments(filter?)`                |
 * | `body.paragraphs.getItems()`        | `getContent(opts?)`                   |
 * | `document.body.text`                | `getContentAsText(opts?)`             |
 * | `revisionCollection.getItems()`     | `getChanges(filter?)`                 |
 * | `Document.onContentChanged.add(...)`| `onContentChange(listener)`           |
 * | `Document.onSelectionChanged.add()` | `onSelectionChange(listener)`         |
 * | `Range.font.bold` / `italic` / `color` / `size` / `name` | `applyFormatting({paraId, search?, marks})` |
 * | `ParagraphFormat.style` / `applyStyle(...)` | `setParagraphStyle({paraId, styleId})` |
 * | `body.insertTable(...)` / `range.insertTable(...)` | `insertTable({rows, columns, data?, paraId?})` |
 * | `range.insertInlinePictureFromBase64(...)` | `insertImage({src, width?, height?, paraId?})` |
 *
 * ## Beyond Word's surface (paged-document affordances)
 *
 *  - `getPage(n)` / `getPages({from, to})` / `getTotalPages()` — Word's JS API
 *    doesn't model pages as first-class addressable units. We do, because the
 *    editor is paged. Backed by the layout-painter's page boundary state.
 *
 * ## Differences (intentional, documented)
 *
 *  - All "insertText" overloads collapse into `proposeChange` with empty-string
 *    semantics: replaceWith="" deletes; search="" inserts at paragraph end;
 *    both non-empty replaces. Word has separate `insertText(...,'Replace')`,
 *    `insertText(...,'Before')`, etc.; we found three modes were enough and
 *    serialize cleaner for LLM tool calls.
 *  - Tracked changes are always *suggestions* in our world. Word lets the
 *    range mutate directly; we always go through the tracked-change path so
 *    the human keeps the final say. (This matches the agent UX.)
 *  - Word's `Range.context.sync()` is unnecessary — every call is one PM
 *    transaction.
 *  - Word's `Range` is a stateful object. Ours is a plain `{paraId, search?}`
 *    JSON value, so it survives JSON.stringify and works for tool calls / MCP.
 *
 * ## Out of scope (gaps we deliberately don't implement)
 *
 *  - Paragraph creation (`body.insertParagraph`). Out of scope for v1.
 *  - Table mutation after insertion (insert row/col, delete cell).
 *  - Headers / footers / sections.
 *  - `Range.getOoxml()` / `getHtml()`. Plain text only.
 *  - `customXmlParts` / `contentControls`.
 *  - Accept / reject tracked changes. Human-only by design.
 *
 * Future versions can grow these by extending this interface — typecheck
 * will then enforce the new contract.
 */

import type { EditorBridge, ContentChangeEvent, SelectionChangeEvent } from './bridge';
import type {
  ContentBlock,
  GetContentOptions,
  ReviewComment,
  ReviewChange,
  ChangeFilter,
  CommentFilter,
  AddCommentByParaIdOptions,
  ReplyOptions,
  ProposeChangeOptions,
  FoundMatch,
  SelectionInfo,
  ApplyFormattingOptions,
  SetParagraphStyleOptions,
  InsertTableOptions,
  InsertImageOptions,
  PageContent,
} from './types';

/**
 * The formal Word-JS-API parity surface. Each method maps to one or more
 * Word API methods (named in the JSDoc above each member).
 *
 * If you change the EditorBridge surface, this interface must change too —
 * the static assertion at the bottom of the file enforces it.
 */
export interface WordCompatBridge {
  // ── Read ────────────────────────────────────────────────────────────────

  /** Word: `document.body.text` (stringified, indexed lines). */
  getContentAsText(options?: GetContentOptions): string;

  /** Word: `body.paragraphs.getItems()`. */
  getContent(options?: GetContentOptions): ContentBlock[];

  /** Word: `commentCollection.getItems()`. */
  getComments(filter?: CommentFilter): ReviewComment[];

  /** Word: `revisionCollection.getItems()`. */
  getChanges(filter?: ChangeFilter): ReviewChange[];

  /** Word: `body.search(text)` returning Range[]. */
  findText(query: string, options?: { caseSensitive?: boolean; limit?: number }): FoundMatch[];

  /** Word: `document.getSelection()`. */
  getSelection(): SelectionInfo | null;

  // ── Mutate (always tracked-change-style suggestions) ─────────────────────

  /** Word: `range.insertComment(text)`. Anchored by paraId (Range handle). */
  addComment(options: AddCommentByParaIdOptions): number | null;

  /** Word: `comment.reply(text)`. */
  replyTo(commentId: number, options: ReplyOptions): number | null;

  /** Word: `comment.resolved = true`. */
  resolveComment(commentId: number): void;

  /**
   * Word: `range.insertText(text, location)` collapsed into one verb.
   *  - replacement: search non-empty, replaceWith non-empty
   *  - deletion:    search non-empty, replaceWith ""
   *  - insertion:   search "",        replaceWith non-empty (paragraph end)
   */
  proposeChange(options: ProposeChangeOptions): boolean;

  /**
   * Word: `range.font.bold = true` / `range.font.italic = true` / etc.
   * Applied to a paragraph or to a unique phrase within it. Direct edit,
   * not a tracked change.
   */
  applyFormatting(options: ApplyFormattingOptions): boolean;

  /** Word: `paragraph.styleBuiltIn = ...` / `paragraph.style = 'Heading 1'`. */
  setParagraphStyle(options: SetParagraphStyleOptions): boolean;

  /**
   * Word: `body.insertTable(...)` / `range.insertTable(...)`.
   * Inserts at the cursor, or after `paraId` when supplied.
   */
  insertTable(options: InsertTableOptions): boolean;

  /**
   * Word: `range.insertInlinePictureFromBase64(...)`.
   * Accepts a data URL so DOCX export can embed the binary image.
   */
  insertImage(options: InsertImageOptions): boolean;

  // ── Paged document affordances (no Word equivalent) ─────────────────────

  /**
   * Read one rendered page (1-indexed). Word's JS API does not expose pages
   * as first-class objects; we do because the editor is paged.
   */
  getPage(pageNumber: number): PageContent | null;

  /** Read a range of rendered pages (1-indexed, inclusive). */
  getPages(options: { from: number; to: number }): PageContent[];

  /** Total number of pages currently rendered. */
  getTotalPages(): number;

  // ── Navigate ─────────────────────────────────────────────────────────────

  /** Word: `range.scrollIntoView()`. */
  scrollTo(paraId: string): boolean;

  // ── Events ───────────────────────────────────────────────────────────────

  /** Word: `Document.onContentChanged.add(handler)`. Returns unsubscribe. */
  onContentChange(listener: (event: ContentChangeEvent) => void): () => void;

  /** Word: `Document.onSelectionChanged.add(handler)`. Returns unsubscribe. */
  onSelectionChange(listener: (event: SelectionChangeEvent) => void): () => void;
}

// ── Compile-time enforcement ────────────────────────────────────────────────
//
// Mirrors `EditorBridge` against the parity contract in both directions:
// (a) every method declared in WordCompatBridge must exist on EditorBridge
//     with a compatible signature (forward check)
// (b) EditorBridge can have extra methods (it does — DocxReviewer-shaped APIs).
//
// If the bridge ever drops a method or changes its shape incompatibly, this
// fails to typecheck. That's the formal parity check.

type ImplementsWordCompat<T extends WordCompatBridge> = T;
/** Static assertion: EditorBridge satisfies the parity contract. Failing this
 * means the bridge dropped or changed a method that's part of the public
 * Word-API mirror. Either fix the bridge or update WordCompatBridge to
 * narrow scope deliberately. The unused-export comment keeps tsc quiet. */

export type _AssertEditorBridgeImplementsWordCompat = ImplementsWordCompat<EditorBridge>;
