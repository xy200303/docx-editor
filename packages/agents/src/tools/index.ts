/**
 * Agent tool definitions and execution.
 *
 * Tools use OpenAI function-calling format. The pattern mirrors Word's JS API:
 * locate first (`read_document` / `find_text` / `read_selection` return paraId
 * handles), then mutate (`comment` / `suggest_change` / `reply_comment` /
 * `resolve_comment` / `scroll`). paraId anchors are stable across edits.
 */

export type { AgentToolDefinition, AgentToolResult } from './types';
import type { AgentToolDefinition, AgentToolResult } from './types';
import type { EditorBridge } from '../bridge';
import type {
  ContentControlFilter,
  ContentControlType,
  InsertTextPlacement,
  InsertTextPosition,
} from '../types';
import { applyFormatting, setParagraphStyle } from './formatting';
import { readPage, readPages } from './pages';

const INSERT_TEXT_POSITIONS: InsertTextPosition[] = [
  'cursor',
  'paragraph_start',
  'paragraph_end',
  'before_paragraph',
  'after_paragraph',
];

const INSERT_TEXT_PLACEMENTS: InsertTextPlacement[] = ['before', 'after', 'replace'];

const CONTENT_CONTROL_TYPES: ContentControlType[] = [
  'richText',
  'plainText',
  'date',
  'dropDownList',
  'comboBox',
  'checkbox',
  'picture',
  'buildingBlockGallery',
  'group',
  'equation',
  'citation',
  'bibliography',
  'unknown',
];

function nonEmptyString(value: unknown): value is string {
  return typeof value === 'string' && value.trim().length > 0;
}

function stringOrUndefined(value: unknown): string | undefined {
  return typeof value === 'string' ? value : undefined;
}

function parseContentControlFilter(
  input: Record<string, unknown>,
  requireAnchor: boolean
): { filter: ContentControlFilter; error?: string } {
  const filter: ContentControlFilter = {};
  if (input.tag !== undefined) {
    if (!nonEmptyString(input.tag)) return { filter, error: '`tag` must be a non-empty string.' };
    filter.tag = input.tag;
  }
  if (input.alias !== undefined) {
    if (!nonEmptyString(input.alias)) {
      return { filter, error: '`alias` must be a non-empty string.' };
    }
    filter.alias = input.alias;
  }
  if (input.id !== undefined) {
    const id = Number(input.id);
    if (!Number.isInteger(id)) return { filter, error: '`id` must be an integer.' };
    filter.id = id;
  }
  if (input.type !== undefined) {
    if (!CONTENT_CONTROL_TYPES.includes(input.type as ContentControlType)) {
      return { filter, error: '`type` must be a supported content-control type.' };
    }
    filter.type = input.type as ContentControlType;
  }

  if (
    requireAnchor &&
    filter.tag === undefined &&
    filter.alias === undefined &&
    filter.id === undefined &&
    filter.type === undefined
  ) {
    return {
      filter,
      error: 'Provide at least one content-control anchor: `tag`, `alias`, `id`, or `type`.',
    };
  }

  return { filter };
}

// ── Locate tools ────────────────────────────────────────────────────────────

const readDocument: AgentToolDefinition<{ fromIndex?: number; toIndex?: number }> = {
  name: 'read_document',
  displayName: 'Reading document',
  description:
    'Read the document content. Returns lines tagged with a stable paragraph id, e.g. ' +
    '"[2A1F3B] First paragraph". Use the bracketed id as `paraId` when commenting or ' +
    'suggesting changes — it survives edits, unlike ordinal indices. ' +
    'Returns the vanilla document (the doc as it exists right now, before any tracked ' +
    'suggestions are accepted): pending insertions are HIDDEN, pending deletions are ' +
    'shown as plain text (still part of the document until accepted), and comment ' +
    'markers are stripped. Use read_changes / read_comments to inspect what is pending.',
  inputSchema: {
    type: 'object',
    properties: {
      fromIndex: { type: 'number', description: 'Start ordinal index (inclusive). Optional.' },
      toIndex: { type: 'number', description: 'End ordinal index (inclusive). Optional.' },
    },
  },
  handler: (input, bridge) => {
    const text = bridge.getContentAsText({
      fromIndex: input.fromIndex,
      toIndex: input.toIndex,
      includeTrackedChanges: false,
      includeCommentAnchors: false,
    });
    return { success: true, data: text };
  },
};

const readSelection: AgentToolDefinition = {
  name: 'read_selection',
  displayName: 'Reading selection',
  description:
    "Read the user's current cursor or selection. Returns the selected text, the " +
    "paragraph it lives in, and that paragraph's `paraId`. Use this when the user " +
    'asks "fix this" or "review what I have selected".',
  inputSchema: { type: 'object', properties: {} },
  handler: (_input, bridge) => {
    const sel = bridge.getSelection();
    if (!sel) return { success: false, error: 'No selection (editor not focused).' };
    return { success: true, data: sel };
  },
};

const findText: AgentToolDefinition<{
  query: string;
  caseSensitive?: boolean;
  limit?: number;
}> = {
  name: 'find_text',
  displayName: 'Finding text',
  description:
    'Locate paragraphs containing `query`. Returns up to `limit` handles, each with ' +
    '`paraId`, the matched substring, and surrounding context. Pass any returned ' +
    '`paraId` (and the `match` as `search`) to add_comment / suggest_change.',
  inputSchema: {
    type: 'object',
    properties: {
      query: { type: 'string', description: 'Text to find (substring match).' },
      caseSensitive: { type: 'boolean', description: 'Default: false.' },
      limit: { type: 'number', description: 'Max paragraphs to return. Default: 20.' },
    },
    required: ['query'],
  },
  handler: (input, bridge) => {
    const matches = bridge.findText(input.query, {
      caseSensitive: input.caseSensitive,
      limit: input.limit,
    });
    if (matches.length === 0) return { success: true, data: 'No matches.' };
    return { success: true, data: matches };
  },
};

const readComments: AgentToolDefinition = {
  name: 'read_comments',
  displayName: 'Reading comments',
  description: 'List all comments in the document with their paragraph anchors.',
  inputSchema: { type: 'object', properties: {} },
  handler: (_input, bridge) => {
    const comments = bridge.getComments();
    if (comments.length === 0) return { success: true, data: 'No comments.' };
    const text = comments
      .map(
        (c) =>
          `[Comment #${c.id}] ${c.author}: "${c.text}"` +
          (c.anchoredText ? ` (anchored to: "${c.anchoredText}")` : '') +
          (c.replies.length > 0
            ? '\n' + c.replies.map((r) => `  Reply by ${r.author}: "${r.text}"`).join('\n')
            : '')
      )
      .join('\n');
    return { success: true, data: text };
  },
};

const readChanges: AgentToolDefinition = {
  name: 'read_changes',
  displayName: 'Reading changes',
  description: 'List tracked changes (insertions / deletions) currently in the document.',
  inputSchema: { type: 'object', properties: {} },
  handler: (_input, bridge) => {
    const changes = bridge.getChanges();
    if (changes.length === 0) return { success: true, data: 'No tracked changes.' };
    const text = changes
      .map((c) => `[Change #${c.id}] ${c.type} by ${c.author}: "${c.text}"`)
      .join('\n');
    return { success: true, data: text };
  },
};

const readContentControls: AgentToolDefinition<{
  tag?: string;
  alias?: string;
  id?: number;
  type?: ContentControlType;
}> = {
  name: 'read_content_controls',
  displayName: 'Reading content controls',
  description:
    'List Word content controls / SDTs in the live document. Use this for template-style ' +
    'documents where stable `tag`, `alias`, or `id` anchors are better than paragraph text search.',
  inputSchema: {
    type: 'object',
    properties: {
      tag: { type: 'string', description: 'Optional Word content-control tag.' },
      alias: { type: 'string', description: 'Optional content-control alias/title.' },
      id: { type: 'number', description: 'Optional numeric content-control id.' },
      type: {
        type: 'string',
        enum: CONTENT_CONTROL_TYPES,
        description: 'Optional content-control type filter.',
      },
    },
  },
  handler: (input, bridge) => {
    const { filter, error } = parseContentControlFilter(input, false);
    if (error) return { success: false, error };

    const controls = bridge.getContentControls(filter);
    if (controls.length === 0) return { success: true, data: 'No content controls.' };
    return {
      success: true,
      data: controls.map((control) => ({
        tag: control.tag,
        alias: control.alias,
        id: control.id,
        type: control.sdtType,
        lock: control.lock,
        showingPlaceholder: control.showingPlaceholder,
        text: control.text,
      })),
    };
  },
};

// ── Mutate tools ────────────────────────────────────────────────────────────

const addComment: AgentToolDefinition<{
  paraId: string;
  text: string;
  search?: string;
}> = {
  name: 'add_comment',
  displayName: 'Adding comment',
  description:
    'Attach a comment to a paragraph, optionally anchored to a unique phrase within ' +
    'it. The user sees it instantly in the comments sidebar.',
  inputSchema: {
    type: 'object',
    properties: {
      paraId: { type: 'string', description: 'Paragraph id from read_document / find_text.' },
      text: { type: 'string', description: 'Comment body.' },
      search: {
        type: 'string',
        description: 'Optional: anchor to this exact phrase within the paragraph. Must be unique.',
      },
    },
    required: ['paraId', 'text'],
  },
  handler: (input, bridge) => {
    const id = bridge.addComment({
      paraId: input.paraId,
      text: input.text,
      search: input.search,
    });
    if (id === null) {
      return {
        success: false,
        error:
          'Could not add comment. The paraId may not exist, or `search` is missing / ambiguous.',
      };
    }
    return { success: true, data: `Comment ${id} added on ${input.paraId}.` };
  },
};

const suggestChange: AgentToolDefinition<{
  paraId: string;
  search: string;
  replaceWith: string;
}> = {
  name: 'suggest_change',
  displayName: 'Suggesting change',
  description:
    'Suggest a tracked change. Three modes: ' +
    '(1) replacement — `search` non-empty, `replaceWith` non-empty; ' +
    '(2) deletion — `search` non-empty, `replaceWith` empty; ' +
    '(3) insertion at paragraph end — `search` empty, `replaceWith` non-empty. ' +
    'The user can accept or reject in the editor UI.',
  inputSchema: {
    type: 'object',
    properties: {
      paraId: { type: 'string', description: 'Paragraph id from read_document / find_text.' },
      search: {
        type: 'string',
        description: 'Phrase to find (must be unique). Empty string = insert at paragraph end.',
      },
      replaceWith: {
        type: 'string',
        description: 'Replacement text. Empty string = delete the matched phrase.',
      },
    },
    required: ['paraId', 'search', 'replaceWith'],
  },
  handler: (input, bridge) => {
    const ok = bridge.proposeChange({
      paraId: input.paraId,
      search: input.search,
      replaceWith: input.replaceWith,
    });
    if (!ok) {
      return {
        success: false,
        error:
          'Could not propose change. Possible causes: paraId not found; search missing or ' +
          'ambiguous; or the target overlaps an existing tracked change.',
      };
    }
    if (!input.search) return { success: true, data: `Insertion proposed on ${input.paraId}.` };
    if (!input.replaceWith) {
      return { success: true, data: `Deletion proposed: "${input.search}" on ${input.paraId}.` };
    }
    return {
      success: true,
      data: `Replacement proposed: "${input.search}" → "${input.replaceWith}" on ${input.paraId}.`,
    };
  },
};

const insertTextTool: AgentToolDefinition<{
  text: string;
  paraId?: string;
  position?: InsertTextPosition;
  search?: string;
  placement?: InsertTextPlacement;
}> = {
  name: 'insert_text',
  displayName: 'Inserting text',
  description:
    'Directly insert text into the live document without creating a comment or tracked-change ' +
    'suggestion. Use this for normal edit requests like "add this paragraph", "write an intro", ' +
    'or "insert this sentence". Omit `paraId` to insert at the current cursor/selection. With ' +
    '`paraId`, default insertion is at the end of that paragraph. With `search`, insert before, ' +
    'after, or replace that unique phrase inside the paragraph.',
  inputSchema: {
    type: 'object',
    properties: {
      text: { type: 'string', description: 'Text to insert.' },
      paraId: {
        type: 'string',
        description: 'Optional paragraph id from read_document / find_text.',
      },
      position: {
        type: 'string',
        enum: INSERT_TEXT_POSITIONS,
        description:
          'Where to insert when `search` is omitted. Defaults to cursor without paraId, paragraph_end with paraId.',
      },
      search: {
        type: 'string',
        description: 'Optional unique phrase inside `paraId` to insert around or replace.',
      },
      placement: {
        type: 'string',
        enum: INSERT_TEXT_PLACEMENTS,
        description: 'Search-relative placement. Defaults to after.',
      },
    },
    required: ['text'],
  },
  handler: (input, bridge) => {
    if (typeof input.text !== 'string' || input.text.length === 0) {
      return { success: false, error: '`text` must be a non-empty string.' };
    }
    const paraId = stringOrUndefined(input.paraId);
    if (input.paraId !== undefined && !nonEmptyString(input.paraId)) {
      return { success: false, error: '`paraId` must be a non-empty string when provided.' };
    }
    if (
      input.position !== undefined &&
      !INSERT_TEXT_POSITIONS.includes(input.position as InsertTextPosition)
    ) {
      return { success: false, error: '`position` must be a supported insert position.' };
    }
    if (
      input.placement !== undefined &&
      !INSERT_TEXT_PLACEMENTS.includes(input.placement as InsertTextPlacement)
    ) {
      return { success: false, error: '`placement` must be before, after, or replace.' };
    }
    if (input.search !== undefined && !nonEmptyString(input.search)) {
      return { success: false, error: '`search` must be a non-empty string when provided.' };
    }
    if (input.search !== undefined && !paraId) {
      return { success: false, error: '`search` requires a `paraId` anchor.' };
    }

    const ok = bridge.insertText({
      text: input.text,
      paraId,
      position: input.position,
      search: input.search,
      placement: input.placement,
    });
    if (!ok) {
      return {
        success: false,
        error:
          'Could not insert text. Possible causes: editor not ready; paraId not found; ' +
          'search missing / ambiguous; or this adapter does not support direct text edits.',
      };
    }
    return { success: true, data: paraId ? `Inserted text in ${paraId}.` : 'Inserted text.' };
  },
};

const replaceTextTool: AgentToolDefinition<{
  paraId: string;
  search: string;
  replaceWith: string;
}> = {
  name: 'replace_text',
  displayName: 'Replacing text',
  description:
    'Directly replace or delete a unique phrase in a paragraph without creating a comment or ' +
    'tracked-change suggestion. Use suggest_change instead only when the user explicitly wants reviewable changes.',
  inputSchema: {
    type: 'object',
    properties: {
      paraId: { type: 'string', description: 'Paragraph id from read_document / find_text.' },
      search: {
        type: 'string',
        description: 'Unique phrase to replace within the paragraph.',
      },
      replaceWith: {
        type: 'string',
        description: 'Replacement text. Empty string deletes the matched phrase.',
      },
    },
    required: ['paraId', 'search', 'replaceWith'],
  },
  handler: (input, bridge) => {
    if (!nonEmptyString(input.paraId)) {
      return { success: false, error: '`paraId` must be a non-empty string.' };
    }
    if (!nonEmptyString(input.search)) {
      return { success: false, error: '`search` must be a non-empty string.' };
    }
    if (typeof input.replaceWith !== 'string') {
      return { success: false, error: '`replaceWith` must be a string.' };
    }

    const ok = bridge.replaceText({
      paraId: input.paraId,
      search: input.search,
      replaceWith: input.replaceWith,
    });
    if (!ok) {
      return {
        success: false,
        error:
          'Could not replace text. Possible causes: paraId not found; search missing / ambiguous; ' +
          'or this adapter does not support direct text edits.',
      };
    }
    return {
      success: true,
      data: input.replaceWith
        ? `Replaced "${input.search}" on ${input.paraId}.`
        : `Deleted "${input.search}" on ${input.paraId}.`,
    };
  },
};

const setContentControlTool: AgentToolDefinition<{
  tag?: string;
  alias?: string;
  id?: number;
  type?: ContentControlType;
  text: string;
  force?: boolean;
}> = {
  name: 'set_content_control',
  displayName: 'Filling content control',
  description:
    'Replace the text content of a Word content control / SDT by stable template metadata ' +
    '(`tag`, `alias`, `id`, or `type`). Prefer this over paragraph search when the document is a template.',
  inputSchema: {
    type: 'object',
    properties: {
      tag: { type: 'string', description: 'Word content-control tag.' },
      alias: { type: 'string', description: 'Content-control alias/title.' },
      id: { type: 'number', description: 'Numeric content-control id.' },
      type: {
        type: 'string',
        enum: CONTENT_CONTROL_TYPES,
        description: 'Content-control type filter.',
      },
      text: {
        type: 'string',
        description: 'Replacement text. Newlines become paragraphs for rich-text controls.',
      },
      force: {
        type: 'boolean',
        description:
          'Override type/binding checks where the editor adapter allows it. Defaults to false.',
      },
    },
    required: ['text'],
  },
  handler: (input, bridge) => {
    const { filter, error } = parseContentControlFilter(input, true);
    if (error) return { success: false, error };
    if (typeof input.text !== 'string') {
      return { success: false, error: '`text` must be a string.' };
    }
    if (input.force !== undefined && typeof input.force !== 'boolean') {
      return { success: false, error: '`force` must be a boolean when provided.' };
    }

    const ok = bridge.setContentControl({
      ...filter,
      text: input.text,
      force: input.force,
    });
    if (!ok) {
      return {
        success: false,
        error:
          'Could not set content control. Possible causes: no matching tag/alias/id/type, ' +
          'locked or unsupported control, or this adapter does not support SDT editing.',
      };
    }
    return { success: true, data: 'Content control updated.' };
  },
};

const insertTableTool: AgentToolDefinition<{
  rows: number;
  columns: number;
  data?: string[][];
  hasHeader?: boolean;
  paraId?: string;
}> = {
  name: 'insert_table',
  displayName: 'Inserting table',
  description:
    'Insert a table into the live document. By default it inserts at the user cursor. ' +
    'Pass `paraId` to insert after a specific paragraph returned by read_document / find_text. ' +
    '`data` is an optional 2D string array; missing cells are left empty and extra cells are ignored.',
  inputSchema: {
    type: 'object',
    properties: {
      rows: { type: 'number', description: 'Number of rows. Must be an integer from 1 to 20.' },
      columns: {
        type: 'number',
        description: 'Number of columns. Must be an integer from 1 to 10.',
      },
      data: {
        type: 'array',
        description: 'Optional table cell contents as rows of strings.',
        items: { type: 'array', items: { type: 'string' } },
      },
      hasHeader: {
        type: 'boolean',
        description: 'Whether the first row should be treated as a table header.',
      },
      paraId: {
        type: 'string',
        description: 'Optional paragraph id. When supplied, inserts after that paragraph.',
      },
    },
    required: ['rows', 'columns'],
  },
  handler: (input, bridge) => {
    const rows = Number(input.rows);
    const columns = Number(input.columns);
    if (!Number.isInteger(rows) || rows < 1 || rows > 20) {
      return { success: false, error: '`rows` must be an integer from 1 to 20.' };
    }
    if (!Number.isInteger(columns) || columns < 1 || columns > 10) {
      return { success: false, error: '`columns` must be an integer from 1 to 10.' };
    }
    if (
      input.data !== undefined &&
      (!Array.isArray(input.data) ||
        !input.data.every(
          (row) => Array.isArray(row) && row.every((cell) => typeof cell === 'string')
        ))
    ) {
      return { success: false, error: '`data` must be a 2D array of strings.' };
    }
    if (input.hasHeader !== undefined && typeof input.hasHeader !== 'boolean') {
      return { success: false, error: '`hasHeader` must be a boolean when provided.' };
    }
    if (input.paraId !== undefined && typeof input.paraId !== 'string') {
      return { success: false, error: '`paraId` must be a string when provided.' };
    }

    const ok = bridge.insertTable({
      rows,
      columns,
      data: input.data,
      hasHeader: input.hasHeader,
      paraId: input.paraId,
    });
    if (!ok) {
      return {
        success: false,
        error:
          'Could not insert table. Possible causes: the editor is not ready, this adapter does ' +
          'not support structural insertions, or the supplied paraId was not found.',
      };
    }
    return { success: true, data: `Inserted ${rows}x${columns} table.` };
  },
};

const insertImageTool: AgentToolDefinition<{
  src: string;
  alt?: string;
  width?: number;
  height?: number;
  paraId?: string;
}> = {
  name: 'insert_image',
  displayName: 'Inserting image',
  description:
    'Insert an inline image into the live document. The `src` must be a base64 data URL ' +
    '(for example `data:image/png;base64,...`) so DOCX export can embed it. By default it ' +
    'inserts at the user cursor; pass `paraId` to insert at the end of that paragraph.',
  inputSchema: {
    type: 'object',
    properties: {
      src: {
        type: 'string',
        description: 'Base64 image data URL, e.g. data:image/png;base64,...',
      },
      alt: { type: 'string', description: 'Alt text / description for the image.' },
      width: {
        type: 'number',
        description: 'Rendered width in pixels. Defaults to 320 when omitted.',
      },
      height: {
        type: 'number',
        description: 'Rendered height in pixels. Defaults to 180 when omitted.',
      },
      paraId: {
        type: 'string',
        description: 'Optional paragraph id. When supplied, inserts at the end of that paragraph.',
      },
    },
    required: ['src'],
  },
  handler: (input, bridge) => {
    if (typeof input.src !== 'string' || !/^data:image\/[^;]+;base64,/.test(input.src)) {
      return {
        success: false,
        error: '`src` must be a base64 image data URL such as data:image/png;base64,...',
      };
    }
    if (input.alt !== undefined && typeof input.alt !== 'string') {
      return { success: false, error: '`alt` must be a string when provided.' };
    }
    if (input.paraId !== undefined && typeof input.paraId !== 'string') {
      return { success: false, error: '`paraId` must be a string when provided.' };
    }

    const width = input.width ?? 320;
    const height = input.height ?? 180;
    if (!Number.isFinite(width) || width <= 0 || !Number.isFinite(height) || height <= 0) {
      return { success: false, error: '`width` and `height` must be positive numbers.' };
    }

    const ok = bridge.insertImage({
      src: input.src,
      alt: input.alt,
      width,
      height,
      paraId: input.paraId,
    });
    if (!ok) {
      return {
        success: false,
        error:
          'Could not insert image. Possible causes: the editor is not ready, this adapter does ' +
          'not support image insertion, or the supplied paraId was not found.',
      };
    }
    return { success: true, data: 'Image inserted.' };
  },
};

const replyComment: AgentToolDefinition<{ commentId: number; text: string }> = {
  name: 'reply_comment',
  displayName: 'Replying to comment',
  description: 'Reply to an existing comment by id. Threaded under the original.',
  inputSchema: {
    type: 'object',
    properties: {
      commentId: { type: 'number', description: 'Comment id from read_comments.' },
      text: { type: 'string', description: 'Reply body.' },
    },
    required: ['commentId', 'text'],
  },
  handler: (input, bridge) => {
    const id = bridge.replyTo(input.commentId, { text: input.text });
    if (id === null) return { success: false, error: `Comment #${input.commentId} not found.` };
    return { success: true, data: `Reply ${id} added to comment ${input.commentId}.` };
  },
};

const resolveComment: AgentToolDefinition<{ commentId: number }> = {
  name: 'resolve_comment',
  displayName: 'Resolving comment',
  description: 'Mark a comment as resolved (done).',
  inputSchema: {
    type: 'object',
    properties: {
      commentId: { type: 'number', description: 'Comment id from read_comments.' },
    },
    required: ['commentId'],
  },
  handler: (input, bridge) => {
    bridge.resolveComment(input.commentId);
    return { success: true, data: `Comment ${input.commentId} resolved.` };
  },
};

// ── Navigate ────────────────────────────────────────────────────────────────

const scroll: AgentToolDefinition<{ paraId: string }> = {
  name: 'scroll',
  displayName: 'Scrolling',
  description: "Scroll the editor to a paragraph by paraId. Does not move the user's cursor.",
  inputSchema: {
    type: 'object',
    properties: {
      paraId: { type: 'string', description: 'Paragraph id from read_document / find_text.' },
    },
    required: ['paraId'],
  },
  handler: (input, bridge) => {
    const ok = bridge.scrollTo(input.paraId);
    if (!ok) return { success: false, error: `paraId ${input.paraId} not found.` };
    return { success: true, data: `Scrolled to ${input.paraId}.` };
  },
};

// ── Registry ────────────────────────────────────────────────────────────────

/**
 * All built-in agent tools — read/write document content, comments, and
 * tracked changes. Use `getToolSchemas()` to feed them to an LLM and
 * `executeToolCall()` to run the handlers against an `EditorBridge`.
 *
 * @public
 */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
export const agentTools: AgentToolDefinition<any>[] = [
  readDocument,
  readSelection,
  readPage,
  readPages,
  findText,
  readComments,
  readChanges,
  readContentControls,
  addComment,
  suggestChange,
  insertTextTool,
  replaceTextTool,
  setContentControlTool,
  insertTableTool,
  insertImageTool,
  applyFormatting,
  setParagraphStyle,
  replyComment,
  resolveComment,
  scroll,
];

/**
 * Execute a tool call against an EditorBridge.
 * Returns the result (never throws).
 *
 * @public
 */
export function executeToolCall(
  toolName: string,
  input: Record<string, unknown>,
  bridge: EditorBridge
): AgentToolResult {
  const tool = agentTools.find((t) => t.name === toolName);
  if (!tool) return { success: false, error: `Unknown tool: ${toolName}` };
  try {
    return tool.handler(input, bridge);
  } catch (e) {
    return { success: false, error: e instanceof Error ? e.message : String(e) };
  }
}

/**
 * Friendly UI label for a tool — sourced from the registry's `displayName`,
 * falling back to a sentence-case version of the snake_case name. Used by
 * `<AgentTimeline>` and any other UI that lists running / completed tools.
 *
 * @example getToolDisplayName('add_comment') // → 'Adding comment'
 * @example getToolDisplayName('fetch_clause_template') // → 'Fetch clause template'
 *
 * @public
 */
export function getToolDisplayName(name: string): string {
  const def = agentTools.find((t) => t.name === name);
  if (def?.displayName) return def.displayName;
  const spaced = name.replace(/_/g, ' ');
  return spaced.charAt(0).toUpperCase() + spaced.slice(1);
}

/**
 * Get tool schemas in OpenAI function-calling format. Works directly
 * with the OpenAI SDK and Anthropic's tools API. For Vercel AI SDK,
 * LangChain, or other agent runtimes, transform this output to that
 * runtime's required shape — see `examples/agent-chat-demo/` for a
 * Vercel AI SDK example. The package stays runtime-agnostic.
 *
 * @public
 */
export function getToolSchemas() {
  return agentTools.map((t) => ({
    type: 'function' as const,
    function: {
      name: t.name,
      description: t.description,
      parameters: t.inputSchema,
    },
  }));
}
