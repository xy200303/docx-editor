/**
 * Ref-API assembler — takes the foundational primitives from
 * `useDocxEditor` plus the action objects from every other composable
 * in `DocxEditor.vue` and returns the `DocxEditorRef`-shaped
 * `exposed` object the parent feeds into `defineExpose`. Owns the
 * small local helpers (print, scrollToPage, scrollToPosition,
 * getEditorRef, getTotalPages, getCurrentPage, findInDocument,
 * getSelectionInfo, getComments, getPageContent, scrollToParaId,
 * onContentChange, onSelectionChange) that previously sat inline in
 * the SFC just to populate the `exposed` literal.
 *
 * `satisfies DocxEditorRef` enforces signatures against EditorRefLike
 * at typecheck time (Decision 10 in the 1.0 spec) without affecting
 * the runtime shape.
 */

import type { Ref, ShallowRef } from 'vue';
import { TextSelection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import type { Schema, Node as PMNode } from 'prosemirror-model';
import type { Document } from '@eigenpal/docx-editor-core/types/document';
import type { Comment } from '@eigenpal/docx-editor-core/types/content';
import type { DocxInput } from '@eigenpal/docx-editor-core/utils';
import type { Layout } from '@eigenpal/docx-editor-core/layout-engine';
import { findPageIndexContainingPmPos } from '@eigenpal/docx-editor-core/layout-engine';
import { insertImageNode, insertTable } from '@eigenpal/docx-editor-core/prosemirror/commands';
import {
  findContentControlsInPM,
  findContentControlPos,
  setContentControlContentTr,
  removeContentControlTr,
  type PMContentControl,
} from '@eigenpal/docx-editor-core/prosemirror';
import {
  ContentControlNotFoundError,
  type ContentControlFilter,
} from '@eigenpal/docx-editor-core/agent';
import {
  findInDocument as findInDocumentImpl,
  getSelectionInfo as getSelectionInfoImpl,
  getPageContent as getPageContentImpl,
} from '../utils/refApiQueries';
import { findParaIdRange, findTextInPmParagraph } from '../utils/paraTextHelpers';
import type { DocxEditorRef } from '../components/DocxEditor/types';
import type { ApplyFormattingOptions } from './useFormattingActions';

export interface UseDocxEditorRefApiOptions {
  // Foundational refs / accessors (useDocxEditor)
  editorView: Ref<EditorView | null>;
  layout: Ref<Layout | null>;
  pagesRef: Ref<HTMLElement | null>;
  pagesViewportRef: Ref<HTMLElement | null>;
  zoom: Ref<number>;
  comments: ShallowRef<Comment[]>;
  // Action handles
  focus: () => void;
  destroy: () => void;
  getDocument: () => Document | null;
  setZoom: (zoom: number) => void;
  save: () => Promise<ArrayBuffer | null>;
  loadDocument: (doc: Document) => void;
  loadDocumentBuffer: (buffer: DocxInput) => Promise<void>;
  addComment: (options: {
    paraId: string;
    text: string;
    author: string;
    search?: string;
  }) => number | null;
  replyToComment: (commentId: number, text: string, author: string) => number | null;
  resolveComment: (commentId: number) => void;
  proposeChange: (options: {
    paraId: string;
    search: string;
    replaceWith: string;
    author: string;
  }) => boolean;
  applyFormatting: (options: ApplyFormattingOptions) => boolean;
  setParagraphStyle: (options: { paraId: string; styleId: string }) => boolean;
  scrollVisiblePositionIntoView: (pmPos: number) => void;
  // Subscriber sets (used by onContentChange / onSelectionChange)
  contentChangeSubscribers: Set<(document: unknown) => void>;
  selectionChangeSubscribers: Set<(selection: unknown) => void>;
  // Optional host hook for print
  onPrint?: () => void;
}

export function useDocxEditorRefApi(opts: UseDocxEditorRefApiOptions): {
  exposed: DocxEditorRef;
} {
  function print() {
    opts.onPrint?.();
    window.print();
  }

  function openPrintPreview() {
    print();
  }

  function getZoom() {
    return opts.zoom.value;
  }

  function scrollToPage(pageNumber: number) {
    if (!Number.isInteger(pageNumber) || pageNumber < 1) return;
    const viewport = opts.pagesViewportRef.value;
    const pageEl = opts.pagesRef.value?.querySelector<HTMLElement>(
      `[data-page-number="${pageNumber}"]`
    );
    if (!viewport || !pageEl) return;
    const viewportRect = viewport.getBoundingClientRect();
    const pageRect = pageEl.getBoundingClientRect();
    viewport.scrollTo({
      top: pageRect.top - viewportRect.top + viewport.scrollTop - 24,
      behavior: 'smooth',
    });
  }

  function scrollToPosition(pmPos: number) {
    if (!Number.isFinite(pmPos)) return;
    opts.scrollVisiblePositionIntoView(pmPos);
  }

  function getEditorRef() {
    if (!opts.editorView.value) return null;
    return {
      getDocument: opts.getDocument,
      getView: () => opts.editorView.value,
      getState: () => opts.editorView.value?.state ?? null,
    };
  }

  function getTotalPages(): number {
    return opts.layout.value?.pages.length ?? 0;
  }

  function getCurrentPage(): number {
    const currentLayout = opts.layout.value;
    const view = opts.editorView.value;
    if (!currentLayout || !view) return 0;
    const pageIndex = findPageIndexContainingPmPos(currentLayout, view.state.selection.from);
    return pageIndex == null ? 0 : pageIndex + 1;
  }

  function scrollToParaId(paraId: string): boolean {
    const view = opts.editorView.value;
    if (!view) return false;
    const range = findParaIdRange(view.state.doc, paraId);
    if (!range) return false;
    opts.scrollVisiblePositionIntoView(range.from + 1);
    return true;
  }

  function findInDocument(query: string, findOpts?: { caseSensitive?: boolean; limit?: number }) {
    return findInDocumentImpl(opts.editorView.value, query, findOpts);
  }

  function getSelectionInfo() {
    return getSelectionInfoImpl(opts.editorView.value);
  }

  function getComments() {
    return opts.comments.value;
  }

  function getContentControls(filter?: ContentControlFilter): PMContentControl[] {
    const view = opts.editorView.value;
    return view ? findContentControlsInPM(view.state.doc, filter ?? {}) : [];
  }

  function scrollToContentControl(filter: ContentControlFilter): boolean {
    const view = opts.editorView.value;
    if (!view) return false;
    const pos = findContentControlPos(view.state.doc, filter);
    if (pos == null) return false;
    scrollToPosition(pos);
    return true;
  }

  function setContentControlContent(
    filter: ContentControlFilter,
    text: string,
    options?: { force?: boolean }
  ): boolean {
    const view = opts.editorView.value;
    if (!view) return false;
    try {
      view.dispatch(setContentControlContentTr(view.state, filter, text, options));
      return true;
    } catch (err) {
      // Not-found is a soft miss; a lock refusal surfaces to the caller.
      if (err instanceof ContentControlNotFoundError) return false;
      throw err;
    }
  }

  function removeContentControl(
    filter: ContentControlFilter,
    options?: { force?: boolean; keepContent?: boolean }
  ): boolean {
    const view = opts.editorView.value;
    if (!view) return false;
    try {
      view.dispatch(removeContentControlTr(view.state, filter, options));
      return true;
    } catch (err) {
      if (err instanceof ContentControlNotFoundError) return false;
      throw err;
    }
  }

  function getPageContent(pageNumber: number) {
    return getPageContentImpl(opts.editorView.value, opts.layout.value, pageNumber);
  }

  function insertTableFromRef(options: {
    rows: number;
    columns: number;
    data?: string[][];
    hasHeader?: boolean;
    paraId?: string;
  }): boolean {
    const view = opts.editorView.value;
    if (!view) return false;
    if (
      !Number.isInteger(options.rows) ||
      !Number.isInteger(options.columns) ||
      options.rows < 1 ||
      options.columns < 1
    ) {
      return false;
    }

    let state = view.state;
    if (options.paraId) {
      const range = findParaIdRange(state.doc, options.paraId);
      if (!range) return false;
      const pos = Math.max(range.from + 1, range.to - 1);
      state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, pos)));
    }

    const ok = insertTable(options.rows, options.columns, {
      data: options.data,
      hasHeader: options.hasHeader,
    })(state, view.dispatch);
    if (ok) view.focus();
    return ok;
  }

  function insertImageFromRef(options: {
    src: string;
    alt?: string;
    width?: number;
    height?: number;
    paraId?: string;
  }): boolean {
    const view = opts.editorView.value;
    if (!view) return false;
    const imageNodeType = view.state.schema.nodes.image;
    if (!imageNodeType || !options.src) return false;

    let pos = view.state.selection.from;
    if (options.paraId) {
      const range = findParaIdRange(view.state.doc, options.paraId);
      if (!range) return false;
      pos = range.to - 1;
    }

    const node = imageNodeType.create({
      src: options.src,
      alt: options.alt,
      width: options.width ?? 320,
      height: options.height ?? 180,
      rId: `rId_img_${Date.now()}`,
      wrapType: 'inline',
      displayMode: 'inline',
    });
    const ok = insertImageNode(view.state, view.dispatch, node, pos);
    if (ok) view.focus();
    return ok;
  }

  function createParagraphNodes(schema: Schema, text: string): PMNode[] {
    return text
      .split(/\r?\n/)
      .map((line) => schema.nodes.paragraph.create(null, line ? schema.text(line) : null));
  }

  function insertTextFromRef(options: {
    text: string;
    paraId?: string;
    position?:
      | 'cursor'
      | 'paragraph_start'
      | 'paragraph_end'
      | 'before_paragraph'
      | 'after_paragraph';
    search?: string;
    placement?: 'before' | 'after' | 'replace';
  }): boolean {
    const view = opts.editorView.value;
    if (!view || typeof options.text !== 'string' || options.text.length === 0) return false;

    const { state } = view;
    let tr = state.tr;

    if (options.search) {
      if (!options.paraId) return false;
      const range = findParaIdRange(state.doc, options.paraId);
      if (!range) return false;
      const textRange = findTextInPmParagraph(state.doc, range.from, range.to, options.search);
      if (!textRange) return false;

      const placement = options.placement ?? 'after';
      if (placement === 'replace') {
        tr = tr.insertText(options.text, textRange.from, textRange.to);
      } else if (placement === 'before') {
        tr = tr.insertText(options.text, textRange.from, textRange.from);
      } else if (placement === 'after') {
        tr = tr.insertText(options.text, textRange.to, textRange.to);
      } else {
        return false;
      }
      view.dispatch(tr.scrollIntoView());
      view.focus();
      return true;
    }

    const position = options.position ?? (options.paraId ? 'paragraph_end' : 'cursor');
    if (position === 'cursor') {
      tr = tr.insertText(options.text, state.selection.from, state.selection.to);
    } else {
      if (!options.paraId) return false;
      const range = findParaIdRange(state.doc, options.paraId);
      if (!range) return false;

      if (position === 'paragraph_start') {
        tr = tr.insertText(options.text, range.from + 1, range.from + 1);
      } else if (position === 'paragraph_end') {
        tr = tr.insertText(options.text, range.to - 1, range.to - 1);
      } else if (position === 'before_paragraph') {
        tr = tr.insert(range.from, createParagraphNodes(state.schema, options.text));
      } else if (position === 'after_paragraph') {
        tr = tr.insert(range.to, createParagraphNodes(state.schema, options.text));
      } else {
        return false;
      }
    }

    view.dispatch(tr.scrollIntoView());
    view.focus();
    return true;
  }

  function replaceTextFromRef(options: {
    paraId: string;
    search: string;
    replaceWith: string;
  }): boolean {
    const view = opts.editorView.value;
    if (!view || !options.paraId || !options.search || typeof options.replaceWith !== 'string') {
      return false;
    }

    const range = findParaIdRange(view.state.doc, options.paraId);
    if (!range) return false;
    const textRange = findTextInPmParagraph(view.state.doc, range.from, range.to, options.search);
    if (!textRange) return false;

    const tr = options.replaceWith
      ? view.state.tr.insertText(options.replaceWith, textRange.from, textRange.to)
      : view.state.tr.delete(textRange.from, textRange.to);
    view.dispatch(tr.scrollIntoView());
    view.focus();
    return true;
  }

  function onContentChange(listener: (document: unknown) => void): () => void {
    opts.contentChangeSubscribers.add(listener);
    return () => opts.contentChangeSubscribers.delete(listener);
  }

  function onSelectionChange(listener: (selection: unknown) => void): () => void {
    opts.selectionChangeSubscribers.add(listener);
    return () => opts.selectionChangeSubscribers.delete(listener);
  }

  const exposed = {
    getAgent: () => null,
    save: opts.save,
    setZoom: opts.setZoom,
    getZoom,
    focus: opts.focus,
    scrollToPage,
    scrollToPosition,
    openPrintPreview,
    print,
    loadDocument: opts.loadDocument,
    loadDocumentBuffer: opts.loadDocumentBuffer,
    destroy: opts.destroy,
    getDocument: opts.getDocument,
    getEditorRef,
    addComment: opts.addComment,
    replyToComment: opts.replyToComment,
    resolveComment: opts.resolveComment,
    proposeChange: opts.proposeChange,
    scrollToParaId,
    findInDocument,
    getSelectionInfo,
    getComments,
    getContentControls,
    scrollToContentControl,
    setContentControlContent,
    removeContentControl,
    applyFormatting: opts.applyFormatting,
    setParagraphStyle: opts.setParagraphStyle,
    insertTable: insertTableFromRef,
    insertImage: insertImageFromRef,
    insertText: insertTextFromRef,
    replaceText: replaceTextFromRef,
    getPageContent,
    getTotalPages,
    getCurrentPage,
    onContentChange,
    onSelectionChange,
  } satisfies DocxEditorRef;

  return { exposed };
}
