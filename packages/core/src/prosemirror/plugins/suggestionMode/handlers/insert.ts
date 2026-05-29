/**
 * Suggesting-mode insertion path. Handles both pure inserts and the
 * replace case (selection non-empty → mark selection deleted, then insert).
 */

import type { EditorView } from 'prosemirror-view';

import { findAdjacentRevision } from '../adjacency';
import { makeMarkAttrs } from '../markAttrs';
import { SUGGESTION_META, type SuggestionModeState } from '../state';
import { markRangeAsDeleted } from './delete';

/**
 * Insert text as a tracked insertion, optionally marking replaced selection as deletion.
 */
export function applySuggestionInsert(
  view: EditorView,
  from: number,
  to: number,
  text: string,
  pluginState: SuggestionModeState
): boolean {
  const insertionType = view.state.schema.marks.insertion;
  if (!insertionType) return false;

  const tr = view.state.tr;
  tr.setMeta(SUGGESTION_META, true);

  const insertAttrs =
    findAdjacentRevision(view.state.doc, from, 'insertion', pluginState.author) ||
    makeMarkAttrs(pluginState);

  if (from !== to) {
    const deletionType = view.state.schema.marks.deletion;
    if (deletionType) {
      // Replace op: pass the insertion's date down so the deletion shares
      // the (author, date) triple — that's what extractTrackedChanges
      // uses to detect adjacent del+ins and fold them into one
      // 'replacement' card. The `w:id` stays distinct so we don't trip
      // the OOXML move-pair serializer (fromProseDoc/paragraph.ts:340).
      markRangeAsDeleted(
        tr,
        view.state.doc,
        from,
        to,
        insertionType,
        deletionType,
        pluginState,
        insertAttrs.date
      );
    }
  }

  const insertAt = tr.mapping.map(to);
  tr.insertText(text, insertAt, insertAt);

  // Strip inherited deletion marks — new text must never be marked as deleted.
  const deletionType = view.state.schema.marks.deletion;
  if (deletionType) {
    tr.removeMark(insertAt, insertAt + text.length, deletionType);
  }

  // Apply the correct insertion mark. If the cursor was inside an existing
  // insertion by the same author, insertText already inherited that mark and
  // insertAttrs will match — addMark is effectively a no-op that preserves
  // the continuous mark span. We intentionally do NOT removeMark(insertionType)
  // first, because that fragments the mark span and creates a nested change.
  tr.addMark(insertAt, insertAt + text.length, insertionType.create(insertAttrs));

  view.dispatch(tr.scrollIntoView());
  return true;
}
