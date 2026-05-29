/**
 * Image commands — thin re-exports from the extension system.
 *
 * Wrap-type transitions for floating images. Inline↔anchor conversions are
 * structural and live in a follow-up; this surface only covers anchor↔anchor.
 */

import type { Command, EditorState, Transaction } from 'prosemirror-state';
import type { Node as PMNode } from 'prosemirror-model';
import { singletonManager } from '../schema';
import { makeRevisionInfo } from '../plugins/revisionIds';
import type {
  ImageLayoutTarget,
  SetImageWrapTypeOptions,
} from '../extensions/nodes/ImageExtension';

const cmds = singletonManager.getCommands();

/**
 * Insert an image node at `pos`, wrapping with the `insertion` mark when
 * suggesting mode is active. Centralizes the tracked-image-insert flow
 * so React `useFileIO`, Vue `useImageActions`, and clipboard-paste
 * paths all share one source of truth — adding a fresh image in
 * suggesting mode always round-trips as `<w:ins>{run with drawing}</w:ins>`.
 *
 * Caller responsibility: produce the `image` node via `schema.nodes.image.create`.
 * This helper handles the dispatch + optional mark application.
 *
 * @public
 */
export function insertImageNode(
  state: EditorState,
  dispatch: ((tr: Transaction) => void) | undefined,
  imageNode: PMNode,
  pos: number
): boolean {
  if (!dispatch) return true;
  const tr = state.tr.insert(pos, imageNode);
  const info = makeRevisionInfo(state);
  const insertionType = state.schema.marks.insertion;
  if (info && insertionType) {
    tr.addMark(
      pos,
      pos + imageNode.nodeSize,
      insertionType.create({
        revisionId: info.revisionId,
        author: info.author,
        date: info.date,
      })
    );
  }
  dispatch(tr.scrollIntoView());
  return true;
}

/**
 * Change a floating image's wrap layout. `pos` is the PM document position of
 * the image node; `target` is either an OOXML wrap type (square / tight /
 * topAndBottom / behind / inFront / inline) or a directional convenience
 * choice (`squareLeft` / `squareRight`).
 *
 * `opts.initialPositionEmu` is used when promoting an inline image to an
 * anchor — the caller measures the image's current rendered offset relative
 * to the column origin in EMUs and passes it through, so the new float lands
 * exactly where the inline glyph used to sit (matches Word's behavior).
 */
export function setImageWrapType(
  pos: number,
  target: ImageLayoutTarget,
  opts?: SetImageWrapTypeOptions
): Command {
  return cmds.setImageWrapType(pos, target, opts);
}

export type {
  AnchorWrapType,
  ImageLayoutTarget,
  SetImageWrapTypeOptions,
} from '../extensions/nodes/ImageExtension';
