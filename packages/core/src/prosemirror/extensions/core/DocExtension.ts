/**
 * Doc Extension — top-level document node.
 *
 * Doc-level attrs ride along with the PM state through undo/redo and
 * transactions, so adapters don't need a separate prop to thread.
 */

import { createNodeExtension } from '../create';

export const DocExtension = createNodeExtension({
  name: 'doc',
  schemaNodeName: 'doc',
  nodeSpec: {
    content: '(paragraph | horizontalRule | pageBreak | table | textBox | blockSdt)+',
    attrs: {
      /** `w:defaultTabStop` (§17.6.13) in twips; null = OOXML default 720. */
      defaultTabStopTwips: { default: null },
    },
  },
});
