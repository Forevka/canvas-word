// Addressing: offsets are UTF-16 code unit indices into a paragraph's concatenated
// run text — the same unit pretext / Intl.Segmenter use, so no conversion layer.

export interface DocPosition {
  blockId: string;
  offset: number;
}

export interface DocSelection {
  anchor: DocPosition;
  focus: DocPosition;
  /** Preserved X column for Up/Down caret movement across lines of differing length. */
  goalX?: number;
}

export const isCollapsed = (s: DocSelection): boolean =>
  s.anchor.blockId === s.focus.blockId && s.anchor.offset === s.focus.offset;
