// Heading-level resolution for the accessibility mirror. This is what decides
// whether a paragraph reaches a screen reader as a heading and at which level, so
// a wrong answer here is a silently broken document outline.
//
// Only the pure resolution is covered: the rest of the mirror builds and patches
// real DOM, and this workspace runs vitest in the default node environment (no
// jsdom), so the node-building and live-region paths are browser-verified instead.

import { describe, expect, it } from "vitest";
import type { Document, NamedStyle } from "@cw/shared";
import { headingLevelOf } from "./mirror";

const docWith = (styles: NamedStyle[]): Document =>
  ({ blocks: [], stylesheet: { styles } }) as unknown as Document;

const style = (id: string, name?: string, basedOn?: string): NamedStyle =>
  ({ id, ...(name ? { name } : {}), ...(basedOn ? { basedOn } : {}) }) as NamedStyle;

describe("headingLevelOf", () => {
  it("reads the level straight off a built-in Heading id", () => {
    expect(headingLevelOf(docWith([style("Heading2", "Heading 2")]), "Heading2")).toBe(2);
  });

  it("treats Title as level 0", () => {
    expect(headingLevelOf(docWith([style("Title", "Title")]), "Title")).toBe(0);
  });

  it("inherits the level through the basedOn chain", () => {
    // A custom style whose own id/name says nothing about headings still has to
    // read as level 2 because it is based on Heading 2.
    const doc = docWith([style("SectionTitle", "Section title", "Heading2"), style("Heading2", "Heading 2")]);
    expect(headingLevelOf(doc, "SectionTitle")).toBe(2);
  });

  it("follows a multi-step chain to the heading at its root", () => {
    const doc = docWith([
      style("Sub", "Sub", "SectionTitle"),
      style("SectionTitle", "Section title", "Heading3"),
      style("Heading3", "Heading 3"),
    ]);
    expect(headingLevelOf(doc, "Sub")).toBe(3);
  });

  it("terminates on a basedOn cycle instead of hanging", () => {
    const doc = docWith([style("A", "A", "B"), style("B", "B", "A")]);
    expect(headingLevelOf(doc, "A")).toBeNull();
  });

  it("is null for a body style, an unknown id, and no style at all", () => {
    const doc = docWith([style("Normal", "Normal")]);
    expect(headingLevelOf(doc, "Normal")).toBeNull();
    expect(headingLevelOf(doc, "NoSuchStyle")).toBeNull();
    expect(headingLevelOf(doc, undefined)).toBeNull();
  });

  it("does not mistake a style merely containing 'heading' for a heading", () => {
    // Every rule anchors "heading" to a word boundary, so "Subheading" is a body
    // style — it must not be announced as a heading just for containing the word.
    expect(headingLevelOf(docWith([style("Subheading", "Subheading")]), "Subheading")).toBeNull();
  });
});
