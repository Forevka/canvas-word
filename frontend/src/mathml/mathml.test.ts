// Conversion correctness for the MathML <-> AST <-> OMML bridges. These are the
// de-risking core (M0): everything downstream (typesetter, editor, docx) trusts
// that an equation survives every hop. We assert three things:
//   1. MathML string -> AST -> MathML string is STABLE (idempotent canonical form).
//   2. AST -> OMML golden output (the exact markup Word expects).
//   3. OMML -> AST -> OMML round-trips (import then re-export).

import { describe, expect, it } from "vitest";
import { parse, isElementNode, type TNode } from "txml/txml";
import type { MathEquation } from "@cw/shared";
import { serializeMathml } from "./serialize";
import { parseMathml } from "./parse";
import { mathmlToOmml } from "./toOmml";
import { ommlToMathml } from "./fromOmml";

/** Re-parse a serialized equation and re-serialize: the canonical form must be a
 *  fixed point (parse∘serialize∘parse == parse). */
const canonical = (mathml: string): string => serializeMathml(parseMathml(mathml));

// Parse the way the real docx importer does (import/docx/xml.ts): keepWhitespace
// matters — a math run whose text is a single/multiple space must survive.
const firstEl = (xml: string): TNode =>
  parse(xml, { keepWhitespace: true, decodeEntities: true }).filter(isElementNode)[0] as TNode;

const ommlOf = (mathml: string): string => mathmlToOmml(parseMathml(mathml));

describe("MathML parse <-> serialize is a stable canonical form", () => {
  const cases: Record<string, string> = {
    fraction: "<math><mfrac><mn>1</mn><mn>2</mn></mfrac></math>",
    superscript: "<math><msup><mi>x</mi><mn>2</mn></msup></math>",
    subscript: "<math><msub><mi>a</mi><mi>n</mi></msub></math>",
    subsup: "<math><msubsup><mi>x</mi><mn>1</mn><mn>2</mn></msubsup></math>",
    sqrt: "<math><msqrt><mrow><mi>x</mi><mo>+</mo><mn>1</mn></mrow></msqrt></math>",
    nthRoot: "<math><mroot><mi>x</mi><mn>3</mn></mroot></math>",
    fenced: '<math><mfenced close=")" open="("><mi>x</mi></mfenced></math>',
    overAccent: '<math><mover accent="true"><mi>x</mi><mo>^</mo></mover></math>',
    matrix:
      "<math><mtable><mtr><mtd><mn>1</mn></mtd><mtd><mn>0</mn></mtd></mtr>" +
      "<mtr><mtd><mn>0</mn></mtd><mtd><mn>1</mn></mtd></mtr></mtable></math>",
    blackboard: '<math><mi mathvariant="double-struck">R</mi></math>',
  };

  for (const [name, mathml] of Object.entries(cases)) {
    it(`${name} is idempotent`, () => {
      const once = canonical(mathml);
      const twice = canonical(once);
      expect(twice).toBe(once);
    });
  }

  it("preserves quadratic-formula structure end to end", () => {
    const quad =
      "<math><mi>x</mi><mo>=</mo><mfrac>" +
      "<mrow><mo>-</mo><mi>b</mi><mo>±</mo><msqrt><mrow><msup><mi>b</mi><mn>2</mn></msup>" +
      "<mo>-</mo><mn>4</mn><mi>a</mi><mi>c</mi></mrow></msqrt></mrow>" +
      "<mrow><mn>2</mn><mi>a</mi></mrow></mfrac></math>";
    const out = canonical(quad);
    // Round-trips to a fixed point...
    expect(canonical(out)).toBe(out);
    // ...and still contains the key structure.
    expect(out).toContain("<mfrac>");
    expect(out).toContain("<msqrt>");
    expect(out).toContain("<msup>");
  });
});

describe("AST -> OMML golden output", () => {
  it("fraction emits m:f / m:num / m:den", () => {
    expect(ommlOf("<math><mfrac><mn>1</mn><mn>2</mn></mfrac></math>")).toBe(
      "<m:oMath><m:f><m:num><m:r><m:rPr><m:sty m:val=\"p\"/></m:rPr><m:t>1</m:t></m:r></m:num>" +
        "<m:den><m:r><m:rPr><m:sty m:val=\"p\"/></m:rPr><m:t>2</m:t></m:r></m:den></m:f></m:oMath>",
    );
  });

  it("identifier defaults to italic (no explicit m:sty)", () => {
    expect(ommlOf("<math><mi>x</mi></math>")).toBe("<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>");
  });

  it("superscript emits m:sSup", () => {
    const omml = ommlOf("<math><msup><mi>x</mi><mn>2</mn></msup></math>");
    expect(omml).toContain("<m:sSup>");
    expect(omml).toContain("<m:sup>");
    expect(omml).toContain("<m:e>");
  });

  it("sqrt hides the degree", () => {
    const omml = ommlOf("<math><msqrt><mi>x</mi></msqrt></math>");
    expect(omml).toContain("<m:rad>");
    expect(omml).toContain('<m:degHide m:val="1"/>');
  });

  it("display math is wrapped in m:oMathPara", () => {
    const omml = mathmlToOmml(parseMathml('<math display="block"><mi>x</mi></math>'));
    expect(omml.startsWith("<m:oMathPara><m:oMath>")).toBe(true);
  });

  it("n-ary sum emits m:nary with the operator char", () => {
    // Build via AST round-trip from OMML below; here just assert the emit shape.
    const eq: MathEquation = {
      display: false,
      root: {
        type: "row",
        children: [
          {
            type: "nary",
            op: "∑",
            sub: { type: "row", children: [{ type: "ident", text: "i" }] },
            sup: { type: "ident", text: "n" },
            body: { type: "ident", text: "i" },
          },
        ],
      },
    };
    const omml = mathmlToOmml(eq);
    expect(omml).toContain("<m:nary>");
    expect(omml).toContain('<m:chr m:val="∑"/>');
  });
});

describe("OMML -> AST -> OMML round-trip (import then re-export)", () => {
  const roundtrip = (omml: string): string => {
    const root = ommlToMathml(firstEl(omml));
    return mathmlToOmml({ root, display: false });
  };

  it("fraction survives", () => {
    const omml =
      "<m:oMath><m:f><m:num><m:r><m:t>1</m:t></m:r></m:num>" +
      "<m:den><m:r><m:t>2</m:t></m:r></m:den></m:f></m:oMath>";
    const back = roundtrip(omml);
    expect(back).toContain("<m:f>");
    expect(back).toContain("<m:num>");
    expect(back).toContain("<m:den>");
  });

  it("classifies digits as numbers and letters as identifiers", () => {
    const root = ommlToMathml(
      firstEl("<m:oMath><m:r><m:t>x</m:t></m:r><m:r><m:t>2</m:t></m:r></m:oMath>"),
    );
    expect(root.children[0]).toMatchObject({ type: "ident", text: "x" });
    expect(root.children[1]).toMatchObject({ type: "number", text: "2" });
  });

  it("classifies operator characters as operators", () => {
    const root = ommlToMathml(firstEl("<m:oMath><m:r><m:t>+</m:t></m:r></m:oMath>"));
    expect(root.children[0]).toMatchObject({ type: "op", text: "+" });
  });

  it("n-ary integral survives with its operator and bounds", () => {
    const omml =
      "<m:oMath><m:nary><m:naryPr><m:chr m:val=\"∫\"/></m:naryPr>" +
      "<m:sub><m:r><m:t>0</m:t></m:r></m:sub><m:sup><m:r><m:t>1</m:t></m:r></m:sup>" +
      "<m:e><m:r><m:t>f</m:t></m:r></m:e></m:nary></m:oMath>";
    const root = ommlToMathml(firstEl(omml));
    expect(root.children[0]).toMatchObject({ type: "nary", op: "∫" });
    const back = mathmlToOmml({ root, display: false });
    expect(back).toContain('<m:chr m:val="∫"/>');
    expect(back).toContain("<m:sub>");
    expect(back).toContain("<m:sup>");
  });

  it("delimiter (fences) survives with begChr/endChr", () => {
    const omml =
      "<m:oMath><m:d><m:dPr><m:begChr m:val=\"[\"/><m:endChr m:val=\"]\"/></m:dPr>" +
      "<m:e><m:r><m:t>x</m:t></m:r></m:e></m:d></m:oMath>";
    const root = ommlToMathml(firstEl(omml));
    expect(root.children[0]).toMatchObject({ type: "fenced", open: "[", close: "]" });
  });

  // ── #11 fidelity gaps ──────────────────────────────────────────────────────

  it("munderover round-trips (nested limUpp/limLow collapses back to one limit)", () => {
    const eq = parseMathml("<math><munderover><mi>x</mi><mn>1</mn><mn>2</mn></munderover></math>");
    const root = ommlToMathml(firstEl(mathmlToOmml(eq)));
    expect(root.children[0]).toMatchObject({
      type: "limit",
      under: { children: [{ type: "number", text: "1" }] },
      over: { children: [{ type: "number", text: "2" }] },
    });
    // Merged, not left nested: the base is the identifier, not another limit.
    const lim = root.children[0] as { base: { type: string } };
    expect(lim.base.type).not.toBe("limit");
    expect(JSON.stringify(lim.base)).toContain('"x"');
  });

  it("double-struck identifier round-trips via m:scr", () => {
    const omml = mathmlToOmml(parseMathml('<math><mi mathvariant="double-struck">R</mi></math>'));
    expect(omml).toContain('<m:scr m:val="double-struck"/>');
    const root = ommlToMathml(firstEl(omml));
    expect(root.children[0]).toMatchObject({ type: "ident", text: "R", variant: "double-struck" });
  });

  it("fraktur / script / sans-serif / monospace variants survive m:scr", () => {
    const cases: [string, string][] = [
      ["fraktur", "g"],
      ["script", "L"],
      ["sans-serif", "A"],
      ["monospace", "x"],
    ];
    for (const [variant, ch] of cases) {
      const omml = mathmlToOmml(parseMathml(`<math><mi mathvariant="${variant}">${ch}</mi></math>`));
      expect(omml).toContain(`<m:scr m:val="${variant}"/>`);
      expect(ommlToMathml(firstEl(omml)).children[0]).toMatchObject({ type: "ident", text: ch, variant });
    }
  });

  it("mspace width survives approximately through OMML", () => {
    const eq = parseMathml('<math><mi>a</mi><mspace width="0.5em"/><mi>b</mi></math>');
    const root = ommlToMathml(firstEl(mathmlToOmml(eq)));
    const space = root.children.find((c) => c.type === "space") as { widthEm?: number } | undefined;
    expect(space).toBeTruthy();
    expect(space!.widthEm).toBeGreaterThan(0);
  });
});
