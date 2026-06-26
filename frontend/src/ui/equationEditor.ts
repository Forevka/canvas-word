// The equation editor — a floating, non-blocking panel for inserting/editing a
// display equation. You edit Presentation MathML (the canonical form) with a
// template/symbol palette, and a live preview renders the result with the SAME
// typesetter the page uses (equationBox + paintMathBoxCanvas), so what you see is
// what lands. Pattern mirrors ui/fieldConstructor.ts + makeFloatingDialog.

import type { MathEquation } from "@cw/shared";
import { MATH_FONT_FAMILY } from "../fonts/clones";
import { injectCssOnce } from "./styles";
import { makeFloatingDialog } from "./floatingDialog";
import { parseMathml } from "../mathml/parse";
import { serializeMathml } from "../mathml/serialize";
import { latexToMath } from "../mathml/latex";
import { mathToLatex } from "../mathml/toLatex";
import { equationBox } from "../layout/math/equationLayout";
import { paintMathBoxCanvas } from "../paint/paintMath";

export interface EquationEditorOptions {
  /** Seed MathML when editing an existing equation; empty for a fresh insert. */
  initialMathml?: string;
  /** Called with the parsed equation when the user clicks Insert/Apply. The
   *  equation's `display` reflects the Display/Inline toggle. */
  onApply: (equation: MathEquation) => void;
  /** True when editing an existing equation (button label becomes "Apply"). */
  editing?: boolean;
  /** Initial Display (block, own line) vs inline. Default true (display). When
   *  editing, the toggle is hidden (kind is fixed by what was clicked). */
  initialDisplay?: boolean;
}

export interface EquationEditorHandle {
  close(): void;
}

const FAMILY = MATH_FONT_FAMILY;
const PREVIEW_SIZE = 26;

const LATEX_PLACEHOLDER = "e.g.  \\frac{-b \\pm \\sqrt{b^2-4ac}}{2a}";
const MML_PLACEHOLDER = "<math> … </math>";
const LATEX_HINT = "Type LaTeX math: \\frac{}{}, ^{}, _{}, \\sqrt{}, \\sum, \\int, \\alpha, \\left( \\right), \\begin{bmatrix}…\\end{bmatrix}. Click a template or symbol to insert.";
const MML_HINT = "Edit Presentation MathML directly: <mi> variables, <mn> numbers, <mo> operators, <mfrac>, <msup>, <msqrt>…";

const CSS = `
.cw-eqe{position:fixed;z-index:1002;width:520px;max-width:94vw;background:#fff;border:1px solid #d0d5dd;border-radius:10px;box-shadow:0 12px 40px rgba(0,0,0,.22);font:13px/1.4 system-ui,sans-serif;color:#1a1a2e;display:flex;flex-direction:column;overflow:hidden;}
.cw-eqe-head{display:flex;align-items:center;gap:6px;padding:9px 12px;background:#f5f7fa;border-bottom:1px solid #e6e9ee;font-weight:600;}
.cw-eqe-head .t{flex:1;}
.cw-eqe-x{cursor:pointer;border:none;background:none;font-size:18px;line-height:1;color:#8a9099;padding:2px 4px;border-radius:4px;}
.cw-eqe-x:hover{background:#e6e9ee;color:#1a1a2e;}
.cw-eqe-body{padding:12px;display:flex;flex-direction:column;gap:10px;}
.cw-eqe-palette{display:flex;flex-wrap:wrap;gap:4px;}
.cw-eqe-palette button{cursor:pointer;border:1px solid #d0d5dd;background:#fff;border-radius:5px;padding:3px 7px;font-size:14px;min-width:30px;}
.cw-eqe-palette button:hover{background:#eef2f7;border-color:#9aa4b2;}
.cw-eqe-ta{width:100%;box-sizing:border-box;min-height:120px;font:12px/1.45 Consolas,ui-monospace,monospace;border:1px solid #d0d5dd;border-radius:6px;padding:8px;resize:vertical;}
.cw-eqe-prevwrap{border:1px solid #e6e9ee;border-radius:6px;background:#fafbfc;min-height:64px;display:flex;align-items:center;justify-content:center;padding:8px;overflow:auto;}
.cw-eqe-err{color:#b23b34;font-size:12px;min-height:0;}
.cw-eqe-foot{display:flex;justify-content:flex-end;gap:8px;padding:10px 12px;border-top:1px solid #e6e9ee;background:#fafbfc;}
.cw-eqe-foot button{cursor:pointer;border-radius:6px;padding:6px 14px;font-size:13px;border:1px solid #d0d5dd;background:#fff;}
.cw-eqe-foot .primary{background:#2f6fed;border-color:#2f6fed;color:#fff;font-weight:600;}
.cw-eqe-hint{font-size:11px;color:#8a9099;}
`;

type InputMode = "latex" | "mathml";

/** Palette: label + snippet for each mode (LaTeX / MathML), inserted at the caret. */
const TEMPLATES: { label: string; title: string; latex: string; mml: string }[] = [
  { label: "x/y", title: "Fraction", latex: "\\frac{}{}", mml: "<mfrac><mrow></mrow><mrow></mrow></mfrac>" },
  { label: "x²", title: "Superscript", latex: "^{}", mml: "<msup><mrow></mrow><mrow></mrow></msup>" },
  { label: "xₙ", title: "Subscript", latex: "_{}", mml: "<msub><mrow></mrow><mrow></mrow></msub>" },
  { label: "√", title: "Square root", latex: "\\sqrt{}", mml: "<msqrt><mrow></mrow></msqrt>" },
  { label: "ⁿ√", title: "nth root", latex: "\\sqrt[]{}", mml: "<mroot><mrow></mrow><mrow></mrow></mroot>" },
  { label: "(◻)", title: "Parentheses", latex: "\\left(\\right)", mml: "<mfenced open=\"(\" close=\")\"><mrow></mrow></mfenced>" },
  { label: "∑", title: "Summation with limits", latex: "\\sum_{}^{}", mml: "<munderover><mo>∑</mo><mrow></mrow><mrow></mrow></munderover>" },
  { label: "∫", title: "Integral with limits", latex: "\\int_{}^{}", mml: "<munderover><mo>∫</mo><mrow></mrow><mrow></mrow></munderover>" },
  { label: "[ ]", title: "2×2 matrix", latex: "\\begin{bmatrix} & \\\\ & \\end{bmatrix}", mml: "<mfenced open=\"[\" close=\"]\"><mtable><mtr><mtd><mrow></mrow></mtd><mtd><mrow></mrow></mtd></mtr><mtr><mtd><mrow></mrow></mtd><mtd><mrow></mrow></mtd></mtr></mtable></mfenced>" },
];

const SYMBOLS = ["π", "∞", "±", "×", "÷", "≤", "≥", "≠", "≈", "→", "α", "β", "θ", "λ", "μ", "Σ", "Δ", "∂", "∇", "∈"];

export function showEquationEditor(opts: EquationEditorOptions): EquationEditorHandle {
  injectCssOnce("cw-equation-editor", CSS);
  const ac = new AbortController();

  const backdrop = document.createElement("div");
  backdrop.style.cssText = "position:fixed;inset:0;z-index:1001;";
  const modal = document.createElement("div");
  modal.className = "cw-eqe";
  // Keep canvas selection/focus from being stolen when interacting with the panel.
  modal.addEventListener("mousedown", (e) => e.stopPropagation());

  const head = document.createElement("div");
  head.className = "cw-eqe-head";
  head.innerHTML = `<span class="t">${opts.editing ? "Edit equation" : "Insert equation"}</span>`;
  const x = document.createElement("button");
  x.className = "cw-eqe-x";
  x.textContent = "×";
  x.title = "Close";
  head.appendChild(x);

  const body = document.createElement("div");
  body.className = "cw-eqe-body";

  // Input mode: LaTeX is the friendly default for both new and existing equations
  // (the stored MathML is converted to LaTeX to seed). If that conversion fails,
  // fall back to showing the raw MathML so nothing is lost.
  let inputMode: InputMode = "latex";
  let seedText = "";
  if (opts.initialMathml) {
    try {
      seedText = mathToLatex(parseMathml(opts.initialMathml).root);
    } catch {
      seedText = opts.initialMathml;
      inputMode = "mathml";
    }
  }

  // Input-mode toggle (LaTeX / MathML).
  const inputRow = document.createElement("div");
  inputRow.className = "cw-eqe-hint";
  const mkInput = (label: string, val: InputMode): HTMLLabelElement => {
    const l = document.createElement("label");
    l.style.cssText = "margin-right:14px;cursor:pointer;font-weight:600;";
    const r = document.createElement("input");
    r.type = "radio";
    r.name = "cw-eqe-input";
    r.checked = inputMode === val;
    r.style.marginRight = "4px";
    r.addEventListener("change", () => {
      if (!r.checked || inputMode === val) return;
      // Convert the current content into the new representation (both directions),
      // so switching views never loses or corrupts the equation.
      const text = ta.value.trim();
      if (text) {
        try {
          const root = inputMode === "latex" ? latexToMath(text) : parseMathml(text).root;
          ta.value = val === "latex" ? mathToLatex(root) : serializeMathml({ root, display: displayMode });
        } catch {
          /* unparseable in the current mode — leave the text for the user to fix */
        }
      }
      inputMode = val;
      ta.placeholder = inputMode === "latex" ? LATEX_PLACEHOLDER : MML_PLACEHOLDER;
      updateHint();
      update();
    });
    l.append(r, document.createTextNode(label));
    return l;
  };
  inputRow.append(mkInput("LaTeX", "latex"), mkInput("MathML", "mathml"));

  // Template palette
  const palette = document.createElement("div");
  palette.className = "cw-eqe-palette";
  const ta = document.createElement("textarea");
  ta.className = "cw-eqe-ta";
  ta.spellcheck = false;
  ta.value = seedText;
  ta.placeholder = inputMode === "latex" ? LATEX_PLACEHOLDER : MML_PLACEHOLDER;

  const insertAtCaret = (text: string): void => {
    const s = ta.selectionStart ?? ta.value.length;
    const e = ta.selectionEnd ?? ta.value.length;
    ta.value = ta.value.slice(0, s) + text + ta.value.slice(e);
    // Place the caret inside the first empty {} / <mrow> so the next keystroke fills it.
    const hole = inputMode === "latex" ? text.indexOf("{}") : text.indexOf("></");
    const caret = hole >= 0 ? s + hole + 1 : s + text.length;
    ta.setSelectionRange(caret, caret);
    ta.focus();
    update();
  };
  for (const t of TEMPLATES) {
    const b = document.createElement("button");
    b.textContent = t.label;
    b.title = t.title;
    b.addEventListener("click", () => insertAtCaret(inputMode === "latex" ? t.latex : t.mml));
    palette.appendChild(b);
  }
  for (const sym of SYMBOLS) {
    const b = document.createElement("button");
    b.textContent = sym;
    b.title = `Insert ${sym}`;
    b.addEventListener("click", () => insertAtCaret(inputMode === "latex" ? sym : `<mo>${sym}</mo>`));
    palette.appendChild(b);
  }

  // Display vs inline toggle (hidden when editing — kind is fixed by the click).
  let displayMode = opts.initialDisplay ?? true;
  const modeRow = document.createElement("div");
  modeRow.className = "cw-eqe-hint";
  if (!opts.editing) {
    const mk = (label: string, val: boolean): HTMLLabelElement => {
      const l = document.createElement("label");
      l.style.cssText = "margin-right:14px;cursor:pointer;";
      const r = document.createElement("input");
      r.type = "radio";
      r.name = "cw-eqe-mode";
      r.checked = displayMode === val;
      r.style.marginRight = "4px";
      r.addEventListener("change", () => { if (r.checked) { displayMode = val; update(); } });
      l.append(r, document.createTextNode(label));
      return l;
    };
    modeRow.append(mk("Display (own line)", true), mk("Inline (in text)", false));
  }

  const hint = document.createElement("div");
  hint.className = "cw-eqe-hint";
  const updateHint = (): void => { hint.textContent = inputMode === "latex" ? LATEX_HINT : MML_HINT; };
  updateHint();

  const prevWrap = document.createElement("div");
  prevWrap.className = "cw-eqe-prevwrap";
  const canvas = document.createElement("canvas");
  prevWrap.appendChild(canvas);
  const err = document.createElement("div");
  err.className = "cw-eqe-err";

  body.append(inputRow, palette, modeRow, ta, hint, prevWrap, err);

  const foot = document.createElement("div");
  foot.className = "cw-eqe-foot";
  const cancel = document.createElement("button");
  cancel.textContent = "Cancel";
  const apply = document.createElement("button");
  apply.className = "primary";
  apply.textContent = opts.editing ? "Apply" : "Insert";
  foot.append(cancel, apply);

  modal.append(head, body, foot);
  backdrop.appendChild(modal);
  document.body.appendChild(backdrop);

  // ---- live preview --------------------------------------------------------
  function currentEquation(): MathEquation | null {
    const xml = ta.value.trim();
    if (!xml) return null;
    try {
      const root = inputMode === "latex" ? latexToMath(xml) : parseMathml(xml).root;
      return { root, display: displayMode };
    } catch {
      return null;
    }
  }

  function update(): void {
    const eq = currentEquation();
    const ctx = canvas.getContext("2d");
    if (!eq || !ctx) {
      err.textContent = eq ? "" : "Could not parse — check the syntax.";
      // Clear any stale render so a parse failure doesn't show the previous result.
      if (ctx) ctx.clearRect(0, 0, canvas.width, canvas.height);
      return;
    }
    err.textContent = "";
    try {
      const box = equationBox(eq, FAMILY, PREVIEW_SIZE);
      const pad = 14;
      const dpr = window.devicePixelRatio || 1;
      const w = Math.max(40, box.width + 2 * pad);
      const h = Math.max(36, box.ascent + box.descent + 2 * pad);
      canvas.width = Math.ceil(w * dpr);
      canvas.height = Math.ceil(h * dpr);
      canvas.style.width = `${w}px`;
      canvas.style.height = `${h}px`;
      ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
      ctx.clearRect(0, 0, w, h);
      paintMathBoxCanvas(ctx, { box, width: box.width, height: box.ascent + box.descent, baseline: box.ascent }, pad, pad + box.ascent, FAMILY, "#1a1a2e");
    } catch (e) {
      err.textContent = `Render error: ${e instanceof Error ? e.message : String(e)}`;
    }
  }
  ta.addEventListener("input", update);

  // ---- lifecycle -----------------------------------------------------------
  const close = (): void => {
    ac.abort();
    backdrop.remove();
  };
  x.addEventListener("click", close);
  cancel.addEventListener("click", close);
  apply.addEventListener("click", () => {
    const eq = currentEquation();
    if (!eq) {
      err.textContent = "Fix the MathML before inserting.";
      return;
    }
    opts.onApply(eq);
    close();
  });
  document.addEventListener("keydown", (e) => { if (e.key === "Escape") close(); }, { signal: ac.signal, capture: true });

  makeFloatingDialog({ backdrop, modal, handle: head, signal: ac.signal, noDrag: ".cw-eqe-x" });
  update();
  ta.focus();

  return { close };
}

/** Serialize an equation back to MathML for seeding the editor when editing. */
export function equationToMathmlString(eq: MathEquation): string {
  return serializeMathml(eq);
}
