// MathML / OMML conversion layer. The canonical equation form is the MathML AST
// (shared/src/model/math.ts); these functions bridge it to the two serialized
// worlds it must interoperate with:
//   - MathML strings  (paste, the equation editor, a11y)  <-> parse / serialize
//   - OMML (.docx math) (import / export)                 <-> fromOmml / toOmml

export { serializeMathml, mathNodeToMathml } from "./serialize";
export { parseMathml } from "./parse";
export { mathmlToOmml } from "./toOmml";
export { ommlToMathml } from "./fromOmml";
