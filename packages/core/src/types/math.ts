// OMML AST shared by docx/xlsx/pptx parsers. Mirrors ECMA-376 §22.1.2.
// Each Rust parser emits JSON matching these types; the core math engine consumes them.
export type MathStyle = 'roman' | 'italic' | 'bold' | 'boldItalic';

export interface MathRun {
  kind: 'run';
  text: string;
  style: MathStyle;
}
export interface MathFraction {
  kind: 'fraction';
  num: MathNode[];
  den: MathNode[];
  /** false = no rule (e.g. binomial); defaults to true. */
  bar?: boolean;
}
export interface MathScript {
  kind: 'sup' | 'sub' | 'subSup';
  base: MathNode[];
  sup?: MathNode[];
  sub?: MathNode[];
}
export interface MathNary {
  kind: 'nary';
  /** operator char, e.g. '∑', '∫', '∏'. */
  op: string;
  /** limit location (`m:limLoc`): 'subSup' = beside the op, 'undOvr' = above/below.
   *  Empty/omitted = default by operator class (integrals → subSup, others → undOvr). */
  limLoc?: string;
  sub?: MathNode[];
  sup?: MathNode[];
  body: MathNode[];
}
export interface MathDelimiter {
  kind: 'delimiter';
  /** opening char (default '('). */
  begChar: string;
  /** closing char (default ')'). */
  endChar: string;
  /** separated groups (e.g. for cases / multiple args). */
  items: MathNode[][];
}
export interface MathRadical {
  kind: 'radical';
  /** optional index (e.g. cube root); empty/omitted = square root. */
  index?: MathNode[];
  radicand: MathNode[];
}
/** Lower/upper limit (`m:limLow` / `m:limUpp`), e.g. lim under n→∞. */
export interface MathLimit {
  kind: 'limit';
  base: MathNode[];
  lower?: MathNode[];
  upper?: MathNode[];
}
/** Matrix (`m:m`) or aligned equation array (`m:eqArr`). rows → cells → nodes. */
export interface MathArray {
  kind: 'array';
  rows: MathNode[][][];
  /** 'eq' = alternating right/left (eqArr); 'center' = matrix; 'left'. */
  align: 'eq' | 'center' | 'left';
}
/** Group character (`m:groupChr`), e.g. under/over brace. */
export interface MathGroupChr {
  kind: 'groupChr';
  char: string;
  pos: 'top' | 'bot';
  base: MathNode[];
}
/** Over/under bar (`m:bar`). */
export interface MathBar {
  kind: 'bar';
  pos: 'top' | 'bot';
  base: MathNode[];
}
/** Accent (`m:acc`), e.g. hat, bar, vector arrow over the base. */
export interface MathAccent {
  kind: 'accent';
  char: string;
  base: MathNode[];
}
export interface MathFunc {
  kind: 'func';
  name: MathNode[];
  arg: MathNode[];
}
export interface MathGroup {
  kind: 'group';
  items: MathNode[];
}
/** Phantom object (`m:phant`, §22.1.2.81): contributes the spacing of `base`
 *  while optionally hiding it and/or zeroing individual dimensions. */
export interface MathPhant {
  kind: 'phant';
  /** §22.1.2.96 `m:show` — `false` hides the base (invisible but occupies space,
   *  i.e. `<mphantom>`); `true` (default) shows it and the phant only tweaks
   *  spacing. */
  show: boolean;
  /** §22.1.2 zeroWid / zeroAsc / zeroDesc — suppress width / ascent / descent so
   *  the base takes no space along that axis. Omitted ⇒ false. */
  zeroWid?: boolean;
  zeroAsc?: boolean;
  zeroDesc?: boolean;
  base: MathNode[];
}
/** Pre-sub-superscript object (`m:sPre`, §22.1.2.99): sub + sup to the LEFT of
 *  the base (e.g. ²₁A). */
export interface MathSPre {
  kind: 'sPre';
  sub: MathNode[];
  sup: MathNode[];
  base: MathNode[];
}
/** Box object (`m:box`, §22.1.2.13): a logical grouping (operator emulator /
 *  line-break control). Draws NO border — a transparent group around `base`. */
export interface MathBox {
  kind: 'box';
  base: MathNode[];
}
/** Border-box object (`m:borderBox`, §22.1.2.11): a border/strikes around the
 *  base. Absent flags ⇒ a full rectangular box. */
export interface MathBorderBox {
  kind: 'borderBox';
  /** §22.1.2 hide* — when true the corresponding edge is NOT drawn. */
  hideTop?: boolean;
  hideBot?: boolean;
  hideLeft?: boolean;
  hideRight?: boolean;
  /** §22.1.2 strike* — strikeBLTR = bottom-left→top-right, strikeTLBR =
   *  top-left→bottom-right diagonal. */
  strikeH?: boolean;
  strikeV?: boolean;
  strikeBltr?: boolean;
  strikeTlbr?: boolean;
  base: MathNode[];
}

export type MathNode =
  | MathRun
  | MathFraction
  | MathScript
  | MathNary
  | MathDelimiter
  | MathRadical
  | MathLimit
  | MathArray
  | MathGroupChr
  | MathBar
  | MathAccent
  | MathFunc
  | MathGroup
  | MathPhant
  | MathSPre
  | MathBox
  | MathBorderBox;

const KINDS = new Set([
  'run',
  'fraction',
  'sup',
  'sub',
  'subSup',
  'nary',
  'delimiter',
  'radical',
  'limit',
  'array',
  'groupChr',
  'bar',
  'accent',
  'func',
  'group',
  'phant',
  'sPre',
  'box',
  'borderBox',
]);

export function isMathNode(v: unknown): v is MathNode {
  return (
    !!v &&
    typeof v === 'object' &&
    KINDS.has((v as { kind?: string }).kind ?? '')
  );
}
