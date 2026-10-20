import type { MathNode, MathStyle } from '../types/math';
import { ownValue } from '../internal/own-key';

// Convert the shared OMML AST to MathML for MathJax. MathJax owns the typesetting
// (spacing, italic/upright, stretchy glyphs), so this layer only maps structure.

const ESC: Record<string, string> = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' };
const esc = (s: string) => s.replace(/[&<>"]/g, (c) => ESC[c]);

type MathScriptVariant = 'double-struck' | 'fraktur' | 'monospace' | 'roman' | 'sans-serif' | 'script';
// Issue #1037 keeps the published MathNode declaration byte-for-byte compatible.
// The parser therefore carries ECMA-376 m:scr as an internal wire until that API
// parity gate is lifted; it is consumed only at the MathML resource boundary.
type InternalMathRun = Extract<MathNode, { kind: 'run' }> & Readonly<{ __script?: MathScriptVariant }>;

function variant(style: MathStyle, script: MathScriptVariant = 'roman'): string | null {
  if (script === 'double-struck' || script === 'monospace') return script;
  if (script === 'fraktur') return style === 'bold' || style === 'boldItalic' ? 'bold-fraktur' : 'fraktur';
  if (script === 'script') return style === 'bold' || style === 'boldItalic' ? 'bold-script' : 'script';
  if (script === 'sans-serif') {
    switch (style) {
      case 'roman':
        return 'sans-serif';
      case 'italic':
        return 'sans-serif-italic';
      case 'bold':
        return 'bold-sans-serif';
      case 'boldItalic':
        return 'sans-serif-bold-italic';
    }
  }
  switch (style) {
    case 'roman':
      return 'normal';
    case 'bold':
      return 'bold';
    case 'boldItalic':
      return 'bold-italic';
    case 'italic':
      return null; // MathML default for <mi> single letters is italic
  }
}

// Group characters that stretch to span their base (braces, over/under brackets).
// Other group chars (arrows, etc.) are drawn at a fixed accent size.
const GROUP_BRACES = new Set([...'⏞⏟⎴⎵︷︸⏜⏝{}[]()¯_‾']);

const BIN = '+−±∓×÷·∗⋅∘∙*/';
const REL = '=≠<>≤≥≈≡∼≅≃→←↔⇒∈∉⊂⊆⊃⊇∝≪≫⊥≔';
const OPEN = '([{⟨⌈⌊';
const CLOSE = ')]}⟩⌉⌋';
const PUNCT = ',;';

/** Tokenize a run's text into MathML token elements (mi/mn/mo) by character class. */
function runToMathML(text: string, style: MathStyle, script?: MathScriptVariant): string {
  const mv = variant(style, script);
  const mvAttr = mv ? ` mathvariant="${mv}"` : '';
  let out = '';
  let numBuf = '';
  const flushNum = () => {
    if (numBuf) {
      out += `<mn${mvAttr}>${esc(numBuf)}</mn>`;
      numBuf = '';
    }
  };
  for (const ch of text) {
    if (ch === ' ') {
      flushNum();
      continue;
    }
    if (ch >= '0' && ch <= '9') {
      numBuf += ch;
      continue;
    }
    flushNum();
    if (BIN.includes(ch) || REL.includes(ch) || PUNCT.includes(ch)) {
      out += `<mo>${esc(ch)}</mo>`;
    } else if (OPEN.includes(ch)) {
      out += `<mo fence="true" stretchy="false">${esc(ch)}</mo>`;
    } else if (CLOSE.includes(ch)) {
      out += `<mo fence="true" stretchy="false">${esc(ch)}</mo>`;
    } else {
      out += `<mi${mvAttr}>${esc(ch)}</mi>`;
    }
  }
  flushNum();
  return out;
}

function seq(nodes: MathNode[]): string {
  return nodes.map(nodeToMathML).join('');
}

/** Wrap a node list in an mrow so it forms one argument. */
function row(nodes: MathNode[]): string {
  return `<mrow>${seq(nodes)}</mrow>`;
}

function nodeToMathML(node: MathNode): string {
  switch (node.kind) {
    case 'run':
      return runToMathML(node.text, node.style, (node as InternalMathRun).__script);
    case 'group':
      return row(node.items);
    case 'fraction':
      return `<mfrac${node.bar === false ? ' linethickness="0"' : ''}>${row(node.num)}${row(node.den)}</mfrac>`;
    case 'sup':
      return `<msup>${row(node.base)}${row(node.sup ?? [])}</msup>`;
    case 'sub':
      return `<msub>${row(node.base)}${row(node.sub ?? [])}</msub>`;
    case 'subSup':
      return `<msubsup>${row(node.base)}${row(node.sub ?? [])}${row(node.sup ?? [])}</msubsup>`;
    case 'nary':
      return naryToMathML(node);
    case 'delimiter':
      return delimiterToMathML(node);
    case 'radical':
      return node.index && node.index.length
        ? `<mroot>${row(node.radicand)}${row(node.index)}</mroot>`
        : `<msqrt>${seq(node.radicand)}</msqrt>`;
    case 'limit':
      return limitToMathML(node);
    case 'array':
      return arrayToMathML(node);
    case 'groupChr': {
      const b = row(node.base);
      // Brace-like group chars stretch to span the base (overbrace / underbrace).
      // Arrows and other marks render at a fixed size as an accent — Word does not
      // shrink them to a narrow base (a stretchy arrow over a single letter looks
      // cramped / clipped).
      const stretchy = GROUP_BRACES.has(node.char);
      const tag = node.pos === 'top' ? 'mover' : 'munder';
      const mo = `<mo stretchy="${stretchy}">${esc(node.char)}</mo>`;
      return stretchy ? `<${tag}>${b}${mo}</${tag}>` : `<${tag} accent="true">${b}${mo}</${tag}>`;
    }
    case 'bar': {
      // Tight stretchy overline/underline (like \overline). NOTE: accent="true" adds a
      // gap above the base in MathJax, so we omit it — the stretchy bar hugs the base.
      const b = row(node.base);
      const barOp = '<mo stretchy="true">&#x2015;</mo>';
      return node.pos === 'bot' ? `<munder>${b}${barOp}</munder>` : `<mover>${b}${barOp}</mover>`;
    }
    case 'accent':
      return accentToMathML(node);
    case 'func':
      return `<mrow>${row(node.name)}<mo>&#x2061;</mo>${row(node.arg)}</mrow>`;
    case 'phant':
      return phantToMathML(node);
    case 'sPre':
      // §22.1.2.99 pre-sub-superscript: base, then <mprescripts/>, then the
      // prescript pair (sub before sup). No postscripts.
      return `<mmultiscripts>${row(node.base)}<mprescripts/>${row(node.sub)}${row(node.sup)}</mmultiscripts>`;
    case 'box':
      // §22.1.2.13 box: a logical grouping with NO border — a transparent mrow.
      return row(node.base);
    case 'borderBox':
      return borderBoxToMathML(node);
  }
}

/** §22.1.2.81 phant → MathML. `show=false` hides the base (`<mphantom>`, which
 *  renders invisibly but reserves space). A shown phant that zeroes a dimension
 *  suppresses that extent with `<mpadded>` (width/height/depth = 0). A plain
 *  shown phant with no zero* flags is just its base. */
function phantToMathML(node: Extract<MathNode, { kind: 'phant' }>): string {
  const inner = node.show ? seq(node.base) : `<mphantom>${seq(node.base)}</mphantom>`;
  // Map zeroWid/zeroAsc/zeroDesc to <mpadded> extent overrides. In MathML,
  // width = advance, height = ascent (above baseline), depth = descent (below).
  const attrs: string[] = [];
  if (node.zeroWid) attrs.push('width="0"');
  if (node.zeroAsc) attrs.push('height="0"');
  if (node.zeroDesc) attrs.push('depth="0"');
  return attrs.length ? `<mpadded ${attrs.join(' ')}>${inner}</mpadded>` : `<mrow>${inner}</mrow>`;
}

/** §22.1.2.11 borderBox → `<menclose>`. The hide* flags REMOVE edges from the
 *  default full box, so we build the notation from the edges that remain (plus
 *  any strikes). With no edges and no strikes, emit a plain mrow (no enclosure). */
function borderBoxToMathML(node: Extract<MathNode, { kind: 'borderBox' }>): string {
  const notation: string[] = [];
  // Full box unless every edge is hidden — otherwise list the surviving edges.
  const top = !node.hideTop;
  const bot = !node.hideBot;
  const left = !node.hideLeft;
  const right = !node.hideRight;
  if (top && bot && left && right) {
    notation.push('box');
  } else {
    if (top) notation.push('top');
    if (bot) notation.push('bottom');
    if (left) notation.push('left');
    if (right) notation.push('right');
  }
  if (node.strikeH) notation.push('horizontalstrike');
  if (node.strikeV) notation.push('verticalstrike');
  // strikeBLTR = bottom-left→top-right = updiagonalstrike;
  // strikeTLBR = top-left→bottom-right = downdiagonalstrike.
  if (node.strikeBltr) notation.push('updiagonalstrike');
  if (node.strikeTlbr) notation.push('downdiagonalstrike');
  const inner = seq(node.base);
  return notation.length
    ? `<menclose notation="${notation.join(' ')}">${inner}</menclose>`
    : `<mrow>${inner}</mrow>`;
}

// Map an OMML accent character to a MathJax-friendly accent. Combining marks (zero
// advance) float when used as <mo> content, so we translate to spacing equivalents.
const ACCENT_MAP: Record<string, string> = {
  '̀': '`', // grave
  '́': '´', // acute
  '̂': '^', // circumflex / hat
  '̃': '~', // tilde
  '̆': '˘', // breve
  '̇': '˙', // dot above
  '̈': '¨', // diaeresis
  '̌': 'ˇ', // caron
  '⃗': '→', // vector arrow
  '⃖': '←', // left arrow
};
// Overline / macron family → rendered as a tight stretchy bar (\overline), no accent gap.
const OVERLINE_CHARS = new Set(['̅', '̄', '¯', '‾', '̲', '̳']);

function accentToMathML(node: Extract<MathNode, { kind: 'accent' }>): string {
  const b = row(node.base);
  if (OVERLINE_CHARS.has(node.char)) {
    return `<mover>${b}<mo stretchy="true">&#x2015;</mo></mover>`;
  }
  const ch = ownValue(ACCENT_MAP, node.char) ?? node.char;
  const stretchy = ch === '→' || ch === '←' ? 'true' : 'false';
  return `<mover accent="true">${b}<mo stretchy="${stretchy}">${esc(ch)}</mo></mover>`;
}

function limitToMathML(node: Extract<MathNode, { kind: 'limit' }>): string {
  const base = row(node.base);
  const lower = node.lower && node.lower.length ? row(node.lower) : null;
  const upper = node.upper && node.upper.length ? row(node.upper) : null;
  if (lower && upper) return `<munderover>${base}${lower}${upper}</munderover>`;
  if (lower) return `<munder>${base}${lower}</munder>`;
  if (upper) return `<mover>${base}${upper}</mover>`;
  return base;
}

function arrayToMathML(node: Extract<MathNode, { kind: 'array' }>): string {
  const cols = Math.max(1, ...node.rows.map((r) => r.length));
  let colalign: string;
  if (node.align === 'eq') {
    // alternating right/left so cells align at the relation (& marker)
    colalign = Array.from({ length: cols }, (_, i) => (i % 2 === 0 ? 'right' : 'left')).join(' ');
  } else if (node.align === 'left') {
    colalign = 'left';
  } else {
    colalign = 'center';
  }
  const rows = node.rows
    .map((cells) => `<mtr>${cells.map((c) => `<mtd>${seq(c)}</mtd>`).join('')}</mtr>`)
    .join('');
  return `<mtable columnalign="${colalign}" rowspacing="0.2em" columnspacing="0.3em">${rows}</mtable>`;
}

// Integrals place limits beside the operator; sums/products/etc. place them above/below.
const INTEGRAL_OPS = new Set([...'∫∬∭∮∯∰∱∲∳⨌⨍⨎⨏⨐⨑⨒⨓⨔⨕⨖⨗']);

function naryToMathML(node: Extract<MathNode, { kind: 'nary' }>): string {
  const side =
    node.limLoc === 'subSup'
      ? true
      : node.limLoc === 'undOvr'
        ? false
        : INTEGRAL_OPS.has(node.op);
  const op = `<mo largeop="true">${esc(node.op)}</mo>`;
  const sub = node.sub ?? [];
  const sup = node.sup ?? [];
  let operator: string;
  if (side) {
    // limits to the right of the operator (integral style)
    if (sub.length && sup.length) operator = `<msubsup>${op}${row(sub)}${row(sup)}</msubsup>`;
    else if (sub.length) operator = `<msub>${op}${row(sub)}</msub>`;
    else if (sup.length) operator = `<msup>${op}${row(sup)}</msup>`;
    else operator = op;
  } else {
    // limits above/below (sum/product style)
    if (sub.length && sup.length) operator = `<munderover>${op}${row(sub)}${row(sup)}</munderover>`;
    else if (sub.length) operator = `<munder>${op}${row(sub)}</munder>`;
    else if (sup.length) operator = `<mover>${op}${row(sup)}</mover>`;
    else operator = op;
  }
  return `<mrow>${operator}${seq(node.body)}</mrow>`;
}

function delimiterToMathML(node: Extract<MathNode, { kind: 'delimiter' }>): string {
  // An EMPTY begChr/endChr ("") means "no delimiter on that side" (e.g. cases use
  // '{' on the left and nothing on the right). The parser already defaults an
  // *absent* dPr to '(' / ')', so an empty string here is intentional → emit an
  // invisible stretchy fence rather than falling back to a paren.
  const fence = (ch: string) => `<mo fence="true" stretchy="true">${esc(ch)}</mo>`;
  const inner = node.items.map((g) => row(g)).join('<mo separator="true">,</mo>');
  return `<mrow>${fence(node.begChar)}${inner}${fence(node.endChar)}</mrow>`;
}

/** Full MathML document string for a formula. `display` selects block vs inline. */
export function mathToMathML(nodes: MathNode[], display: boolean): string {
  const mode = display ? 'block' : 'inline';
  return `<math xmlns="http://www.w3.org/1998/Math/MathML" display="${mode}">${seq(nodes)}</math>`;
}
