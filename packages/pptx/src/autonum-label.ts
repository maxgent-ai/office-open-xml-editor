// Auto-number bullet labels, shared by the renderer (paint) and the font
// requests (which scripts a marker's text proves for its family).

/**
 * Format an autoNum bullet label from a counter value and OOXML numType.
 * Spec: ECMA-376 §20.1.10.61 (ST_TextAutonumberScheme).
 *
 * Supports the symmetric Plain / Period / ParenR / ParenBoth variants for
 * the four core scripts (arabic, alphaLc/Uc, romanLc/Uc) plus arabicDb
 * (full-width Arabic digits). CJK / Thai / Hebrew / Arabic-Abjad schemes
 * are intentionally not handled here — they fall through to the default
 * `arabicPeriod` rendering rather than emitting a wrong-looking glyph.
 */
export function formatAutoNum(counter: number, numType: string): string {
  const arabic = `${counter}`;
  const alphaLc = counter >= 1 && counter <= 26
    ? String.fromCharCode(96 + counter)
    : arabic;
  const alphaUc = counter >= 1 && counter <= 26
    ? String.fromCharCode(64 + counter)
    : arabic;
  const romanLc = toRoman(counter).toLowerCase();
  const romanUc = toRoman(counter);
  // ECMA-376 §20.1.10.61 lists `arabicDb*` as full-width Arabic digits
  // (U+FF10–U+FF19), used in East Asian numbered lists.
  const arabicDb = arabic.replace(/[0-9]/g, (d) =>
    String.fromCharCode(0xff10 + (d.charCodeAt(0) - 0x30)),
  );

  switch (numType) {
    case 'arabicPlain':       return arabic;
    case 'arabicPeriod':      return `${arabic}.`;
    case 'arabicParenR':      return `${arabic})`;
    case 'arabicParenBoth':   return `(${arabic})`;
    case 'arabicDbPlain':     return arabicDb;
    case 'arabicDbPeriod':    return `${arabicDb}.`;
    case 'alphaLcPlain':      return alphaLc;
    case 'alphaLcPeriod':     return `${alphaLc}.`;
    case 'alphaLcParenR':     return `${alphaLc})`;
    case 'alphaLcParenBoth':  return `(${alphaLc})`;
    case 'alphaUcPlain':      return alphaUc;
    case 'alphaUcPeriod':     return `${alphaUc}.`;
    case 'alphaUcParenR':     return `${alphaUc})`;
    case 'alphaUcParenBoth':  return `(${alphaUc})`;
    case 'romanLcPlain':      return romanLc;
    case 'romanLcPeriod':     return `${romanLc}.`;
    case 'romanLcParenR':     return `${romanLc})`;
    case 'romanLcParenBoth':  return `(${romanLc})`;
    case 'romanUcPlain':      return romanUc;
    case 'romanUcPeriod':     return `${romanUc}.`;
    case 'romanUcParenR':     return `${romanUc})`;
    case 'romanUcParenBoth':  return `(${romanUc})`;
    default:                  return `${arabic}.`;
  }
}

function toRoman(n: number): string {
  const vals = [1000,900,500,400,100,90,50,40,10,9,5,4,1];
  const syms = ['M','CM','D','CD','C','XC','L','XL','X','IX','V','IV','I'];
  let result = '';
  for (let i = 0; i < vals.length; i++) {
    while (n >= vals[i]) { result += syms[i]; n -= vals[i]; }
  }
  return result;
}
