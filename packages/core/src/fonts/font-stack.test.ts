import { describe, expect, it } from 'vitest';
import {
  ARABIC_SUBSTITUTE_FACES,
  NASKH_ARABIC_SUBSTITUTE_FACES,
  officeFaceWebAlias,
} from './font-stack';

describe('font-stack substitution policy views', () => {
  it('aliases only Office faces whose substitute is another family', () => {
    expect(officeFaceWebAlias('Calibri')).toBe('Carlito');
    expect(officeFaceWebAlias('CAMBRIA')).toBe('Caladea');
    expect(officeFaceWebAlias('Franklin Gothic Medium')).toBe('Libre Franklin');
    expect(officeFaceWebAlias('Univers Next Arabic')).toBe('Noto Sans Arabic');
    // Families served under their own name and the Noto faces are not aliases.
    expect(officeFaceWebAlias('Roboto')).toBeUndefined();
    expect(officeFaceWebAlias('Noto Naskh Arabic')).toBeUndefined();
    expect(officeFaceWebAlias('constructor')).toBeUndefined();
  });

  it('derives the Arabic substitute face sets from the substitution policy', () => {
    expect([...ARABIC_SUBSTITUTE_FACES].sort()).toEqual([
      'arabic typesetting', 'noto naskh arabic', 'noto sans arabic', 'sakkal majalla',
      'simplified arabic', 'traditional arabic', 'univers next arabic',
    ]);
    expect([...NASKH_ARABIC_SUBSTITUTE_FACES].sort()).toEqual([
      'arabic typesetting', 'noto naskh arabic', 'sakkal majalla',
      'simplified arabic', 'traditional arabic',
    ]);
  });
});
