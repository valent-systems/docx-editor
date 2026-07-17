import { describe, test, expect } from 'bun:test';
import { decomposeWeightedFontName, resolveFontFamily } from '../fontResolver';

// Word names weighted faces by suffixing the family ("Onest SemiBold",
// "Roboto Light"). Treating those as literal families 400s the Google Fonts
// fetch and every run measured in fallback metrics — pagination drifts from
// Word across a whole document.

describe('decomposeWeightedFontName', () => {
  test('recognizes single-word weight suffixes', () => {
    expect(decomposeWeightedFontName('Onest SemiBold')).toEqual({
      baseFamily: 'Onest',
      weight: 600,
    });
    expect(decomposeWeightedFontName('Roboto Light')).toEqual({
      baseFamily: 'Roboto',
      weight: 300,
    });
    expect(decomposeWeightedFontName('Segoe UI Semilight')).toEqual({
      baseFamily: 'Segoe UI',
      weight: 350,
    });
  });

  test('recognizes two-word suffixes ("Semi Bold", "Extra Light")', () => {
    expect(decomposeWeightedFontName('Onest Semi Bold')).toEqual({
      baseFamily: 'Onest',
      weight: 600,
    });
    expect(decomposeWeightedFontName('Roboto Extra Light')).toEqual({
      baseFamily: 'Roboto',
      weight: 200,
    });
  });

  test('leaves genuine families alone', () => {
    // "Bold"/"Black" are not decomposed (Arial Black is a real family; Word
    // encodes bold via w:b) and single words never are.
    expect(decomposeWeightedFontName('Arial Black')).toBeNull();
    expect(decomposeWeightedFontName('Times New Roman')).toBeNull();
    expect(decomposeWeightedFontName('Light')).toBeNull();
    expect(decomposeWeightedFontName('Calibri')).toBeNull();
  });
});

describe('resolveFontFamily with weight-suffixed names', () => {
  test('keeps the original face first, then the base family stack', () => {
    const r = resolveFontFamily('Onest SemiBold');
    expect(r.cssFallback.startsWith('"Onest SemiBold", Onest')).toBe(true);
    // Loader fetches the base family (a real Google Fonts family name).
    expect(r.googleFont).toBe('Onest');
  });

  test('inherits the base family mapping (Calibri Light → Carlito)', () => {
    const r = resolveFontFamily('Calibri Light');
    expect(r.googleFont).toBe('Carlito');
    expect(r.hasGoogleEquivalent).toBe(true);
    // Metrics come from the base family's OS/2 ratio, not the generic default.
    expect(r.singleLineRatio).toBe(resolveFontFamily('Calibri').singleLineRatio);
  });

  test('directly-mapped names are untouched by decomposition', () => {
    expect(resolveFontFamily('Calibri').googleFont).toBe('Carlito');
    expect(resolveFontFamily('Arial').googleFont).toBe('Arimo');
  });
});
