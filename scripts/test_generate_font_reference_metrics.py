"""Focused generator regression: uv run --with fonttools python scripts/test_generate_font_reference_metrics.py."""

import runpy
import unittest
from pathlib import Path

from fontTools.feaLib.builder import addOpenTypeFeaturesFromString
from fontTools.ttLib import TTFont, newTable
from fontTools.ttLib.tables._c_m_a_p import CmapSubtable

generator = runpy.run_path(str(Path(__file__).with_name('generate-font-reference-metrics.py')))
stacked_vertical_glyphs = generator['stacked_vertical_glyphs']
stacked_vrt2_only_glyphs = generator['stacked_vrt2_only_glyphs']


class StackedVerticalCoverageTest(unittest.TestCase):
    def test_unions_vert_features_and_extensions_but_not_vrt2_or_alternates(self):
        font = TTFont()
        font.setGlyphOrder([
            '.notdef', 'paren', 'paren.vert', 'bracket', 'bracket.vert',
            'dots', 'dots.vert', 'parallel', 'parallel.vert', 'letter', 'letter.vert',
            'square', 'square.alt',
        ])
        cmap = CmapSubtable.newSubtable(4)
        cmap.platformID, cmap.platEncID, cmap.language = 3, 1, 0
        # Both code points sharing a glyph must survive the inverse cmap query.
        cmap.cmap = {
            ord('('): 'paren', ord(')'): 'paren', ord('「'): 'bracket',
            ord('…'): 'dots', ord('∥'): 'parallel', ord('A'): 'letter',
            ord('［'): 'square',
        }
        font['cmap'] = newTable('cmap')
        font['cmap'].tableVersion = 0
        font['cmap'].tables = [cmap]
        addOpenTypeFeaturesFromString(font, '''
            languagesystem DFLT dflt;
            languagesystem hani dflt;
            lookup Direct {
                sub paren by paren.vert;
                sub letter by letter.vert;
            } Direct;
            lookup Extended useExtension {
                sub bracket by bracket.vert;
                sub dots by dots.vert;
            } Extended;
            feature vert {
                script DFLT;
                lookup Direct;
                script hani;
                lookup Extended;
                sub square from [square.alt];
            } vert;
            feature vrt2 {
                sub parallel by parallel.vert;
                sub bracket by bracket.vert;
            } vrt2;
        ''')
        self.assertEqual(stacked_vertical_glyphs(font), '()…「')
        # vrt2 coverage is recorded apart, minus what vert already covers.
        self.assertEqual(stacked_vrt2_only_glyphs(font), '∥')

    def test_missing_gsub_has_no_coverage(self):
        self.assertEqual(stacked_vertical_glyphs(TTFont()), '')
        self.assertEqual(stacked_vrt2_only_glyphs(TTFont()), '')


if __name__ == '__main__':
    unittest.main()
