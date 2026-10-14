"""Properties that SaveAsText leaves out must not render as 0.

Run: python -m unittest discover -s tests
"""

import os
import re
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import accform_render as ar  # noqa: E402


def form(defaults: str, header: str, detail: str) -> str:
    return f"""Version =20
VersionRequired =20
Begin Form
    Width =6000
    Begin
{defaults}
        Begin FormHeader
            Height =400
            Name ="FormHeader"
            Begin
{header}
            End
        End
        Begin Section
            Height =600
            Name ="Detail"
            Begin
{detail}
            End
        End
    End
End
"""


COMBO = """                Begin ComboBox
                    Left =100
                    Top =100
                    Width =1500
                    Name ="cboVertrag"
                End"""

LINE = """                Begin Line
                    Left =3000
                    Width =0
                    Height =500
                    Name ="Trenner"
                End
                Begin Line
                    Left =100
                    Top =550
                    Name ="Waagerecht"
                End"""


def rects(svg: str) -> list[tuple[float, float, float, float]]:
    out = []
    for m in re.finditer(r'<rect x="([\d.]+)" y="([\d.]+)" width="([\d.]+)" height="([\d.]+)"', svg):
        out.append(tuple(float(v) for v in m.groups()))
    return out


def combo_height(svg: str) -> float:
    # x=100 twips -> 6.667 px; the section backgrounds start at x=0
    return next(h for x, y, w, h in rects(svg) if abs(x - 100 / 15) < 0.01)


class DefaultHeight(unittest.TestCase):
    def test_combo_without_height_gets_access_builtin(self):
        svg = ar.render_svg(form("", "", COMBO))
        # no FontSize anywhere -> MS Sans Serif 8: 8 * 24 + 48 = 240 twips = 16 px
        self.assertAlmostEqual(combo_height(svg), 16.0, places=2)

    def test_combo_inherits_height_from_default_style(self):
        defaults = """        Begin ComboBox
            Width =1701
            Height =300
        End"""
        svg = ar.render_svg(form(defaults, "", COMBO))
        self.assertAlmostEqual(combo_height(svg), 20.0, places=2)

    def test_font_size_from_default_style(self):
        defaults = """        Begin ComboBox
            FontSize =11
        End"""
        svg = ar.render_svg(form(defaults, "", COMBO))
        self.assertAlmostEqual(combo_height(svg), (11 * 24 + 48) / 15, places=2)

    def test_width_inherits_from_default_style(self):
        defaults = """        Begin ComboBox
            Width =1701
        End"""
        combo_ohne_breite = COMBO.replace("                    Width =1500\n", "")
        svg = ar.render_svg(form(defaults, "", combo_ohne_breite))
        breite = next(w for x, y, w, h in rects(svg) if abs(x - 100 / 15) < 0.01)
        self.assertAlmostEqual(breite, 1701 / 15, places=2)

    def test_lines_get_no_font_height(self):
        defaults = """        Begin Line
            Width =1701
        End"""
        svg = ar.render_svg(form(defaults, "", LINE))
        lines = re.findall(r'<line x1="([\d.]+)" y1="([\d.]+)" x2="([\d.]+)" y2="([\d.]+)"', svg)
        senkrecht, waagerecht = [tuple(float(v) for v in l) for l in lines]
        self.assertEqual(senkrecht[0], senkrecht[2])            # explicit Width =0 stays
        self.assertEqual(waagerecht[1], waagerecht[3])          # no Height -> horizontal
        self.assertAlmostEqual(waagerecht[2] - waagerecht[0], 1701 / 15, places=2)


class Sections(unittest.TestCase):
    def test_header_is_rendered_above_detail(self):
        header = """                Begin Label
                    Left =100
                    Top =100
                    Width =1500
                    Height =226
                    Name ="lblVertrag"
                    Caption ="move to Contract"
                End"""
        svg = ar.render_svg(form("", header, COMBO))
        self.assertIn("move to Contract", svg)
        self.assertLess(svg.index('id="FormHeader"'), svg.index('id="Detail"'))
        # total height = header + detail
        self.assertIn('height="66.667"', svg.split("\n")[1])


if __name__ == "__main__":
    unittest.main()
