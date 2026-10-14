"""
accform_render.py

Renders a Microsoft Access SaveAsText form/report definition (the *.bas layout
file exported by the joyfullservice msaccess-vcs-addin, "Split VBA from designer
layouts" mode) into SVG, without needing Access.

Ported from the add-in's VBA classes:
  - clsFormLayoutParser  -> LayoutParser
  - clsLayoutNode        -> LayoutNode
  - clsFormLayoutSvgWriter -> SvgWriter

Faithful to the add-in for: tree parsing, twips->px scaling (px = twips/15 @96dpi),
section stacking, painter's order, and the Label/TextBox/ComboBox/ListBox/
CommandButton/Rectangle/Line rendering. Reconstructed (add-in source was
truncated when this was written): CheckBox/OptionButton, Image, Tab, Subform,
OptionGroup, generic placeholder, and colour resolution. Colour resolution
handles absolute OLE colours exactly; theme-referenced colours fall back to the
supplied default (see resolve_node_color).

Properties a control does not write itself are taken from the form's default
control styles, and a missing height of text-like controls falls back to a
font-based estimate of the Access built-in (see ACCESS_DEFAULT_FONT_SIZE).
"""

from __future__ import annotations

TWIPS_PER_PX = 15.0  # px = twips * 96 / 1440 = twips / 15
GENERATOR_VERSION = "py-port-1.1"

# SaveAsText omits every property that equals its default. The default comes
# from the form's default control style (the "Begin <Type>" blocks without a
# Name at the top of the container) and, failing that, from Access itself.
# A control without its own Height therefore has a height - it is just not
# written. These constants cover the Access built-ins the renderer needs.
#
# Access's built-in font for controls is MS Sans Serif 8 pt: a form whose
# default styles set no FontSize uses it (modern forms write FontSize =11
# explicitly in their default styles).
ACCESS_DEFAULT_FONT_SIZE = 8.0
# Controls whose built-in height follows the font. The height is approximated
# as font size * 24 + 48 twips: 8 pt -> 240 (0.423 cm), 11 pt -> 312, close to
# the 240/315 Access uses. Lines, rectangles, images, subforms etc. have no
# font-based height and keep 0 if nothing is written.
FONT_HEIGHT_TYPES = frozenset(
    {"TextBox", "ComboBox", "Label", "ListBox", "CommandButton", "ToggleButton"}
)
# Section keywords in file order of appearance; forms and reports.
SECTION_TYPES = frozenset(
    {
        "FormHeader", "PageHeader", "ReportHeader", "BreakHeader", "Section",
        "BreakFooter", "PageFooter", "FormFooter", "ReportFooter",
    }
)


# --------------------------------------------------------------------------- #
# LayoutNode  (port of clsLayoutNode)
# --------------------------------------------------------------------------- #
class LayoutNode:
    __slots__ = ("control_type", "props", "children")

    def __init__(self, control_type: str = "") -> None:
        self.control_type = control_type
        self.props: dict[str, str] = {}
        self.children: list["LayoutNode"] = []

    def add_child(self, child: "LayoutNode") -> None:
        self.children.append(child)

    def has_prop(self, key: str) -> bool:
        return key in self.props

    def prop_val(self, key: str, default=None):
        return self.props.get(key, default)

    def _twip(self, key: str) -> float:
        raw = self.props.get(key)
        if raw is None:
            return 0.0
        try:
            return float(str(raw).strip())
        except (TypeError, ValueError):
            return 0.0

    @property
    def twip_left(self) -> float:
        return self._twip("Left")

    @property
    def twip_top(self) -> float:
        return self._twip("Top")

    @property
    def twip_width(self) -> float:
        return self._twip("Width")

    @property
    def twip_height(self) -> float:
        return self._twip("Height")


# --------------------------------------------------------------------------- #
# LayoutParser  (port of clsFormLayoutParser.Parse)
# --------------------------------------------------------------------------- #
class LayoutParser:
    def parse(self, content: str) -> LayoutNode:
        lines = content.replace("\r\n", "\n").replace("\r", "\n").split("\n")

        root = LayoutNode("Root")
        stack: list[LayoutNode] = [root]
        current = root

        in_hex = False
        hex_key = ""
        hex_parts: list[str] = []

        for raw in lines:
            trimmed = raw.strip()

            # Stop at the VBA code boundary (present only if code not split out)
            if trimmed == "CodeBehindForm":
                break

            # Multi-line hex data blocks (ImageData, RecSrcDt, ...)
            if in_hex:
                if trimmed == "End":
                    current.props[hex_key] = "".join(hex_parts)
                    hex_parts = []
                    in_hex = False
                else:
                    hex_parts.append(trimmed.replace(",", "").replace(" ", ""))
                continue

            if not trimmed:
                continue

            # Begin <Type>  -> new child node
            if trimmed.startswith("Begin "):
                child = LayoutNode(trimmed[6:].strip())
                current.add_child(child)
                stack.append(child)
                current = child
                continue

            # bare "Begin" -> anonymous Defaults block
            if trimmed == "Begin":
                child = LayoutNode("Defaults")
                current.add_child(child)
                stack.append(child)
                current = child
                continue

            # End -> pop
            if trimmed == "End":
                if len(stack) > 1:
                    stack.pop()
                    current = stack[-1]
                continue

            # Property "Key =Value"
            pos = trimmed.find(" =")
            if pos > 0:
                key = trimmed[:pos].strip()
                value = trimmed[pos + 2:].strip()
                if value == "Begin":  # start of a hex data block
                    in_hex = True
                    hex_key = key
                    hex_parts = []
                    continue
                value = _strip_quotes(value)
                current.props[key] = value
                continue

            # Standalone flag "Key=Value"
            eq = trimmed.find("=")
            if eq > 0:
                key = trimmed[:eq].strip()
                value = _strip_quotes(trimmed[eq + 1:].strip())
                current.props[key] = value

        return root


def _strip_quotes(value: str) -> str:
    if len(value) >= 2 and value[0] == '"' and value[-1] == '"':
        return value[1:-1]
    return value


# --------------------------------------------------------------------------- #
# Colour helpers
# --------------------------------------------------------------------------- #
def ole_to_hex(value) -> str | None:
    """Convert an Access/OLE colour long (BGR) to #RRGGBB. Returns None for
    system colours (negative) or unparseable values."""
    try:
        v = int(str(value).strip())
    except (TypeError, ValueError):
        return None
    if v < 0:  # system colour reference (e.g. -2147483633) - not resolvable here
        return None
    b = (v >> 16) & 0xFF
    g = (v >> 8) & 0xFF
    r = v & 0xFF
    return f"#{r:02X}{g:02X}{b:02X}"


# --------------------------------------------------------------------------- #
# SvgWriter  (port of clsFormLayoutSvgWriter)
# --------------------------------------------------------------------------- #
class SvgWriter:
    def __init__(self, scale_mode: str = "Universal", image_embed: str = "Small"):
        self.scale_mode = scale_mode
        self.image_embed = image_embed
        self._lines: list[str] = []
        self._indent = 0
        self._styles: dict[str, dict[str, str]] = {}

    # -- public ------------------------------------------------------------- #
    def generate(self, root: LayoutNode) -> str:
        self._lines = []
        self._indent = 0

        form = self._find_form_node(root)
        if form is None:
            return ""

        self._styles = self._collect_default_styles(form)
        form_w = _to_float(form.prop_val("Width", 10000))
        total_h = self._calc_total_height(form)
        dbl_w = form_w / TWIPS_PER_PX
        dbl_h = total_h / TWIPS_PER_PX

        self._add(
            f"<!-- generator: accform layout-svg {GENERATOR_VERSION}, "
            f"mode: {self.scale_mode.lower()}, scale: 96dpi -->"
        )
        self._add(
            '<svg xmlns="http://www.w3.org/2000/svg" '
            f'width="{fmt_px(dbl_w)}" height="{fmt_px(dbl_h)}" '
            f'viewBox="0 0 {fmt_px(dbl_w)} {fmt_px(dbl_h)}" '
            'font-family="Calibri, sans-serif">'
        )
        self._indent += 1

        sections: list[LayoutNode] = []
        self._collect_sections(form, sections)

        y_off = 0.0
        for section in sections:
            self._render_section(section, y_off, dbl_w)
            y_off += section.twip_height / TWIPS_PER_PX

        self._indent -= 1
        self._add("</svg>")
        return "\n".join(self._lines)

    # -- tree helpers ------------------------------------------------------- #
    @staticmethod
    def _find_form_node(root: LayoutNode):
        for child in root.children:
            if child.control_type in ("Form", "Report"):
                return child
        return None

    def _calc_total_height(self, form: LayoutNode) -> float:
        sections: list[LayoutNode] = []
        self._collect_sections(form, sections)
        return sum(s.twip_height for s in sections)

    def _collect_sections(self, node: LayoutNode, out: list[LayoutNode]) -> None:
        # All section kinds, in file order - the order Access stacks them in
        # design view. Collecting only "Section" (the detail) dropped the form
        # header with its column captions and the footer with its buttons.
        for child in node.children:
            if child.control_type in SECTION_TYPES:
                out.append(child)
            elif child.control_type == "Defaults":
                self._collect_sections(child, out)

    @staticmethod
    def _collect_default_styles(form: LayoutNode) -> dict[str, dict[str, str]]:
        """The form's default control styles: control type -> properties.

        They live in the anonymous container next to the sections, as
        "Begin <Type>" blocks without a Name.
        """
        styles: dict[str, dict[str, str]] = {}
        for container in form.children:
            if container.control_type != "Defaults":
                continue
            for child in container.children:
                if child.control_type in SECTION_TYPES or child.has_prop("Name"):
                    continue
                styles.setdefault(child.control_type, {}).update(child.props)
        return styles

    # -- emit --------------------------------------------------------------- #
    def _add(self, line: str) -> None:
        self._lines.append(("  " * self._indent) + line)

    def _render_section(self, section: LayoutNode, y_off: float, form_w: float) -> None:
        h = section.twip_height / TWIPS_PER_PX
        name = section.prop_val("Name", "Section")
        back = self._resolve_color(section, "Back", "#F0F0F0")

        self._add(f'<g id="{escape_xml(name)}">')
        self._indent += 1
        self._add(
            f'<rect x="0" y="{fmt_px(y_off)}" width="{fmt_px(form_w)}" '
            f'height="{fmt_px(h)}" fill="{back}" />'
        )
        self._render_children(section, y_off)
        self._indent -= 1
        self._add("</g>")

    def _render_children(self, parent: LayoutNode, section_y: float) -> None:
        for child in parent.children:
            if child.control_type == "Defaults":
                self._render_children(child, section_y)
            else:
                self._render_control(child, section_y)

    def _render_control(self, node: LayoutNode, section_y: float) -> None:
        ct = node.control_type
        dispatch = {
            "Label": self._render_label,
            "TextBox": self._render_textbox,
            "ComboBox": self._render_combobox,
            "ListBox": self._render_textbox,
            "CommandButton": self._render_button,
            "Rectangle": self._render_rectangle,
            "Line": self._render_line,
            "CheckBox": self._render_checkbox,
            "OptionButton": self._render_checkbox,
            "ToggleButton": self._render_button,
            "Image": self._render_image,
            "Tab": self._render_tab,
            "Subform": self._render_subform,
            "OptionGroup": self._render_optiongroup,
        }
        fn = dispatch.get(ct)
        if fn is not None:
            fn(node, section_y)
        elif ct in ("Section", "Defaults", "Page"):
            pass
        elif node.has_prop("Left") and node.has_prop("Top"):
            self._render_placeholder(node, section_y)

        # Attached child controls (e.g. a label bound to a textbox), unless the
        # control renders its own children.
        if ct not in ("Tab", "OptionGroup", "Section", "Defaults", "Page"):
            if node.children:
                self._render_children(node, section_y)

    # -- geometry / colour -------------------------------------------------- #
    def _prop(self, node: LayoutNode, key: str, default=None):
        """A property as Access sees it: the control's own value, else the
        form's default style for its type, else ``default``."""
        if node.has_prop(key):
            return node.prop_val(key)
        return self._styles.get(node.control_type, {}).get(key, default)

    def _twips(self, node: LayoutNode, key: str) -> float:
        return _to_float(self._prop(node, key, 0), 0.0)

    def _height_twips(self, node: LayoutNode) -> float:
        h = self._twips(node, "Height")
        if h <= 0 and node.control_type in FONT_HEIGHT_TYPES:
            size = _to_float(self._prop(node, "FontSize", ACCESS_DEFAULT_FONT_SIZE),
                             ACCESS_DEFAULT_FONT_SIZE)
            h = size * 24 + 48
        return h

    def _rect(self, node: LayoutNode, section_y: float):
        x = node.twip_left / TWIPS_PER_PX
        y = section_y + node.twip_top / TWIPS_PER_PX
        w = self._twips(node, "Width") / TWIPS_PER_PX
        h = self._height_twips(node) / TWIPS_PER_PX
        return x, y, w, h

    def _resolve_color(self, node: LayoutNode, prefix: str, default: str) -> str:
        # Exact path: absolute OLE colour on "<prefix>Color".
        if node.has_prop(prefix + "Color"):
            hx = ole_to_hex(node.prop_val(prefix + "Color"))
            if hx is not None:
                return hx
        # Theme-referenced colour: not resolvable without the theme palette;
        # fall back to the supplied default (documented approximation).
        return default

    def _opacity(self, node: LayoutNode) -> float:
        return 0.3 if node.prop_val("Visible", "") == "NotDefault" else 1.0

    # -- control renderers -------------------------------------------------- #
    def _render_label(self, node, section_y):
        x, y, w, h = self._rect(node, section_y)
        caption = str(node.prop_val("Caption", ""))
        fore = self._resolve_color(node, "Fore", "#000000")
        font_size = _to_float(node.prop_val("FontSize", 11)) * 0.75
        weight = "bold" if _to_int(node.prop_val("FontWeight", 400)) >= 700 else "normal"
        op = self._opacity(node)

        if node.prop_val("BackStyle", "0") == "1":
            back = self._resolve_color(node, "Back", "#FFFFFF")
            self._add(
                f'<rect x="{fmt_px(x)}" y="{fmt_px(y)}" width="{fmt_px(w)}" '
                f'height="{fmt_px(h)}" fill="{back}"{opacity_attr(op)} />'
            )
        if caption:
            self._add(
                f'<text x="{fmt_px(x + 2)}" y="{fmt_px(y + h * 0.75)}" '
                f'font-size="{fmt_px(font_size)}" font-weight="{weight}" '
                f'fill="{fore}"{opacity_attr(op)} clip-path="inset(0)">'
                f"{escape_xml(caption)}</text>"
            )

    def _render_textbox(self, node, section_y):
        x, y, w, h = self._rect(node, section_y)
        back = self._resolve_color(node, "Back", "#FFFFFF")
        border = self._resolve_color(node, "Border", "#ABABAB")
        op = self._opacity(node)
        self._add(
            f'<rect x="{fmt_px(x)}" y="{fmt_px(y)}" width="{fmt_px(w)}" '
            f'height="{fmt_px(h)}" fill="{back}" stroke="{border}" '
            f'stroke-width="1"{opacity_attr(op)} />'
        )

    def _render_combobox(self, node, section_y):
        x, y, w, h = self._rect(node, section_y)
        back = self._resolve_color(node, "Back", "#FFFFFF")
        border = self._resolve_color(node, "Border", "#ABABAB")
        op = self._opacity(node)
        btn_w = h * 0.8
        self._add(
            f'<rect x="{fmt_px(x)}" y="{fmt_px(y)}" width="{fmt_px(w)}" '
            f'height="{fmt_px(h)}" fill="{back}" stroke="{border}" '
            f'stroke-width="1"{opacity_attr(op)} />'
        )
        self._add(
            f'<line x1="{fmt_px(x + w - btn_w)}" y1="{fmt_px(y)}" '
            f'x2="{fmt_px(x + w - btn_w)}" y2="{fmt_px(y + h)}" '
            f'stroke="{border}" stroke-width="1"{opacity_attr(op)} />'
        )

    def _render_button(self, node, section_y):
        x, y, w, h = self._rect(node, section_y)
        caption = str(node.prop_val("Caption", ""))
        back = self._resolve_color(node, "Back", "#E1E1E1")
        fore = self._resolve_color(node, "Fore", "#000000")
        border = self._resolve_color(node, "Border", "#ABABAB")
        font_size = _to_float(node.prop_val("FontSize", 11)) * 0.75
        op = self._opacity(node)
        self._add(
            f'<rect x="{fmt_px(x)}" y="{fmt_px(y)}" width="{fmt_px(w)}" '
            f'height="{fmt_px(h)}" rx="{fmt_px(3)}" fill="{back}" '
            f'stroke="{border}" stroke-width="1"{opacity_attr(op)} />'
        )
        if caption:
            caption = caption.replace("&&", "&").replace("&", "")
            if caption.startswith(" "):
                caption = caption[1:]
            self._add(
                f'<text x="{fmt_px(x + w / 2)}" '
                f'y="{fmt_px(y + h / 2 + font_size / 3)}" '
                f'font-size="{fmt_px(font_size)}" fill="{fore}" '
                f'text-anchor="middle"{opacity_attr(op)}>'
                f"{escape_xml(caption)}</text>"
            )

    def _render_rectangle(self, node, section_y):
        x, y, w, h = self._rect(node, section_y)
        op = self._opacity(node)
        if node.prop_val("BackStyle", "0") == "1":
            back = self._resolve_color(node, "Back", "#FFFFFF")
        else:
            back = "none"
        border = self._resolve_color(node, "Border", "none") if node.has_prop("BorderColor") else "none"
        self._add(
            f'<rect x="{fmt_px(x)}" y="{fmt_px(y)}" width="{fmt_px(w)}" '
            f'height="{fmt_px(h)}" fill="{back}" stroke="{border}" '
            f'stroke-width="1"{opacity_attr(op)} />'
        )

    def _render_line(self, node, section_y):
        x1 = node.twip_left / TWIPS_PER_PX
        y1 = section_y + node.twip_top / TWIPS_PER_PX
        x2 = x1 + self._twips(node, "Width") / TWIPS_PER_PX
        y2 = y1 + self._twips(node, "Height") / TWIPS_PER_PX
        color = self._resolve_color(node, "Border", "#000000")
        width = _to_float(node.prop_val("BorderWidth", 1)) / TWIPS_PER_PX
        if width < 1:
            width = 1
        op = self._opacity(node)
        self._add(
            f'<line x1="{fmt_px(x1)}" y1="{fmt_px(y1)}" x2="{fmt_px(x2)}" '
            f'y2="{fmt_px(y2)}" stroke="{color}" stroke-width="{fmt_px(width)}"'
            f'{opacity_attr(op)} />'
        )

    # ---- reconstructed renderers (add-in source was truncated) ------------ #
    def _render_checkbox(self, node, section_y):
        x, y, w, h = self._rect(node, section_y)
        border = self._resolve_color(node, "Border", "#ABABAB")
        op = self._opacity(node)
        size = min(h, 12.0) if h else 12.0
        self._add(
            f'<rect x="{fmt_px(x)}" y="{fmt_px(y)}" width="{fmt_px(size)}" '
            f'height="{fmt_px(size)}" fill="#FFFFFF" stroke="{border}" '
            f'stroke-width="1"{opacity_attr(op)} />'
        )

    def _render_image(self, node, section_y):
        x, y, w, h = self._rect(node, section_y)
        op = self._opacity(node)
        self._add(
            f'<rect x="{fmt_px(x)}" y="{fmt_px(y)}" width="{fmt_px(w)}" '
            f'height="{fmt_px(h)}" fill="#EDEDED" stroke="#ABABAB" '
            f'stroke-width="1"{opacity_attr(op)} />'
        )
        self._add(
            f'<line x1="{fmt_px(x)}" y1="{fmt_px(y)}" x2="{fmt_px(x + w)}" '
            f'y2="{fmt_px(y + h)}" stroke="#ABABAB" stroke-width="1"'
            f'{opacity_attr(op)} />'
        )

    def _render_tab(self, node, section_y):
        x, y, w, h = self._rect(node, section_y)
        op = self._opacity(node)
        strip = min(h, 24.0) if h else 24.0
        self._add(
            f'<rect x="{fmt_px(x)}" y="{fmt_px(y)}" width="{fmt_px(w)}" '
            f'height="{fmt_px(h)}" fill="none" stroke="#ABABAB" '
            f'stroke-width="1"{opacity_attr(op)} />'
        )
        self._add(
            f'<line x1="{fmt_px(x)}" y1="{fmt_px(y + strip)}" '
            f'x2="{fmt_px(x + w)}" y2="{fmt_px(y + strip)}" stroke="#ABABAB" '
            f'stroke-width="1"{opacity_attr(op)} />'
        )
        # Render controls that live on the tab pages.
        self._render_children(node, section_y)

    def _render_subform(self, node, section_y):
        x, y, w, h = self._rect(node, section_y)
        op = self._opacity(node)
        self._add(
            f'<rect x="{fmt_px(x)}" y="{fmt_px(y)}" width="{fmt_px(w)}" '
            f'height="{fmt_px(h)}" fill="#F7F7F7" stroke="#ABABAB" '
            f'stroke-width="1" stroke-dasharray="4 2"{opacity_attr(op)} />'
        )
        src = node.prop_val("SourceObject", "")
        if src:
            self._add(
                f'<text x="{fmt_px(x + 4)}" y="{fmt_px(y + 14)}" font-size="9" '
                f'fill="#777777"{opacity_attr(op)}>{escape_xml(str(src))}</text>'
            )

    def _render_optiongroup(self, node, section_y):
        x, y, w, h = self._rect(node, section_y)
        op = self._opacity(node)
        self._add(
            f'<rect x="{fmt_px(x)}" y="{fmt_px(y)}" width="{fmt_px(w)}" '
            f'height="{fmt_px(h)}" fill="none" stroke="#ABABAB" '
            f'stroke-width="1"{opacity_attr(op)} />'
        )
        self._render_children(node, section_y)

    def _render_placeholder(self, node, section_y):
        x, y, w, h = self._rect(node, section_y)
        op = self._opacity(node)
        self._add(
            f'<rect x="{fmt_px(x)}" y="{fmt_px(y)}" width="{fmt_px(w)}" '
            f'height="{fmt_px(h)}" fill="none" stroke="#CCCCCC" '
            f'stroke-width="1" stroke-dasharray="3 3"{opacity_attr(op)} />'
        )


# --------------------------------------------------------------------------- #
# formatting helpers
# --------------------------------------------------------------------------- #
def fmt_px(v: float) -> str:
    s = f"{float(v):.3f}".rstrip("0").rstrip(".")
    return s if s not in ("", "-0") else "0"


def opacity_attr(op: float) -> str:
    return "" if op >= 1 else f' opacity="{fmt_px(op)}"'


def escape_xml(text: str) -> str:
    return (
        str(text)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def _to_float(v, default: float = 0.0) -> float:
    try:
        return float(str(v).strip())
    except (TypeError, ValueError):
        return default


def _to_int(v, default: int = 0) -> int:
    try:
        return int(float(str(v).strip()))
    except (TypeError, ValueError):
        return default


# --------------------------------------------------------------------------- #
# top-level API
# --------------------------------------------------------------------------- #
def render_svg(source_text: str, scale_mode: str = "Universal",
               image_embed: str = "Small") -> str:
    root = LayoutParser().parse(source_text)
    return SvgWriter(scale_mode, image_embed).generate(root)


def render_png(source_text: str, scale_mode: str = "Universal") -> bytes:
    import cairosvg  # imported lazily so SVG-only use needs no cairo
    svg = render_svg(source_text, scale_mode)
    if not svg:
        raise ValueError("No Form/Report node found in source (empty SVG).")
    return cairosvg.svg2png(bytestring=svg.encode("utf-8"))


if __name__ == "__main__":
    import sys
    path = sys.argv[1]
    with open(path, "r", encoding="utf-8-sig") as fh:
        print(render_svg(fh.read()))
