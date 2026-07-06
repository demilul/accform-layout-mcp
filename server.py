"""
server.py - local MCP server that renders an Access form/report layout file
(the *.bas SaveAsText designer file exported by the msaccess-vcs-addin) to an
image, so an AI coding agent (GitHub Copilot agent mode, etc.) can visually
verify UI edits it made to the source without round-tripping through Access.

Run:
    python server.py                     # stdio transport (for Copilot/VS Code)

Env:
    ACCFORM_SRC_ROOT   optional base folder; relative paths passed to the tool
                       are resolved against it (e.g. your M1.accdb.src folder).

Tools:
    render_form(form)          -> saved SVG file path of the current on-disk layout
    render_form_svg(form)      -> SVG text and saved SVG file
    list_forms()               -> available *.bas layout files under the root
"""

from __future__ import annotations

import argparse
import os
import sys
from datetime import date
from pathlib import Path

from fastmcp import FastMCP

import accform_render as ar

__version__ = "0.1.0"

SRC_ROOT = Path(os.environ.get("ACCFORM_SRC_ROOT", ".")).expanduser().resolve()
TOOL_ROOT = Path(__file__).resolve().parent
OUTPUT_ROOT = TOOL_ROOT / "rendered"

mcp = FastMCP("accform-layout")


def _resolve(form: str, src_root: Path | None = None) -> Path:
    """Resolve a form argument to a *.bas layout file path.

    Accepts an absolute path, a path relative to ACCFORM_SRC_ROOT, or a bare
    form name (searched for as '<name>.bas' anywhere under the root).
    """
    root = src_root or SRC_ROOT
    p = Path(form).expanduser()
    if p.is_absolute() and p.exists():
        return p

    cand = (root / form).resolve()
    if cand.exists():
        return cand

    if not form.lower().endswith(".bas"):
        cand = (root / f"{form}.bas").resolve()
        if cand.exists():
            return cand

    stem = Path(form).stem
    matches = list(root.rglob(f"{stem}.bas"))
    if len(matches) == 1:
        return matches[0]
    if len(matches) > 1:
        rels = ", ".join(str(m.relative_to(root)) for m in matches[:10])
        raise FileNotFoundError(
            f"Ambiguous form '{form}'. Matches: {rels}. Pass a more specific path."
        )
    raise FileNotFoundError(
        f"No layout file found for '{form}' under {root}. "
        f"Use list_forms() to see what is available."
    )


def _read(path: Path) -> str:
    # Layout files are exported UTF-8 (no BOM); utf-8-sig is tolerant of either.
    return path.read_text(encoding="utf-8-sig")


def _artifact_path(
    source_path: Path,
    src_root: Path | None = None,
    output_root: Path | None = None,
    dated: bool = True,
) -> Path:
    """Map a source layout file to a dated SVG artifact path."""
    root = src_root or SRC_ROOT
    out_root = output_root or OUTPUT_ROOT
    try:
        rel = source_path.relative_to(root)
    except ValueError:
        rel = Path(source_path.name)

    target_dir = out_root / rel.parent
    if dated:
        target_dir = out_root / date.today().isoformat() / rel.parent
    return target_dir / rel.with_suffix(".svg").name


def _render_svg(
    form: str,
    scale_mode: str = "Universal",
    src_root: Path | None = None,
    output_root: Path | None = None,
    dated_output: bool = True,
) -> tuple[Path, str]:
    root = src_root or SRC_ROOT
    path = _resolve(form, src_root=root)
    svg = ar.render_svg(_read(path), scale_mode=scale_mode)
    if not svg:
        raise ValueError(
            f"No Form/Report node found in '{path.name}'. Is this a layout .bas?"
        )

    artifact = _artifact_path(
        path,
        src_root=root,
        output_root=output_root,
        dated=dated_output,
    )
    artifact.parent.mkdir(parents=True, exist_ok=True)
    artifact.write_text(svg, encoding="utf-8")
    return artifact, svg


@mcp.tool
def render_form(form: str, scale_mode: str = "Universal") -> str:
    """Render an Access form/report layout to SVG and store it under a dated
    artifact folder so the result is inspectable without a native Cairo stack.

    Pass the form name (e.g. 'A_Lot_Liste') or the path to its *.bas layout
    file. The render reflects only the declarative designer layout on disk:
    positions, sizes, captions, absolute colours, hidden controls (dimmed). It
    does NOT reflect runtime VBA behaviour (Form_Load hiding controls, dynamic
    RecordSource, conditional formatting).
    """
    artifact, _svg = _render_svg(form, scale_mode=scale_mode)
    try:
        return str(artifact.relative_to(SRC_ROOT))
    except ValueError:
        return str(artifact)


@mcp.tool
def render_form_svg(form: str, scale_mode: str = "Universal") -> str:
    """Render an Access form/report layout to SVG markup (text). Useful when you
    want to inspect exact control coordinates, sizes and captions rather than
    just look at a picture. Same layout caveats as render_form."""
    _artifact, svg = _render_svg(form, scale_mode=scale_mode)
    return svg


@mcp.tool
def list_forms() -> list[str]:
    """List the Access form/report layout files (*.bas designer files) available
    under the configured source root, as paths relative to it."""
    out: list[str] = []
    for p in sorted(SRC_ROOT.rglob("*.bas")):
        try:
            head = p.read_text(encoding="utf-8-sig", errors="ignore")[:200]
        except OSError:
            continue
        # A designer layout starts with the SaveAsText header, not a VBA module.
        if "Begin Form" in head or "Begin Report" in head or head.lstrip().startswith("Version ="):
            out.append(str(p.relative_to(SRC_ROOT)))
    return out


def mcp_main() -> None:
    """Start MCP server on stdio transport."""
    mcp.run()


def cli_main() -> int:
    """Render one or more Access SaveAsText layout files to SVG from the CLI."""
    parser = argparse.ArgumentParser(
        prog="accform-cli",
        description="Render Access form/report layout .bas files to SVG.",
    )
    parser.add_argument(
        "forms",
        nargs="+",
        help="Form name, relative .bas path, or absolute .bas path (supports multiple).",
    )
    parser.add_argument(
        "-s",
        "--scale-mode",
        default="Universal",
        help="Scale mode passed to renderer (default: Universal).",
    )
    parser.add_argument(
        "-o",
        "--output-dir",
        default=None,
        help="Output directory. Default keeps MCP style: rendered/YYYY-MM-DD/...",
    )
    parser.add_argument(
        "--src-root",
        default=None,
        help="Source root for relative paths and form-name lookup (default: ACCFORM_SRC_ROOT or .).",
    )
    parser.add_argument(
        "--png",
        action="store_true",
        help="Additionally write PNG files (requires cairosvg).",
    )
    parser.add_argument(
        "--quiet",
        action="store_true",
        help="Only print errors.",
    )
    parser.add_argument(
        "--version",
        action="version",
        version=f"%(prog)s {__version__}",
    )

    args = parser.parse_args()

    src_root = Path(args.src_root or os.environ.get("ACCFORM_SRC_ROOT", ".")).expanduser().resolve()
    if args.output_dir:
        out_root = Path(args.output_dir).expanduser().resolve()
        dated_output = False
    else:
        out_root = OUTPUT_ROOT
        dated_output = True

    had_error = False
    png_missing_dep = False

    for form in args.forms:
        try:
            artifact, svg = _render_svg(
                form,
                scale_mode=args.scale_mode,
                src_root=src_root,
                output_root=out_root,
                dated_output=dated_output,
            )
            if not args.quiet:
                print(f"OK  {form} -> {artifact}")

            if args.png:
                try:
                    png_bytes = ar.render_png(_read(_resolve(form, src_root=src_root)), scale_mode=args.scale_mode)
                    png_path = artifact.with_suffix(".png")
                    png_path.write_bytes(png_bytes)
                    if not args.quiet:
                        print(f"OK  {form} -> {png_path}")
                except ImportError:
                    png_missing_dep = True
                    had_error = True
                    print(
                        "ERR PNG output requires 'cairosvg'. Install with 'pip install accform-layout-mcp[png]'.",
                        file=sys.stderr,
                    )
                except Exception as ex:
                    had_error = True
                    print(f"ERR {form}: {ex}", file=sys.stderr)
        except Exception as ex:
            had_error = True
            print(f"ERR {form}: {ex}", file=sys.stderr)

    if had_error and png_missing_dep:
        return 2
    if had_error:
        return 1
    return 0


if __name__ == "__main__":
    mcp_main()
