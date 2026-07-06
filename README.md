# accform-layout-mcp

Local MCP server and CLI for Access layout files (`*.bas`, SaveAsText), rendered
to SVG without Microsoft Access. This lets you visually verify form/report
layout changes directly.

## Features

- MCP tools: `render_form`, `render_form_svg`, `list_forms`
- CLI: `accform-cli` for single-file and batch rendering
- Default output under `rendered/YYYY-MM-DD/`
- Optional PNG export via `cairosvg`

## Installation

### Install from repository

```bash
python -m venv .venv
.venv\Scripts\activate
pip install .
```

Optional PNG support:

```bash
pip install ".[png]"
```

## Start the MCP server

After installation:

```bash
accform-layout-mcp
```

Alternative (directly from repository):

```bash
python server.py
```

## Use the CLI

Render a single file:

```bash
accform-cli path/to/F_Contract.bas
```

Render multiple files in one run:

```bash
accform-cli forms/F_Contract.bas forms/F_Customer.bas
```

Use a custom output directory:

```bash
accform-cli forms/F_Contract.bas -o out/svg
```

Generate optional PNG output:

```bash
accform-cli forms/F_Contract.bas --png
```

Show all options:

```bash
accform-cli --help
```

## MCP in VS Code Copilot (Agent Mode)

Example for `.vscode/mcp.json`:

```json
{
  "servers": {
    "accform-layout": {
      "type": "stdio",
      "command": "python",
      "args": ["${workspaceFolder}/server.py"],
      "env": {
        "ACCFORM_SRC_ROOT": "${workspaceFolder}"
      }
    }
  }
}
```

`ACCFORM_SRC_ROOT` limits file discovery to your project folder.

## Tool behavior

- `render_form(form, scale_mode="Universal")`
  - renders SVG to `rendered/YYYY-MM-DD/...`
  - returns the path to the generated file
- `render_form_svg(form, scale_mode="Universal")`
  - returns SVG markup as text
  - also writes the same SVG file to disk
- `list_forms()`
  - lists all detected layout `*.bas` files under `ACCFORM_SRC_ROOT`

## What rendering includes

Static rendering includes declarative designer state:

- positions, sizes, captions
- absolute colors
- visible/hidden controls (hidden controls are dimmed)

It does not include runtime VBA behavior, for example:

- `Form_Load` logic
- dynamic `RecordSource`
- conditional formatting
- data-driven subforms

## Fidelity

Faithfully ported: parse tree, twips-to-pixel conversion, section stacking,
painter order, and renderers for
Label/TextBox/ComboBox/ListBox/CommandButton/Rectangle/Line, including OLE
color conversion.

Reconstructed: CheckBox/OptionButton, Image, Tab, Subform, OptionGroup,
placeholder rendering, and theme color references.

## Development

Install dependencies from `requirements.txt`:

```bash
pip install -r requirements.txt
```

## License

MIT. See `LICENSE`.
