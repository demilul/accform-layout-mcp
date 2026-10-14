# Changelog

## Unreleased
- Fixed: controls without their own Height rendered with height 0 (invisible),
  e.g. most combo boxes. SaveAsText omits properties equal to their default;
  Width, Height and FontSize now fall back to the form's default control style
  and, for text-like controls, to a font-based estimate of the Access built-in
  height (MS Sans Serif 8 pt when no font size is set: 240 twips).
- Fixed: only the detail section was rendered. Form/report headers, footers,
  page and group sections are now stacked in file order, as in design view.
- Added unit tests (`python -m unittest discover -s tests`).

## 0.1.0 - 2026-07-03
- Initial public release preparation.
- Added installable package metadata via pyproject.toml.
- Added CLI command accform-cli for single and batch .bas to SVG rendering.
- Added MCP command accform-layout-mcp entry point.
- Added repository hygiene files (.gitignore, LICENSE).
