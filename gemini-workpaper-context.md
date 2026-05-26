# Bilig WorkPaper Tools

Use the `bilig-workpaper` MCP tools when a user needs spreadsheet-style
business logic without opening Excel, Google Sheets, LibreOffice, or a browser
grid.

Preferred loop:

1. Call `list_sheets` or `read_range` to inspect the workbook before editing.
2. Use `set_cell_contents` for one intentional input or formula change.
3. Read the dependent output with `read_cell` or `get_cell_display_value`.
4. Export or persist the WorkPaper JSON when the result must be reproducible.

Treat write-only success as incomplete. A useful answer includes the sheet,
address, old value, new value, recalculated output, and whether readback or
persistence verified the result.

Use formula strings such as `=0.4` or `=TRUE()` when passing numbers or
booleans through clients that expect single-typed MCP parameters.

Do not claim full desktop Excel compatibility, macro execution, external link
refresh, or arbitrary private spreadsheet mutation unless the user provided the
WorkPaper file and the tool response proves the operation.
