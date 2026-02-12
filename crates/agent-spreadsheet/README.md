# agent-spreadsheet

Stateless spreadsheet CLI for AI agents.

Typical workflow:

```bash
agent-spreadsheet list-sheets workbook.xlsx
agent-spreadsheet copy workbook.xlsx /tmp/workbook-edit.xlsx
agent-spreadsheet edit /tmp/workbook-edit.xlsx Sheet1 "A1=42"
agent-spreadsheet recalculate /tmp/workbook-edit.xlsx
agent-spreadsheet diff workbook.xlsx /tmp/workbook-edit.xlsx
```

This crate reuses shared primitives from `spreadsheet-kit` and tool logic from `spreadsheet-mcp`.
