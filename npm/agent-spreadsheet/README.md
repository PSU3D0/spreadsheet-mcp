# agent-spreadsheet (npm)

Stateless spreadsheet CLI for AI agents.

This package installs a native prebuilt `agent-spreadsheet` binary for your platform.

## Install

```bash
npm i -g agent-spreadsheet
agent-spreadsheet --help
```

## Quickstart

```bash
agent-spreadsheet list-sheets workbook.xlsx
agent-spreadsheet copy workbook.xlsx /tmp/workbook-edit.xlsx
agent-spreadsheet edit /tmp/workbook-edit.xlsx Sheet1 "A1=42"
agent-spreadsheet diff workbook.xlsx /tmp/workbook-edit.xlsx
```

## Environment

- `AGENT_SPREADSHEET_DOWNLOAD_BASE_URL` (optional): override release download host.
