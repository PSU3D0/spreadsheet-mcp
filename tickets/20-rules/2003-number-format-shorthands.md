# Ticket: 2003 Number Format Shorthands (style_batch normalization)

## Why (Human Operator Replacement)
Humans constantly apply "currency", "percent", and "date" formats. Raw Excel format codes are fiddly and error-prone for agents.

## Scope
- Extend `style_batch` normalization to accept a friendly shorthand:
  - `number_format: { kind: "currency" | "percent" | "date_iso" | "accounting" | "integer" }`
- Expand to concrete Excel format codes in the patch.

## Non-Goals
- Locale-specific formatting.
- A large catalog of formats.

## Proposed Tool Surface
```json
{
  "fork_id": "fork-123",
  "ops": [
    {
      "sheet_name": "Inputs",
      "range": "B3:B100",
      "style": {
        "number_format": {"kind":"currency"}
      }
    }
  ]
}
```

## Implementation Notes
- Implement expansion in style normalization (where other shorthands live).
- Codes (initial):
  - currency: "$#,##0.00"
  - percent: "0.00%"
  - date_iso: "yyyy-mm-dd"
  - integer: "0"
  - accounting: "_($* #,##0.00_)" (confirm desired)

## Tests
- Shorthand expands and persists in stylesheet.
- No changes when explicit format_code is provided.

## Definition of Done
- Agents can apply common formats without memorizing format code syntax.
