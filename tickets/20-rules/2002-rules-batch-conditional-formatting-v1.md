# Ticket: 2002 rules_batch (Conditional Formatting v1)

## Why (Human Operator Replacement)
Human operators use conditional formatting for dashboards and review sheets. Without CF write support, AI outputs are visually incomplete and require manual polish.

## Scope
Extend `rules_batch` with v1 op: `add_conditional_format` supporting:
- `cell_is` rules with operator + formula
- `expression` rules with formula
- simple style payload (fill/font/bold) mapped to a dxf

## Non-Goals
- Color scales, icon sets, data bars (later).
- Structural rewrite of CF formulas (deferred).

## Proposed Tool Surface
```json
{
  "fork_id": "fork-123",
  "ops": [
    {
      "kind": "add_conditional_format",
      "sheet_name": "Dashboard",
      "target_range": "D3:D100",
      "rule": {
        "kind": "cell_is",
        "operator": "lessThan",
        "formula": "0",
        "style": {"fill_color":"FFFFE0E0","font_color":"FF8A0000","bold":true}
      }
    }
  ],
  "mode": "apply"
}
```

Warnings:
- `WARN_CF_FORMULA_NOT_ADJUSTED_ON_STRUCTURE` (static warning in docs + tool description).

## Implementation Notes
- Create or reuse a dxf entry in stylesheet.
- Append a `ConditionalFormatting` block with sqref and one rule.
- Ensure rule `priority` is assigned consistently (e.g., max+1).

## Tests
- CF written with correct sqref and rule type/operator.
- dxf emitted and referenced by dxfId.
- Preview staging if supported.

## Definition of Done
- Basic red/yellow/green thresholds can be authored.
- Excel opens and shows the rule.
