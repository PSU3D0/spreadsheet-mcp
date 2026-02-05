# Ticket: 2001 rules_batch (Data Validation v1)

## Why (Human Operator Replacement)
Humans add dropdowns and constraints to prevent bad inputs. For template-driven automation, validation is a core "operator" feature that reduces downstream damage.

## Scope
- Add a new write tool: `rules_batch`.
- Implement v1 op: `set_data_validation` supporting:
  - list dropdowns
  - numeric constraints (whole/decimal)
  - date constraints
  - custom formula constraints
- Support prompt/error messages.

## Non-Goals
- Structural edit rewrite of validations (deferred).
- Full parity for all DV flags (errorStyle, showDropDown, etc.).

## Proposed Tool Surface
```json
{
  "fork_id": "fork-123",
  "ops": [
    {
      "kind": "set_data_validation",
      "sheet_name": "Inputs",
      "target_range": "B3:B100",
      "validation": {
        "kind": "list",
        "formula1": "=Lists!$A$1:$A$10",
        "allow_blank": false,
        "prompt": {"title":"Choose a category","message":"Pick from dropdown"},
        "error": {"title":"Invalid","message":"Choose a listed value"}
      }
    }
  ],
  "mode": "preview"
}
```

Warnings:
- `WARN_VALIDATION_FORMULA_NOT_PARSED` if formula is accepted as opaque string.

## Implementation Notes
- Use umya `DataValidation` + `DataValidations`.
- Strategy:
  - Add/replace validation entries whose `sqref` matches `target_range`.
  - Keep operation idempotent for repeated calls.
- Consider a `replace_mode` later (merge vs replace validations).

## Tests
- List validation persists and Excel shows dropdown (verify by OOXML read-back).
- Prompt/error fields persist.
- Preview/apply staging works.

## Definition of Done
- Common template validations can be added reliably with one call.
- Output is stable and idempotent.
