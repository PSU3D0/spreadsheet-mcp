# Design: 4303 Region Append + Footer-Aware Insert Helpers

## Problem

Common row-oriented spreadsheet tasks still require too many primitive operations:

- find insertion point manually
- avoid subtotal/footer rows manually
- insert rows manually
- patch formulas manually
- verify totals manually

This is too much substrate work for a moderate scenario.

## Goals

- Add generic row-oriented helpers that reduce multi-step edit sequences.
- Respect detected regions and footer/subtotal structure where confidence is sufficient.
- Make preview/plan output first-class.

## Non-Goals

- Domain-specific business logic.
- Solving every structural edge case in one release.
- Full structural workbook model.

## Proposed helper families

### 1) append rows into region
Inputs:
- sheet / region reference
- rows payload
- matching mode (header-based, positional, explicit mapping)

Outputs:
- insertion point
- rows appended
- affected range summary
- warnings

### 2) append CSV into region
Same as above, but with import and optional header matching rules.

### 3) insert before footer/subtotal row
Inputs:
- detected region or explicit anchor
- row count or rows payload
- footer policy

Outputs:
- inserted span
- identified footer row(s)
- impacted total/formula zones
- warnings / confidence level

## Detection / confidence model

Helpers should return confidence metadata, for example:
- `high` — clear region/footer pattern
- `medium` — likely pattern, but ambiguous cases exist
- `low` — unsafe to continue automatically

Low-confidence cases should fail or require explicit override.

## Preview-first principle

Every helper should support a dry-run/plan mode that reports:

- insertion point
- affected ranges
- candidate footer/subtotal rows
- formula/totals that may need rebuild or extension
- confidence/warnings

## Relation to existing primitives

These helpers should compile down to existing lower-level transforms/structure ops where possible.

The value is not a new core semantics fork; the value is a better agent-facing workflow layer.

## Success criteria

- Typical append/insert workflows require a few high-level steps.
- Unsafe ambiguity is surfaced explicitly.
- Agents do less manual row arithmetic and subtotal hunting.
