# Design: 4301 Post-Edit Verification + Provenance

## Problem

After edits and recalc, the current surfaces make it too expensive to answer simple but critical questions:

- Did my target outputs change as intended?
- Which errors are new versus pre-existing?
- What changed directly versus only as a downstream consequence?

Raw diffs and recalc output are necessary but not sufficient.

## Goals

- Make post-edit proof a first-class surface.
- Compare a baseline and current state with explicit target and error summaries.
- Separate direct edits, recalculated values, and newly introduced errors.
- Keep the response compact enough for agent use.

## Non-Goals

- Full dependency-cone causality (that is strategic follow-on work).
- A visual workbook diff renderer.

## Proposed comparison model

### Inputs
- baseline reference
  - file path, session snapshot, fork checkpoint, or explicit workbook ref
- current reference
- optional scope selectors
  - target cells
  - named ranges
  - sheets
  - error-only mode

### Core outputs
- `target_deltas`
  - before, after, changed, classification
- `new_errors`
  - cell, error type, formula/value context, likely source class
- `preexisting_errors`
- `named_range_deltas`
- `summary`
  - total direct edits
  - total recalculated cells
  - total new errors
  - total changed targets

### Classification hints
Suggested top-level change classes:
- `direct_edit`
- `formula_shift`
- `recalc_result`
- `new_error`
- `named_range_change`
- `unchanged`

## API shape principles

- summary-first
- explicit empty arrays
- no hidden inference from omitted fields
- filters for scope and verbosity

## Surface projection

### CLI
- likely `asp verify ...`
- support baseline/current comparison and target lists
- compact JSON default

### MCP
- verification tool with structured output and filters

### SDK
- helper methods returning normalized structures
- convenience functions for target-only or error-only checks

## Implementation notes

This design should reuse existing compare/read/diff primitives where possible, but project them into a verification-specific contract.

The main product move is not inventing new raw data; it is packaging the right proof surface.

## Success criteria

- A moderate edit flow can answer "what changed and what newly broke?" in one verification step.
- Agents stop reconstructing verification manually from multiple noisy outputs.
