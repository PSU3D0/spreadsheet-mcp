# Design: 4502 Dependency-Cone Verification Engine

## Problem

Current verification can show direct changes, but the stronger headless advantage is explaining downstream consequences: what changed because of an edit, what newly broke in the affected cone, and what stayed outside the blast radius.

## Goals

- Build dependency-scoped verification over a baseline/current comparison.
- Distinguish direct edits from propagated effects.
- Report newly broken outputs within relevant dependency cones.

## Non-Goals

- Full formula-theorem correctness.
- Perfect support for every unsupported engine edge case on day one.

## Proposed model

### Inputs
- baseline state
- current state
- optional seed cells/ranges/named ranges
- optional scope filters

### Analysis stages
1. identify direct edits
2. expand dependency cones from those edits or target seeds
3. compare baseline vs current within the scoped cones
4. classify direct vs propagated changes
5. summarize newly introduced errors and target deltas

### Candidate outputs
- direct edits
- propagated impacts
- unchanged protected areas
- new errors in cone
- affected targets
- confidence / unsupported-analysis notes

## Key product value

This is the path to answering:
- "what did my edit cause?"
- "what newly broke downstream?"
- "which targets changed because of this insertion?"

## Success criteria

- Dependency-scoped verification is usable on representative scenarios.
- Output stays compact enough for agent workflows.
- Provenance is more useful than generic diff alone.
