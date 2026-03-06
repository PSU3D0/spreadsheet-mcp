# Ticket: 4401 Task-Oriented MCP Workflow Helpers + Recipes

## Depends On
- Tranches 42 and 43 foundations
- Existing MCP session/fork/staging surface

## Why
MCP is powerful, but still too plumbing-oriented for common workflows. Agents need fewer, sharper tools for safe stateful operations.

## Owner / Effort / Risk
- Owner (proposed): MCP / Agent Workflow
- Effort: L
- Risk: High

## Scope
Add a task-oriented recipe layer on top of existing MCP orchestration.

### Candidate helpers
- safe edit/recalc/verify workflow bundle
- append/insert workflow recipes
- checkpoint + verify + summarize recipe
- target-delta verification recipe

### Design constraints
- Must not fork shared semantics unnecessarily.
- Must respect boundary rules: MCP owns orchestration, core owns semantics.

## Non-Goals
- Replacing the existing low-level MCP tools.
- Hiding every advanced option behind recipes.

## Tests
- Recipes produce equivalent outcomes to the underlying low-level sequences.
- Failure surfaces remain explicit and inspectable.
- Capability boundaries stay documented.

## Definition of Done
- Common stateful agent workflows can be expressed with fewer orchestration round-trips.
