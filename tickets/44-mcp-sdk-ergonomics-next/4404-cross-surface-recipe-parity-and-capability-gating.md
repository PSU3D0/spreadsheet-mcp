# Ticket: 4404 Cross-Surface Recipe Parity + Capability Gating

## Depends On
- 4401
- 4402
- Existing surface capability matrix / boundary rules

## Why
As higher-level helpers are added, the project needs a disciplined way to express where they are shared, where they are MCP-only, and how SDK callers discover those differences.

## Owner / Effort / Risk
- Owner (proposed): Surface Architecture / SDK
- Effort: M
- Risk: Med

## Scope
Extend the capability matrix and parity harness to cover recipe/helper-level functionality.

### Required work
- classify helpers as shared vs MCP-only vs SDK-only convenience
- expose capability detection cleanly
- add parity/drift coverage where shared behavior is intended

## Non-Goals
- Making every helper available on every surface.
- Weakening the existing boundary rules.

## Tests
- Capability drift is caught automatically.
- Shared helpers produce equivalent semantics across intended surfaces.
- Unsupported helper calls return explicit capability errors.

## Definition of Done
- Higher-level helpers do not become an uncontrolled semantics fork.
