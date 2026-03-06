# Ticket: 4403 CLI vNext Command Grouping + Workflow Help Surface

## Depends On
- 4201
- 4205

## Why
The CLI has accumulated too many top-level commands. Even with better docs, the surface remains harder to navigate than it should be.

## Owner / Effort / Risk
- Owner (proposed): CLI UX
- Effort: L
- Risk: High

## Scope
Design a vNext CLI grouping strategy and workflow-oriented help surface.

### Candidate groups
- `read`
- `analyze`
- `write`
- `verify`
- `session`

### Migration concerns
- preserve backward compatibility where possible
- provide aliases or a clear deprecation path
- ensure docs/help/examples remain coherent during transition

## Non-Goals
- Rewriting all command implementations from scratch.
- Shipping a breaking change without migration planning.

## Tests
- Help output remains discoverable and coherent.
- Existing commands either continue to work or fail with explicit migration guidance.
- Docs/examples stay synchronized.

## Definition of Done
- The CLI surface becomes workflow-oriented rather than a flat wall of commands.
