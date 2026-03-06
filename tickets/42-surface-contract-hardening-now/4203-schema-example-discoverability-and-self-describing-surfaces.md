# Ticket: 4203 Schema / Example Discoverability + Self-Describing Surfaces

## Depends On
- 4202

## Why
Agents should not have to inspect source or infer payload shapes from trial and error. The surfaces should expose canonical schemas and examples directly.

## Owner / Effort / Risk
- Owner (proposed): CLI / MCP / SDK
- Effort: M
- Risk: Med

## Scope
Add discoverability pathways for canonical request/response shapes.

### CLI
- Add commands such as:
  - `asp schema <command-or-op>`
  - `asp example <command-or-op>`
- Keep output compact, copy-pastable, and machine-usable.

### MCP
- Expose tool examples and canonical payload notes in tool metadata or sidecar docs.
- Ensure examples match actual validation behavior.

### SDK
- Export example builders or typed helpers where appropriate.
- Make capability-specific examples explicit.

## Non-Goals
- Full interactive TUI help.
- Replacing formal documentation with examples alone.

## Tests
- Example outputs validate against actual command/tool inputs.
- Drift tests fail when examples diverge from the canonical contract.
- Help/docs link to the same canonical examples.

## Definition of Done
- An agent can ask the product itself for the right shape instead of reverse-engineering it.
