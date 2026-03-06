# Ticket: 4201 `asp` Command Obviousness, Install, and Help Surface

## Depends On
- Existing `agent-spreadsheet` binary packaging
- Existing npm CLI wrapper packaging

## Why
The benchmark burned significant time/tokens just discovering how to invoke the tool. That is unacceptable for an agent-first benchmark surface.

The command name, installation story, README examples, and `--help` surface need to agree.

## Owner / Effort / Risk
- Owner (proposed): CLI / Packaging
- Effort: M
- Risk: Low

## Scope
Make the operator and agent entrypoint obvious and self-consistent.

### CLI / Packaging
- Decide the supported invocation contract:
  - `asp`
  - `agent-spreadsheet`
  - or both via first-class aliasing
- Ensure install flows actually land the supported command on `PATH`.
- Align cargo, npm, docs, and examples.

### Help Surface
- Add a concise top-level quickstart to CLI help.
- Ensure help text points to canonical workflows rather than only enumerating commands.
- Reduce hidden discovery cost for common tasks.

### Docs / Examples
- Remove examples that assume unavailable aliases.
- Add one install verification step:
  - `asp --version`
  - `asp --help`

## Non-Goals
- Full CLI command-tree redesign.
- Workflow/task helper addition.

## Tests
### Packaging / Smoke
- Installed CLI exposes the documented binary name(s).
- README install paths work on supported environments.

### CLI UX
- Top-level help contains a short workflow-oriented quickstart.
- Example commands in docs match actual binary names.

## Definition of Done
- An agent/operator can discover the correct executable immediately.
- No official docs reference a command alias that does not actually exist.
