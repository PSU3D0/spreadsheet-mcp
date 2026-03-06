# Ticket: 4205 Thin Canonical Docs, Skills, and Agent Workflows

## Depends On
- 4202
- 4203
- 4204

## Why
The current guidance contains good instincts, but the benchmark feedback shows it is too repetitive and too expensive to consume.

We need thinner, more canonical guidance.

## Owner / Effort / Risk
- Owner (proposed): Docs / Agent UX
- Effort: M
- Risk: Low

## Scope
Refactor docs and skills around short operational checklists backed by exact examples.

### Content Goals
- One short workflow per task family.
- One exact JSON example per op family.
- One troubleshooting section for common failure modes.
- Minimal overlap between skills and reference docs.

### Priority Areas
- session mutation workflow
- read/discover workflow
- post-edit verification workflow
- benchmark/operator quickstart

## Non-Goals
- Rewriting all docs in one sweep.
- Adding new product semantics.

## Tests
- Docs/examples are exercised by smoke tests where practical.
- Skill links resolve and point to canonical examples.
- No duplicate conflicting examples remain for targeted workflows.

## Definition of Done
- Agents can follow short, canonical workflows without rereading large overlapping guidance.
