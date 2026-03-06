# Design: 4501 Structural Workbook Model

## Problem

Headless workflows currently reason mostly in terms of cells, ranges, and ad hoc heuristics. That is not enough to reliably support higher-level workflow helpers and proof surfaces.

## Goals

- Create a reusable structural representation for workbook patterns relevant to editing and verification.
- Make uncertainty explicit via confidence/ambiguity metadata.
- Support downstream helpers such as footer-aware insert, template cloning, and grouped verification.

## Non-Goals

- Perfect understanding of every workbook.
- Full semantic understanding of business intent.

## Proposed model layers

### Layer 1: raw detected structures
- regions
- header rows/bands
- footer/subtotal rows
- merged layout cues
- formula density zones

### Layer 2: derived workflow anchors
- append targets
- insert-before-footer anchors
- template row candidates
- likely totals zones
- named structural anchors

### Layer 3: confidence / ambiguity metadata
- confidence score or band
- reasons / evidence
- ambiguity notes

## Design principles

- preserve the lower-level evidence, not only the conclusion
- support partial understanding
- prefer explicit uncertainty over false precision
- keep the model machine-consumable and testable

## Consumers

- workflow helpers
- grouped diff/verification
- future contract/assertion systems
- benchmark harnesses

## Success criteria

- Higher-level helpers stop rediscovering structure independently.
- Structural assumptions become inspectable and testable.
- Ambiguous workbooks fail safely rather than pretending certainty.
