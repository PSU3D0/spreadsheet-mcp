# Design: 4403 CLI vNext Command Grouping + Workflow Help Surface

## Problem

The CLI has grown into a large flat wall of commands. That is tolerable for human operators skimming a help page, but it is poor for agent-driven discovery.

Agents benefit from progressive discovery:

- `asp --help`
- `asp read --help`
- `asp write --help`
- `asp write batch --help`

Today the surface does not support that pattern well because too many unrelated commands live at the top level.

## Architectural constraints

This ticket is a **CLI taxonomy refactor**, not a cross-surface rename.

### CLI boundary
The CLI is a stateless, path-driven adapter.

- parser / routing: `crates/spreadsheet-kit/src/cli/mod.rs`
- path-driven adapters: `crates/spreadsheet-kit/src/cli/commands/**`

### MCP boundary
The MCP server does **not** shell out to the CLI.
It exposes tool handlers directly against Rust tool functions.

- server router: `crates/spreadsheet-mcp/src/server.rs`
- shared tool surface: `crates/spreadsheet-kit/src/tools/**`

### SDK boundary
The JS SDK does **not** shell out to the CLI.
It is a backend abstraction over MCP and WASM.

- MCP backend: `npm/spreadsheet-kit-sdk/src/mcp-backend.js`
- WASM backend: `npm/spreadsheet-kit-sdk/src/wasm-backend.js`
- normalization layer: `npm/spreadsheet-kit-sdk/src/backend.js`

### Consequence
4403 must **not** rename MCP tools or SDK methods just because CLI commands are regrouped.
The runtime blast radius is the CLI only. Cross-surface docs that mention CLI names should be updated, but MCP and SDK runtime semantics stay unchanged.

## Goals

- Land **true nested commands** as the primary CLI surface.
- Make the CLI workflow-oriented and progressively discoverable.
- Keep `schema` / `example` aligned with the nested taxonomy.
- Preserve compatibility temporarily via hidden legacy aliases.
- Avoid rewriting existing command implementations from scratch.

## Non-Goals

- Rewriting MCP tool names.
- Rewriting SDK method names.
- Creating a fourth semantics fork.
- Reimplementing the command business logic in a new layer.

## Proposed final command tree

```text
asp
├── read
│   ├── sheets
│   ├── overview
│   ├── values
│   ├── export
│   ├── cells
│   ├── page
│   ├── table
│   ├── names
│   ├── workbook
│   └── layout
│
├── analyze
│   ├── find-value
│   ├── find-formula
│   ├── formula-map
│   ├── formula-trace
│   ├── scan-volatiles
│   ├── sheet-statistics
│   ├── table-profile
│   └── ref-impact
│
├── write
│   ├── cells
│   ├── import
│   ├── append
│   ├── clone-template-row
│   ├── clone-row-band
│   ├── formulas
│   │   └── replace
│   ├── name
│   │   ├── define
│   │   ├── update
│   │   └── delete
│   └── batch
│       ├── transform
│       ├── style
│       ├── formula-pattern
│       ├── structure
│       ├── column-size
│       ├── sheet-layout
│       └── rules
│
├── workbook
│   ├── create
│   ├── copy
│   └── recalculate
│
├── verify
│   ├── proof
│   └── diff
│
├── schema
│   ├── write batch transform
│   ├── write batch style
│   ├── write batch formula-pattern
│   ├── write batch structure
│   ├── write batch column-size
│   ├── write batch sheet-layout
│   ├── write batch rules
│   └── session op <kind>
│
├── example
│   ├── write batch transform
│   ├── write batch style
│   ├── write batch formula-pattern
│   ├── write batch structure
│   ├── write batch column-size
│   ├── write batch sheet-layout
│   ├── write batch rules
│   └── session op <kind>
│
├── session
│   └── ... existing nested session commands unchanged
│
└── sheetport
    └── ... existing nested sheetport commands unchanged
```

## Old → new mapping

| Old command | New canonical command |
|---|---|
| `list-sheets` | `read sheets` |
| `sheet-overview` | `read overview` |
| `range-values` | `read values` |
| `range-export` | `read export` |
| `inspect-cells` | `read cells` |
| `sheet-page` | `read page` |
| `read-table` | `read table` |
| `named-ranges` | `read names` |
| `describe` | `read workbook` |
| `layout-page` | `read layout` |
| `find-value` | `analyze find-value` |
| `find-formula` | `analyze find-formula` |
| `formula-map` | `analyze formula-map` |
| `formula-trace` | `analyze formula-trace` |
| `scan-volatiles` | `analyze scan-volatiles` |
| `sheet-statistics` | `analyze sheet-statistics` |
| `table-profile` | `analyze table-profile` |
| `check-ref-impact` | `analyze ref-impact` |
| `edit` | `write cells` |
| `range-import` | `write import` |
| `append-region` | `write append` |
| `clone-template-row` | `write clone-template-row` |
| `clone-row-band` | `write clone-row-band` |
| `replace-in-formulas` | `write formulas replace` |
| `transform-batch` | `write batch transform` |
| `style-batch` | `write batch style` |
| `apply-formula-pattern` | `write batch formula-pattern` |
| `structure-batch` | `write batch structure` |
| `column-size-batch` | `write batch column-size` |
| `sheet-layout-batch` | `write batch sheet-layout` |
| `rules-batch` | `write batch rules` |
| `define-name` | `write name define` |
| `update-name` | `write name update` |
| `delete-name` | `write name delete` |
| `create-workbook` | `workbook create` |
| `copy` | `workbook copy` |
| `recalculate` | `workbook recalculate` |
| `verify` | `verify proof` |
| `diff` | `verify diff` |
| `run-manifest` | `sheetport run` |

Already-nested commands remain unchanged:

- `session ...`
- `sheetport manifest ...`
- `sheetport bind-check`
- `sheetport run`

## Alias policy

The nested surface is canonical.

### Canonical / visible
Only the nested surface appears in:

- `asp --help`
- subcommand help
- docs / READMEs
- skills
- examples
- tests (except explicit legacy alias coverage)

### Legacy compatibility
Legacy flat commands remain supported temporarily through a **hidden argv normalizer**.

Examples:

- `asp list-sheets file.xlsx` → `asp read sheets file.xlsx`
- `asp append-region ...` → `asp write append ...`
- `asp diff a.xlsx b.xlsx` → `asp verify diff a.xlsx b.xlsx`
- `asp schema transform-batch` → `asp schema write batch transform`
- `asp example session-op transform.write_matrix` → `asp example session op transform.write_matrix`

### Warning policy
When a legacy alias is used:

- the command still runs
- a warning is emitted to stderr unless `--quiet`
- the warning includes the exact canonical replacement

Example:

```text
warning: 'append-region' is deprecated; use 'write append'
```

### Removal policy
Keep hidden legacy aliases for one follow-up tranche / one minor release, then remove them and fail with explicit migration guidance.

## Schema / example syntax

`schema` and `example` remain global discoverability commands, but their targets become nested.

### Canonical batch payload targets

```bash
asp schema write batch transform
asp schema write batch style
asp schema write batch formula-pattern
asp schema write batch structure
asp schema write batch column-size
asp schema write batch sheet-layout
asp schema write batch rules
```

```bash
asp example write batch transform
asp example write batch style
asp example write batch formula-pattern
asp example write batch structure
asp example write batch column-size
asp example write batch sheet-layout
asp example write batch rules
```

### Canonical session payload targets

```bash
asp schema session op transform.write_matrix
asp example session op transform.write_matrix
```

### Invalid intermediate targets
Intermediate group targets should fail with explicit guidance.

Examples:

```bash
asp schema write
asp schema write batch
```

Should return an error listing valid leaf targets.

### Legacy discoverability aliases
Temporarily supported through normalization:

- `schema transform-batch`
- `example structure-batch`
- `schema session-op <kind>`
- `example session-op <kind>`

## Parser structure

The flat `Commands` enum should be replaced by nested enums along these lines:

- `Commands::Read(ReadCommands)`
- `Commands::Analyze(AnalyzeCommands)`
- `Commands::Write(WriteCommands)`
- `Commands::Workbook(WorkbookCommands)`
- `Commands::Verify(VerifyCommands)`
- `Commands::Schema(DiscoverabilityRootCommands)`
- `Commands::Example(DiscoverabilityRootCommands)`
- `Commands::Session(Box<SessionCommands>)`
- `Commands::Sheetport(Box<SheetportCommands>)`

Sub-groups inside `WriteCommands`:

- `WriteCommands::Cells`
- `WriteCommands::Import`
- `WriteCommands::Append`
- `WriteCommands::CloneTemplateRow`
- `WriteCommands::CloneRowBand`
- `WriteCommands::Formulas(WriteFormulaCommands)`
- `WriteCommands::Name(WriteNameCommands)`
- `WriteCommands::Batch(WriteBatchCommands)`

Sub-groups for discoverability:

- `DiscoverabilityRootCommands::Write(DiscoverabilityWriteCommands)`
- `DiscoverabilityRootCommands::Session(DiscoverabilitySessionCommands)`

## Implementation strategy

### Reuse existing command functions
The business logic functions in `crates/spreadsheet-kit/src/cli/commands/**` stay in place.

This ticket primarily changes:

- clap parser topology
- dispatch routing
- help surface
- schema / example target parsing
- docs and tests

### Legacy argv normalization
Add a normalization layer before `Cli::parse_from(...)` that handles:

- flat command aliases
- legacy `schema` / `example` targets
- deprecated `run-manifest`

This is preferred over keeping duplicate visible parser variants.

## Affected files

### Must change
- `crates/spreadsheet-kit/src/cli/mod.rs`
- `crates/spreadsheet-kit/tests/cli_integration.rs`
- `README.md`
- `npm/agent-spreadsheet/README.md`
- `skills/SAFE_EDITING_SKILL.md`
- `skills/CLI_BATCH_WRITE_SKILL.md`
- `docs/architecture/surface-capability-matrix.md`
- `tickets/44-mcp-sdk-ergonomics-next/4403-cli-vnext-command-grouping-and-workflow-help-surface.md`

### Intentionally runtime-unaffected
- `crates/spreadsheet-mcp/src/server.rs`
- `crates/spreadsheet-kit/src/tools/**`
- `npm/spreadsheet-kit-sdk/src/backend.js`
- `npm/spreadsheet-kit-sdk/src/mcp-backend.js`
- `npm/spreadsheet-kit-sdk/src/wasm-backend.js`
- `npm/spreadsheet-kit-sdk/src/index.js`

## Tests

Required coverage:

- top-level help shows only nested groups
- subgroup help remains discoverable and coherent
- canonical nested commands parse correctly
- `schema` / `example` nested targets parse correctly
- hidden legacy aliases still work and warn
- docs/examples stay synchronized with the nested surface

## Definition of done

4403 is done when:

1. the canonical surface is truly nested
2. `asp --help` no longer shows a flat wall of commands
3. docs / examples / skills reflect only the nested surface
4. `schema` / `example` use nested target syntax
5. legacy flat aliases still function temporarily with warnings
6. MCP and SDK runtime semantics remain unchanged
