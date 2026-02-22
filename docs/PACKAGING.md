# Packaging & Naming Conventions

## Package names

| Surface | Crate name | npm name | Binary name |
| --- | --- | --- | --- |
| Core primitives | `spreadsheet-kit` | — | — |
| WASM adapter/runtime | `spreadsheet-kit-wasm` | `spreadsheet-kit-wasm` (planned distribution package) | — |
| MCP server | `spreadsheet-mcp` | — | `spreadsheet-mcp` |
| CLI | `spreadsheet-kit` | `agent-spreadsheet` | `agent-spreadsheet` |
| JS SDK backend abstraction | — | `spreadsheet-kit-sdk` | — |

The workspace umbrella name is **spreadsheet-kit**. The GitHub repo is `PSU3D0/spreadsheet-mcp` (historical — predates the workspace split).

## Versioning

- `spreadsheet-kit` — follows semver independently (currently 0.1.x)
- `spreadsheet-kit-wasm` — follows semver independently, but tracks shared core contract changes closely
- `spreadsheet-mcp` — follows semver independently (currently 0.10.x)
- `agent-spreadsheet` (npm) — published from `cli-v*` tags; package version is derived from tag suffix and must match available `v*` GitHub release assets
- `spreadsheet-kit-sdk` (npm) — published from `sdk-v*` tags; semver independent from CLI/binary cadence

Release ordering for tranche-35 surfaces:

1. publish core crates in dependency order (`spreadsheet-kit` → `spreadsheet-kit-wasm` → `spreadsheet-mcp`)
2. publish npm packages (`agent-spreadsheet`, `spreadsheet-kit-sdk`, and `spreadsheet-kit-wasm` when enabled)
3. run smoke tests against SDK MCP + WASM backends before final release promotion

### Tag lanes

- `vX.Y.Z` → Rust release lane (GitHub release assets + crates publish)
- `cli-vX.Y.Z` → npm `agent-spreadsheet` publish lane
- `sdk-vX.Y.Z` → npm `spreadsheet-kit-sdk` publish lane

### Compatibility notes (SDK/MCP/WASM)

| SDK line | MCP compatibility | WASM compatibility | Notes |
| --- | --- | --- | --- |
| `0.1.x` | compatible with `spreadsheet-mcp` `0.10.x` when required capabilities are present | compatible with tranche-35 `spreadsheet-kit-wasm` exports | Capability checks are the source of truth for mixed-version safety |

Policy:

- Shared core contracts follow semver discipline.
- Adapter-only additions should be additive and non-breaking.
- SDK callers must branch on `backend.getCapabilities()` before backend-specific flows (`supportsForkLifecycle`, `supportsStaging`, etc.).
- Capability removals/deprecations require explicit migration notes and release callouts.

## Release artifacts

GitHub Releases include native binaries for operator/server surfaces:

| Asset pattern | Binary |
| --- | --- |
| `spreadsheet-mcp-{target}` | MCP server |
| `agent-spreadsheet-{target}` | CLI |

Targets: `linux-x86_64`, `macos-x86_64`, `macos-aarch64`, `windows-x86_64(.exe)`

WASM + SDK artifacts are published as package artifacts:

| Artifact | Surface |
| --- | --- |
| crate `spreadsheet-kit-wasm` | Rust/WASM adapter crate |
| npm `spreadsheet-kit-sdk` | JS SDK backend abstraction |
| npm `spreadsheet-kit-wasm` (planned) | JS/WASM runtime distribution |

## Default features

`spreadsheet-mcp` ships with `recalc-formualizer` as a default feature. This means:
- `cargo install spreadsheet-mcp` and `cargo install spreadsheet-kit --bin agent-spreadsheet` include the Formualizer recalc engine out of the box
- LibreOffice (`recalc-libreoffice`) is only used in the Docker `:full` image
- To build without recalc: `cargo install spreadsheet-mcp --no-default-features`

## Docker images

Published to `ghcr.io/psu3d0/spreadsheet-mcp`:

| Tag | Contents |
| --- | --- |
| `latest` | Slim read-only image (~15 MB), `spreadsheet-mcp` binary only |
| `full` | Full image (~800 MB), includes LibreOffice + recalc macros |

## `agent-spreadsheet` npm install flow

1. `npm install` triggers `postinstall` → `scripts/install.js`
2. Script resolves platform triple (linux-x64, darwin-x64, darwin-arm64, win32-x64)
3. Downloads `agent-spreadsheet-{asset}` from GitHub Releases `v{version}`
4. Places binary in `vendor/` within the package directory
5. `bin/agent-spreadsheet.js` spawns the vendored binary

Override download source with `AGENT_SPREADSHEET_DOWNLOAD_BASE_URL`. Use a pre-built local binary with `AGENT_SPREADSHEET_LOCAL_BINARY`.

## README structure

| File | Audience | Focus |
| --- | --- | --- |
| Root `README.md` | All users | Umbrella: install, quickstarts, tool surface, deployment, config |
| `crates/spreadsheet-kit/README.md` | Crate consumers | Scope, types, what's excluded |
| `crates/spreadsheet-mcp/README.md` | MCP users | Quickstart configs, feature summary, link to root |
| `crates/spreadsheet-kit/README.md` | CLI users | `agent-spreadsheet` binary usage and command surface |
| `npm/agent-spreadsheet/README.md` | npm CLI users | Install, platform matrix, troubleshooting, env vars |
| `npm/spreadsheet-kit-sdk/README.md` | npm SDK users | Backend abstraction, capabilities, typed errors |
