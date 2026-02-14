# Packaging & Naming Conventions

## Package names

| Surface | Crate name | npm name | Binary name |
| --- | --- | --- | --- |
| Core primitives | `spreadsheet-kit` | — | — |
| MCP server | `spreadsheet-mcp` | — | `spreadsheet-mcp` |
| CLI | `spreadsheet-kit` | `agent-spreadsheet` | `agent-spreadsheet` |

The workspace umbrella name is **spreadsheet-kit**. The GitHub repo is `PSU3D0/spreadsheet-mcp` (historical — predates the workspace split).

## Versioning

- `spreadsheet-kit` — follows semver independently (currently 0.1.x)
- `spreadsheet-mcp` — follows semver independently (currently 0.9.x)
- `agent-spreadsheet` (npm) — version is kept in sync with the release tag

The release workflow publishes crates in dependency order (`spreadsheet-kit` → `spreadsheet-mcp`) and then publishes the npm package.

## Release artifacts

GitHub Releases include binaries for both surfaces:

| Asset pattern | Binary |
| --- | --- |
| `spreadsheet-mcp-{target}` | MCP server |
| `agent-spreadsheet-{target}` | CLI |

Targets: `linux-x86_64`, `macos-x86_64`, `macos-aarch64`, `windows-x86_64(.exe)`

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

## npm install flow

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
| `npm/agent-spreadsheet/README.md` | npm users | Install, platform matrix, troubleshooting, env vars |
