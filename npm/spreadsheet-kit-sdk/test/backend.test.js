const test = require("node:test")
const assert = require("node:assert/strict")

const {
  McpBackend,
  WasmBackend,
  CapabilityError,
  BackendOperationError
} = require("../src")
const { MCP_CAPABILITIES } = require("../src/capabilities")

async function sharedReadFlow(backend) {
  const ctx = { workbookId: "wb-1", sessionId: "session-1" }
  const sheets = await backend.listSheets(ctx)
  const range = await backend.rangeValues({
    ...ctx,
    sheetName: sheets[0],
    ranges: "A1:B2"
  })
  return { sheets, range }
}

test("switching backends keeps shared read callsites stable", async () => {
  const mcp = new McpBackend({
    transport: {
      async invoke(operation, params) {
        if (operation === "list_sheets") {
          assert.equal(params.workbook_id, "wb-1")
          return { sheets: [{ name: "Sheet1" }] }
        }
        if (operation === "range_values") {
          assert.equal(params.workbook_id, "wb-1")
          assert.equal(params.sheet_name, "Sheet1")
          return {
            sheet_name: "Sheet1",
            values: [{ range: "A1:B2", rows: [["v1", "v2"]] }]
          }
        }
        throw new Error(`unexpected op ${operation}`)
      }
    }
  })

  const wasm = new WasmBackend({
    bindings: {
      async listSheets(sessionId) {
        assert.equal(sessionId, "session-1")
        return ["Sheet1"]
      },
      async rangeValues(sessionId, params) {
        assert.equal(sessionId, "session-1")
        assert.equal(params.sheetName, "Sheet1")
        return {
          sheetName: "Sheet1",
          values: [{ range: "A1:B2", rows: [["v1", "v2"]] }]
        }
      }
    }
  })

  const mcpResult = await sharedReadFlow(mcp)
  const wasmResult = await sharedReadFlow(wasm)
  assert.deepEqual(wasmResult, mcpResult)
})

test("capability misuse returns typed capability errors", async () => {
  const wasm = new WasmBackend({ bindings: {} })

  await assert.rejects(
    () => wasm.createFork({ sessionId: "session-1" }),
    (error) => {
      assert.ok(error instanceof CapabilityError)
      assert.equal(error.code, "UNSUPPORTED_CAPABILITY")
      assert.equal(error.backend, "wasm")
      assert.equal(error.capability, "supportsForkLifecycle")
      return true
    }
  )
})

test("backend failures are normalized", async () => {
  const mcp = new McpBackend({
    transport: {
      async invoke() {
        throw new Error("boom")
      }
    }
  })

  await assert.rejects(
    () => mcp.listSheets({ workbookId: "wb-1" }),
    (error) => {
      assert.ok(error instanceof BackendOperationError)
      assert.equal(error.backend, "mcp")
      assert.equal(error.operation, "list_sheets")
      return true
    }
  )
})

test("mcp createFork normalizes workbook id field", async () => {
  const mcp = new McpBackend({
    transport: {
      async invoke(operation, params) {
        assert.equal(operation, "create_fork")
        assert.equal(params.workbook_or_fork_id, "wb-1")
        return { fork_id: "fork-1" }
      }
    }
  })

  const result = await mcp.createFork({ workbookId: "wb-1" })
  assert.equal(result.fork_id, "fork-1")
})

test("backend-specific no-op methods throw explicit unsupported errors", async () => {
  const mcp = new McpBackend({
    transport: {
      async invoke() {
        throw new Error("not used")
      }
    },
    capabilities: {
      ...MCP_CAPABILITIES,
      supportsSessionLifecycle: true,
      supportsExportWorkbook: true
    }
  })

  await assert.rejects(
    () => mcp.createSession(),
    (error) => {
      assert.equal(error.code, "UNSUPPORTED")
      assert.equal(error.backend, "mcp")
      return true
    }
  )
})
