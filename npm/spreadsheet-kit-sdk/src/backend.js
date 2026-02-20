const { CapabilityError, SpreadsheetSdkError } = require("./errors")

/**
 * @typedef {import("./capabilities").BackendCapabilities} BackendCapabilities
 */

/**
 * @typedef {object} SpreadsheetBackend
 * @property {"mcp"|"wasm"} kind
 * @property {() => Readonly<BackendCapabilities>} getCapabilities
 * @property {(input: Record<string, unknown>) => Promise<string[]>} listSheets
 * @property {(input: Record<string, unknown>) => Promise<{ sheetName: string, values: unknown[] }>} rangeValues
 */

/**
 * @param {{ kind: string, getCapabilities: () => Record<string, boolean> }} backend
 * @param {string} capability
 * @param {string} method
 */
function requireCapability(backend, capability, method) {
  const capabilities = backend.getCapabilities()
  if (!capabilities[capability]) {
    throw new CapabilityError({
      backend: backend.kind,
      capability,
      method
    })
  }
}

/**
 * @param {unknown} value
 * @param {string} name
 */
function requiredString(value, name) {
  if (typeof value !== "string" || value.length === 0) {
    throw new SpreadsheetSdkError(`missing required field '${name}'`, {
      code: "INVALID_ARGUMENT"
    })
  }
  return value
}

module.exports = {
  requireCapability,
  requiredString
}
