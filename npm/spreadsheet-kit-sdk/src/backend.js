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

/**
 * @param {unknown} result
 * @param {string} fallbackSheetName
 */
function normalizeSheetPageResult(result, fallbackSheetName) {
  if (!result || typeof result !== "object") {
    throw new SpreadsheetSdkError("invalid sheetPage response", {
      code: "INVALID_RESPONSE"
    })
  }

  return {
    workbookId: typeof result.workbookId === "string"
      ? result.workbookId
      : typeof result.workbook_id === "string"
        ? result.workbook_id
        : undefined,
    sheetName: typeof result.sheetName === "string"
      ? result.sheetName
      : typeof result.sheet_name === "string"
        ? result.sheet_name
        : fallbackSheetName,
    rows: Array.isArray(result.rows) ? result.rows : [],
    nextStartRow: typeof result.nextStartRow === "number"
      ? result.nextStartRow
      : typeof result.next_start_row === "number"
        ? result.next_start_row
        : undefined,
    headerRow: result.headerRow ?? result.header_row,
    compact: result.compact,
    valuesOnly: result.valuesOnly ?? result.values_only,
    format: result.format
  }
}

/**
 * @param {unknown} result
 */
function normalizeGridExportResult(result) {
  if (!result || typeof result !== "object") {
    throw new SpreadsheetSdkError("invalid gridExport response", {
      code: "INVALID_RESPONSE"
    })
  }
  return {
    sheet: result.sheet,
    anchor: result.anchor,
    columns: result.columns,
    merges: result.merges,
    rows: result.rows
  }
}

/**
 * @param {unknown} result
 */
function normalizeTransformBatchResult(result) {
  if (!result || typeof result !== "object") {
    throw new SpreadsheetSdkError("invalid transformBatch response", {
      code: "INVALID_RESPONSE"
    })
  }
  return {
    opsApplied: result.opsApplied ?? result.ops_applied,
    cellsTouched: result.cellsTouched ?? result.cells_touched,
    cellsValueSet: result.cellsValueSet ?? result.cells_value_set,
    cellsFormulaSet: result.cellsFormulaSet ?? result.cells_formula_set,
    cellsFormulaCleared: result.cellsFormulaCleared ?? result.cells_formula_cleared,
    cellsSkippedKeepFormulas: result.cellsSkippedKeepFormulas ?? result.cells_skipped_keep_formulas,
    formulaParseDiagnostics: result.formulaParseDiagnostics ?? result.formula_parse_diagnostics
  }
}

module.exports = {
  requireCapability,
  requiredString,
  normalizeSheetPageResult,
  normalizeGridExportResult,
  normalizeTransformBatchResult
}
