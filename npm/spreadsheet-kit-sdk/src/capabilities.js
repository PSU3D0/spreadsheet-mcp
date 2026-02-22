/**
 * @typedef {object} BackendCapabilities
 * @property {boolean} supportsListSheets
 * @property {boolean} supportsRangeValues
 * @property {boolean} supportsSheetPage
 * @property {boolean} supportsGridExport
 * @property {boolean} supportsTransformBatch
 * @property {boolean} supportsForkLifecycle
 * @property {boolean} supportsStaging
 * @property {boolean} supportsSessionLifecycle
 * @property {boolean} supportsExportWorkbook
 */

/** @type {Readonly<BackendCapabilities>} */
const MCP_CAPABILITIES = Object.freeze({
  supportsListSheets: true,
  supportsRangeValues: true,
  supportsSheetPage: true,
  supportsGridExport: true,
  supportsTransformBatch: true,
  supportsForkLifecycle: true,
  supportsStaging: true,
  supportsSessionLifecycle: false,
  supportsExportWorkbook: false
})

/** @type {Readonly<BackendCapabilities>} */
const WASM_CAPABILITIES = Object.freeze({
  supportsListSheets: true,
  supportsRangeValues: true,
  supportsSheetPage: true,
  supportsGridExport: true,
  supportsTransformBatch: true,
  supportsForkLifecycle: false,
  supportsStaging: false,
  supportsSessionLifecycle: true,
  supportsExportWorkbook: true
})

/**
 * @param {Readonly<BackendCapabilities>} capabilities
 * @returns {Readonly<BackendCapabilities>}
 */
function freezeCapabilities(capabilities) {
  return Object.freeze({ ...capabilities })
}

module.exports = {
  MCP_CAPABILITIES,
  WASM_CAPABILITIES,
  freezeCapabilities
}
