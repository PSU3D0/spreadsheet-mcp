use serde::{Deserialize, Serialize};
use spreadsheet_kit::core::session::{
    SessionApplySummary, SessionRangeSelection, SessionSheetPageParams, SessionTransformOp,
    WorkbookSession,
};
use spreadsheet_kit::model::{GridPayload, RangeValuesEntry, SheetPageFormat, SheetPageResponse};
use std::collections::HashMap;
use std::sync::{Arc, Mutex, MutexGuard};

#[derive(Debug, thiserror::Error, Clone, PartialEq, Eq)]
pub enum SessionApiError {
    #[error("session '{session_id}' not found")]
    SessionNotFound { session_id: String },
    #[error("invalid argument: {message}")]
    InvalidArgument { message: String },
    #[error("unsupported in wasm mvp: {message}")]
    Unsupported { message: String },
    #[error("internal error: {message}")]
    Internal { message: String },
}

impl SessionApiError {
    pub fn code(&self) -> &'static str {
        match self {
            SessionApiError::SessionNotFound { .. } => "SESSION_NOT_FOUND",
            SessionApiError::InvalidArgument { .. } => "INVALID_ARGUMENT",
            SessionApiError::Unsupported { .. } => "UNSUPPORTED",
            SessionApiError::Internal { .. } => "INTERNAL",
        }
    }

    fn internal(message: impl Into<String>) -> Self {
        Self::Internal {
            message: message.into(),
        }
    }
}

pub type SessionResult<T> = Result<T, SessionApiError>;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SessionApiErrorPayload {
    pub code: String,
    pub message: String,
}

impl From<SessionApiError> for SessionApiErrorPayload {
    fn from(value: SessionApiError) -> Self {
        Self {
            code: value.code().to_string(),
            message: value.to_string(),
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(untagged)]
pub enum RangeSelectionInput {
    Single(String),
    Multi(Vec<String>),
}

impl From<RangeSelectionInput> for SessionRangeSelection {
    fn from(value: RangeSelectionInput) -> Self {
        match value {
            RangeSelectionInput::Single(range) => SessionRangeSelection::Single(range),
            RangeSelectionInput::Multi(ranges) => SessionRangeSelection::Multi(ranges),
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct RangeValuesParams {
    #[serde(alias = "sheet_name")]
    pub sheet_name: String,
    pub ranges: RangeSelectionInput,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct RangeValuesResult {
    pub sheet_name: String,
    pub values: Vec<RangeValuesEntry>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct GridExportParams {
    #[serde(alias = "sheet_name")]
    pub sheet_name: String,
    pub range: String,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct TransformBatchOptions {
    #[serde(default)]
    pub dry_run: bool,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SheetPageParams {
    #[serde(alias = "sheet_name")]
    pub sheet_name: String,
    #[serde(default)]
    pub start_row: Option<u32>,
    #[serde(default)]
    pub page_size: Option<u32>,
    #[serde(default)]
    pub columns: Option<Vec<String>>,
    #[serde(default)]
    pub columns_by_header: Option<Vec<String>>,
    #[serde(default)]
    pub include_formulas: Option<bool>,
    #[serde(default)]
    pub include_styles: Option<bool>,
    #[serde(default)]
    pub include_header: Option<bool>,
    #[serde(default)]
    pub format: Option<SheetPageFormat>,
}

impl From<SheetPageParams> for SessionSheetPageParams {
    fn from(value: SheetPageParams) -> Self {
        SessionSheetPageParams {
            sheet_name: value.sheet_name,
            start_row: value.start_row.unwrap_or(1),
            page_size: value.page_size.unwrap_or(50),
            columns: value.columns,
            columns_by_header: value.columns_by_header,
            include_formulas: value.include_formulas.unwrap_or(true),
            include_styles: value.include_styles.unwrap_or(false),
            include_header: value.include_header.unwrap_or(true),
            format: value.format.unwrap_or_default(),
        }
    }
}

#[derive(Default)]
struct SessionStore {
    next_id: u64,
    sessions: HashMap<String, WorkbookSession>,
}

#[derive(Clone, Default)]
pub struct SessionApi {
    store: Arc<Mutex<SessionStore>>,
}

impl SessionApi {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn create_session(&self, workbook_bytes: &[u8]) -> SessionResult<String> {
        let session = WorkbookSession::from_bytes(workbook_bytes).map_err(|err| {
            SessionApiError::InvalidArgument {
                message: err.to_string(),
            }
        })?;

        let mut store = self.lock_store()?;
        store.next_id += 1;
        let session_id = format!("session-{:016x}", store.next_id);
        store.sessions.insert(session_id.clone(), session);
        Ok(session_id)
    }

    pub fn list_sheets(&self, session_id: &str) -> SessionResult<Vec<String>> {
        let store = self.lock_store()?;
        let session =
            store
                .sessions
                .get(session_id)
                .ok_or_else(|| SessionApiError::SessionNotFound {
                    session_id: session_id.to_string(),
                })?;

        Ok(session.list_sheets())
    }

    pub fn range_values(
        &self,
        session_id: &str,
        params: RangeValuesParams,
    ) -> SessionResult<RangeValuesResult> {
        let store = self.lock_store()?;
        let session =
            store
                .sessions
                .get(session_id)
                .ok_or_else(|| SessionApiError::SessionNotFound {
                    session_id: session_id.to_string(),
                })?;

        let values = session
            .range_values(
                &params.sheet_name,
                SessionRangeSelection::from(params.ranges),
            )
            .map_err(|err| SessionApiError::InvalidArgument {
                message: err.to_string(),
            })?;

        Ok(RangeValuesResult {
            sheet_name: params.sheet_name,
            values,
        })
    }

    pub fn sheet_page(
        &self,
        session_id: &str,
        params: SheetPageParams,
    ) -> SessionResult<SheetPageResponse> {
        let store = self.lock_store()?;
        let session =
            store
                .sessions
                .get(session_id)
                .ok_or_else(|| SessionApiError::SessionNotFound {
                    session_id: session_id.to_string(),
                })?;

        let mut response =
            session
                .sheet_page(params.into())
                .map_err(|err| SessionApiError::InvalidArgument {
                    message: err.to_string(),
                })?;
        response.workbook_id = spreadsheet_kit::model::WorkbookId(session_id.to_string());
        Ok(response)
    }

    pub fn grid_export(
        &self,
        session_id: &str,
        params: GridExportParams,
    ) -> SessionResult<GridPayload> {
        let store = self.lock_store()?;
        let session =
            store
                .sessions
                .get(session_id)
                .ok_or_else(|| SessionApiError::SessionNotFound {
                    session_id: session_id.to_string(),
                })?;

        session
            .grid_export(&params.sheet_name, &params.range)
            .map_err(|err| SessionApiError::InvalidArgument {
                message: err.to_string(),
            })
    }

    pub fn transform_batch(
        &self,
        session_id: &str,
        ops: Vec<SessionTransformOp>,
        options: TransformBatchOptions,
    ) -> SessionResult<SessionApplySummary> {
        if ops.is_empty() {
            return Err(SessionApiError::InvalidArgument {
                message: "at least one transform op is required".to_string(),
            });
        }

        let mut store = self.lock_store()?;

        if options.dry_run {
            let bytes = store
                .sessions
                .get(session_id)
                .ok_or_else(|| SessionApiError::SessionNotFound {
                    session_id: session_id.to_string(),
                })?
                .to_bytes()
                .map_err(|err| SessionApiError::Internal {
                    message: err.to_string(),
                })?;

            let mut preview =
                WorkbookSession::from_bytes(bytes).map_err(|err| SessionApiError::Internal {
                    message: err.to_string(),
                })?;

            return preview
                .apply_ops(&ops)
                .map_err(|err| SessionApiError::InvalidArgument {
                    message: err.to_string(),
                });
        }

        let session =
            store
                .sessions
                .get_mut(session_id)
                .ok_or_else(|| SessionApiError::SessionNotFound {
                    session_id: session_id.to_string(),
                })?;

        session
            .apply_ops(&ops)
            .map_err(|err| SessionApiError::InvalidArgument {
                message: err.to_string(),
            })
    }

    pub fn export_workbook(&self, session_id: &str) -> SessionResult<Vec<u8>> {
        let store = self.lock_store()?;
        let session =
            store
                .sessions
                .get(session_id)
                .ok_or_else(|| SessionApiError::SessionNotFound {
                    session_id: session_id.to_string(),
                })?;

        session.to_bytes().map_err(|err| SessionApiError::Internal {
            message: err.to_string(),
        })
    }

    pub fn dispose_session(&self, session_id: &str) -> SessionResult<bool> {
        let mut store = self.lock_store()?;
        Ok(store.sessions.remove(session_id).is_some())
    }

    fn lock_store(&self) -> SessionResult<MutexGuard<'_, SessionStore>> {
        self.store
            .lock()
            .map_err(|_| SessionApiError::internal("session store lock poisoned"))
    }
}

#[cfg(target_arch = "wasm32")]
mod wasm_bindings {
    use super::*;
    use wasm_bindgen::prelude::*;

    fn api() -> &'static SessionApi {
        static API: std::sync::OnceLock<SessionApi> = std::sync::OnceLock::new();
        API.get_or_init(SessionApi::new)
    }

    fn to_js_error(err: SessionApiError) -> JsValue {
        let payload = SessionApiErrorPayload::from(err);
        serde_wasm_bindgen::to_value(&payload)
            .unwrap_or_else(|_| JsValue::from_str(&payload.message))
    }

    fn to_js_value<T: Serialize>(value: &T) -> Result<JsValue, JsValue> {
        serde_wasm_bindgen::to_value(value).map_err(|err| {
            to_js_error(SessionApiError::Internal {
                message: format!("failed to serialize response: {err}"),
            })
        })
    }

    fn from_js_value<T: for<'de> Deserialize<'de>>(value: JsValue) -> SessionResult<T> {
        serde_wasm_bindgen::from_value(value).map_err(|err| SessionApiError::InvalidArgument {
            message: format!("invalid params: {err}"),
        })
    }

    #[wasm_bindgen(js_name = createSession)]
    pub fn create_session_js(workbook_bytes: Vec<u8>) -> Result<String, JsValue> {
        api().create_session(&workbook_bytes).map_err(to_js_error)
    }

    #[wasm_bindgen(js_name = listSheets)]
    pub fn list_sheets_js(session_id: String) -> Result<JsValue, JsValue> {
        let sheets = api().list_sheets(&session_id).map_err(to_js_error)?;
        to_js_value(&sheets)
    }

    #[wasm_bindgen(js_name = rangeValues)]
    pub fn range_values_js(session_id: String, params: JsValue) -> Result<JsValue, JsValue> {
        let params: RangeValuesParams = from_js_value(params).map_err(to_js_error)?;
        let result = api()
            .range_values(&session_id, params)
            .map_err(to_js_error)?;
        to_js_value(&result)
    }

    #[wasm_bindgen(js_name = sheetPage)]
    pub fn sheet_page_js(session_id: String, params: JsValue) -> Result<JsValue, JsValue> {
        let params: SheetPageParams = from_js_value(params).map_err(to_js_error)?;
        let result = api().sheet_page(&session_id, params).map_err(to_js_error)?;
        to_js_value(&result)
    }

    #[wasm_bindgen(js_name = gridExport)]
    pub fn grid_export_js(session_id: String, params: JsValue) -> Result<JsValue, JsValue> {
        let params: GridExportParams = from_js_value(params).map_err(to_js_error)?;
        let payload = api()
            .grid_export(&session_id, params)
            .map_err(to_js_error)?;
        to_js_value(&payload)
    }

    #[wasm_bindgen(js_name = transformBatch)]
    pub fn transform_batch_js(
        session_id: String,
        ops: JsValue,
        options: Option<JsValue>,
    ) -> Result<JsValue, JsValue> {
        let ops: Vec<SessionTransformOp> = from_js_value(ops).map_err(to_js_error)?;
        let options = match options {
            Some(value) => from_js_value(value).map_err(to_js_error)?,
            None => TransformBatchOptions::default(),
        };

        let summary = api()
            .transform_batch(&session_id, ops, options)
            .map_err(to_js_error)?;
        to_js_value(&summary)
    }

    #[wasm_bindgen(js_name = exportWorkbook)]
    pub fn export_workbook_js(session_id: String) -> Result<Vec<u8>, JsValue> {
        api().export_workbook(&session_id).map_err(to_js_error)
    }

    #[wasm_bindgen(js_name = disposeSession)]
    pub fn dispose_session_js(session_id: String) -> Result<bool, JsValue> {
        api().dispose_session(&session_id).map_err(to_js_error)
    }
}
