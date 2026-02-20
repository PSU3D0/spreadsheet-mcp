use serde::{Deserialize, Serialize};

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

#[cfg(not(target_arch = "wasm32"))]
mod host {
    use super::*;
    use spreadsheet_kit::core::session::{
        SessionApplySummary, SessionRangeSelection, SessionTransformOp, WorkbookSession,
    };
    use spreadsheet_kit::model::{GridPayload, RangeValuesEntry};
    use std::collections::HashMap;
    use std::sync::{Arc, Mutex, MutexGuard};

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
            _params: SheetPageParams,
        ) -> SessionResult<serde_json::Value> {
            let store = self.lock_store()?;
            let _session =
                store
                    .sessions
                    .get(session_id)
                    .ok_or_else(|| SessionApiError::SessionNotFound {
                        session_id: session_id.to_string(),
                    })?;

            Err(SessionApiError::Unsupported {
                message: "sheetPage is deferred in tranche-35 wasm MVP; use rangeValues/gridExport"
                    .to_string(),
            })
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

                let mut preview = WorkbookSession::from_bytes(bytes).map_err(|err| {
                    SessionApiError::Internal {
                        message: err.to_string(),
                    }
                })?;

                return preview
                    .apply_ops(&ops)
                    .map_err(|err| SessionApiError::InvalidArgument {
                        message: err.to_string(),
                    });
            }

            let session = store.sessions.get_mut(session_id).ok_or_else(|| {
                SessionApiError::SessionNotFound {
                    session_id: session_id.to_string(),
                }
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
}

#[cfg(not(target_arch = "wasm32"))]
pub use host::*;

#[cfg(target_arch = "wasm32")]
mod wasm_bindings {
    use super::*;
    use std::collections::HashMap;
    use std::sync::{Mutex, OnceLock};
    use wasm_bindgen::prelude::*;

    #[derive(Default)]
    struct ByteSessionStore {
        next_id: u64,
        sessions: HashMap<String, Vec<u8>>,
    }

    fn store() -> &'static Mutex<ByteSessionStore> {
        static STORE: OnceLock<Mutex<ByteSessionStore>> = OnceLock::new();
        STORE.get_or_init(|| Mutex::new(ByteSessionStore::default()))
    }

    fn to_js_error(err: SessionApiError) -> JsValue {
        let payload = SessionApiErrorPayload::from(err);
        serde_wasm_bindgen::to_value(&payload)
            .unwrap_or_else(|_| JsValue::from_str(&payload.message))
    }

    fn from_js_value<T: for<'de> Deserialize<'de>>(value: JsValue) -> SessionResult<T> {
        serde_wasm_bindgen::from_value(value).map_err(|err| SessionApiError::InvalidArgument {
            message: format!("invalid params: {err}"),
        })
    }

    fn unsupported() -> JsValue {
        to_js_error(SessionApiError::Unsupported {
            message:
                "native workbook engine is host-only in this build; use MCP backend or host runtime"
                    .to_string(),
        })
    }

    #[wasm_bindgen(js_name = createSession)]
    pub fn create_session_js(workbook_bytes: Vec<u8>) -> Result<String, JsValue> {
        let mut guard = store()
            .lock()
            .map_err(|_| to_js_error(SessionApiError::internal("session store lock poisoned")))?;
        guard.next_id += 1;
        let session_id = format!("session-{:016x}", guard.next_id);
        guard.sessions.insert(session_id.clone(), workbook_bytes);
        Ok(session_id)
    }

    #[wasm_bindgen(js_name = listSheets)]
    pub fn list_sheets_js(_session_id: String) -> Result<JsValue, JsValue> {
        Err(unsupported())
    }

    #[wasm_bindgen(js_name = rangeValues)]
    pub fn range_values_js(_session_id: String, params: JsValue) -> Result<JsValue, JsValue> {
        let _: serde_json::Value = from_js_value(params).map_err(to_js_error)?;
        Err(unsupported())
    }

    #[wasm_bindgen(js_name = sheetPage)]
    pub fn sheet_page_js(_session_id: String, params: JsValue) -> Result<JsValue, JsValue> {
        let _: serde_json::Value = from_js_value(params).map_err(to_js_error)?;
        Err(unsupported())
    }

    #[wasm_bindgen(js_name = gridExport)]
    pub fn grid_export_js(_session_id: String, params: JsValue) -> Result<JsValue, JsValue> {
        let _: serde_json::Value = from_js_value(params).map_err(to_js_error)?;
        Err(unsupported())
    }

    #[wasm_bindgen(js_name = transformBatch)]
    pub fn transform_batch_js(
        _session_id: String,
        ops: JsValue,
        options: Option<JsValue>,
    ) -> Result<JsValue, JsValue> {
        let _: serde_json::Value = from_js_value(ops).map_err(to_js_error)?;
        if let Some(options) = options {
            let _: serde_json::Value = from_js_value(options).map_err(to_js_error)?;
        }
        Err(unsupported())
    }

    #[wasm_bindgen(js_name = exportWorkbook)]
    pub fn export_workbook_js(session_id: String) -> Result<Vec<u8>, JsValue> {
        let guard = store()
            .lock()
            .map_err(|_| to_js_error(SessionApiError::internal("session store lock poisoned")))?;
        guard
            .sessions
            .get(&session_id)
            .cloned()
            .ok_or_else(|| to_js_error(SessionApiError::SessionNotFound { session_id }))
    }

    #[wasm_bindgen(js_name = disposeSession)]
    pub fn dispose_session_js(session_id: String) -> Result<bool, JsValue> {
        let mut guard = store()
            .lock()
            .map_err(|_| to_js_error(SessionApiError::internal("session store lock poisoned")))?;
        Ok(guard.sessions.remove(&session_id).is_some())
    }
}
