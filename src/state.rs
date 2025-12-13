use crate::config::ServerConfig;
#[cfg(feature = "recalc")]
use crate::fork::{ForkConfig, ForkRegistry};
use crate::model::{WorkbookId, WorkbookListResponse};
#[cfg(feature = "recalc")]
use crate::recalc::{
    GlobalRecalcLock, GlobalScreenshotLock, LibreOfficeBackend, RecalcBackend, RecalcConfig,
    create_executor,
};
use crate::tools::filters::WorkbookFilter;
use crate::utils::{hash_path_metadata, make_short_workbook_id};
use crate::workbook::{WorkbookContext, build_workbook_list};
use anyhow::{Result, anyhow};
use lru::LruCache;
use parking_lot::RwLock;
use std::collections::HashMap;
use std::fs;
use std::num::NonZeroUsize;
use std::path::{Path, PathBuf};
use std::sync::Arc;
use tokio::task;
use walkdir::WalkDir;

pub struct AppState {
    config: Arc<ServerConfig>,
    cache: RwLock<LruCache<WorkbookId, Arc<WorkbookContext>>>,
    index: RwLock<HashMap<WorkbookId, PathBuf>>,
    alias_index: RwLock<HashMap<String, WorkbookId>>,
    #[cfg(feature = "recalc")]
    fork_registry: Option<Arc<ForkRegistry>>,
    #[cfg(feature = "recalc")]
    recalc_backend: Option<Arc<dyn RecalcBackend>>,
    #[cfg(feature = "recalc")]
    recalc_semaphore: Option<GlobalRecalcLock>,
    #[cfg(feature = "recalc")]
    screenshot_semaphore: Option<GlobalScreenshotLock>,
}

impl AppState {
    pub fn new(config: Arc<ServerConfig>) -> Self {
        let capacity = NonZeroUsize::new(config.cache_capacity.max(1)).unwrap();

        #[cfg(feature = "recalc")]
        let (fork_registry, recalc_backend, recalc_semaphore, screenshot_semaphore) =
            if config.recalc_enabled {
                let fork_config = ForkConfig::default();
                let registry = ForkRegistry::new(fork_config)
                    .map(Arc::new)
                    .map_err(|e| tracing::warn!("failed to init fork registry: {}", e))
                    .ok();

                if let Some(registry) = &registry {
                    registry.clone().start_cleanup_task();
                }

                let executor = create_executor(&RecalcConfig::default());
                let backend: Arc<dyn RecalcBackend> = Arc::new(LibreOfficeBackend::new(executor));
                let backend = if backend.is_available() {
                    Some(backend)
                } else {
                    tracing::warn!("recalc backend not available (soffice not found)");
                    None
                };

                let semaphore = GlobalRecalcLock::new(config.max_concurrent_recalcs);
                let screenshot_semaphore = GlobalScreenshotLock::new();

                (
                    registry,
                    backend,
                    Some(semaphore),
                    Some(screenshot_semaphore),
                )
            } else {
                (None, None, None, None)
            };

        Self {
            config,
            cache: RwLock::new(LruCache::new(capacity)),
            index: RwLock::new(HashMap::new()),
            alias_index: RwLock::new(HashMap::new()),
            #[cfg(feature = "recalc")]
            fork_registry,
            #[cfg(feature = "recalc")]
            recalc_backend,
            #[cfg(feature = "recalc")]
            recalc_semaphore,
            #[cfg(feature = "recalc")]
            screenshot_semaphore,
        }
    }

    pub fn config(&self) -> Arc<ServerConfig> {
        self.config.clone()
    }

    #[cfg(feature = "recalc")]
    pub fn fork_registry(&self) -> Option<&Arc<ForkRegistry>> {
        self.fork_registry.as_ref()
    }

    #[cfg(feature = "recalc")]
    pub fn recalc_backend(&self) -> Option<&Arc<dyn RecalcBackend>> {
        self.recalc_backend.as_ref()
    }

    #[cfg(feature = "recalc")]
    pub fn recalc_semaphore(&self) -> Option<&GlobalRecalcLock> {
        self.recalc_semaphore.as_ref()
    }

    #[cfg(feature = "recalc")]
    pub fn screenshot_semaphore(&self) -> Option<&GlobalScreenshotLock> {
        self.screenshot_semaphore.as_ref()
    }

    pub fn list_workbooks(&self, filter: WorkbookFilter) -> Result<WorkbookListResponse> {
        let response = build_workbook_list(&self.config, &filter)?;
        {
            let mut index = self.index.write();
            let mut aliases = self.alias_index.write();
            for descriptor in &response.workbooks {
                let abs_path = self.config.resolve_path(PathBuf::from(&descriptor.path));
                index.insert(descriptor.workbook_id.clone(), abs_path);
                aliases.insert(
                    descriptor.short_id.to_ascii_lowercase(),
                    descriptor.workbook_id.clone(),
                );
            }
        }
        Ok(response)
    }

    pub async fn open_workbook(&self, workbook_id: &WorkbookId) -> Result<Arc<WorkbookContext>> {
        let canonical = self.canonicalize_workbook_id(workbook_id)?;
        {
            let mut cache = self.cache.write();
            if let Some(entry) = cache.get(&canonical) {
                return Ok(entry.clone());
            }
        }

        let path = self.resolve_workbook_path(&canonical)?;
        let config = self.config.clone();
        let path_buf = path.clone();
        let workbook_id_clone = canonical.clone();
        let workbook =
            task::spawn_blocking(move || WorkbookContext::load(&config, &path_buf)).await??;
        let workbook = Arc::new(workbook);

        {
            let mut aliases = self.alias_index.write();
            aliases.insert(
                workbook.short_id.to_ascii_lowercase(),
                workbook_id_clone.clone(),
            );
        }

        let mut cache = self.cache.write();
        cache.put(workbook_id_clone, workbook.clone());
        Ok(workbook)
    }

    pub fn close_workbook(&self, workbook_id: &WorkbookId) -> Result<()> {
        let canonical = self.canonicalize_workbook_id(workbook_id)?;
        let mut cache = self.cache.write();
        cache.pop(&canonical);
        Ok(())
    }

    pub fn evict_by_path(&self, path: &Path) {
        let index = self.index.read();
        let workbook_id = index
            .iter()
            .find(|(_, p)| *p == path)
            .map(|(id, _)| id.clone());
        drop(index);

        if let Some(id) = workbook_id {
            let mut cache = self.cache.write();
            cache.pop(&id);
        }
    }

    fn resolve_workbook_path(&self, workbook_id: &WorkbookId) -> Result<PathBuf> {
        #[cfg(feature = "recalc")]
        if let Some(registry) = &self.fork_registry
            && let Some(fork_path) = registry.get_fork_path(workbook_id.as_str())
        {
            return Ok(fork_path);
        }

        if let Some(path) = self.index.read().get(workbook_id).cloned() {
            return Ok(path);
        }

        let located = self.scan_for_workbook(workbook_id.as_str())?;
        self.register_location(&located);
        Ok(located.path)
    }

    fn canonicalize_workbook_id(&self, workbook_id: &WorkbookId) -> Result<WorkbookId> {
        #[cfg(feature = "recalc")]
        if let Some(registry) = &self.fork_registry
            && registry.get_fork_path(workbook_id.as_str()).is_some()
        {
            return Ok(workbook_id.clone());
        }

        if self.index.read().contains_key(workbook_id) {
            return Ok(workbook_id.clone());
        }
        let aliases = self.alias_index.read();
        if let Some(mapped) = aliases.get(workbook_id.as_str()).cloned() {
            return Ok(mapped);
        }
        let lowered = workbook_id.as_str().to_ascii_lowercase();
        if lowered != workbook_id.as_str()
            && let Some(mapped) = aliases.get(&lowered).cloned()
        {
            return Ok(mapped);
        }

        let located = self.scan_for_workbook(workbook_id.as_str())?;
        let canonical = located.workbook_id.clone();
        self.register_location(&located);
        Ok(canonical)
    }

    fn scan_for_workbook(&self, candidate: &str) -> Result<LocatedWorkbook> {
        let candidate_lower = candidate.to_ascii_lowercase();

        if let Some(single) = self.config.single_workbook() {
            let metadata = fs::metadata(single)?;
            let canonical = WorkbookId(hash_path_metadata(single, &metadata));
            let slug = single
                .file_stem()
                .map(|s| s.to_string_lossy().to_string())
                .unwrap_or_else(|| "workbook".to_string());
            let short_id = make_short_workbook_id(&slug, canonical.as_str());
            if candidate == canonical.as_str() || candidate_lower == short_id {
                return Ok(LocatedWorkbook {
                    workbook_id: canonical,
                    short_id,
                    path: single.to_path_buf(),
                });
            }
            return Err(anyhow!(
                "workbook id {} not found in single-workbook mode (expected {} or {})",
                candidate,
                canonical.as_str(),
                short_id
            ));
        }

        for entry in WalkDir::new(&self.config.workspace_root) {
            let entry = entry?;
            if !entry.file_type().is_file() {
                continue;
            }
            let path = entry.path();
            if !has_supported_extension(&self.config.supported_extensions, path) {
                continue;
            }
            let metadata = entry.metadata()?;
            let canonical = WorkbookId(hash_path_metadata(path, &metadata));
            let slug = path
                .file_stem()
                .map(|s| s.to_string_lossy().to_string())
                .unwrap_or_else(|| "workbook".to_string());
            let short_id = make_short_workbook_id(&slug, canonical.as_str());
            if candidate == canonical.as_str() || candidate_lower == short_id {
                return Ok(LocatedWorkbook {
                    workbook_id: canonical,
                    short_id,
                    path: path.to_path_buf(),
                });
            }
        }
        Err(anyhow!("workbook id {} not found", candidate))
    }

    fn register_location(&self, located: &LocatedWorkbook) {
        self.index
            .write()
            .insert(located.workbook_id.clone(), located.path.clone());
        self.alias_index.write().insert(
            located.short_id.to_ascii_lowercase(),
            located.workbook_id.clone(),
        );
    }
}

struct LocatedWorkbook {
    workbook_id: WorkbookId,
    short_id: String,
    path: PathBuf,
}

fn has_supported_extension(allowed: &[String], path: &Path) -> bool {
    path.extension()
        .and_then(|ext| ext.to_str())
        .map(|ext| {
            let lower = ext.to_ascii_lowercase();
            allowed.iter().any(|candidate| candidate == &lower)
        })
        .unwrap_or(false)
}
