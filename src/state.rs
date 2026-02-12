use crate::config::ServerConfig;
#[cfg(feature = "recalc")]
use crate::fork::{ForkConfig, ForkRegistry};
use crate::model::{WorkbookId, WorkbookListResponse};
#[cfg(feature = "recalc")]
use crate::recalc::{
    GlobalRecalcLock, GlobalScreenshotLock, LibreOfficeBackend, RecalcBackend, RecalcConfig,
    create_executor,
};
use crate::repository::{PathWorkspaceRepository, WorkbookRepository};
use crate::tools::filters::WorkbookFilter;
use crate::workbook::WorkbookContext;
use anyhow::Result;
use lru::LruCache;
use parking_lot::RwLock;
use std::num::NonZeroUsize;
use std::path::Path;
use std::sync::Arc;
use tokio::task;

pub struct AppState {
    config: Arc<ServerConfig>,
    repository: Arc<dyn WorkbookRepository>,
    cache: RwLock<LruCache<WorkbookId, Arc<WorkbookContext>>>,
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
        #[cfg(feature = "recalc")]
        let (fork_registry, recalc_backend, recalc_semaphore, screenshot_semaphore) =
            init_recalc_components(&config);

        #[cfg(feature = "recalc")]
        let repository: Arc<dyn WorkbookRepository> = Arc::new(PathWorkspaceRepository::new(
            config.clone(),
            fork_registry.clone(),
        ));

        #[cfg(not(feature = "recalc"))]
        let repository: Arc<dyn WorkbookRepository> =
            Arc::new(PathWorkspaceRepository::new(config.clone()));

        let capacity = NonZeroUsize::new(config.cache_capacity.max(1)).unwrap();

        Self {
            config,
            repository,
            cache: RwLock::new(LruCache::new(capacity)),
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

    pub fn new_with_repository(
        config: Arc<ServerConfig>,
        repository: Arc<dyn WorkbookRepository>,
    ) -> Self {
        let capacity = NonZeroUsize::new(config.cache_capacity.max(1)).unwrap();

        #[cfg(feature = "recalc")]
        let (fork_registry, recalc_backend, recalc_semaphore, screenshot_semaphore) =
            init_recalc_components(&config);

        Self {
            config,
            repository,
            cache: RwLock::new(LruCache::new(capacity)),
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
        self.repository.list(&filter)
    }

    pub async fn open_workbook(&self, workbook_id: &WorkbookId) -> Result<Arc<WorkbookContext>> {
        let resolved = self.repository.resolve(workbook_id)?;
        let canonical = resolved.workbook_id.clone();
        {
            let mut cache = self.cache.write();
            if let Some(entry) = cache.get(&canonical) {
                return Ok(entry.clone());
            }
        }

        let repo = self.repository.clone();
        let workbook = task::spawn_blocking(move || repo.load_context(&resolved)).await??;
        let workbook = Arc::new(workbook);

        let mut cache = self.cache.write();
        cache.put(canonical, workbook.clone());
        Ok(workbook)
    }

    pub fn close_workbook(&self, workbook_id: &WorkbookId) -> Result<()> {
        let canonical = self.repository.resolve(workbook_id)?.workbook_id;
        let mut cache = self.cache.write();
        cache.pop(&canonical);
        Ok(())
    }

    pub fn evict_by_path(&self, path: &Path) {
        let evict_ids: Vec<WorkbookId> = self
            .cache
            .read()
            .iter()
            .filter_map(|(id, ctx)| {
                if ctx.path == path {
                    Some(id.clone())
                } else {
                    None
                }
            })
            .collect();

        if evict_ids.is_empty() {
            return;
        }

        let mut cache = self.cache.write();
        for id in evict_ids {
            cache.pop(&id);
        }
    }
}

#[cfg(feature = "recalc")]
fn init_recalc_components(
    config: &Arc<ServerConfig>,
) -> (
    Option<Arc<ForkRegistry>>,
    Option<Arc<dyn RecalcBackend>>,
    Option<GlobalRecalcLock>,
    Option<GlobalScreenshotLock>,
) {
    if !config.recalc_enabled {
        return (None, None, None, None);
    }

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
}
