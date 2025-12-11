use anyhow::{Result, anyhow};
use chrono::{DateTime, Utc};
use parking_lot::Mutex;
use sha2::{Digest, Sha256};
use std::collections::HashMap;
use std::fs;
use std::path::{Path, PathBuf};
use std::sync::Arc;
use std::time::{Duration, Instant};

const FORK_DIR: &str = "/tmp/mcp-forks";
const DEFAULT_TTL_SECS: u64 = 3600;
const DEFAULT_MAX_FORKS: usize = 10;
const CLEANUP_TASK_CHECK_SECS: u64 = 60;
const MAX_FILE_SIZE: u64 = 100 * 1024 * 1024; // 100MB

#[derive(Debug, Clone)]
pub struct EditOp {
    pub timestamp: DateTime<Utc>,
    pub sheet: String,
    pub address: String,
    pub value: String,
    pub is_formula: bool,
}

#[derive(Debug)]
pub struct ForkContext {
    pub fork_id: String,
    pub base_path: PathBuf,
    pub work_path: PathBuf,
    pub created_at: Instant,
    pub edits: Vec<EditOp>,
    base_hash: String,
    base_modified: std::time::SystemTime,
}

impl ForkContext {
    fn new(fork_id: String, base_path: PathBuf, work_path: PathBuf) -> Result<Self> {
        let metadata = fs::metadata(&base_path)?;
        let base_modified = metadata.modified()?;
        let base_hash = hash_file(&base_path)?;

        Ok(Self {
            fork_id,
            base_path,
            work_path,
            created_at: Instant::now(),
            edits: Vec::new(),
            base_hash,
            base_modified,
        })
    }

    pub fn is_expired(&self, ttl: Duration) -> bool {
        self.created_at.elapsed() > ttl
    }

    pub fn validate_base_unchanged(&self) -> Result<()> {
        let metadata = fs::metadata(&self.base_path)?;
        let current_modified = metadata.modified()?;

        if current_modified != self.base_modified {
            return Err(anyhow!("base file modified since fork creation"));
        }

        let current_hash = hash_file(&self.base_path)?;
        if current_hash != self.base_hash {
            return Err(anyhow!("base file content changed since fork creation"));
        }

        Ok(())
    }
}

fn hash_file(path: &Path) -> Result<String> {
    let contents = fs::read(path)?;
    let mut hasher = Sha256::new();
    hasher.update(&contents);
    Ok(format!("{:x}", hasher.finalize()))
}

#[derive(Debug, Clone)]
pub struct ForkConfig {
    pub ttl: Duration,
    pub max_forks: usize,
    pub fork_dir: PathBuf,
}

impl Default for ForkConfig {
    fn default() -> Self {
        Self {
            ttl: Duration::from_secs(DEFAULT_TTL_SECS),
            max_forks: DEFAULT_MAX_FORKS,
            fork_dir: PathBuf::from(FORK_DIR),
        }
    }
}

pub struct ForkRegistry {
    forks: Mutex<HashMap<String, ForkContext>>,
    config: ForkConfig,
}

impl ForkRegistry {
    pub fn new(config: ForkConfig) -> Result<Self> {
        fs::create_dir_all(&config.fork_dir)?;
        Ok(Self {
            forks: Mutex::new(HashMap::new()),
            config,
        })
    }

    pub fn start_cleanup_task(self: Arc<Self>) {
        tokio::spawn(async move {
            let mut interval = tokio::time::interval(Duration::from_secs(CLEANUP_TASK_CHECK_SECS));
            loop {
                interval.tick().await;
                self.evict_expired();
            }
        });
    }

    pub fn create_fork(&self, base_path: &Path, workspace_root: &Path) -> Result<String> {
        self.evict_expired();

        {
            let forks = self.forks.lock();
            if forks.len() >= self.config.max_forks {
                return Err(anyhow!(
                    "max forks ({}) reached, discard existing forks first",
                    self.config.max_forks
                ));
            }
        }

        let ext = base_path
            .extension()
            .and_then(|e| e.to_str())
            .map(|e| e.to_ascii_lowercase());

        if ext.as_deref() != Some("xlsx") {
            return Err(anyhow!(
                "only .xlsx files supported for fork/recalc (got {:?})",
                ext
            ));
        }

        if !base_path.starts_with(workspace_root) {
            return Err(anyhow!("base path must be within workspace root"));
        }

        if !base_path.exists() {
            return Err(anyhow!("base file does not exist: {:?}", base_path));
        }

        let metadata = fs::metadata(base_path)?;
        if metadata.len() > MAX_FILE_SIZE {
            return Err(anyhow!(
                "base file too large: {} bytes (max {} MB)",
                metadata.len(),
                MAX_FILE_SIZE / 1024 / 1024
            ));
        }

        let fork_id = uuid::Uuid::new_v4().to_string();
        let work_path = self.config.fork_dir.join(format!("{}.xlsx", fork_id));

        fs::copy(base_path, &work_path)?;

        let context = ForkContext::new(fork_id.clone(), base_path.to_path_buf(), work_path)?;

        self.forks.lock().insert(fork_id.clone(), context);

        Ok(fork_id)
    }

    pub fn get_fork(&self, fork_id: &str) -> Result<Arc<ForkContext>> {
        self.evict_expired();

        let forks = self.forks.lock();
        forks
            .get(fork_id)
            .map(|ctx| Arc::new(ctx.clone()))
            .ok_or_else(|| anyhow!("fork not found: {}", fork_id))
    }

    pub fn with_fork_mut<F, R>(&self, fork_id: &str, f: F) -> Result<R>
    where
        F: FnOnce(&mut ForkContext) -> Result<R>,
    {
        let mut forks = self.forks.lock();
        let ctx = forks
            .get_mut(fork_id)
            .ok_or_else(|| anyhow!("fork not found: {}", fork_id))?;
        f(ctx)
    }

    pub fn discard_fork(&self, fork_id: &str) -> Result<()> {
        let mut forks = self.forks.lock();
        if let Some(ctx) = forks.remove(fork_id) {
            let _ = fs::remove_file(&ctx.work_path);
        }
        Ok(())
    }

    pub fn save_fork(
        &self,
        fork_id: &str,
        target_path: &Path,
        workspace_root: &Path,
    ) -> Result<()> {
        if !target_path.starts_with(workspace_root) {
            return Err(anyhow!("target path must be within workspace root"));
        }

        let ext = target_path
            .extension()
            .and_then(|e| e.to_str())
            .map(|e| e.to_ascii_lowercase());

        if ext.as_deref() != Some("xlsx") {
            return Err(anyhow!("target must be .xlsx"));
        }

        let mut forks = self.forks.lock();
        let ctx = forks
            .get(fork_id)
            .ok_or_else(|| anyhow!("fork not found: {}", fork_id))?;

        ctx.validate_base_unchanged()?;

        fs::copy(&ctx.work_path, target_path)?;

        if let Some(ctx) = forks.remove(fork_id) {
            let _ = fs::remove_file(&ctx.work_path);
        }

        Ok(())
    }

    pub fn list_forks(&self) -> Vec<ForkInfo> {
        self.evict_expired();

        let forks = self.forks.lock();
        forks
            .values()
            .map(|ctx| ForkInfo {
                fork_id: ctx.fork_id.clone(),
                base_path: ctx.base_path.display().to_string(),
                created_at: ctx.created_at,
                edit_count: ctx.edits.len(),
            })
            .collect()
    }

    fn evict_expired(&self) {
        let mut forks = self.forks.lock();
        let expired: Vec<String> = forks
            .iter()
            .filter(|(_, ctx)| ctx.is_expired(self.config.ttl))
            .map(|(id, _)| id.clone())
            .collect();

        for id in expired {
            if let Some(ctx) = forks.remove(&id) {
                let _ = fs::remove_file(&ctx.work_path);
                tracing::debug!(fork_id = %id, "evicted expired fork");
            }
        }
    }
}

impl Clone for ForkContext {
    fn clone(&self) -> Self {
        Self {
            fork_id: self.fork_id.clone(),
            base_path: self.base_path.clone(),
            work_path: self.work_path.clone(),
            created_at: self.created_at,
            edits: self.edits.clone(),
            base_hash: self.base_hash.clone(),
            base_modified: self.base_modified,
        }
    }
}

#[derive(Debug, Clone)]
pub struct ForkInfo {
    pub fork_id: String,
    pub base_path: String,
    pub created_at: Instant,
    pub edit_count: usize,
}
