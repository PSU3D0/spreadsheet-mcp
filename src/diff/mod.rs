pub mod address;
pub mod cells;
pub mod hash;
pub mod merge;
pub mod sst;

use anyhow::Result;
use cells::CellIterator;
use merge::{CellDiff, diff_streams};
use quick_xml::events::Event;
use quick_xml::reader::Reader;
use schemars::JsonSchema;
use serde::Serialize;
use sst::Sst;
use std::collections::HashMap;
use std::fs::File;
use std::io::BufReader;
use std::path::Path;
use zip::ZipArchive;

#[derive(Debug, Serialize, JsonSchema)]
pub struct DiffResult {
    pub sheet: String,
    #[serde(flatten)]
    pub diff: CellDiff,
}

pub fn calculate_changeset(
    base_path: &Path,
    fork_path: &Path,
    sheet_filter: Option<&str>,
) -> Result<Vec<DiffResult>> {
    let mut base_zip = ZipArchive::new(File::open(base_path)?)?;
    let mut fork_zip = ZipArchive::new(File::open(fork_path)?)?;

    // Load SSTs
    let base_sst = load_sst(&mut base_zip).ok();
    let fork_sst = load_sst(&mut fork_zip).ok();

    // Load Sheet Map (Name -> Path)
    // We assume structure hasn't changed drastically (sheet rename/reorder).
    // If it has, we might miss things or fail.
    // For V1, we rely on base structure.
    let sheet_map = load_sheet_map(&mut base_zip)?;

    let mut all_changes = Vec::new();

    for (name, path) in sheet_map {
        if let Some(filter) = sheet_filter
            && name != filter {
                continue;
            }

        // 1. Hash Check
        let base_hash = if let Ok(f) = base_zip.by_name(&path) {
            hash::compute_hash(f)?
        } else {
            0 // Missing file
        };

        let fork_hash = if let Ok(f) = fork_zip.by_name(&path) {
            hash::compute_hash(f)?
        } else {
            0
        };

        if base_hash != 0 && base_hash == fork_hash {
            continue;
        }

        // 2. Diff
        let base_iter = if let Ok(f) = base_zip.by_name(&path) {
            Some(CellIterator::new(BufReader::new(f), base_sst.as_ref()))
        } else {
            None
        };

        let fork_iter = if let Ok(f) = fork_zip.by_name(&path) {
            Some(CellIterator::new(BufReader::new(f), fork_sst.as_ref()))
        } else {
            None
        };

        let diffs = match (base_iter, fork_iter) {
            (Some(b), Some(f)) => diff_streams(b, f)?,
            (Some(b), None) => diff_streams(b, std::iter::empty())?,
            (None, Some(f)) => diff_streams(std::iter::empty(), f)?,
            (None, None) => Vec::new(),
        };

        for d in diffs {
            all_changes.push(DiffResult {
                sheet: name.clone(),
                diff: d,
            });
        }
    }

    Ok(all_changes)
}

fn load_sst(zip: &mut ZipArchive<File>) -> Result<Sst> {
    let f = zip.by_name("xl/sharedStrings.xml")?;
    Sst::from_reader(BufReader::new(f))
}

fn load_sheet_map(zip: &mut ZipArchive<File>) -> Result<HashMap<String, String>> {
    // 1. Parse workbook.xml for name -> rId
    let mut name_to_rid = HashMap::new();
    {
        let workbook_xml = zip.by_name("xl/workbook.xml")?;
        let mut reader = Reader::from_reader(BufReader::new(workbook_xml));

        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                    if e.name().as_ref() == b"sheet" {
                        let mut name = String::new();
                        let mut rid = String::new();
                        for attr in e.attributes() {
                            let attr = attr?;
                            if attr.key.as_ref() == b"name" {
                                name = String::from_utf8_lossy(&attr.value).to_string();
                            } else if attr.key.as_ref() == b"r:id" {
                                rid = String::from_utf8_lossy(&attr.value).to_string();
                            }
                        }
                        if !name.is_empty() && !rid.is_empty() {
                            name_to_rid.insert(rid, name);
                        }
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }
    }

    // 2. Parse _rels/workbook.xml.rels for rId -> Target
    let mut rid_to_target = HashMap::new();
    {
        let rels_xml = zip.by_name("xl/_rels/workbook.xml.rels")?;
        let mut reader = Reader::from_reader(BufReader::new(rels_xml));
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                    if e.name().as_ref() == b"Relationship" {
                        let mut id = String::new();
                        let mut target = String::new();
                        for attr in e.attributes() {
                            let attr = attr?;
                            if attr.key.as_ref() == b"Id" {
                                id = String::from_utf8_lossy(&attr.value).to_string();
                            } else if attr.key.as_ref() == b"Target" {
                                target = String::from_utf8_lossy(&attr.value).to_string();
                            }
                        }
                        rid_to_target.insert(id, target);
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                _ => {}
            }
            buf.clear();
        }
    }

    // 3. Join
    let mut map = HashMap::new();
    for (rid, name) in name_to_rid {
        if let Some(target) = rid_to_target.get(&rid) {
            // Target is usually "worksheets/sheet1.xml" or "/xl/worksheets/sheet1.xml"
            // Zip entries usually don't have leading slash, and start from root (which is where workbook.xml is relative to?)
            // Actually, workbook.xml is in xl/. Target is relative to xl/.
            // So "worksheets/sheet1.xml" -> "xl/worksheets/sheet1.xml"
            let path = if target.starts_with('/') {
                target.trim_start_matches('/').to_string()
            } else {
                format!("xl/{}", target)
            };
            map.insert(name, path);
        }
    }

    Ok(map)
}
