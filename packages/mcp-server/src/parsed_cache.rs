use std::any::{Any, TypeId};
use std::collections::{HashMap, VecDeque};
use std::fs;
use std::path::PathBuf;
use std::sync::{Arc, Mutex, OnceLock};
use std::time::SystemTime;

use docx_model::Document;
use pptx_model::Presentation;
use xlsx_model::{Workbook, Worksheet};

const CACHE_ENTRIES: usize = 8;

#[derive(Clone, Hash, PartialEq, Eq)]
struct Key {
    path: PathBuf,
    modified: SystemTime,
    len: u64,
}

type CachedModel = Arc<dyn Any + Send + Sync>;
struct Entry {
    key: Key,
    model: Option<CachedModel>,
    markdown: HashMap<MarkdownKind, Arc<String>>,
}

type Flight = Arc<OnceLock<Result<CachedModel, String>>>;

#[derive(Default)]
struct CacheState {
    entries: VecDeque<Entry>,
    model_flights: HashMap<(Key, TypeId), Flight>,
    markdown_flights: HashMap<(Key, MarkdownKind), Flight>,
}

impl CacheState {
    fn upsert(&mut self, key: Key, update: impl FnOnce(&mut Entry)) {
        let mut entry = self
            .entries
            .iter()
            .position(|entry| entry.key == key)
            .and_then(|position| self.entries.remove(position))
            .unwrap_or_else(|| Entry {
                key,
                model: None,
                markdown: HashMap::new(),
            });
        update(&mut entry);
        self.entries.push_back(entry);
        while self.entries.len() > CACHE_ENTRIES {
            self.entries.pop_front();
        }
    }
}

#[derive(Clone, Copy, PartialEq, Eq, Hash)]
pub enum MarkdownKind {
    Docx,
    Xlsx,
    Pptx,
}

static CACHE: OnceLock<Mutex<CacheState>> = OnceLock::new();

fn cache() -> &'static Mutex<CacheState> {
    CACHE.get_or_init(|| Mutex::new(CacheState::default()))
}

fn identity(path: &str) -> Result<Key, String> {
    let canonical = fs::canonicalize(path).map_err(|e| format!("Cannot read '{}': {}", path, e))?;
    let meta = fs::metadata(&canonical).map_err(|e| format!("Cannot read '{}': {}", path, e))?;
    if meta.is_dir() {
        // Preserve the previous fs::read error for directories.
        let detail = fs::read(&canonical).err().map_or_else(
            || "not a regular file".to_string(),
            |error| error.to_string(),
        );
        return Err(format!("Cannot read '{}': {}", path, detail));
    }
    if !meta.is_file() {
        return Err(format!("Cannot read '{}': not a regular file", path));
    }
    Ok(Key {
        path: canonical,
        modified: meta
            .modified()
            .map_err(|e| format!("Cannot read '{}': {}", path, e))?,
        len: meta.len(),
    })
}

fn read_checked(path: &str, key: &Key) -> Result<Vec<u8>, String> {
    let data = fs::read(&key.path).map_err(|e| format!("Cannot read '{}': {}", path, e))?;
    if identity(path)? != *key || data.len() as u64 != key.len {
        return Err(format!("Cannot read '{}': file changed during read", path));
    }
    Ok(data)
}

fn get<T: Any + Send + Sync>(
    path: &str,
    parse: impl FnOnce(&[u8], &Key) -> Result<T, String>,
) -> Result<Arc<T>, String> {
    get_in(cache(), path, parse)
}

fn get_in<T: Any + Send + Sync>(
    entries_lock: &Mutex<CacheState>,
    path: &str,
    parse: impl FnOnce(&[u8], &Key) -> Result<T, String>,
) -> Result<Arc<T>, String> {
    let key = identity(path)?;
    let flight_key = (key.clone(), TypeId::of::<T>());
    let (found, flight) = {
        let mut state = entries_lock
            .lock()
            .unwrap_or_else(|poison| poison.into_inner());
        let found = state
            .entries
            .iter()
            .position(|entry| {
                entry.key == key && entry.model.as_ref().is_some_and(|model| model.is::<T>())
            })
            .and_then(|position| state.entries.remove(position))
            .map(|entry| {
                let found = Arc::clone(entry.model.as_ref().expect("matched model"));
                state.entries.push_back(entry);
                found
            });
        let flight = if found.is_none() {
            Some(Arc::clone(
                state.model_flights.entry(flight_key.clone()).or_default(),
            ))
        } else {
            None
        };
        (found, flight)
    };
    if let Some(found) = found {
        return found
            .downcast::<T>()
            .map_err(|_| "cache type mismatch".to_string());
    }
    let flight = flight.expect("miss registered an in-flight parse");
    // OnceLock makes all callers of this key share the same parse and result.
    // A failed parse is removed from the flight map and is never cached.
    let result = flight.get_or_init(|| {
        let result = read_checked(path, &key)
            .and_then(|data| parse(&data, &key))
            .map(|model| Arc::new(model) as CachedModel);
        let mut state = entries_lock
            .lock()
            .unwrap_or_else(|poison| poison.into_inner());
        if let Ok(model) = &result {
            state.upsert(key.clone(), |entry| entry.model = Some(Arc::clone(model)));
        }
        state.model_flights.remove(&flight_key);
        result
    });
    result.clone().and_then(|model| {
        model
            .downcast::<T>()
            .map_err(|_| "cache type mismatch".into())
    })
}

pub fn docx(path: &str) -> Result<Arc<Document>, String> {
    get(path, |bytes, _| docx_parser::parse_docx_model_native(bytes))
}

pub fn pptx(path: &str) -> Result<Arc<Presentation>, String> {
    get(path, |bytes, _| pptx_parser::parse_pptx_model_native(bytes))
}

pub struct XlsxDocument {
    pub workbook: Workbook,
    sheets: Mutex<SheetCache>,
    key: Key,
}

#[derive(Default)]
struct SheetCache {
    entries: VecDeque<(u32, Arc<Worksheet>)>,
    flights: HashMap<u32, SheetFlight>,
}

type SheetFlight = Arc<OnceLock<Result<Arc<Worksheet>, String>>>;

pub fn xlsx(path: &str) -> Result<Arc<XlsxDocument>, String> {
    get(path, |bytes, key| {
        Ok(XlsxDocument {
            workbook: xlsx_parser::parse_workbook_model_native(bytes)?,
            sheets: Mutex::new(SheetCache::default()),
            key: key.clone(),
        })
    })
}

pub fn xlsx_sheet(
    path: &str,
    document: &XlsxDocument,
    index: u32,
    name: &str,
) -> Result<Arc<Worksheet>, String> {
    let key = identity(path)?;
    if key != document.key {
        return Err(format!("Cannot read '{}': file changed during read", path));
    }
    let (found, flight) = {
        let mut cache = document
            .sheets
            .lock()
            .unwrap_or_else(|poison| poison.into_inner());
        let found = cache
            .entries
            .iter()
            .position(|(candidate, _)| *candidate == index)
            .and_then(|position| cache.entries.remove(position))
            .map(|entry| {
                let sheet = Arc::clone(&entry.1);
                cache.entries.push_back(entry);
                sheet
            });
        let flight = if found.is_none() {
            Some(Arc::clone(cache.flights.entry(index).or_default()))
        } else {
            None
        };
        (found, flight)
    };
    if let Some(sheet) = found {
        return Ok(sheet);
    }
    let flight = flight.expect("miss registered an in-flight sheet parse");
    flight
        .get_or_init(|| {
            let result = read_checked(path, &key)
                .and_then(|bytes| xlsx_parser::parse_sheet_model_native(&bytes, index, name))
                .map(Arc::new);
            let mut cache = document
                .sheets
                .lock()
                .unwrap_or_else(|poison| poison.into_inner());
            if let Ok(sheet) = &result {
                cache.entries.retain(|(candidate, _)| *candidate != index);
                cache.entries.push_back((index, Arc::clone(sheet)));
                while cache.entries.len() > CACHE_ENTRIES {
                    cache.entries.pop_front();
                }
            }
            cache.flights.remove(&index);
            result
        })
        .clone()
}

pub fn markdown(
    path: &str,
    kind: MarkdownKind,
    render: impl FnOnce(&[u8]) -> Result<String, String>,
) -> Result<String, String> {
    let key = identity(path)?;
    let flight_key = (key.clone(), kind);
    let (found, flight) = {
        let mut state = cache().lock().unwrap_or_else(|poison| poison.into_inner());
        let found = state
            .entries
            .iter()
            .position(|entry| entry.key == key && entry.markdown.contains_key(&kind))
            .and_then(|position| state.entries.remove(position))
            .map(|entry| {
                let found = Arc::clone(entry.markdown.get(&kind).expect("matched markdown"));
                state.entries.push_back(entry);
                found
            });
        let flight = if found.is_none() {
            Some(Arc::clone(
                state
                    .markdown_flights
                    .entry(flight_key.clone())
                    .or_default(),
            ))
        } else {
            None
        };
        (found, flight)
    };
    if let Some(found) = found {
        return Ok((*found).clone());
    }
    let flight = flight.expect("miss registered an in-flight markdown parse");
    let result = flight.get_or_init(|| {
        let result = read_checked(path, &key)
            .and_then(|bytes| render(&bytes))
            .map(|output| Arc::new(output) as CachedModel);
        let mut state = cache().lock().unwrap_or_else(|poison| poison.into_inner());
        if let Ok(output) = &result {
            let output = output
                .clone()
                .downcast::<String>()
                .expect("markdown flight type");
            state.upsert(key.clone(), |entry| {
                entry.markdown.insert(kind, output);
            });
        }
        state.markdown_flights.remove(&flight_key);
        result
    });
    result
        .clone()
        .and_then(|output| {
            output
                .downcast::<String>()
                .map_err(|_| "cache type mismatch".into())
        })
        .map(|output| (*output).clone())
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::sync::atomic::{AtomicUsize, Ordering};
    use std::sync::Barrier;
    use std::time::Duration;

    #[test]
    fn concurrent_identical_misses_share_one_parse_and_failure_is_retryable() {
        let entries = Mutex::new(CacheState::default());
        let root = std::env::temp_dir().join(format!(
            "ooxml-mcp-flight-{}-{}",
            std::process::id(),
            SystemTime::now()
                .duration_since(SystemTime::UNIX_EPOCH)
                .unwrap()
                .as_nanos()
        ));
        fs::create_dir(&root).unwrap();
        let path = root.join("model.bin");
        fs::write(&path, b"model").unwrap();
        let path = path.to_str().unwrap();
        let calls = AtomicUsize::new(0);
        let start = Barrier::new(17);
        std::thread::scope(|scope| {
            let requests: Vec<_> = (0..16)
                .map(|_| {
                    let start = &start;
                    let calls = &calls;
                    let entries = &entries;
                    scope.spawn(move || {
                        start.wait();
                        get_in(entries, path, |bytes, _| {
                            calls.fetch_add(1, Ordering::SeqCst);
                            std::thread::sleep(Duration::from_millis(40));
                            Ok::<_, String>(bytes.to_vec())
                        })
                        .unwrap()
                    })
                })
                .collect();
            start.wait();
            let results: Vec<_> = requests
                .into_iter()
                .map(|request| request.join().unwrap())
                .collect();
            assert!(results
                .iter()
                .all(|result| Arc::ptr_eq(result, &results[0])));
        });
        assert_eq!(calls.load(Ordering::SeqCst), 1);

        fs::write(path, b"invalid model").unwrap();
        let failed_calls = AtomicUsize::new(0);
        let start = Barrier::new(17);
        std::thread::scope(|scope| {
            let requests: Vec<_> = (0..16)
                .map(|_| {
                    let start = &start;
                    let calls = &failed_calls;
                    let entries = &entries;
                    scope.spawn(move || {
                        start.wait();
                        get_in::<Vec<u8>>(entries, path, |_, _| {
                            calls.fetch_add(1, Ordering::SeqCst);
                            std::thread::sleep(Duration::from_millis(40));
                            Err("invalid model".into())
                        })
                    })
                })
                .collect();
            start.wait();
            for request in requests {
                assert_eq!(request.join().unwrap().unwrap_err(), "invalid model");
            }
        });
        assert_eq!(failed_calls.load(Ordering::SeqCst), 1);
        let retried = get_in(&entries, path, |bytes, _| {
            failed_calls.fetch_add(1, Ordering::SeqCst);
            Ok::<_, String>(bytes.to_vec())
        })
        .unwrap();
        assert_eq!(*retried, b"invalid model");
        assert_eq!(failed_calls.load(Ordering::SeqCst), 2);

        let markdown_calls = AtomicUsize::new(0);
        let start = Barrier::new(17);
        std::thread::scope(|scope| {
            let requests: Vec<_> = (0..16)
                .map(|_| {
                    let start = &start;
                    let calls = &markdown_calls;
                    scope.spawn(move || {
                        start.wait();
                        markdown(path, MarkdownKind::Docx, |bytes| {
                            calls.fetch_add(1, Ordering::SeqCst);
                            std::thread::sleep(Duration::from_millis(40));
                            Ok(String::from_utf8(bytes.to_vec()).unwrap())
                        })
                        .unwrap()
                    })
                })
                .collect();
            start.wait();
            for request in requests {
                assert_eq!(request.join().unwrap(), "invalid model");
            }
        });
        assert_eq!(markdown_calls.load(Ordering::SeqCst), 1);
        fs::remove_file(path).unwrap();
        fs::remove_dir(root).unwrap();
    }

    #[test]
    fn cache_reuses_unchanged_model_evicts_old_entries_and_checks_identity() {
        let entries = Mutex::new(CacheState::default());
        let root = std::env::temp_dir().join(format!(
            "ooxml-mcp-cache-{}-{}",
            std::process::id(),
            SystemTime::now()
                .duration_since(SystemTime::UNIX_EPOCH)
                .unwrap()
                .as_nanos()
        ));
        fs::create_dir(&root).unwrap();
        let path = root.join("model.bin");
        fs::write(&path, b"a").unwrap();
        let path_str = path.to_str().unwrap();
        let parses = AtomicUsize::new(0);
        let parse = |bytes: &[u8], _: &Key| {
            parses.fetch_add(1, Ordering::Relaxed);
            Ok::<_, String>(bytes.to_vec())
        };
        let first = get_in(&entries, path_str, parse).unwrap();
        let reused = get_in(&entries, path_str, parse).unwrap();
        assert!(Arc::ptr_eq(&first, &reused));
        assert_eq!(parses.load(Ordering::Relaxed), 1);

        fs::write(&path, b"changed").unwrap();
        let changed = get_in(&entries, path_str, parse).unwrap();
        assert_eq!(*changed, b"changed");
        assert_eq!(parses.load(Ordering::Relaxed), 2);

        for index in 0..CACHE_ENTRIES {
            let other = root.join(format!("other-{index}.bin"));
            fs::write(&other, b"other").unwrap();
            get_in(&entries, other.to_str().unwrap(), parse).unwrap();
        }
        let reloaded = get_in(&entries, path_str, parse).unwrap();
        assert!(!Arc::ptr_eq(&changed, &reloaded));
        assert_eq!(parses.load(Ordering::Relaxed), CACHE_ENTRIES + 3);

        let previous_identity = identity(path_str).unwrap();
        fs::write(&path, b"changed again").unwrap();
        assert!(read_checked(path_str, &previous_identity)
            .unwrap_err()
            .contains("file changed during read"));
        assert_eq!(parses.load(Ordering::Relaxed), CACHE_ENTRIES + 3);
        for index in 0..CACHE_ENTRIES {
            fs::remove_file(root.join(format!("other-{index}.bin"))).unwrap();
        }
        fs::remove_file(path).unwrap();
        fs::remove_dir(root).unwrap();
    }
}
