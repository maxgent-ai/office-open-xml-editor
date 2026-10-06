//! Per-archive, immutable worksheet resource policy.
//!
//! A policy is chosen at most once for one admitted archive. It is chosen after
//! package admission and before any parse operation, and it is never stored in
//! module-global state. Every value is a finite, positive integer no greater
//! than `Number.MAX_SAFE_INTEGER - 1`, so the JS boundary can represent it
//! exactly.
//!
//! The limits are admission checks on measured usage. They are never used to
//! size allocations, so a very large limit does not reserve memory. The policy
//! adds no extra grid or byte caps. The defaults are the generated hard XLSX
//! worksheet constants, so legacy callers keep the same thresholds.

use ooxml_common::resource::{
    HARD_MAX_XLSX_WORKSHEET_CELLS, HARD_MAX_XLSX_WORKSHEET_CELL_CONTENT_UTF8_BYTES,
    HARD_MAX_XLSX_WORKSHEET_JSON_BYTES, HARD_MAX_XLSX_WORKSHEET_ROWS,
};

/// Largest accepted policy value: `Number.MAX_SAFE_INTEGER - 1`.
pub(crate) const MAX_WORKSHEET_POLICY_LIMIT: u64 = 9_007_199_254_740_990;

/// Caller-adjustable worksheet limits for one archive.
///
/// The fields are private and there are no setters. A value cannot change
/// after construction. To use different limits, build a new policy and pass it
/// to a fresh archive.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct WorksheetResourcePolicy {
    max_rows: u64,
    max_cells: u64,
    max_owned_utf8_bytes: u64,
    max_json_bytes: u64,
}

fn validate_limit(name: &str, value: u64) -> Result<u64, String> {
    if value == 0 || value > MAX_WORKSHEET_POLICY_LIMIT {
        return Err(format!(
            "worksheet resource policy {name} must be a positive integer no greater than {MAX_WORKSHEET_POLICY_LIMIT} (got {value})"
        ));
    }
    Ok(value)
}

impl WorksheetResourcePolicy {
    /// Validate and build an immutable policy.
    pub fn new(
        max_rows: u64,
        max_cells: u64,
        max_owned_utf8_bytes: u64,
        max_json_bytes: u64,
    ) -> Result<Self, String> {
        Ok(Self {
            max_rows: validate_limit("maxRows", max_rows)?,
            max_cells: validate_limit("maxCells", max_cells)?,
            max_owned_utf8_bytes: validate_limit("maxOwnedUtf8Bytes", max_owned_utf8_bytes)?,
            max_json_bytes: validate_limit("maxJsonBytes", max_json_bytes)?,
        })
    }

    pub(crate) fn max_rows(&self) -> u64 {
        self.max_rows
    }

    pub(crate) fn max_cells(&self) -> u64 {
        self.max_cells
    }

    pub(crate) fn max_owned_utf8_bytes(&self) -> u64 {
        self.max_owned_utf8_bytes
    }

    pub(crate) fn max_json_bytes(&self) -> u64 {
        self.max_json_bytes
    }
}

impl Default for WorksheetResourcePolicy {
    /// The generated hard XLSX worksheet constants, used as policy defaults.
    fn default() -> Self {
        Self {
            max_rows: HARD_MAX_XLSX_WORKSHEET_ROWS,
            max_cells: HARD_MAX_XLSX_WORKSHEET_CELLS,
            max_owned_utf8_bytes: HARD_MAX_XLSX_WORKSHEET_CELL_CONTENT_UTF8_BYTES,
            max_json_bytes: HARD_MAX_XLSX_WORKSHEET_JSON_BYTES,
        }
    }
}
