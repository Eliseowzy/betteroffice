//! Port of `packages/docx/src/yrs/tableRowSpans.ts`.
#![allow(unused_variables, unused_imports, dead_code, clippy::all)]

use std::collections::{HashMap, HashSet};
use std::rc::Rc;
use indexmap::IndexMap;
use crate::js::{Js, Result, SaveError};
use crate::session::{Inheritance, Session};
use crate::{arr, obj, throw};

/// `calculateRowSpans` (src/yrs/tableRowSpans.ts:7).
pub(crate) fn calculate_row_spans(table: &Js) -> Result<Js> {
    todo!("calculateRowSpans")
}

/// `seededCellIndices` (src/yrs/tableRowSpans.ts:49).
pub(crate) fn seeded_cell_indices(table: &Js) -> Result<Js> {
    todo!("seededCellIndices")
}
