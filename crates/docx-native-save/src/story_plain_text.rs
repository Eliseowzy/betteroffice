//! Port of `packages/docx/src/yrs/storyPlainText.ts`.
#![allow(unused_variables, unused_imports, dead_code, clippy::all)]

use std::collections::{HashMap, HashSet};
use std::rc::Rc;
use indexmap::IndexMap;
use crate::js::{Js, Result, SaveError};
use crate::session::{Inheritance, Session};
use crate::{arr, obj, throw};

/// `positiveSpan` (src/yrs/storyPlainText.ts:43).
pub(crate) fn positive_span(value: &Js) -> Result<f64> {
    todo!("positiveSpan")
}

/// `objectValue` (src/yrs/storyPlainText.ts:47).
pub(crate) fn object_value(value: &Js) -> Result<Js> {
    todo!("objectValue")
}

/// `embedPlainText` (src/yrs/storyPlainText.ts:54).
pub(crate) fn embed_plain_text(reader: &Session, kind: &str, payload: &Js) -> Result<String> {
    todo!("embedPlainText")
}

/// `cellPlainText` (src/yrs/storyPlainText.ts:90).
pub(crate) fn cell_plain_text(reader: &Session, story: &str) -> Result<String> {
    todo!("cellPlainText")
}

/// `tablePlainText` (src/yrs/storyPlainText.ts:96).
pub(crate) fn table_plain_text(reader: &Session, payload: &Js, range: &Js) -> Result<Vec<String>> {
    todo!("tablePlainText")
}

/// `storyPlainText` (src/yrs/storyPlainText.ts:135).
pub(crate) fn story_plain_text(reader: &Session, story: &str, from: f64, to: f64) -> Result<String> {
    todo!("storyPlainText")
}

/// `tableRowGrid` (src/yrs/storyPlainText.ts:154).
pub(crate) fn table_row_grid(tr_pr: &Js) -> Result<Js> {
    todo!("tableRowGrid")
}

/// `tableAnchors` (src/yrs/storyPlainText.ts:165).
pub(crate) fn table_anchors(payload: &Js) -> Result<Js> {
    todo!("tableAnchors")
}

/// `storyOffsetForLoc` (src/yrs/storyPlainText.ts:207).
pub(crate) fn story_offset_for_loc(reader: &Js, loc: &Js) -> Result<f64> {
    todo!("storyOffsetForLoc")
}
