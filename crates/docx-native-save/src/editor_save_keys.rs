//! Port of `packages/docx/src/yrs/editorSaveKeys.ts`.
#![allow(unused_variables, unused_imports, dead_code, clippy::all)]

use std::collections::{HashMap, HashSet};
use std::rc::Rc;
use indexmap::IndexMap;
use crate::js::{Js, Result, SaveError};
use crate::session::{Inheritance, Session};
use crate::{arr, obj, throw};

/// `editorSaveKeys` (src/yrs/editorSaveKeys.ts:5).
pub(crate) fn editor_save_keys(document: &Js, comments: &Js) -> Result<Js> {
    todo!("editorSaveKeys")
}

/// `noteSeparatorKeys` (src/yrs/editorSaveKeys.ts:39).
pub(crate) fn note_separator_keys(document: &Js) -> Result<Js> {
    todo!("noteSeparatorKeys")
}

/// `stringify` (src/yrs/editorSaveKeys.ts:46).
pub(crate) fn stringify(value: &Js) -> Result<String> {
    todo!("stringify")
}
