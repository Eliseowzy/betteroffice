//! Port of `packages/docx/src/yrs/sections.ts`.
#![allow(unused_variables, unused_imports, dead_code, clippy::all)]

use std::collections::{HashMap, HashSet};
use std::rc::Rc;
use indexmap::IndexMap;
use crate::js::{Js, Result, SaveError};
use crate::session::{Inheritance, Session};
use crate::{arr, obj, throw};

/// `sectionPropertiesEqual` (src/yrs/sections.ts:24).
pub(crate) fn section_properties_equal(left: &Js, right: &Js) -> Result<bool> {
    todo!("sectionPropertiesEqual")
}

/// `class DocxSectionError`.
pub(crate) struct DocxSectionError {
    _todo: (),
}

impl DocxSectionError {
    /// `DocxSectionError.constructor` (src/yrs/sections.ts:18).
    pub(crate) fn new(code: &str, message: &str) -> Result<Self> {
        todo!("DocxSectionError.constructor")
    }

}
