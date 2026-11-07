//! Port of `packages/docx/src/utils/fontFamilyMerge.ts`.
#![allow(unused_variables, unused_imports, dead_code, clippy::all)]

use std::collections::{HashMap, HashSet};
use std::rc::Rc;
use indexmap::IndexMap;
use crate::js::{Js, Result, SaveError};
use crate::session::{Inheritance, Session};
use crate::{arr, obj, throw};

/// `mergeFontFamily` (src/utils/fontFamilyMerge.ts:18).
pub(crate) fn merge_font_family(target: &Js, source: &Js) -> Result<Js> {
    todo!("mergeFontFamily")
}
