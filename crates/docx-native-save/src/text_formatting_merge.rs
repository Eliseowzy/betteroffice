//! Port of `packages/docx/src/utils/textFormattingMerge.ts`.
#![allow(unused_variables, unused_imports, dead_code, clippy::all)]

use std::collections::{HashMap, HashSet};
use std::rc::Rc;
use indexmap::IndexMap;
use crate::js::{Js, Result, SaveError};
use crate::session::{Inheritance, Session};
use crate::{arr, obj, throw};

/// `mergeTextFormatting` (src/utils/textFormattingMerge.ts:23).
pub(crate) fn merge_text_formatting(target: &Js, source: &Js) -> Result<Js> {
    todo!("mergeTextFormatting")
}

/// `mergeStyleRunFormatting` (src/utils/textFormattingMerge.ts:66).
pub(crate) fn merge_style_run_formatting(target: &Js, source: &Js) -> Result<Js> {
    todo!("mergeStyleRunFormatting")
}
