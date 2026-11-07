//! Port of `packages/docx/src/yrs/saveFormatting.ts`.
#![allow(unused_variables, unused_imports, dead_code, clippy::all)]

use std::collections::{HashMap, HashSet};
use std::rc::Rc;
use indexmap::IndexMap;
use crate::js::{Js, Result, SaveError};
use crate::session::{Inheritance, Session};
use crate::{arr, obj, throw};

/// `isStyleSourcedNumPr` (src/yrs/saveFormatting.ts:45).
pub(crate) fn is_style_sourced_num_pr(attrs: &Js) -> Result<bool> {
    todo!("isStyleSourcedNumPr")
}

/// `sameValue` (src/yrs/saveFormatting.ts:53).
pub(crate) fn same_value(left: &Js, right: &Js) -> Result<bool> {
    todo!("sameValue")
}

/// `inheritsValue` (src/yrs/saveFormatting.ts:58).
pub(crate) fn inherits_value(attrs: &Js, orig: &Js, inherited: &Js, key: &Js) -> Result<bool> {
    todo!("inheritsValue")
}

/// `withoutInherited` (src/yrs/saveFormatting.ts:75).
pub(crate) fn without_inherited(formatting: &Js, inherited: &Js) -> Result<Js> {
    todo!("withoutInherited")
}

/// `paragraphAttrsToFormatting` (src/yrs/saveFormatting.ts:110).
pub(crate) fn paragraph_attrs_to_formatting(attrs: &Js, inherited: &Js) -> Result<Js> {
    todo!("paragraphAttrsToFormatting")
}

/// `tableAttrsToFormatting` (src/yrs/saveFormatting.ts:241).
pub(crate) fn table_attrs_to_formatting(attrs: &Js) -> Result<Js> {
    todo!("tableAttrsToFormatting")
}

/// `measurementMargins` (src/yrs/saveFormatting.ts:302).
pub(crate) fn measurement_margins(margins: &Js) -> Result<Js> {
    todo!("measurementMargins")
}

/// `tableRowAttrsToFormatting` (src/yrs/saveFormatting.ts:327).
pub(crate) fn table_row_attrs_to_formatting(attrs: &Js) -> Result<Js> {
    todo!("tableRowAttrsToFormatting")
}

/// `tableCellAttrsToFormatting` (src/yrs/saveFormatting.ts:379).
pub(crate) fn table_cell_attrs_to_formatting(attrs: &Js) -> Result<Js> {
    todo!("tableCellAttrsToFormatting")
}

/// `cellMargins` (src/yrs/saveFormatting.ts:438).
pub(crate) fn cell_margins(margins: &Js) -> Result<Js> {
    todo!("cellMargins")
}
