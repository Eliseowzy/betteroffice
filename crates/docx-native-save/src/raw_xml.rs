//! Port of `packages/docx/src/types/content/rawXml.ts`.
#![allow(unused_variables, unused_imports, dead_code, clippy::all)]

use std::collections::{HashMap, HashSet};
use std::rc::Rc;
use indexmap::IndexMap;
use crate::js::{Js, Result, SaveError};
use crate::session::{Inheritance, Session};
use crate::{arr, obj, throw};

/// `isRawXml` (src/types/content/rawXml.ts:8).
pub(crate) fn is_raw_xml(value: &Js) -> Result<bool> {
    todo!("isRawXml")
}
