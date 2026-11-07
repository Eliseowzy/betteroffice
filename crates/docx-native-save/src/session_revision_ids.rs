//! Port of `packages/docx/src/docx/sessionRevisionIds.ts`.
#![allow(unused_variables, unused_imports, dead_code, clippy::all)]

use std::collections::{HashMap, HashSet};
use std::rc::Rc;
use indexmap::IndexMap;
use crate::js::{Js, Result, SaveError};
use crate::session::{Inheritance, Session};
use crate::{arr, obj, throw};

/// `markRevisionInfo` (src/docx/sessionRevisionIds.ts:20).
pub(crate) fn mark_revision_info(info: &Js, session: bool) -> Result<Js> {
    todo!("markRevisionInfo")
}

/// `withSessionRevisionIds` (src/docx/sessionRevisionIds.ts:39).
pub(crate) fn with_session_revision_ids(tree: &Js, write: &dyn Fn() -> Result<Js>) -> Result<Js> {
    todo!("withSessionRevisionIds")
}
