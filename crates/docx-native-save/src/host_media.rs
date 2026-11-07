//! Port of `packages/docx/src/yrs/hostMedia.ts`.
#![allow(unused_variables, unused_imports, dead_code, clippy::all)]

use std::collections::{HashMap, HashSet};
use std::rc::Rc;
use indexmap::IndexMap;
use crate::js::{Js, Result, SaveError};
use crate::session::{Inheritance, Session};
use crate::{arr, obj, throw};

/// `resolveCommentMedia` (src/yrs/hostMedia.ts:6).
pub(crate) fn resolve_comment_media(comments: &Js, data_url: &dyn Fn(&str) -> Result<Option<String>>) -> Result<()> {
    todo!("resolveCommentMedia")
}

/// `resolveHostJsonCommentMedia` (src/yrs/hostMedia.ts:22).
pub(crate) fn resolve_host_json_comment_media(json: &str, data_url: &dyn Fn(&str) -> Result<Option<String>>) -> Result<String> {
    todo!("resolveHostJsonCommentMedia")
}
