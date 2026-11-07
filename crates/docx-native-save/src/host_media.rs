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
    if comments.is_arr() {
        for item in comments.items() {
            resolve_comment_media(&item, data_url)?;
        }
    } else if comments.truthy() && comments.is_obj() {
        for (key, field) in comments.entries() {
            if key == "src" && field.is_str() {
                let token = field.str_or_empty().to_owned();
                comments.set(&key, data_url(&token)?.map_or(field, Js::from));
            } else {
                resolve_comment_media(&field, data_url)?;
            }
        }
    }
    Ok(())
}

/// `resolveHostJsonCommentMedia` (src/yrs/hostMedia.ts:22).
pub(crate) fn resolve_host_json_comment_media(json: &str, data_url: &dyn Fn(&str) -> Result<Option<String>>) -> Result<String> {
    if !json.contains("\"media:") {
        return Ok(json.to_owned());
    }
    let host = Js::parse(json)?;
    let comments = host.path(&["envelope", "document", "package", "document", "comments"]);
    if comments.is_undefined() {
        return Ok(json.to_owned());
    }
    resolve_comment_media(&comments, data_url)?;
    Ok(host.stringify().unwrap_or_default())
}
