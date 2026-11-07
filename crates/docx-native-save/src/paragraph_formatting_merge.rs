//! Port of `packages/docx/src/utils/paragraphFormattingMerge.ts`.
#![allow(unused_variables, unused_imports, dead_code, clippy::all)]

use std::collections::{HashMap, HashSet};
use std::rc::Rc;
use indexmap::IndexMap;
use crate::js::{Js, Result, SaveError};
use crate::session::{Inheritance, Session};
use crate::{arr, obj, throw};
use crate::text_formatting_merge::merge_text_formatting;

/// `mergeParagraphFormatting` (src/utils/paragraphFormattingMerge.ts:4).
pub(crate) fn merge_paragraph_formatting(target: &Js, source: &Js) -> Result<Js> {
    if !source.truthy() {
        return Ok(target.clone());
    }
    if !target.truthy() {
        return Ok(source.spread());
    }
    let result = target.spread();
    for key in source.keys() {
        let value = source.get(&key);
        if value.is_undefined() {
            continue;
        }
        if key == "runProperties" {
            result.set("runProperties", merge_text_formatting(&result.get("runProperties"), &source.get("runProperties"))?);
        } else if key == "borders" || key == "numPr" || key == "frame" {
            let base = result.get(&key);
            let merged = if base.truthy() { base.spread() } else { Js::new_obj() };
            if value.truthy() {
                merged.assign(&value);
            }
            result.set(&key, merged);
        } else if key == "tabs" && value.is_arr() {
            result.set("tabs", value.copy_arr());
        } else {
            result.set(&key, value);
        }
    }
    Ok(result)
}
