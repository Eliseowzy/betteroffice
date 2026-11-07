//! Port of `packages/docx/src/utils/textFormattingMerge.ts`.
#![allow(unused_variables, unused_imports, dead_code, clippy::all)]

use std::collections::{HashMap, HashSet};
use std::rc::Rc;
use indexmap::IndexMap;
use crate::js::{Js, Result, SaveError};
use crate::session::{Inheritance, Session};
use crate::{arr, obj, throw};
use crate::font_family_merge::merge_font_family;

/// `mergeTextFormatting` (src/utils/textFormattingMerge.ts:23).
pub(crate) fn merge_text_formatting(target: &Js, source: &Js) -> Result<Js> {
    if !source.truthy() && !target.truthy() {
        return Ok(Js::Undefined);
    }
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
        let is_object = value.is_object_like() && !value.is_null();
        if key == "fontFamily" && is_object {
            result.set("fontFamily", merge_font_family(&target.get("fontFamily"), &value)?);
            continue;
        }
        if key == "color" && is_object {
            let c = &value;
            let has_explicit = c.get("rgb").truthy() || c.get("themeColor").truthy() || c.get("themeTint").truthy() || c.get("themeShade").truthy();
            if !c.get("auto").truthy() || has_explicit {
                result.set("color", c.clone());
            }
            continue;
        }
        if is_object && !value.is_arr() {
            let merged = target.get(&key).nullish_or(Js::new_obj()).spread();
            merged.assign(&value);
            result.set(&key, merged);
            continue;
        }
        result.set(&key, value);
    }
    Ok(result)
}

/// `mergeStyleRunFormatting` (src/utils/textFormattingMerge.ts:66).
pub(crate) fn merge_style_run_formatting(target: &Js, source: &Js) -> Result<Js> {
    let result = merge_text_formatting(target, source)?;
    if !result.truthy() || !source.truthy() {
        return Ok(result);
    }
    for key in [
        "bold", "boldCs", "italic", "italicCs", "allCaps", "smallCaps", "strike", "doubleStrike", "outline", "shadow", "emboss", "imprint", "hidden",
    ] {
        if let Some(b) = source.get(key).as_bool() {
            result.set(key, target.get(key).truthy() != b);
        }
    }
    Ok(result)
}
