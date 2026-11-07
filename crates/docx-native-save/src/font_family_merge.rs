//! Port of `packages/docx/src/utils/fontFamilyMerge.ts`.
#![allow(unused_variables, unused_imports, dead_code, clippy::all)]

use crate::js::{Js, Result, SaveError};
use crate::session::{Inheritance, Session};
use crate::{arr, obj, throw};
use indexmap::IndexMap;
use std::collections::{HashMap, HashSet};
use std::rc::Rc;

/// `mergeFontFamily` (src/utils/fontFamilyMerge.ts:18).
pub(crate) fn merge_font_family(target: &Js, source: &Js) -> Result<Js> {
    let result = if target.truthy() {
        target.spread()
    } else {
        Js::new_obj()
    };
    let src = source;
    let pairs = [
        ("ascii", "asciiTheme"),
        ("hAnsi", "hAnsiTheme"),
        ("eastAsia", "eastAsiaTheme"),
        ("cs", "csTheme"),
    ];
    let mut pair_keys: Vec<&str> = Vec::new();
    for (explicit, theme) in pairs {
        pair_keys.push(explicit);
        pair_keys.push(theme);
        if !src.get(explicit).is_undefined() || !src.get(theme).is_undefined() {
            result.delete(explicit);
            result.delete(theme);
            if !src.get(explicit).is_undefined() {
                result.set(explicit, src.get(explicit));
            }
            if !src.get(theme).is_undefined() {
                result.set(theme, src.get(theme));
            }
        }
    }
    for key in src.keys() {
        if !pair_keys.contains(&key.as_str()) && !src.get(&key).is_undefined() {
            result.set(&key, src.get(&key));
        }
    }
    Ok(result)
}
