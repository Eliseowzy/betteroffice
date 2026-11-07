//! Port of `packages/docx/src/yrs/sections.ts`.
#![allow(unused_variables, unused_imports, dead_code, clippy::all)]

use crate::js::{Js, Result, SaveError};
use crate::session::{Inheritance, Session};
use crate::{arr, obj, throw};
use indexmap::IndexMap;
use std::collections::{HashMap, HashSet};
use std::rc::Rc;

/// `a.localeCompare(b)` for property names — approximated: case-insensitive first,
/// lowercase before uppercase on ties.
fn locale_compare(a: &str, b: &str) -> std::cmp::Ordering {
    let primary = a.to_lowercase().cmp(&b.to_lowercase());
    if primary != std::cmp::Ordering::Equal {
        return primary;
    }
    for (x, y) in a.chars().zip(b.chars()) {
        if x != y {
            return if x.is_lowercase() {
                std::cmp::Ordering::Less
            } else {
                std::cmp::Ordering::Greater
            };
        }
    }
    a.len().cmp(&b.len())
}

fn sorted(value: &Js) -> Js {
    if value.is_arr() {
        return Js::new_arr(value.items().iter().map(sorted).collect());
    }
    if value.truthy() && (value.is_obj() || value.is_map() || value.is_bytes()) {
        let mut entries: Vec<(String, Js)> = value
            .entries()
            .into_iter()
            .filter(|(_, v)| !v.is_undefined())
            .collect();
        entries.sort_by(|(a, _), (b, _)| locale_compare(a, b));
        let out = Js::new_obj();
        for (k, v) in entries {
            out.set(&k, sorted(&v));
        }
        return out;
    }
    value.clone()
}

/// `sectionPropertiesEqual` (src/yrs/sections.ts:24).
pub(crate) fn section_properties_equal(left: &Js, right: &Js) -> Result<bool> {
    Ok(sorted(left).stringify() == sorted(right).stringify())
}

/// `class DocxSectionError`.
pub(crate) struct DocxSectionError {
    pub(crate) code: String,
    pub(crate) message: String,
}

impl DocxSectionError {
    /// `DocxSectionError.constructor` (src/yrs/sections.ts:18).
    pub(crate) fn new(code: &str, message: &str) -> Result<Self> {
        Ok(Self {
            code: code.to_owned(),
            message: message.to_owned(),
        })
    }

    /// The thrown error (`SaveError::section_unsupported` for code `unsupported`).
    pub(crate) fn into_error(self) -> SaveError {
        if self.code == "unsupported" {
            SaveError::section_unsupported(self.message)
        } else {
            SaveError::new(self.message)
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn section_properties_equal_matches_ts() {
        let a = Js::parse(
            r#"{"pageWidth":12240,"margins":{"top":1440,"left":1800},"cols":[{"w":1,"b":2}]}"#,
        )
        .unwrap();
        let b = Js::parse(
            r#"{"margins":{"left":1800,"top":1440},"cols":[{"b":2,"w":1}],"pageWidth":12240}"#,
        )
        .unwrap();
        assert!(section_properties_equal(&a, &b).unwrap());
        let c = b.spread();
        c.set("extra", Js::Undefined);
        assert!(section_properties_equal(&a, &c).unwrap());
        let d = Js::parse(
            r#"{"margins":{"left":1800,"top":1440},"cols":[{"w":1},{"b":2}],"pageWidth":12240}"#,
        )
        .unwrap();
        assert!(!section_properties_equal(&a, &d).unwrap());
        assert!(section_properties_equal(&Js::Undefined, &Js::Undefined).unwrap());
        assert!(!section_properties_equal(&Js::Null, &Js::new_obj()).unwrap());
    }
}
