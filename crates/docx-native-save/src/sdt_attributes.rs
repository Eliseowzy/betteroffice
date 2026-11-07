//! Port of `packages/docx/src/types/sdtAttributes.ts`.
#![allow(unused_variables, unused_imports, dead_code, clippy::all)]

use std::collections::{HashMap, HashSet};
use std::rc::Rc;
use indexmap::IndexMap;
use crate::js::{Js, Result, SaveError};
use crate::session::{Inheritance, Session};
use crate::{arr, obj, throw};

fn or_null(v: Js) -> Js {
    v.nullish_or(Js::Null)
}

fn json(v: &Js) -> Js {
    v.stringify().map_or(Js::Undefined, Js::from)
}

/// `sdtPropsToAttrs` (src/types/sdtAttributes.ts:4).
pub(crate) fn sdt_props_to_attrs(props: &Js) -> Result<Js> {
    let list_items = props.get("listItems");
    let data_binding = props.get("dataBinding");
    Ok(obj! {
        "sdtType" => props.get("sdtType"),
        "id" => or_null(props.get("id")),
        "alias" => or_null(props.get("alias")),
        "tag" => or_null(props.get("tag")),
        "lock" => or_null(props.get("lock")),
        "placeholder" => or_null(props.get("placeholder")),
        "showingPlaceholder" => props.get("showingPlaceholder").nullish_or(Js::Bool(false)),
        "dateFormat" => or_null(props.get("dateFormat")),
        "listItems" => if list_items.truthy() { json(&list_items) } else { Js::Null },
        "checked" => or_null(props.get("checked")),
        "dataBinding" => if data_binding.truthy() { json(&data_binding) } else { Js::Null },
        "multiLine" => or_null(props.get("multiLine")),
        "rawPropertiesXml" => or_null(props.get("rawPropertiesXml")),
        "rawEndPropertiesXml" => or_null(props.get("rawEndPropertiesXml")),
    })
}

/// `sdtAttrsToProps` (src/types/sdtAttributes.ts:24).
pub(crate) fn sdt_attrs_to_props(attrs: &Js) -> Result<Js> {
    let s = |k: &str| Js::from(crate::js::to_js_string(&attrs.get(k)));
    let props = obj! { "sdtType" => attrs.get("sdtType").nullish_or(Js::from("richText")) };
    if attrs.get("id").is_num() {
        props.set("id", attrs.get("id"));
    }
    if !attrs.get("alias").is_nullish() {
        props.set("alias", s("alias"));
    }
    if !attrs.get("tag").is_nullish() {
        props.set("tag", s("tag"));
    }
    if !attrs.get("lock").is_nullish() {
        props.set("lock", attrs.get("lock"));
    }
    if !attrs.get("placeholder").is_nullish() {
        props.set("placeholder", s("placeholder"));
    }
    if attrs.get("showingPlaceholder").truthy() {
        props.set("showingPlaceholder", true);
    }
    if !attrs.get("dateFormat").is_nullish() {
        props.set("dateFormat", s("dateFormat"));
    }
    if let Some(text) = attrs.get("listItems").as_str().filter(|t| !t.is_empty()) {
        if let Ok(v) = Js::parse(text) {
            props.set("listItems", v);
        }
    }
    if !attrs.get("checked").is_nullish() {
        props.set("checked", attrs.get("checked"));
    }
    if let Some(text) = attrs.get("dataBinding").as_str().filter(|t| !t.is_empty()) {
        if let Ok(v) = Js::parse(text) {
            props.set("dataBinding", v);
        }
    }
    if attrs.get("multiLine").is_bool() {
        props.set("multiLine", attrs.get("multiLine"));
    }
    if !attrs.get("rawPropertiesXml").is_nullish() {
        props.set("rawPropertiesXml", s("rawPropertiesXml"));
    }
    if !attrs.get("rawEndPropertiesXml").is_nullish() {
        props.set("rawEndPropertiesXml", s("rawEndPropertiesXml"));
    }
    Ok(props)
}
