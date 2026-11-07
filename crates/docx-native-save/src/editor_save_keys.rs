//! Port of `packages/docx/src/yrs/editorSaveKeys.ts`.
#![allow(unused_variables, unused_imports, dead_code, clippy::all)]

use std::collections::{HashMap, HashSet};
use std::rc::Rc;
use indexmap::IndexMap;
use crate::js::{Js, Result, SaveError};
use crate::session::{Inheritance, Session};
use crate::{arr, obj, throw};

/// `const { content, ...metadata } = part`.
fn without_content(part: &Js) -> Js {
    let metadata = part.spread();
    metadata.delete("content");
    metadata
}

fn parts(map: &Js) -> Js {
    Js::new_arr(
        map.map_entries()
            .into_iter()
            .map(|(id, part)| arr![id, without_content(&part)])
            .collect(),
    )
}

fn notes(entries: &Js) -> Js {
    Js::new_arr(
        entries
            .items()
            .iter()
            .map(|note| {
                let metadata = without_content(note);
                metadata.delete("verbatimXml");
                metadata.delete("sourceOrdinal");
                metadata
            })
            .collect(),
    )
}

/// `editorSaveKeys` (src/yrs/editorSaveKeys.ts:5). `comments` is `Js::Undefined` when absent.
pub(crate) fn editor_save_keys(document: &Js, comments: &Js) -> Result<Js> {
    let pkg = document.get("package");
    let sections = pkg.path(&["document", "sections"]);
    let metadata = obj! {
        "finalSectionProperties" => pkg.path(&["document", "finalSectionProperties"]),
        "sections" => if sections.is_nullish() {
            Js::Undefined
        } else {
            Js::new_arr(sections.items().iter().map(without_content).collect())
        },
        "headers" => parts(&pkg.get("headers")),
        "footers" => parts(&pkg.get("footers")),
        "footnotes" => notes(&pkg.get("footnotes")),
        "endnotes" => notes(&pkg.get("endnotes")),
        "relationships" => Js::new_arr(pkg.get("relationships").map_entries().into_iter().map(|(k, v)| arr![k, v]).collect()),
    };
    let saved = comments
        .clone()
        .nullish_or(pkg.path(&["document", "comments"]))
        .nullish_or(arr![]);
    Ok(obj! {
        "metadata" => stringify(&metadata)?,
        "comments" => stringify(&saved)?,
        "commentIds" => Js::new_arr(saved.items().iter().map(|c| c.get("id")).collect()),
    })
}

/// `noteSeparatorKeys` (src/yrs/editorSaveKeys.ts:39).
pub(crate) fn note_separator_keys(document: &Js) -> Result<Js> {
    Ok(obj! {
        "footnotes" => stringify(&document.path(&["package", "footnoteSeparators"]).nullish_or(arr![]))?,
        "endnotes" => stringify(&document.path(&["package", "endnoteSeparators"]).nullish_or(arr![]))?,
    })
}

/// `stringify` (src/yrs/editorSaveKeys.ts:46).
pub(crate) fn stringify(value: &Js) -> Result<String> {
    Ok(value.stringify_maps_as_entries().unwrap_or_default())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn editor_save_keys_match_ts() {
        let headers = Js::new_map();
        headers.map_set("rId1", obj! { "type" => "default", "content" => arr![1], "hdrShapes" => Js::new_map() });
        let rels = Js::new_map();
        rels.map_set("rId1", obj! { "target" => "header1.xml" });
        let doc = obj! { "package" => obj! {
            "document" => obj! {
                "content" => arr![],
                "finalSectionProperties" => obj! { "pageWidth" => 12240 },
                "comments" => Js::parse(r#"[{"id":3,"author":"A"},{"id":7}]"#).unwrap(),
            },
            "headers" => headers,
            "footnotes" => Js::parse(r#"[{"id":1,"content":[],"verbatimXml":"<x/>","sourceOrdinal":2,"type":"normal"}]"#).unwrap(),
            "relationships" => rels,
            "footnoteSeparators" => Js::parse(r#"[{"a":1}]"#).unwrap(),
        } };
        let keys = editor_save_keys(&doc, &Js::Undefined).unwrap();
        assert_eq!(
            keys.stringify().unwrap(),
            r#"{"metadata":"{\"finalSectionProperties\":{\"pageWidth\":12240},\"headers\":[[\"rId1\",{\"type\":\"default\",\"hdrShapes\":[]}]],\"footers\":[],\"footnotes\":[{\"id\":1,\"type\":\"normal\"}],\"endnotes\":[],\"relationships\":[[\"rId1\",{\"target\":\"header1.xml\"}]]}","comments":"[{\"id\":3,\"author\":\"A\"},{\"id\":7}]","commentIds":[3,7]}"#
        );
        assert_eq!(note_separator_keys(&doc).unwrap().stringify().unwrap(), r#"{"footnotes":"[{\"a\":1}]","endnotes":"[]"}"#);
    }
}
