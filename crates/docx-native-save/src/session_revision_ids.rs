//! Port of `packages/docx/src/docx/sessionRevisionIds.ts`.

use std::cell::Cell;
use std::collections::{HashMap, HashSet};

use crate::js::{self, Js, Marks, Result, SaveError};

const MAX_ID: f64 = 2147483648.0;
const COMMENT_MARKERS: [&str; 3] = ["commentRangeStart", "commentRangeEnd", "commentReference"];
const ANNOTATIONS: [&str; 9] = [
    "commentRangeStart",
    "commentRangeEnd",
    "commentReference",
    "bookmarkStart",
    "bookmarkEnd",
    "moveFromRangeStart",
    "moveFromRangeEnd",
    "moveToRangeStart",
    "moveToRangeEnd",
];

thread_local! {
    static SESSION_INFOS: Marks<()> = Marks::new();
    static MARKED: Cell<bool> = const { Cell::new(false) };
}

/// `markRevisionInfo` (src/docx/sessionRevisionIds.ts:20).
pub(crate) fn mark_revision_info(info: &Js, session: bool) -> Result<Js> {
    if session {
        SESSION_INFOS.with(|s| s.set(info, ()));
    }
    MARKED.with(|m| m.set(true));
    Ok(info.clone())
}

/// JS `\s`.
fn is_js_space(c: char) -> bool {
    matches!(
        c,
        '\t' | '\n'
            | '\u{b}'
            | '\u{c}'
            | '\r'
            | ' '
            | '\u{a0}'
            | '\u{1680}'
            | '\u{2028}'
            | '\u{2029}'
            | '\u{202f}'
            | '\u{205f}'
            | '\u{3000}'
            | '\u{feff}'
    ) || ('\u{2000}'..='\u{200a}').contains(&c)
}

fn is_word(c: char) -> bool {
    c.is_ascii_alphanumeric() || c == '_'
}

/// The ids of `/\bw:id\s*=\s*["'](\d+)["']/g` in `value`.
fn xml_ids(value: &str) -> Vec<String> {
    let mut ids = Vec::new();
    let mut from = 0;
    while let Some(found) = value[from..].find("w:id") {
        let start = from + found;
        from = start + 1;
        if value[..start].chars().next_back().is_some_and(is_word) {
            continue;
        }
        let mut rest = value[start + 4..].chars().peekable();
        while rest.peek().is_some_and(|c| is_js_space(*c)) {
            rest.next();
        }
        if rest.next() != Some('=') {
            continue;
        }
        while rest.peek().is_some_and(|c| is_js_space(*c)) {
            rest.next();
        }
        if !matches!(rest.next(), Some('"' | '\'')) {
            continue;
        }
        let mut digits = String::new();
        while rest.peek().is_some_and(char::is_ascii_digit) {
            digits.push(rest.next().unwrap());
        }
        if digits.is_empty() || !matches!(rest.next(), Some('"' | '\'')) {
            continue;
        }
        // The match ends past the closing quote; the next search starts there.
        from = start + 4 + value[start + 4..].len() - rest.collect::<String>().len();
        ids.push(digits);
    }
    ids
}

struct Site {
    holder: Js,
    key: SiteKey,
    info: Js,
}

enum SiteKey {
    Prop(String),
    Index(usize),
}

/// A `Map` key with SameValueZero semantics for the primitives an id holds.
fn id_key(value: &Js) -> String {
    match value.ptr() {
        Some(p) => format!("object:{p}"),
        None => match value {
            Js::Num(n) if *n == 0.0 => "number:0".into(),
            other => format!("{}:{}", other.type_of(), js::to_js_string(other)),
        },
    }
}

/// `withSessionRevisionIds` (src/docx/sessionRevisionIds.ts:39).
pub(crate) fn with_session_revision_ids(tree: &Js, write: &dyn Fn() -> Result<Js>) -> Result<Js> {
    if !MARKED.with(Cell::get) {
        return write();
    }
    let mut sites: Vec<Site> = Vec::new();
    let mut max = 0.0f64;
    let mut used: HashSet<u64> = HashSet::new();
    fn note(id: &Js, raises_base: bool, used: &mut HashSet<u64>, max: &mut f64) {
        let Js::Num(n) = id else { return };
        if !js::is_integer(id) || *n < 0.0 || *n >= MAX_ID {
            return;
        }
        used.insert(*n as u64);
        if raises_base && *n > *max {
            *max = *n;
        }
    }
    fn visit(
        value: &Js,
        holder: &Js,
        key: SiteKey,
        sites: &mut Vec<Site>,
        used: &mut HashSet<u64>,
        max: &mut f64,
    ) {
        if let Js::Str(s) = value {
            if s.contains("w:id") {
                for id in xml_ids(s) {
                    note(&Js::Num(js::to_number(&Js::from(id))), true, used, max);
                }
            }
            return;
        }
        if !matches!(value, Js::Obj(_) | Js::Arr(_) | Js::Map(_)) {
            return;
        }
        if SESSION_INFOS.with(|s| s.has(value)) {
            sites.push(Site {
                holder: holder.clone(),
                key,
                info: value.clone(),
            });
            return;
        }
        if value.is_arr() {
            for (index, entry) in value.items().iter().enumerate() {
                visit(entry, value, SiteKey::Index(index), sites, used, max);
            }
            return;
        }
        if value.is_map() {
            return;
        }
        if value.get("author").is_str() {
            let comment = value.get("content").is_arr();
            note(&value.get("id"), !comment, used, max);
            note(&value.get("revisionId"), !comment, used, max);
        } else if let Some(t) = value.get("type").as_str()
            && ANNOTATIONS.contains(&t) {
                note(&value.get("id"), !COMMENT_MARKERS.contains(&t), used, max);
            }
        for child in value.keys() {
            visit(
                &value.get(&child),
                value,
                SiteKey::Prop(child),
                sites,
                used,
                max,
            );
        }
    }
    visit(
        tree,
        &Js::new_obj(),
        SiteKey::Prop(String::new()),
        &mut sites,
        &mut used,
        &mut max,
    );
    if sites.is_empty() {
        return write();
    }
    let mut next = max;
    let mut ids: HashMap<String, f64> = HashMap::new();
    for site in &sites {
        let key = id_key(&site.info.get("id"));
        if ids.contains_key(&key) {
            continue;
        }
        loop {
            next = if next + 1.0 < MAX_ID { next + 1.0 } else { 1.0 };
            if next == max {
                return Err(SaveError::new("no Word revision ids left in range"));
            }
            if !used.contains(&(next as u64)) {
                break;
            }
        }
        ids.insert(key, next);
    }
    let put = |site: &Site, value: Js| match &site.key {
        SiteKey::Prop(k) => site.holder.set(k, value),
        SiteKey::Index(i) => site.holder.set_at(*i, value),
    };
    for site in &sites {
        let id = ids.get(&id_key(&site.info.get("id"))).copied();
        put(site, site.info.with(&[("id", Js::from(id))]));
    }
    let result = write();
    for site in &sites {
        put(site, site.info.clone());
    }
    result
}

#[cfg(test)]
pub(crate) fn reset_marks_for_test() {
    MARKED.with(|m| m.set(false));
    SESSION_INFOS.with(Marks::clear);
}

/// Drops the `SESSION_INFOS` entries of objects that no longer exist.
pub(crate) fn prune_session_infos() {
    SESSION_INFOS.with(|marks| marks.clear());
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{arr, obj};

    #[test]
    fn numbers_session_revisions_above_annotation_ids() {
        // bun: the same tree through withSessionRevisionIds gives
        // {"a":[{"type":"insertion","info":{"id":-1,"author":"me"}},{"type":"bookmarkStart","id":7},{"type":"commentRangeStart","id":40},{"type":"deletion","info":{"id":-1,"author":"me"}},{"type":"insertion","info":{"id":-2,"author":"me"}}],"xml":"<w:ins w:id=\"9\"/>"}
        // => written with ids 10, 10, 11 (9 from the XML raises the base; 40 is a comment marker).
        reset_marks_for_test();
        let i1 = mark_revision_info(&obj! { "id" => -1.0, "author" => "me" }, true).unwrap();
        let i2 = mark_revision_info(&obj! { "id" => -2.0, "author" => "me" }, true).unwrap();
        let tree = obj! {
            "a" => arr![
                obj! { "type" => "insertion", "info" => i1.clone() },
                obj! { "type" => "bookmarkStart", "id" => 7.0 },
                obj! { "type" => "commentRangeStart", "id" => 40.0 },
                obj! { "type" => "deletion", "info" => i1.clone() },
                obj! { "type" => "insertion", "info" => i2.clone() },
            ],
            "xml" => "<w:ins w:id=\"9\"/>",
        };
        let t = tree.clone();
        let out =
            with_session_revision_ids(&tree, &|| Ok(Js::from(t.stringify().unwrap()))).unwrap();
        assert_eq!(
            out.as_str().unwrap(),
            r#"{"a":[{"type":"insertion","info":{"id":10,"author":"me"}},{"type":"bookmarkStart","id":7},{"type":"commentRangeStart","id":40},{"type":"deletion","info":{"id":10,"author":"me"}},{"type":"insertion","info":{"id":11,"author":"me"}}],"xml":"<w:ins w:id=\"9\"/>"}"#
        );
        assert!(tree.get("a").at(0).get("info").strict_eq(&i1));
        reset_marks_for_test();
    }
}
