//! Port of `packages/docx/src/yrs/residentCommentProjection.ts`.
#![allow(unused_variables, unused_imports, dead_code, clippy::all)]

use std::collections::{HashMap, HashSet};
use std::rc::Rc;
use indexmap::IndexMap;
use crate::js::{Js, Result, SaveError};
use crate::session::{Inheritance, Session};
use crate::{arr, obj, throw};

use crate::js::num_to_string;

/// `/^\d+$/u.test(s) ? Number(s) : NaN`.
fn digits_number(s: &str) -> f64 {
    if !s.is_empty() && s.bytes().all(|b| b.is_ascii_digit()) {
        crate::js::to_number(&Js::from(s))
    } else {
        f64::NAN
    }
}

fn is_safe_integer(n: f64) -> bool {
    n.is_finite() && n.fract() == 0.0 && n.abs() <= 9_007_199_254_740_991.0
}

/// A JS `Set<number>` (SameValueZero; insertion order kept for `Math.max(...)`).
#[derive(Default)]
struct NumSet(Vec<f64>);

impl NumSet {
    fn has(&self, n: f64) -> bool {
        self.0.iter().any(|v| *v == n || (v.is_nan() && n.is_nan()))
    }
    fn add(&mut self, n: f64) {
        if !self.has(n) {
            self.0.push(n);
        }
    }
}

/// `commentContent` (src/yrs/residentCommentProjection.ts:17).
pub(crate) fn comment_content(body: &Js, fallback: &Js) -> Result<Js> {
    if body.is_arr() && body.items().iter().all(|e| e.get("type").as_str() == Some("paragraph")) {
        return Ok(body.clone());
    }
    Ok(match body.as_str() {
        Some(text) => arr![obj! {
            "type" => "paragraph",
            "formatting" => obj! {},
            "content" => arr![obj! {
                "type" => "run",
                "formatting" => obj! {},
                "content" => arr![obj! { "type" => "text", "text" => text }],
            }],
        }],
        None => fallback.clone(),
    })
}

/// `savedCommentIds` (src/yrs/residentCommentProjection.ts:27). `previous` is a `Js::Map`
/// (or `Js::Undefined` for the default empty map); returns a `Js::Map` of id → number.
pub(crate) fn saved_comment_ids(source_ids: &[f64], native: &Js, previous: &Js) -> Result<Js> {
    let mut used = NumSet::default();
    for id in source_ids {
        used.add(*id);
    }
    for (_, v) in previous.map_entries() {
        used.add(crate::js::to_number(&v));
    }
    let keys: HashSet<String> = source_ids.iter().map(|n| num_to_string(*n)).collect();
    let ids = Js::new_map();
    for comment in native.items() {
        let id = crate::js::to_js_string(&comment.get("id"));
        if keys.contains(&id) {
            ids.map_set(&id, crate::js::to_number(&comment.get("id")));
        }
    }
    let mut next = used.0.iter().fold(0.0f64, |m, v| if m.is_nan() || v.is_nan() { f64::NAN } else { m.max(*v) }) + 1.0;
    for comment in native.items() {
        let cid = crate::js::to_js_string(&comment.get("id"));
        if ids.map_has(&cid) {
            continue;
        }
        let saved = previous.map_get(&cid);
        if !saved.is_undefined() {
            ids.map_set(&cid, saved);
            continue;
        }
        let mut id = digits_number(&cid);
        if !is_safe_integer(id) || used.has(id) {
            while used.has(next) {
                next += 1.0;
            }
            id = next;
        }
        used.add(id);
        ids.map_set(&cid, id);
    }
    Ok(ids)
}

/// `projectResidentComments` (src/yrs/residentCommentProjection.ts:55).
///
/// `engine` stands for the resident engine session as a snapshot object:
/// `engine.propertyDocument` is what `engine.propertyDocument()` returns and
/// `engine.listComments` what `engine.listComments()` returns. `aliases` is a `Js::Map`
/// (mutated in place, as in TS).
pub(crate) fn project_resident_comments(engine: &Js, aliases: &Js, initialized: bool, fresh: bool) -> Result<Js> {
    let source = engine.path(&["propertyDocument", "package", "document", "comments"]).nullish_or(arr![]).items();
    let native = engine.get("listComments").items();
    let id_of = |c: &Js| crate::js::to_js_string(&c.get("id"));
    for comment in &native {
        let id = digits_number(&id_of(comment));
        if is_safe_integer(id) {
            aliases.map_set(&id_of(comment), id);
        }
    }
    if !fresh {
        let mut next = 1.0f64;
        let mut all: Vec<f64> = source.iter().map(|c| crate::js::to_number(&c.get("id"))).collect();
        all.extend(aliases.map_entries().iter().map(|(_, v)| crate::js::to_number(v)));
        for id in all {
            next = if next.is_nan() || id.is_nan() { f64::NAN } else { next.max(id + 1.0) };
        }
        for comment in &native {
            if !aliases.map_has(&id_of(comment)) {
                aliases.map_set(&id_of(comment), next);
                next += 1.0;
            }
        }
    }
    let numbers = if fresh {
        let source_ids: Vec<f64> = source.iter().map(|c| crate::js::to_number(&c.get("id"))).collect();
        saved_comment_ids(&source_ids, &Js::new_arr(native.clone()), &Js::Undefined)?
    } else {
        aliases.clone()
    };
    let ids = Js::new_obj();
    for comment in &native {
        ids.set(&id_of(comment), numbers.map_get(&id_of(comment)));
    }
    let mut held: IndexMap<String, Js> = IndexMap::new();
    for comment in &native {
        held.insert(crate::js::to_js_string(&numbers.map_get(&id_of(comment))), comment.clone());
    }
    let parent_of = |comment: &Js| -> Js {
        let parent = comment.get("parentId");
        if parent.is_null() {
            return Js::Undefined;
        }
        let p = crate::js::to_js_string(&parent);
        numbers.map_get(&p).nullish_or({
            let n = digits_number(&p);
            if p.bytes().all(|b| b.is_ascii_digit()) && !p.is_empty() { Js::Num(n) } else { Js::Undefined }
        })
    };
    let comments = Js::new_arr(Vec::new());
    for original in &source {
        let Some(comment) = held.get(&crate::js::to_js_string(&original.get("id"))).cloned() else {
            if !aliases.map_has(&crate::js::to_js_string(&original.get("id"))) {
                comments.push(original.clone());
            }
            continue;
        };
        let parent_id = parent_of(&comment);
        let original_done = original.get("done").nullish_or(Js::Bool(false));
        let done = if initialized { comment.get("done") } else { comment.get("done").or(original_done.clone()) };
        let out = original.spread();
        if comment.get("author").truthy() {
            out.set("author", comment.get("author"));
        }
        if comment.get("date").truthy() {
            out.set("date", comment.get("date"));
        }
        if !comment.get("body").is_nullish() {
            out.set("content", comment_content(&comment.get("body"), &original.get("content"))?);
        }
        if !done.strict_eq(&original_done) {
            out.set("done", done.clone());
            if !original.get("status").is_undefined() {
                out.set("status", if done.truthy() { "resolved" } else { "active" });
            }
        }
        if !parent_id.is_undefined() {
            out.set("parentId", parent_id);
        }
        comments.push(out);
    }
    let mut source_ids = NumSet::default();
    for c in &source {
        source_ids.add(crate::js::to_number(&c.get("id")));
    }
    for comment in &native {
        let id = numbers.map_get(&id_of(comment));
        let is_source = match &id {
            Js::Num(n) => source_ids.has(*n),
            _ => false,
        };
        if !is_source {
            let date = comment.get("date");
            comments.push(obj! {
                "id" => id,
                "author" => comment.get("author"),
                "date" => if date.truthy() { date } else { Js::Undefined },
                "content" => comment_content(&comment.get("body"), &arr![])?,
                "done" => comment.get("done"),
                "parentId" => parent_of(comment),
            });
        }
    }
    let session_ids = Js::new_map();
    for comment in &native {
        session_ids.map_set(&crate::js::to_js_string(&numbers.map_get(&id_of(comment))), comment.get("id"));
    }
    Ok(obj! { "comments" => comments, "ids" => ids, "sessionIds" => session_ids })
}

/// `refuse` (src/yrs/residentCommentProjection.ts:99).
pub(crate) fn refuse(reason: &str) -> Result<Js> {
    Ok(obj! { "status" => "refused", "reason" => reason })
}

/// `readCommentInsertionRange` (src/yrs/residentCommentProjection.ts:101).
///
/// Needs the live engine's `geometryReader` and repeated `selectionText` queries, which a
/// `Js` value cannot answer; it is not on the save path. `engine` must carry
/// `paragraphSpans: { [story]: [{ paraId, length }] }` and `paragraphText: { [paraId]: [unit] }`
/// where `paragraphText[paraId][k]` is `selectionText(offset k).before.length` and
/// `paragraphText[paraId + ":after"]` the text after offset 0.
pub(crate) fn read_comment_insertion_range(engine: &Js, para_id: &str, search: Option<&str>) -> Result<Js> {
    let spans = engine.get("paragraphSpans");
    let mut matches: Vec<(String, Js)> = Vec::new();
    for story in spans.keys() {
        for paragraph in spans.get(&story).items() {
            if paragraph.get("paraId").as_str() == Some(para_id) {
                matches.push((story.clone(), paragraph));
            }
        }
    }
    if matches.len() != 1 {
        return refuse(if matches.is_empty() { "missing-target" } else { "ambiguous-target" });
    }
    if search == Some("") {
        return refuse("invalid-step");
    }
    let (story, paragraph) = matches.remove(0);
    let length = crate::js::to_number(&paragraph.get("length"));
    let at = obj! { "story" => story.as_str(), "paraId" => para_id, "offset" => 0 };
    let Some(search) = search else {
        return Ok(obj! { "story" => story.as_str(), "start" => at.clone(), "end" => at.with(&[("offset", Js::Num(length))]) });
    };
    let text = crate::js::to_js_string(&engine.path(&["paragraphText", &format!("{para_id}:after")]));
    let Some(start) = crate::js::index_of16(&text, search, 0) else { return refuse("missing-target") };
    if crate::js::index_of16(&text, search, start + 1).is_some() {
        return refuse("ambiguous-target");
    }
    let before = engine.path(&["paragraphText", para_id]);
    let native_offset = |offset: f64| -> f64 {
        let (mut low, mut high) = (0.0f64, length);
        while low < high {
            let middle = ((low + high) / 2.0).floor();
            if crate::js::to_number(&before.at(middle as usize)) < offset {
                low = middle + 1.0;
            } else {
                high = middle;
            }
        }
        low
    };
    let s = native_offset(start as f64 + 1.0) - 1.0;
    let e = native_offset(start as f64 + crate::js::len16(search) as f64);
    Ok(obj! { "story" => story.as_str(), "start" => at.with(&[("offset", Js::Num(s))]), "end" => at.with(&[("offset", Js::Num(e))]) })
}

#[cfg(test)]
mod tests {
    use super::*;

    fn entries(m: &Js) -> String {
        m.map_entries().iter().map(|(k, v)| format!("{k}={}", crate::js::to_js_string(v))).collect::<Vec<_>>().join(",")
    }

    #[test]
    fn saved_comment_ids_match_ts() {
        let n = |s: &str| Js::parse(s).unwrap();
        assert_eq!(entries(&saved_comment_ids(&[], &n(r#"[{"id":"5"},{"id":"9201:0"}]"#), &Js::Undefined).unwrap()), "5=5,9201:0=1");
        assert_eq!(
            entries(&saved_comment_ids(&[1.0], &n(r#"[{"id":"9201:0"},{"id":"5"},{"id":"1"},{"id":"01"}]"#), &Js::Undefined).unwrap()),
            "1=1,9201:0=2,5=5,01=3"
        );
    }
}
