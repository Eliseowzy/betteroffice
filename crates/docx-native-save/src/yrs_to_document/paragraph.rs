//! Port of `packages/docx/src/yrs/yrsToDocument.ts` (part).
#![allow(unused_variables, unused_imports, dead_code, clippy::all)]

use std::collections::{HashMap, HashSet};
use std::rc::Rc;
use indexmap::IndexMap;
use crate::js::{self, Js, Result, SaveError};
use crate::session::{Inheritance, Session};
use crate::{arr, obj, throw};
use super::*;
use crate::raw_xml::is_raw_xml;
use crate::save_formatting::{paragraph_attrs_to_formatting, table_attrs_to_formatting, table_cell_attrs_to_formatting, table_row_attrs_to_formatting};
use crate::story_plain_text::table_row_grid;
use crate::table_row_spans::seeded_cell_indices;

// ----- private copies of constants declared in part A (lines 1–1454) -----

const PARAGRAPH_ATTR_DEFAULTS_KEYS_B: &[&str] = &[
    "paraId", "textId", "alignment", "spaceBefore", "spaceAfter", "lineSpacing", "lineSpacingRule",
    "spacingExplicit", "indentLeft", "indentRight", "indentFirstLine", "hangingIndent", "numPr",
    "numPrFromStyle", "listNumFmt", "listIsBullet", "listMarker", "listMarkerHidden",
    "listMarkerFontFamily", "listMarkerFontSize", "listMarkerSuffix", "listLevelNumFmts",
    "listAbstractNumId", "listStartOverride", "styleId", "borders", "shading", "tabs",
    "pageBreakBefore", "renderedPageBreakBefore", "keepNext", "keepLines", "widowControl",
    "contextualSpacing", "snapToGrid", "autoSpaceDE", "autoSpaceDN", "adjustRightInd", "cnfStyle",
    "defaultTextFormatting", "sectionBreakType", "bidi", "outlineLevel", "bookmarks",
    "_originalFormatting", "_originalRunBoundaries", "_sectionProperties", "pPrIns", "pPrDel",
    "pPrChange",
];

fn paragraph_attr_defaults_b() -> Js {
    let o = Js::new_obj();
    for key in PARAGRAPH_ATTR_DEFAULTS_KEYS_B {
        // `hangingIndent: false`, every other default null.
        o.set(key, if *key == "hangingIndent" { Js::Bool(false) } else { Js::Null });
    }
    o
}

fn table_attr_defaults_b() -> Js {
    let o = Js::new_obj();
    for key in [
        "styleId", "width", "widthType", "justification", "columnWidths", "tableLayout", "floating",
        "cellMargins", "look", "bidi", "_originalFormatting", "tblPrChange",
    ] {
        o.set(key, Js::Null);
    }
    o
}

fn table_row_attr_defaults_b() -> Js {
    obj! {
        "height" => Js::Null,
        "heightRule" => Js::Null,
        "isHeader" => false,
        "_originalFormatting" => Js::Null,
        "trIns" => Js::Null,
        "trDel" => Js::Null,
        "trPrChange" => Js::Null,
    }
}

const BOOLEAN_MARKS_B: &[&str] = &[
    "bold", "italic", "superscript", "subscript", "allCaps", "smallCaps", "emboss", "imprint",
    "textShadow", "textOutline", "hidden", "rtl",
];

const SOURCE_RUN_PROPERTIES_B: &[&str] =
    &["bold", "italic", "fontSize", "boldCs", "italicCs", "fontSizeCs", "language", "fontFamily"];

fn non_formatting_attrs_b() -> HashSet<String> {
    ["hyperlink", "ins", "del"].iter().map(|s| s.to_string()).collect()
}

const MAX_SAFE_INTEGER_B: f64 = 9007199254740991.0;

// ----- module state owned by this part -----

thread_local! {
    /// `projectedTableCells` (src/yrs/yrsToDocument.ts:2650).
    static PROJECTED_TABLE_CELLS: js::Marks<Js> = js::Marks::new();
}

/// `projectedTableCells.get(table)`, for part C.
pub(crate) fn projected_table_cells_get(table: &Js) -> Option<Js> {
    PROJECTED_TABLE_CELLS.with(|m| m.get(table))
}

/// `projectedTableCells.set(table, cells)`, for part C.
pub(crate) fn projected_table_cells_set(table: &Js, cells: Js) {
    PROJECTED_TABLE_CELLS.with(|m| m.set(table, cells))
}


// ----- small JS helpers -----

fn num(value: &Js) -> f64 {
    js::to_number(value)
}

fn js_min(a: f64, b: f64) -> f64 {
    if a.is_nan() || b.is_nan() { f64::NAN } else { a.min(b) }
}

fn js_max(a: f64, b: f64) -> f64 {
    if a.is_nan() || b.is_nan() { f64::NAN } else { a.max(b) }
}

fn ty(value: &Js) -> String {
    value.get("type").str_or_empty().to_string()
}

fn type_is(value: &Js, t: &str) -> bool {
    value.get("type").as_str() == Some(t)
}

/// `arr[i]` for a numeric index (undefined for negative, fractional or out of range).
fn at_f(array: &Js, index: f64) -> Js {
    if index >= 0.0 && index.fract() == 0.0 { array.at(index as usize) } else { Js::Undefined }
}

fn str_items(value: &Js) -> Vec<String> {
    value.items().iter().map(js::to_js_string).collect()
}

fn concat_arr(target: &mut Vec<Js>, source: &Js) {
    target.extend(source.items());
}

fn includes_str(array: &Js, needle: &str) -> bool {
    array.items().iter().any(|v| v.as_str() == Some(needle))
}

fn comment_marker(boundary: &Js) -> Result<Js> {
    Ok(if boundary.get("kind").as_str() == Some("start") {
        obj! { "type" => "commentRangeStart", "id" => boundary.get("id") }
    } else {
        obj! { "type" => "commentRangeEnd", "id" => boundary.get("id") }
    })
}

/// A property read that throws as JS does on `null`/`undefined`.
fn prop(value: &Js, key: &str) -> Result<Js> {
    if value.is_nullish() {
        throw!("Cannot read properties of {} (reading '{}')", js::to_js_string(value), key);
    }
    Ok(value.get(key))
}

/// `/^(.*:t\d+):r(\d+)c(\d+)(~\d+)?$/` with the tilde suffix allowed or not.
fn match_cell_story(story: &str, allow_tilde: bool) -> Option<(String, String, String)> {
    if story.chars().any(|c| matches!(c, '\n' | '\r' | '\u{2028}' | '\u{2029}')) {
        return None;
    }
    let colon = story.rfind(':')?;
    let (head, tail) = (&story[..colon], &story[colon + 1..]);
    let tail = tail.strip_prefix('r')?;
    let row_len = tail.bytes().take_while(u8::is_ascii_digit).count();
    if row_len == 0 {
        return None;
    }
    let row = &tail[..row_len];
    let tail = tail[row_len..].strip_prefix('c')?;
    let col_len = tail.bytes().take_while(u8::is_ascii_digit).count();
    if col_len == 0 {
        return None;
    }
    let col = &tail[..col_len];
    let rest = &tail[col_len..];
    if !rest.is_empty() {
        if !allow_tilde {
            return None;
        }
        let digits = rest.strip_prefix('~')?;
        if digits.is_empty() || !digits.bytes().all(|b| b.is_ascii_digit()) {
            return None;
        }
    }
    let trimmed = head.trim_end_matches(|c: char| c.is_ascii_digit());
    if trimmed.len() == head.len() || !trimmed.ends_with(":t") {
        return None;
    }
    Some((head.to_string(), row.to_string(), col.to_string()))
}

// ----- ported functions -----

/// `buildParagraphContent` (src/yrs/yrsToDocument.ts:1455).
pub(crate) fn build_paragraph_content(items: &Js, recovery: &Js, format: &dyn Fn(&Js) -> Result<Js>, revision_ids: &Js, in_control: bool) -> Result<Js> {
    let items = restore_projected_field_results(items, recovery, format)?;
    let content = Js::new_arr(Vec::new());
    let mut current_run: Option<Js> = None;
    let mut current_formatting_key: Option<String> = None;
    let mut current_hyperlink: Option<Js> = None;
    let non_formatting = non_formatting_attrs_b();

    fn flush_run(content: &Js, run: &mut Option<Js>, key: &mut Option<String>) {
        if let Some(r) = run.take() {
            content.push(r);
        }
        *key = None;
    }
    fn flush_hyperlink(content: &Js, link: &mut Option<Js>) {
        if let Some(l) = link.take() {
            content.push(l);
        }
    }

    for item in items.items() {
        let attributes = item.get("attributes");
        if item.get("kind").as_str() == Some("embed") && item.get("embedKind").as_str() == Some("noteRef") {
            flush_run(&content, &mut current_run, &mut current_formatting_key);
            flush_hyperlink(&content, &mut current_hyperlink);
            let note = ordinary_content_for_item(&item, recovery, format)?;
            if note.truthy() {
                content.push(note);
            }
            continue;
        }

        let raw = attributes.get("ins").nullish_or(attributes.get("del"));
        let kind = if attributes.get("ins").truthy() { "insertion" } else { "deletion" };
        let mapped = mapped_revision_id(&raw, &Js::from(kind), revision_ids)?;
        let revision = tracked_info(&raw, false, mapped)?;
        if revision.truthy() {
            flush_run(&content, &mut current_run, &mut current_formatting_key);
            flush_hyperlink(&content, &mut current_hyperlink);
            let tracked = tracked_content_for_item(&item, &revision, in_control, recovery, format)?;
            if !tracked.truthy() {
                continue;
            }
            let previous = content.last();
            if is_tracked_wrapper(&previous)?
                && is_tracked_wrapper(&tracked)?
                && previous.get("type").strict_eq(&tracked.get("type"))
                && previous.path(&["info", "id"]).strict_eq(&tracked.path(&["info", "id"]))
                && previous.path(&["info", "author"]).strict_eq(&tracked.path(&["info", "author"]))
                && previous.path(&["info", "date"]).strict_eq(&tracked.path(&["info", "date"]))
            {
                let target = previous.get("content");
                for child in tracked.get("content").items() {
                    target.push(child);
                }
            } else {
                content.push(tracked);
            }
            continue;
        }

        let link_key = hyperlink_key(&attributes)?;
        if let Some(link_key) = link_key {
            flush_run(&content, &mut current_run, &mut current_formatting_key);
            let current_key = current_hyperlink.as_ref().map(|link| {
                let href = link.get("href");
                if href.truthy() {
                    href
                } else if link.get("anchor").truthy() {
                    Js::from(format!("#{}", js::to_js_string(&link.get("anchor"))))
                } else {
                    Js::from("")
                }
            });
            let history = as_object(&attributes.get("hyperlink"))?.get("history");
            let link_history = if history.is_bool() { history } else { Js::Undefined };
            let differs = match (&current_hyperlink, &current_key) {
                (Some(link), Some(key)) => {
                    !key.strict_eq(&Js::from(link_key.as_str())) || !link.get("history").strict_eq(&link_history)
                }
                _ => true,
            };
            if differs {
                flush_hyperlink(&content, &mut current_hyperlink);
                let created = create_hyperlink(&attributes)?;
                current_hyperlink = if created.truthy() { Some(created) } else { None };
            }
            if let Some(link) = &current_hyperlink {
                add_to_hyperlink(link, &item, recovery, format)?;
            }
            continue;
        }

        flush_hyperlink(&content, &mut current_hyperlink);
        if item.get("kind").as_str() == Some("text") {
            let key = stable_stringify_without(&attributes, &non_formatting)?;
            let text = js::to_js_string(&item.get("text"));
            let next_run = create_text_run(&text, &attributes, format)?;
            match (&current_run, &current_formatting_key) {
                (Some(run), Some(k)) if *k == key => append_text_run(run, &next_run)?,
                _ => {
                    flush_run(&content, &mut current_run, &mut current_formatting_key);
                    current_run = Some(next_run);
                    current_formatting_key = Some(key);
                }
            }
            continue;
        }

        flush_run(&content, &mut current_run, &mut current_formatting_key);
        let child = ordinary_content_for_item(&item, recovery, format)?;
        if child.truthy() {
            content.push(child);
        }
    }

    flush_run(&content, &mut current_run, &mut current_formatting_key);
    flush_hyperlink(&content, &mut current_hyperlink);
    Ok(content)
}

/// `marksKeyToYrsAttrs` (src/yrs/yrsToDocument.ts:1554).
pub(crate) fn marks_key_to_yrs_attrs(marks_key: Option<&str>) -> Result<Js> {
    let marks_key = match marks_key {
        Some(k) if !k.is_empty() => k,
        _ => return Ok(Js::new_obj()),
    };
    let attrs = Js::new_obj();
    for part in marks_key.split('|') {
        let colon = match part.find(':') {
            Some(c) if c > 0 => c,
            _ => return Ok(Js::Null),
        };
        let name = &part[..colon];
        let value = match Js::parse(&part[colon + 1..]) {
            Ok(v) => v,
            Err(_) => return Ok(Js::Null),
        };
        if name == "comment" || name == "footnoteRef" {
            continue;
        }
        if BOOLEAN_MARKS_B.contains(&name) {
            attrs.set(name, true);
        } else if name == "boldCs" || name == "italicCs" {
            attrs.set(name, prop(&value, "value")?);
        } else if name == "highlight" {
            attrs.set("highlight", prop(&value, "color")?);
        } else if name == "rPrChange" {
            attrs.set("rPrChange", prop(&value, "changes")?);
        } else if name == "insertion" || name == "deletion" {
            let info = obj! {
                "id" => prop(&value, "revisionId")?,
                "author" => value.get("author"),
                "date" => value.get("date"),
                "isMovePair" => value.get("isMovePair"),
            };
            attrs.set(if name == "insertion" { "ins" } else { "del" }, drop_nulls(&info)?);
        } else {
            attrs.set(name, drop_nulls(&value)?);
        }
    }
    Ok(attrs)
}

/// `noteMarkContent` (src/yrs/yrsToDocument.ts:1587).
pub(crate) fn note_mark_content(boundary: &Js) -> Result<Js> {
    let marks = boundary.get("noteMarks");
    if marks.len() == 0 {
        return Ok(Js::Null);
    }
    Ok(Js::new_arr(
        marks
            .items()
            .iter()
            .map(|mark| {
                obj! { "type" => if mark.as_str() == Some("endnote") { "endnoteRefMark" } else { "footnoteRefMark" } }
            })
            .collect(),
    ))
}

/// `boundaryContent` (src/yrs/yrsToDocument.ts:1598).
pub(crate) fn boundary_content(boundary: &Js, formatting: &Js) -> Result<Js> {
    let notes = note_mark_content(boundary)?;
    if notes.truthy() {
        return Ok(notes);
    }
    let text = js::to_js_string(&boundary.get("text"));
    let breaks = boundary.get("breaks");
    if breaks.len() == 0 {
        return run_content_for_text(&text, formatting);
    }
    let length = js::len16(&text) as f64;
    let mut content = Vec::new();
    let mut cursor = 0.0;
    for entry in breaks.items() {
        let at = js_min(js_max(num(&entry.get("offset")), cursor), length);
        if at > cursor {
            concat_arr(&mut content, &run_content_for_text(&js::slice16(&text, cursor, Some(at)), formatting)?);
        }
        content.push(obj! { "type" => "break", "breakType" => entry.get("type") });
        cursor = at;
    }
    if cursor < length {
        concat_arr(&mut content, &run_content_for_text(&js::slice16(&text, cursor, None), formatting)?);
    }
    Ok(Js::new_arr(content))
}

/// `sourceBreakBoundaries` (src/yrs/yrsToDocument.ts:1618).
pub(crate) fn source_break_boundaries(paragraph: &Js) -> Result<Js> {
    if paragraph.truthy() {
        let recorded = projected_block_source_runs(paragraph);
        if recorded.truthy() {
            return Ok(recorded);
        }
    }
    let mut boundaries = Vec::new();
    for child in paragraph.get("content").items() {
        if !type_is(&child, "run") {
            let t = ty(&child);
            if matches!(
                t.as_str(),
                "commentRangeStart" | "commentRangeEnd" | "bookmarkStart" | "bookmarkEnd" | "rawXml"
            ) {
                continue;
            }
            return Ok(Js::new_arr(Vec::new()));
        }
        let mut text = String::new();
        let mut breaks = Vec::new();
        for entry in child.get("content").items() {
            let t = ty(&entry);
            if t == "text" {
                text.push_str(&js::to_js_string(&entry.get("text")));
            } else if t == "commentReference" {
                text.push('\u{FFFC}');
            } else if t == "break"
                && matches!(entry.get("breakType").as_str(), Some("page") | Some("column"))
            {
                breaks.push(obj! { "offset" => js::len16(&text) as f64, "type" => entry.get("breakType") });
            } else {
                return Ok(Js::new_arr(Vec::new()));
            }
        }
        boundaries.push(obj! { "text" => text, "breaks" => Js::new_arr(breaks) });
    }
    Ok(Js::new_arr(boundaries))
}

fn boundary_sources(boundaries: &Js) -> Result<Js> {
    let mut source = Vec::new();
    for boundary in boundaries.items() {
        source.push(obj! { "type" => "run", "content" => boundary_content(&boundary, &Js::new_obj())? });
    }
    Ok(Js::new_arr(source))
}

/// `liveRunBoundaries` (src/yrs/yrsToDocument.ts:1644).
pub(crate) fn live_run_boundaries(boundaries: &Js, live: &Js) -> Result<Js> {
    let mut count = 0usize;
    for boundary in boundaries.items() {
        count += if note_mark_content(&boundary)?.truthy() { 0 } else { boundary.get("breaks").len() };
    }
    if count == 0 {
        return Ok(boundaries.clone());
    }
    let source = boundary_sources(boundaries)?;
    let hoisted = count as i64
        - trailing_block_breaks(&obj! { "type" => "paragraph", "content" => source.clone() })?.len() as i64;
    let visible = includes_str(&flow_tokens(&source, &Js::new_arr(Vec::new()))?, "visible");
    let leading_live = live.get("leading");
    let trailing_live = live.get("trailing");
    let mut index = 0i64;
    let mut leading = 0usize;
    let mut trailing = 0usize;
    let mut out = Vec::new();
    for boundary in boundaries.items() {
        if boundary.get("breaks").len() == 0 || note_mark_content(&boundary)?.truthy() {
            out.push(boundary);
            continue;
        }
        let all = boundary.get("breaks").items();
        let mut kept = Vec::new();
        for entry in &all {
            let before = index < hoisted;
            index += 1;
            let keep = if before && visible && entry.get("type").as_str() == Some("page") {
                live.get("leadingPage").truthy()
            } else if before {
                if !entry.get("type").strict_eq(&leading_live.at(leading)) {
                    false
                } else {
                    leading += 1;
                    true
                }
            } else if !entry.get("type").strict_eq(&trailing_live.at(trailing)) {
                false
            } else {
                trailing += 1;
                true
            };
            if keep {
                kept.push(entry.clone());
            }
        }
        out.push(if kept.len() == all.len() {
            boundary
        } else {
            boundary.with(&[("breaks", Js::new_arr(kept))])
        });
    }
    Ok(Js::new_arr(out))
}

/// `paragraphBreaksAreLive` (src/yrs/yrsToDocument.ts:1677).
pub(crate) fn paragraph_breaks_are_live(content: &Js, live: &Js) -> Result<bool> {
    let paragraph = obj! { "type" => "paragraph", "content" => content.clone() };
    let tokens = flow_tokens(content, &Js::new_arr(Vec::new()))?;
    let first_ok = !(tokens.at(0).as_str() == Some("page") && includes_str(&tokens, "visible"))
        || live.get("leadingPage").truthy();
    if !first_ok {
        return Ok(false);
    }
    let leading = live.get("leading");
    if !leading_block_breaks(&paragraph)?
        .items()
        .iter()
        .enumerate()
        .all(|(i, kind)| kind.strict_eq(&leading.at(i)))
    {
        return Ok(false);
    }
    let trailing = live.get("trailing");
    Ok(trailing_block_breaks(&paragraph)?
        .items()
        .iter()
        .enumerate()
        .all(|(i, kind)| kind.strict_eq(&trailing.at(i))))
}

/// `restoreRecordedBreaks` (src/yrs/yrsToDocument.ts:1685).
pub(crate) fn restore_recorded_breaks(content: &Js, items: &Js, boundaries: &Js, live_breaks: &Js) -> Result<Js> {
    let mut breaks: Vec<(f64, Js)> = Vec::new();
    let mut start = 0.0;
    for boundary in boundaries.items() {
        let mut cursor = 0.0;
        let text_len = js::len16(&js::to_js_string(&boundary.get("text"))) as f64;
        let entries = if note_mark_content(&boundary)?.truthy() { Vec::new() } else { boundary.get("breaks").items() };
        for entry in entries {
            cursor = js_min(js_max(num(&entry.get("offset")), cursor), text_len);
            breaks.push((start + cursor, entry.get("type")));
        }
        start += text_len;
    }
    if breaks.is_empty() {
        return Ok(content.clone());
    }
    let source = boundary_sources(boundaries)?;
    let trailing = trailing_block_breaks(&obj! { "type" => "paragraph", "content" => source.clone() })?;
    let hoisted = breaks.len() as i64 - trailing.len() as i64;
    let visible = includes_str(&flow_tokens(&source, &Js::new_arr(Vec::new()))?, "visible");
    let kept: Vec<(f64, Js)> = breaks
        .into_iter()
        .enumerate()
        .filter(|(index, entry)| *index as i64 >= hoisted || (visible && entry.1.as_str() == Some("page")))
        .map(|(_, entry)| entry)
        .collect();
    let recorded: Vec<u16> = boundaries
        .items()
        .iter()
        .flat_map(|b| js::utf16(&js::to_js_string(&b.get("text"))))
        .collect();
    let mut live: Vec<u16> = Vec::new();
    for item in items.items() {
        if item.get("kind").as_str() == Some("text") {
            live.extend(js::utf16(&js::to_js_string(&item.get("text"))));
        } else if !is_block_break(&item)? {
            live.push(0xFFFC);
        }
    }
    let mut prefix = 0usize;
    let shared = recorded.len().min(live.len());
    while prefix < shared && recorded[prefix] == live[prefix] {
        prefix += 1;
    }
    let mut suffix = 0usize;
    while suffix < shared - prefix && recorded[recorded.len() - 1 - suffix] == live[live.len() - 1 - suffix] {
        suffix += 1;
    }
    let mut length = 0.0;
    for child in content.items() {
        length += paragraph_content_length(&child)?;
    }
    let (prefix, suffix) = (prefix as f64, suffix as f64);
    let (rlen, llen) = (recorded.len() as f64, live.len() as f64);
    let marks: Vec<Js> = kept
        .iter()
        .enumerate()
        .map(|(id, (offset, _))| {
            let offset = *offset;
            let mapped = if offset <= prefix {
                offset
            } else if offset >= rlen - suffix {
                offset + llen - rlen
            } else {
                js_min(offset, llen - suffix)
            };
            obj! { "id" => id as f64, "kind" => "start", "offset" => js_min(js_max(mapped, 0.0), length) }
        })
        .collect();
    let placed = insert_boundaries(content, &Js::new_arr(marks), &|boundary| {
        let id = num(&boundary.get("id")) as usize;
        Ok(obj! { "type" => "run", "content" => arr![obj! { "type" => "break", "breakType" => kept[id].1.clone() }] })
    })?;
    let placed_trailing = trailing_block_breaks(&obj! { "type" => "paragraph", "content" => placed.clone() })?;
    let same = placed_trailing.len() == trailing.len()
        && placed_trailing.items().iter().enumerate().all(|(i, kind)| kind.strict_eq(&trailing.at(i)));
    Ok(if same && paragraph_breaks_are_live(&placed, live_breaks)? { placed } else { content.clone() })
}

/// `restoreOriginalRuns` (src/yrs/yrsToDocument.ts:1747).
pub(crate) fn restore_original_runs(content: &Js, items: &Js, boundaries: &Js, base_paragraph: &Js, live_breaks: &Js) -> Result<Js> {
    if boundaries.len() == 0 {
        let source = source_break_boundaries(base_paragraph)?;
        return restore_recorded_breaks(content, items, &live_run_boundaries(&source, live_breaks)?, live_breaks);
    }
    let recorded = live_run_boundaries(boundaries, live_breaks)?;
    let fallback = || restore_recorded_breaks(content, items, &recorded, live_breaks);
    let all_text_runs = content.items().iter().all(|child| {
        type_is(child, "run") && child.get("content").items().iter().all(|e| type_is(e, "text"))
    });
    let odd_item = items
        .items()
        .iter()
        .any(|item| item.get("kind").as_str() != Some("text") || item.path(&["attributes", "hyperlink"]).truthy());
    let note_with_text = recorded
        .items()
        .iter()
        .any(|b| b.get("noteMarks").len() > 0 && js::len16(&js::to_js_string(&b.get("text"))) > 0);
    if !all_text_runs || odd_item || note_with_text {
        return fallback();
    }
    let full_text: String = items
        .items()
        .iter()
        .map(|item| if item.get("kind").as_str() == Some("text") { js::to_js_string(&item.get("text")) } else { String::new() })
        .collect();
    let recorded_text: String = recorded.items().iter().map(|b| js::to_js_string(&b.get("text"))).collect();
    if full_text != recorded_text {
        return fallback();
    }

    let item_list = items.items();
    let non_formatting = non_formatting_attrs_b();
    let mut item_index = 0usize;
    let mut item_offset = 0usize;
    let mut restored_attrs = Vec::new();
    for boundary in recorded.items() {
        let marks_key = boundary.get("marksKey");
        let expected = marks_key_to_yrs_attrs(marks_key.as_str())?;
        if !expected.truthy() {
            return fallback();
        }
        restored_attrs.push(expected.clone());
        let mut remaining = js::len16(&js::to_js_string(&boundary.get("text")));
        while remaining > 0 {
            let item = item_list.get(item_index).cloned().unwrap_or_default();
            if item.get("kind").as_str() != Some("text") {
                return fallback();
            }
            if stable_stringify_without(&item.get("attributes"), &non_formatting)? != stable_stringify(&expected)? {
                return fallback();
            }
            let item_len = js::len16(&js::to_js_string(&item.get("text")));
            let available = item_len - item_offset;
            let consumed = remaining.min(available);
            remaining -= consumed;
            item_offset += consumed;
            if item_offset == item_len {
                item_index += 1;
                item_offset = 0;
            }
        }
    }

    if item_index != item_list.len() || item_offset != 0 {
        return fallback();
    }

    let source = if base_paragraph.truthy() {
        source_runs(&base_paragraph.get("content"), false)?
    } else {
        Js::new_arr(Vec::new())
    };
    let recorded_items = recorded.items();
    let source_items = source.items();
    let source_matches = source_items.len() == recorded_items.len()
        && source_items
            .iter()
            .enumerate()
            .all(|(i, run)| run.get("text").strict_eq(&recorded_items[i].get("text")));
    let mut restored = Vec::new();
    for (index, boundary) in recorded_items.iter().enumerate() {
        let run = obj! {
            "type" => "run",
            "content" => boundary_content(boundary, &attrs_to_text_formatting(&restored_attrs[index])?)?,
        };
        let formatting = if source_matches {
            boundary.get("formatting").with(&[
                ("fontFamily", source_items[index].path(&["formatting", "fontFamily"])),
                ("language", source_items[index].path(&["formatting", "language"])),
            ])
        } else {
            boundary.get("formatting")
        };
        if formatting.truthy() && formatting.key_count() > 0 {
            run.set("formatting", formatting);
        }
        if boundary.get("propertyChanges").len() > 0 {
            run.set("propertyChanges", boundary.get("propertyChanges"));
        }
        restored.push(run);
    }
    let restored = Js::new_arr(restored);
    Ok(if paragraph_breaks_are_live(&restored, live_breaks)? { restored } else { content.clone() })
}

/// `runContentLength` (src/yrs/yrsToDocument.ts:1819).
pub(crate) fn run_content_length(content: &Js) -> Result<f64> {
    let t = ty(content);
    if t == "text" || t == "instrText" {
        return Ok(js::len16(&js::to_js_string(&content.get("text"))) as f64);
    }
    if t == "symbol" {
        return Ok(content.get("char").len() as f64);
    }
    let break_type = content.get("breakType");
    if matches!(
        t.as_str(),
        "tab" | "softHyphen" | "noBreakHyphen" | "footnoteRef" | "endnoteRef" | "horizontalRule"
            | "commentReference" | "drawing" | "shape" | "chart" | "opaqueDrawing"
    ) || (t == "break" && (break_type.is_undefined() || break_type.as_str() == Some("textWrapping")))
    {
        return Ok(1.0);
    }
    Ok(0.0)
}

/// `runTextLength` (src/yrs/yrsToDocument.ts:1842).
pub(crate) fn run_text_length(run: &Js) -> Result<f64> {
    let mut length = 0.0;
    for content in run.get("content").items() {
        length += run_content_length(&content)?;
    }
    Ok(length)
}

/// `linkChildLength` (src/yrs/yrsToDocument.ts:1847).
pub(crate) fn link_child_length(child: &Js) -> Result<f64> {
    if type_is(child, "run") {
        return run_text_length(child);
    }
    Ok(if matches!(ty(child).as_str(), "simpleField" | "complexField" | "inlineSdt" | "mathEquation") { 1.0 } else { 0.0 })
}

/// `paragraphContentLength` (src/yrs/yrsToDocument.ts:1862).
pub(crate) fn paragraph_content_length(content: &Js) -> Result<f64> {
    match ty(content).as_str() {
        "run" => run_text_length(content),
        "hyperlink" => {
            let mut sum = 0.0;
            for child in content.get("structuredChildren").nullish_or(content.get("children")).items() {
                sum += link_child_length(&child)?;
            }
            Ok(sum)
        }
        "complexField" => {
            let mut children = content.path(&["structuredCode", "inline"]).items();
            children.extend(content.path(&["structuredResult", "inline"]).items());
            let mut sum = 1.0;
            for child in children {
                if type_is(&child, "hyperlink") || type_is(&child, "simpleField") {
                    sum += paragraph_content_length(&child)?;
                }
            }
            Ok(sum)
        }
        "simpleField" | "inlineSdt" | "mathEquation" => Ok(1.0),
        "insertion" | "deletion" | "moveFrom" | "moveTo" => {
            let mut sum = 0.0;
            for child in content.get("content").items() {
                sum += paragraph_content_length(&child)?;
            }
            Ok(sum)
        }
        _ => Ok(0.0),
    }
}

/// `splitItems` (src/yrs/yrsToDocument.ts:1895).
pub(crate) fn split_items(items: &Js, at: f64, length: &dyn Fn(&Js) -> Result<f64>, split: &dyn Fn(&Js, f64) -> Result<Js>) -> Result<Js> {
    let mut left = Vec::new();
    let mut right = Vec::new();
    let mut offset = 0.0;
    for item in items.items() {
        let width = length(&item)?;
        if offset >= at {
            right.push(item);
        } else if offset + width <= at {
            left.push(item);
        } else {
            let parts = split(&item, at - offset)?;
            if !parts.truthy() {
                return Ok(Js::Null);
            }
            if parts.at(0).truthy() {
                left.push(parts.at(0));
            }
            if parts.at(1).truthy() {
                right.push(parts.at(1));
            }
        }
        offset += width;
    }
    Ok(arr![Js::new_arr(left), Js::new_arr(right)])
}

/// `splitRunContent` (src/yrs/yrsToDocument.ts:1919).
pub(crate) fn split_run_content(content: &Js, at: f64) -> Result<Js> {
    if !type_is(content, "text") && !type_is(content, "instrText") {
        return Ok(Js::Null);
    }
    let text = js::to_js_string(&content.get("text"));
    Ok(arr![
        content.with(&[("text", Js::from(js::slice16(&text, 0.0, Some(at))))]),
        content.with(&[("text", Js::from(js::slice16(&text, at, None)))]),
    ])
}

/// `linkPart` (src/yrs/yrsToDocument.ts:1928).
pub(crate) fn link_part(link: &Js, children: &Js) -> Result<Js> {
    if children.len() == 0 {
        return Ok(Js::Null);
    }
    let legacy = Js::new_arr(
        children
            .items()
            .into_iter()
            .filter(|c| matches!(ty(c).as_str(), "run" | "bookmarkStart" | "bookmarkEnd"))
            .collect(),
    );
    Ok(if link.get("structuredChildren").truthy() {
        link.with(&[("children", legacy), ("structuredChildren", children.clone())])
    } else {
        link.with(&[("children", legacy)])
    })
}

fn split_halves(content: &Js, parts: &Js) -> Js {
    if !parts.truthy() {
        return Js::Null;
    }
    let half = |p: Js| if p.len() > 0 { content.with(&[("content", p)]) } else { Js::Null };
    arr![half(parts.at(0)), half(parts.at(1))]
}

/// `splitContent` (src/yrs/yrsToDocument.ts:1944).
pub(crate) fn split_content(content: &Js, at: f64) -> Result<Js> {
    if type_is(content, "run") {
        let parts = split_items(&content.get("content"), at, &run_content_length, &split_run_content)?;
        return Ok(split_halves(content, &parts));
    }
    if type_is(content, "hyperlink") {
        let children = content.get("structuredChildren").nullish_or(content.get("children"));
        let parts = split_items(&children, at, &link_child_length, &|child, offset| {
            if type_is(child, "run") { split_content(child, offset) } else { Ok(Js::Null) }
        })?;
        if !parts.truthy() {
            return Ok(Js::Null);
        }
        return Ok(arr![link_part(content, &parts.at(0))?, link_part(content, &parts.at(1))?]);
    }
    if is_tracked_wrapper(content)? {
        let parts = split_items(&content.get("content"), at, &paragraph_content_length, &split_content)?;
        return Ok(split_halves(content, &parts));
    }
    Ok(Js::Null)
}

/// `insertBoundaries` (src/yrs/yrsToDocument.ts:1980). Callers wanting the TS default
/// `makeMarker` pass `&comment_marker_default`.
pub(crate) fn insert_boundaries(content: &Js, boundaries: &Js, make_marker: &dyn Fn(&Js) -> Result<Js>) -> Result<Js> {
    if boundaries.len() == 0 {
        return Ok(content.clone());
    }
    let list = boundaries.items();
    let key = |b: &Js| format!("{}:{}", js::to_js_string(&b.get("offset")), js::to_js_string(&b.get("id")));
    let carets: HashSet<String> = list
        .iter()
        .filter(|b| {
            b.get("kind").as_str() == Some("end")
                && list.iter().any(|o| {
                    o.get("kind").as_str() == Some("start")
                        && o.get("id").strict_eq(&b.get("id"))
                        && o.get("offset").strict_eq(&b.get("offset"))
                })
        })
        .map(|b| key(b))
        .collect();
    let rank = |b: &Js| -> f64 {
        if carets.contains(&key(b)) { 1.0 } else if b.get("kind").as_str() == Some("end") { 0.0 } else { 2.0 }
    };
    let mut sorted = list.clone();
    sorted.sort_by(|l, r| {
        let mut c = num(&l.get("offset")) - num(&r.get("offset"));
        if c == 0.0 || c.is_nan() {
            c = rank(l) - rank(r);
        }
        if c == 0.0 || c.is_nan() {
            c = num(&l.get("id")) - num(&r.get("id"));
        }
        if c == 0.0 || c.is_nan() {
            let (lk, rk) = (l.get("kind"), r.get("kind"));
            c = if lk.strict_eq(&rk) { 0.0 } else if lk.as_str() == Some("start") { -1.0 } else { 1.0 };
        }
        if c < 0.0 {
            std::cmp::Ordering::Less
        } else if c > 0.0 {
            std::cmp::Ordering::Greater
        } else {
            std::cmp::Ordering::Equal
        }
    });
    let offset_of = |i: usize| num(&sorted[i].get("offset"));
    let result = Js::new_arr(Vec::new());
    let mut cursor = 0.0;
    let mut bi = 0usize;
    let emit = |offset: f64, bi: &mut usize| -> Result<()> {
        while *bi < sorted.len() && offset_of(*bi) <= offset {
            result.push(make_marker(&sorted[*bi])?);
            *bi += 1;
        }
        Ok(())
    };

    emit(0.0, &mut bi)?;
    for child in content.items() {
        let end = cursor + paragraph_content_length(&child)?;
        let mut item = child.clone();
        while item.truthy() && bi < sorted.len() && offset_of(bi) < end {
            let offset = offset_of(bi);
            let parts = split_content(&item, offset - cursor)?;
            if !parts.truthy() {
                break;
            }
            if parts.at(0).truthy() {
                result.push(parts.at(0));
            }
            cursor = offset;
            emit(cursor, &mut bi)?;
            item = parts.at(1);
        }
        let mut inside = Vec::new();
        while bi < sorted.len() && offset_of(bi) < end {
            inside.push(sorted[bi].clone());
            bi += 1;
        }
        for boundary in &inside {
            if boundary.get("kind").as_str() == Some("start") {
                result.push(make_marker(boundary)?);
            }
        }
        if item.truthy() {
            result.push(item);
        }
        for boundary in &inside {
            if boundary.get("kind").as_str() == Some("end") {
                result.push(make_marker(boundary)?);
            }
        }
        cursor = end;
        emit(cursor, &mut bi)?;
    }
    emit(f64::INFINITY, &mut bi)?;
    Ok(result)
}

/// The TS default `makeMarker` of `insertBoundaries`: comment range markers.
pub(crate) fn comment_marker_default(boundary: &Js) -> Result<Js> {
    comment_marker(boundary)
}

/// `bookmarkBoundaries` (src/yrs/yrsToDocument.ts:2047).
pub(crate) fn bookmark_boundaries(properties: &Js) -> Result<Js> {
    let bookmarks = properties.get("bookmarks");
    if !bookmarks.is_arr() {
        return Ok(Js::new_arr(Vec::new()));
    }
    let result = Js::new_arr(Vec::new());
    for raw in bookmarks.items() {
        let bookmark = as_object(&raw)?;
        let id = as_finite_number(&bookmark.get("id"))?;
        let id = match id {
            Some(id) if bookmark.truthy() => id,
            _ => continue,
        };
        let offset = as_finite_number(&bookmark.get("offset"))?.unwrap_or(0.0);
        let name = Js::from(as_string(&bookmark.get("name"))?);
        let col_first = Js::from(as_finite_number(&bookmark.get("colFirst"))?);
        let col_last = Js::from(as_finite_number(&bookmark.get("colLast"))?);
        let make = |kind: &str, offset: f64| {
            obj! {
                "id" => id, "kind" => kind, "offset" => offset,
                "name" => name.clone(), "colFirst" => col_first.clone(), "colLast" => col_last.clone(),
            }
        };
        match bookmark.get("kind").as_str() {
            Some("start") => result.push(make("start", offset)),
            Some("end") => result.push(make("end", offset)),
            _ => {
                result.push(make("start", 0.0));
                result.push(make("end", MAX_SAFE_INTEGER_B));
            }
        }
    }
    Ok(result)
}

/// `nestedParagraphs` (src/yrs/yrsToDocument.ts:2071).
pub(crate) fn nested_paragraphs(value: &Js, found: &Js) -> Result<Js> {
    if value.is_arr() {
        for item in value.items() {
            nested_paragraphs(&item, found)?;
        }
    } else if value.is_obj() {
        if type_is(value, "paragraph") {
            found.push(value.clone());
        }
        for (_, child) in value.entries() {
            nested_paragraphs(&child, found)?;
        }
    }
    Ok(found.clone())
}

/// `restoreNestedOrdinals` (src/yrs/yrsToDocument.ts:2087).
pub(crate) fn restore_nested_ordinals(content: &Js, base: &Js) -> Result<()> {
    let sources = nested_paragraphs(base, &Js::new_arr(Vec::new()))?.items();
    if !sources.iter().any(|p| p.get("sourceOrdinal").is_num()) {
        return Ok(());
    }
    let targets = nested_paragraphs(content, &Js::new_arr(Vec::new()))?.items();
    if targets.len() != sources.len() {
        return Ok(());
    }
    for (index, target) in targets.iter().enumerate() {
        let ordinal = sources[index].get("sourceOrdinal");
        if ordinal.is_num() {
            target.set("sourceOrdinal", ordinal);
        }
    }
    Ok(())
}

/// `restoreRawInlines` (src/yrs/yrsToDocument.ts:2098).
pub(crate) fn restore_raw_inlines(content: &Js, base: &Js) -> Result<Js> {
    if !base.truthy() {
        return Ok(content.clone());
    }
    let mut length = 0.0;
    for child in content.items() {
        length += paragraph_content_length(&child)?;
    }
    let base_content = base.get("content");
    let mut boundaries = Vec::new();
    let mut offset = 0.0;
    for (index, child) in base_content.items().iter().enumerate() {
        if type_is(child, "rawXml") {
            boundaries.push(obj! { "id" => index as f64, "kind" => "start", "offset" => js_min(offset, length) });
        }
        offset += paragraph_content_length(child)?;
    }
    insert_boundaries(content, &Js::new_arr(boundaries), &|boundary| {
        Ok(base_content.at(num(&boundary.get("id")) as usize))
    })
}

/// `paragraphAttrs` (src/yrs/yrsToDocument.ts:2110).
pub(crate) fn paragraph_attrs(properties: &Js) -> Result<Js> {
    let attrs = paragraph_attr_defaults_b();
    attrs.assign(properties);
    attrs.set("styleId", properties.get("pStyle").nullish_or(Js::Null));
    attrs.set("_sectionProperties", properties.get("sectPr").nullish_or(Js::Null));
    attrs.delete("pStyle");
    attrs.delete("sectPr");
    Ok(attrs)
}

/// `sourceRuns` (src/yrs/yrsToDocument.ts:2119).
pub(crate) fn source_runs(content: &Js, into_fields: bool) -> Result<Js> {
    let mut out = Vec::new();
    for child in content.items() {
        let t = ty(&child);
        if t == "run" {
            let mut text = String::new();
            for entry in child.get("content").items() {
                match ty(&entry).as_str() {
                    "text" | "instrText" => text.push_str(&js::to_js_string(&entry.get("text"))),
                    "symbol" => text.push_str(&js::to_js_string(&entry.get("char"))),
                    "tab" => text.push('\t'),
                    "softHyphen" => text.push('\u{00ad}'),
                    "noBreakHyphen" => text.push('\u{2011}'),
                    _ => text.push_str(&"\u{fffc}".repeat(run_content_length(&entry)? as usize)),
                }
            }
            out.push(obj! { "text" => text, "formatting" => child.get("formatting") });
            continue;
        }
        if t == "hyperlink" {
            out.extend(source_runs(&child.get("children"), into_fields)?.items());
            continue;
        }
        if is_tracked_wrapper(&child)? || t == "inlineSdt" || (into_fields && t == "simpleField") {
            out.extend(source_runs(&child.get("content"), into_fields)?.items());
            continue;
        }
        out.push(obj! { "text" => "\u{fffc}".repeat(paragraph_content_length(&child)? as usize) });
    }
    Ok(Js::new_arr(out))
}

/// `s[i]`: `undefined` (None) out of range.
fn unit(s: &[u16], i: i64) -> Option<u16> {
    if i >= 0 && (i as usize) < s.len() { Some(s[i as usize]) } else { None }
}

/// JS `slice` bound semantics: negative counts from the end, then clamped.
fn rel(v: i64, len: usize) -> usize {
    if v < 0 { (len as i64 + v).max(0) as usize } else { (v as usize).min(len) }
}

fn edit_midpoint16(original: &[u16], current: &[u16]) -> Option<(i64, i64)> {
    let (olen, clen) = (original.len() as i64, current.len() as i64);
    let distance_limit = ((olen + clen) as f64 / 2.0).ceil() as i64;
    let offset = distance_limit + 1;
    let size = (2 * distance_limit + 3) as usize;
    let mut forward = vec![-1i64; size];
    let mut backward = vec![-1i64; size];
    forward[(offset + 1) as usize] = 0;
    backward[(offset + 1) as usize] = 0;
    let delta = olen - clen;
    let odd = delta % 2 != 0;
    let get = |a: &Vec<i64>, i: i64| -> Option<i64> { if i >= 0 && (i as usize) < a.len() { Some(a[i as usize]) } else { None } };
    let (mut forward_start, mut forward_end, mut backward_start, mut backward_end) = (0i64, 0i64, 0i64, 0i64);
    let mut distance = 0i64;
    while distance <= distance_limit {
        let mut diagonal = -distance + forward_start;
        while diagonal <= distance - forward_end {
            let index = offset + diagonal;
            let mut x = if diagonal == -distance
                || (diagonal != distance && forward[(index - 1) as usize] < forward[(index + 1) as usize])
            {
                forward[(index + 1) as usize]
            } else {
                forward[(index - 1) as usize] + 1
            };
            let mut y = x - diagonal;
            while x < olen && y < clen && unit(original, x) == unit(current, y) {
                x += 1;
                y += 1;
            }
            forward[index as usize] = x;
            if x > olen {
                forward_end += 2;
            } else if y > clen {
                forward_start += 2;
            } else if odd {
                if let Some(b) = get(&backward, offset + delta - diagonal) {
                    if b >= 0 && x >= olen - b {
                        return Some((x, y));
                    }
                }
            }
            diagonal += 2;
        }
        let mut diagonal = -distance + backward_start;
        while diagonal <= distance - backward_end {
            let index = offset + diagonal;
            let mut x = if diagonal == -distance
                || (diagonal != distance && backward[(index - 1) as usize] < backward[(index + 1) as usize])
            {
                backward[(index + 1) as usize]
            } else {
                backward[(index - 1) as usize] + 1
            };
            let mut y = x - diagonal;
            while x < olen && y < clen && unit(original, olen - x - 1) == unit(current, clen - y - 1) {
                x += 1;
                y += 1;
            }
            backward[index as usize] = x;
            if x > olen {
                backward_end += 2;
            } else if y > clen {
                backward_start += 2;
            } else if !odd {
                if let Some(f) = get(&forward, offset + delta - diagonal) {
                    if f >= 0 && f >= olen - x {
                        return Some((f, f - delta + diagonal));
                    }
                }
            }
            diagonal += 2;
        }
        distance += 1;
    }
    None
}

/// `editMidpoint` (src/yrs/yrsToDocument.ts:2143).
pub(crate) fn edit_midpoint(original: &str, current: &str) -> Result<Js> {
    Ok(match edit_midpoint16(&js::utf16(original), &js::utf16(current)) {
        Some((x, y)) => arr![x as f64, y as f64],
        None => Js::Null,
    })
}

/// `(original, current, length)` of each `SourceTextMatch`.
fn source_text_matches16(original: &[u16], current: &[u16]) -> Vec<(usize, usize, usize)> {
    fn visit(before: &[u16], after: &[u16], source_offset: usize, live_offset: usize, matches: &mut Vec<(usize, usize, usize)>) {
        let shared = before.len().min(after.len());
        let mut prefix = 0;
        while prefix < shared && before[prefix] == after[prefix] {
            prefix += 1;
        }
        if prefix > 0 {
            matches.push((source_offset, live_offset, prefix));
        }
        let mut suffix = 0;
        while suffix < shared - prefix && before[before.len() - suffix - 1] == after[after.len() - suffix - 1] {
            suffix += 1;
        }
        let old_middle = &before[prefix..before.len() - suffix];
        let new_middle = &after[prefix..after.len() - suffix];
        if !old_middle.is_empty() && !new_middle.is_empty() {
            if let Some((s0, s1)) = edit_midpoint16(old_middle, new_middle) {
                let (a, b) = (rel(s0, old_middle.len()), rel(s1, new_middle.len()));
                if (s0 > 0 || s1 > 0) && (s0 < old_middle.len() as i64 || s1 < new_middle.len() as i64) {
                    visit(&old_middle[..a], &new_middle[..b], source_offset + prefix, live_offset + prefix, matches);
                    visit(
                        &old_middle[a..],
                        &new_middle[b..],
                        (source_offset as i64 + prefix as i64 + s0) as usize,
                        (live_offset as i64 + prefix as i64 + s1) as usize,
                        matches,
                    );
                }
            }
        }
        if suffix > 0 {
            matches.push((source_offset + before.len() - suffix, live_offset + after.len() - suffix, suffix));
        }
    }
    let mut matches = Vec::new();
    visit(original, current, 0, 0, &mut matches);
    matches
}

/// `sourceTextMatches` (src/yrs/yrsToDocument.ts:2203).
pub(crate) fn source_text_matches(original: &str, current: &str) -> Result<Js> {
    Ok(Js::new_arr(
        source_text_matches16(&js::utf16(original), &js::utf16(current))
            .into_iter()
            .map(|(o, c, l)| obj! { "original" => o as f64, "current" => c as f64, "length" => l as f64 })
            .collect(),
    ))
}

/// `preserveSourceRunFormatting` (src/yrs/yrsToDocument.ts:2232).
pub(crate) fn preserve_source_run_formatting(items: &Js, sources: &Js) -> Result<Js> {
    if sources.len() == 0 {
        return Ok(items.clone());
    }
    let source_list = sources.items();
    let original: Vec<u16> = source_list.iter().flat_map(|s| js::utf16(&js::to_js_string(&s.get("text")))).collect();
    let mut current: Vec<u16> = Vec::new();
    for item in items.items() {
        if item.get("kind").as_str() == Some("text") {
            current.extend(js::utf16(&js::to_js_string(&item.get("text"))));
        } else if item.get("embedKind").as_str() == Some("tab") {
            current.push(b'\t' as u16);
        } else {
            current.push(0xFFFC);
        }
    }
    let matches = source_text_matches16(&original, &current);
    let current_len = current.len();
    let mapped = |offset: usize| -> usize {
        if offset == 0 {
            return 0;
        }
        let mut source_end = 0;
        let mut live_end = 0;
        for &(mo, mc, ml) in &matches {
            if offset < mo {
                return if offset == source_end { live_end } else { mc };
            }
            if offset <= mo + ml {
                return mc + offset - mo;
            }
            source_end = mo + ml;
            live_end = mc + ml;
        }
        if offset == source_end { live_end } else { current_len }
    };
    let mut source_offset = 0usize;
    let mut spans: Vec<(usize, usize, Js)> = Vec::new();
    for (index, source) in source_list.iter().enumerate() {
        let start = mapped(source_offset);
        source_offset += js::len16(&js::to_js_string(&source.get("text")));
        let end = if index + 1 == source_list.len() { current_len } else { mapped(source_offset) };
        if end > start {
            spans.push((start, end, source.get("formatting")));
        }
    }
    let mut offset = 0usize;
    let mut out = Vec::new();
    for item in items.items() {
        let is_text = item.get("kind").as_str() == Some("text");
        let text = js::to_js_string(&item.get("text"));
        let text_len = js::len16(&text);
        let end = offset + if is_text { text_len } else { 1 };
        while offset < end {
            let span = spans.iter().find(|(s, e, _)| *s <= offset && offset < *e);
            let next = end.min(span.map_or(end, |s| s.1));
            let source = Js::new_obj();
            if let Some((_, _, formatting)) = span {
                for key in SOURCE_RUN_PROPERTIES_B {
                    let value = formatting.get(key);
                    if !value.is_undefined() {
                        source.set(key, value);
                    }
                }
            }
            let attributes = item.get("attributes").with(&[("_sourceRunFormatting", source)]);
            out.push(if is_text {
                let sliced = js::slice16(
                    &text,
                    text_len as f64 - (end - offset) as f64,
                    Some(text_len as f64 - (end - next) as f64),
                );
                item.with(&[("text", Js::from(sliced)), ("attributes", attributes)])
            } else {
                item.with(&[("attributes", attributes)])
            });
            offset = next;
        }
    }
    Ok(Js::new_arr(out))
}

/// `restoreLineBreaks` (src/yrs/yrsToDocument.ts:2278).
pub(crate) fn restore_line_breaks(content: &Js, base: &Js) -> Result<()> {
    if !base.truthy() {
        return Ok(());
    }
    let source = paragraph_chart_runs(&base.get("content"))?;
    let target = paragraph_chart_runs(content)?;
    let text = |runs: &Js| -> Result<String> {
        Ok(source_runs(runs, false)?.items().iter().map(|r| js::to_js_string(&r.get("text"))).collect())
    };
    let matches = source_text_matches16(&js::utf16(&text(&source)?), &js::utf16(&text(&target)?));
    let mut breaks: HashMap<i64, Js> = HashMap::new();
    let mut offset = 0.0f64;
    for run in source.items() {
        for entry in run.get("content").items() {
            let len = run_content_length(&entry)?;
            if type_is(&entry, "break") && len == 1.0 {
                let o = offset as usize;
                if let Some(&(mo, mc, _)) = matches.iter().find(|&&(mo, _, ml)| mo <= o && o < mo + ml) {
                    breaks.insert((mc + o - mo) as i64, entry.clone());
                }
            }
            offset += len;
        }
    }
    offset = 0.0;
    for run in target.items() {
        let mut next = Vec::new();
        for entry in run.get("content").items() {
            let original = breaks.get(&(offset as i64)).cloned();
            let len = run_content_length(&entry)?;
            offset += len;
            next.push(match original {
                Some(original) if type_is(&entry, "break") && len == 1.0 && original.truthy() => original.spread(),
                _ => entry,
            });
        }
        run.set("content", Js::new_arr(next));
    }
    Ok(())
}

/// `paragraphFromStory` (src/yrs/yrsToDocument.ts:2305).
pub(crate) fn paragraph_from_story(para_id: &str, properties: &Js, items: &Js, comment_boundaries: &Js, base_paragraph: &Js, live_breaks: &Js, recovery: &Js, format: &dyn Fn(&Js) -> Result<Js>, revision_ids: &Js, inherited: &Js) -> Result<Js> {
    let attrs = paragraph_attrs(properties)?;
    let has_base = base_paragraph.truthy();
    if has_base && attrs.get("_originalFormatting").truthy() {
        let original = attrs.get("_originalFormatting");
        let base_rpr = base_paragraph.path(&["formatting", "runProperties"]);
        let run_properties = if original.get("runProperties").truthy() || base_rpr.truthy() {
            original.get("runProperties").with(&[
                ("fontFamily", base_rpr.get("fontFamily")),
                ("language", base_rpr.get("language")),
            ])
        } else {
            Js::Undefined
        };
        attrs.set("_originalFormatting", original.with(&[("runProperties", run_properties)]));
    }
    let boundaries_attr = attrs.get("_originalRunBoundaries");
    let sources = if has_base {
        source_runs(&base_paragraph.get("content"), false)?
    } else if boundaries_attr.is_arr() {
        boundaries_attr.clone()
    } else if attrs.path(&["_originalFormatting", "runProperties"]).truthy() {
        let rpr = attrs.path(&["_originalFormatting", "runProperties"]);
        let text: String = items
            .items()
            .iter()
            .map(|item| if item.get("kind").as_str() == Some("text") { js::to_js_string(&item.get("text")) } else { "\u{fffc}".to_string() })
            .collect();
        arr![obj! {
            "text" => text,
            "formatting" => obj! { "fontFamily" => rpr.get("fontFamily"), "language" => rpr.get("language") },
        }]
    } else {
        Js::new_arr(Vec::new())
    };
    let mut content = build_paragraph_content(&preserve_source_run_formatting(items, &sources)?, recovery, format, revision_ids, false)?;
    content = restore_original_runs(
        &content,
        items,
        &if boundaries_attr.is_arr() { boundaries_attr.clone() } else { Js::Undefined },
        base_paragraph,
        live_breaks,
    )?;
    restore_line_breaks(&content, base_paragraph)?;
    content = restore_raw_inlines(&content, base_paragraph)?;
    if has_base {
        restore_nested_ordinals(&content, &base_paragraph.get("content"))?;
        restore_text_box_bookmarks(&content, &base_paragraph.get("content"))?;
    }
    content = insert_boundaries(&content, comment_boundaries, &comment_marker)?;

    let mut bookmarks = Vec::new();
    for boundary in bookmark_boundaries(properties)?.items() {
        let offset = if num(&boundary.get("offset")) == MAX_SAFE_INTEGER_B {
            let mut sum = 0.0;
            for child in content.items() {
                sum += paragraph_content_length(&child)?;
            }
            Js::from(sum)
        } else {
            boundary.get("offset")
        };
        bookmarks.push(boundary.with(&[("offset", offset)]));
    }
    if !bookmarks.is_empty() {
        content = insert_boundaries(&content, &Js::new_arr(bookmarks), &|boundary| {
            Ok(if boundary.get("kind").as_str() == Some("start") {
                let name = boundary.get("name");
                let marker = obj! {
                    "type" => "bookmarkStart",
                    "id" => boundary.get("id"),
                    "name" => if name.truthy() { name } else { Js::from("") },
                };
                if !boundary.get("colFirst").is_undefined() {
                    marker.set("colFirst", boundary.get("colFirst"));
                }
                if !boundary.get("colLast").is_undefined() {
                    marker.set("colLast", boundary.get("colLast"));
                }
                marker.set("position", obj! { "offset" => boundary.get("offset") });
                marker
            } else {
                obj! {
                    "type" => "bookmarkEnd",
                    "id" => boundary.get("id"),
                    "position" => obj! { "offset" => boundary.get("offset") },
                }
            })
        })?;
    }

    let paragraph = obj! {
        "type" => "paragraph",
        "paraId" => if para_id.is_empty() { Js::Undefined } else { Js::from(para_id) },
        "textId" => base_paragraph.get("textId"),
    };
    if base_paragraph.get("extraAttributes").truthy() {
        paragraph.set("extraAttributes", base_paragraph.get("extraAttributes"));
    }
    if base_paragraph.get("paraIdAttribute").truthy() {
        paragraph.set("paraIdAttribute", base_paragraph.get("paraIdAttribute"));
    }
    paragraph.set("formatting", paragraph_attrs_to_formatting(&attrs, inherited)?);
    paragraph.set("content", content);
    if base_paragraph.get("renderedPageBreakBefore").truthy() {
        paragraph.set("renderedPageBreakBefore", true);
    }

    let p_pr_ins = tracked_info(&properties.get("pPrIns"), true, None)?;
    let p_pr_del = tracked_info(&properties.get("pPrDel"), true, None)?;
    if p_pr_ins.truthy() {
        paragraph.set("pPrIns", p_pr_ins);
    }
    if p_pr_del.truthy() {
        paragraph.set("pPrDel", p_pr_del);
    }
    let p_pr_change = properties.get("pPrChange");
    if p_pr_change.is_arr() && p_pr_change.len() > 0 {
        paragraph.set("propertyChanges", property_changes(&p_pr_change)?);
    }
    if properties.get("sectPr").truthy() {
        paragraph.set("sectionProperties", properties.get("sectPr"));
    } else if properties.get("sectionBreakType").truthy() {
        paragraph.set("sectionProperties", obj! { "sectionStart" => properties.get("sectionBreakType") });
    }
    Ok(paragraph)
}

/// `inferTableBorders` (src/yrs/yrsToDocument.ts:2414).
pub(crate) fn infer_table_borders(rows: &Js) -> Result<Js> {
    let first_row = rows.at(0).get("cells");
    let last_row = if rows.len() > 0 { rows.at(rows.len() - 1).get("cells") } else { Js::Undefined };
    let corner = first_row.at(0).path(&["formatting", "borders"]);
    if !first_row.truthy() || !last_row.truthy() || !corner.truthy() {
        return Ok(Js::Undefined);
    }
    let last_of_first = if first_row.len() > 0 { first_row.at(first_row.len() - 1) } else { Js::Undefined };
    let entries = [
        ("top", corner.get("top")),
        ("left", corner.get("left")),
        ("bottom", last_row.at(0).path(&["formatting", "borders", "bottom"])),
        ("right", last_of_first.path(&["formatting", "borders", "right"])),
        ("insideH", if rows.len() > 1 { corner.get("bottom") } else { Js::Undefined }),
        ("insideV", if first_row.len() > 1 { corner.get("right") } else { Js::Undefined }),
    ];
    let borders = Js::new_obj();
    for (key, value) in entries {
        if !value.is_undefined() {
            borders.set(key, value);
        }
    }
    Ok(if borders.key_count() > 0 { borders } else { Js::Undefined })
}

/// `normalizeVMergeRuns` (src/yrs/yrsToDocument.ts:2431).
pub(crate) fn normalize_v_merge_runs(rows: &Js) -> Result<()> {
    let mut columns: IndexMap<String, Vec<(usize, Js)>> = IndexMap::new();
    for (row_index, row) in rows.items().iter().enumerate() {
        let mut column = num(&row.path(&["formatting", "gridBefore"]).nullish_or(Js::from(0.0)));
        for cell in row.get("cells").items() {
            let start = column;
            column += num(&cell.path(&["formatting", "gridSpan"]).nullish_or(Js::from(1.0)));
            if cell.path(&["formatting", "vMerge"]).truthy() {
                columns.entry(js::num_to_string(start)).or_default().push((row_index, cell.clone()));
            }
        }
    }

    let clear = |cell: &Js| {
        let formatting = cell.get("formatting");
        if !formatting.truthy() {
            return;
        }
        formatting.delete("vMerge");
        if formatting.key_count() == 0 {
            cell.set("formatting", Js::Undefined);
        }
    };

    for entries in columns.values() {
        let mut start: Option<Js> = None;
        let mut length = 0usize;
        let mut last_row: i64 = -1;
        let close = |start: &mut Option<Js>, length: &mut usize| {
            if let Some(s) = start.as_ref() {
                if *length < 2 {
                    clear(s);
                }
            }
            *start = None;
            *length = 0;
        };
        for (row_index, cell) in entries {
            let marker = cell.path(&["formatting", "vMerge"]);
            if marker.as_str() == Some("restart") {
                close(&mut start, &mut length);
                start = Some(cell.clone());
                length = 1;
                last_row = *row_index as i64;
            } else if marker.as_str() == Some("continue") {
                if start.is_some() && *row_index as i64 == last_row + 1 {
                    length += 1;
                    last_row = *row_index as i64;
                } else {
                    close(&mut start, &mut length);
                    clear(cell);
                }
            }
        }
        close(&mut start, &mut length);
    }
    Ok(())
}

/// `tableCellFromPayload` (src/yrs/yrsToDocument.ts:2482).
pub(crate) fn table_cell_from_payload(context: &crate::yrs_to_document::SaveContext, payload: &Js) -> Result<Js> {
    let attrs = context.cell_attrs(payload)?;
    let story = payload.get("story");
    let content = match story.as_str() {
        Some(s) if !s.is_empty() && context.story_ids.contains(s) => context.story_to_blocks(s)?,
        _ => Js::new_arr(Vec::new()),
    };
    let cell = obj! {
        "type" => "tableCell",
        "formatting" => table_cell_attrs_to_formatting(&attrs)?,
        "content" => content,
    };
    let marker = attrs.get("cellMarker");
    if marker.truthy() {
        let info = tracked_info(&marker.get("info"), true, None)?
            .nullish_or(obj! { "id" => 0.0, "author" => "Unknown" });
        match marker.get("kind").as_str() {
            Some("ins") => cell.set("structuralChange", obj! { "type" => "tableCellInsertion", "info" => info }),
            Some("del") => cell.set("structuralChange", obj! { "type" => "tableCellDeletion", "info" => info }),
            _ => {
                let change = obj! { "type" => "tableCellMerge", "info" => info };
                if marker.get("vMerge").truthy() {
                    change.set("vMerge", marker.get("vMerge"));
                }
                if marker.get("vMergeOrig").truthy() {
                    change.set("vMergeOrig", marker.get("vMergeOrig"));
                }
                cell.set("structuralChange", change);
            }
        }
    }
    let tc_pr_change = attrs.get("tcPrChange");
    if tc_pr_change.is_arr() && tc_pr_change.len() > 0 {
        cell.set("propertyChanges", tc_pr_change);
    }
    Ok(cell)
}

struct Anchor {
    row: f64,
    col: f64,
    rowspan: f64,
    colspan: f64,
    cell: Js,
    story: Option<String>,
}

fn grid_key(row: f64, col: f64) -> String {
    format!("{}-{}", js::num_to_string(row), js::num_to_string(col))
}

/// `tableFromPayload` (src/yrs/yrsToDocument.ts:2513).
pub(crate) fn table_from_payload(context: &crate::yrs_to_document::SaveContext, payload: &Js) -> Result<Js> {
    let rows_raw = payload.get("rows");
    let row_payloads = if rows_raw.is_arr() { rows_raw.items() } else { Vec::new() };
    let mut occupied: HashSet<String> = HashSet::new();
    let mut anchors: Vec<Anchor> = Vec::new();
    let mut total_columns = 0.0f64;

    for (row_index, row) in row_payloads.iter().enumerate() {
        let row_index = row_index as f64;
        let mut column = num(&table_row_grid(&row.get("trPr"))?.get("before"));
        let cells = row.get("cells");
        for cell_payload in if cells.is_arr() { cells.items() } else { Vec::new() } {
            while occupied.contains(&grid_key(row_index, column)) {
                column += 1.0;
            }
            let tc_pr = cell_payload.get("tcPr").nullish_or(Js::new_obj());
            let rowspan = as_finite_number(&tc_pr.get("rowspan"))?.filter(|v| *v != 0.0).unwrap_or(1.0);
            let colspan = as_finite_number(&tc_pr.get("colspan"))?.filter(|v| *v != 0.0).unwrap_or(1.0);
            anchors.push(Anchor {
                row: row_index,
                col: column,
                rowspan,
                colspan,
                cell: table_cell_from_payload(context, &cell_payload)?,
                story: as_string(&cell_payload.get("story"))?,
            });
            let mut r = row_index;
            while r < row_index + rowspan {
                let mut c = column;
                while c < column + colspan {
                    occupied.insert(grid_key(r, c));
                    c += 1.0;
                }
                r += 1.0;
            }
            column += colspan;
            total_columns = js_max(total_columns, column);
        }
    }

    let mut by_start: HashMap<String, usize> = HashMap::new();
    for (i, anchor) in anchors.iter().enumerate() {
        by_start.insert(grid_key(anchor.row, anchor.col), i);
    }
    let mut by_covered: HashMap<String, usize> = HashMap::new();
    for (i, anchor) in anchors.iter().enumerate() {
        let mut row = anchor.row;
        while row < anchor.row + anchor.rowspan {
            let mut col = anchor.col;
            while col < anchor.col + anchor.colspan {
                by_covered.insert(grid_key(row, col), i);
                col += 1.0;
            }
            row += 1.0;
        }
    }

    let mut cell_stories: Vec<Js> = Vec::new();
    let mut rows = Vec::new();
    for (row_index, row_payload) in row_payloads.iter().enumerate() {
        let ri = row_index as f64;
        let mut cells = Vec::new();
        let mut stories = Vec::new();
        let mut col = num(&table_row_grid(&row_payload.get("trPr"))?.get("before"));
        while col < total_columns {
            if let Some(&i) = by_start.get(&grid_key(ri, col)) {
                let anchor = &anchors[i];
                let formatting = anchor.cell.get("formatting").nullish_or(Js::new_obj()).spread();
                if anchor.colspan > 1.0 {
                    formatting.set("gridSpan", anchor.colspan);
                } else {
                    formatting.delete("gridSpan");
                }
                if anchor.rowspan > 1.0 {
                    formatting.set("vMerge", "restart");
                } else {
                    let v = formatting.get("vMerge");
                    if v.as_str() != Some("restart") && v.as_str() != Some("continue") {
                        formatting.delete("vMerge");
                    }
                }
                let formatting = if formatting.key_count() > 0 { formatting } else { Js::Undefined };
                cells.push(anchor.cell.with(&[("formatting", formatting)]));
                stories.push(Js::from(anchor.story.clone()));
                col += anchor.colspan;
                continue;
            }
            let covering = match by_covered.get(&grid_key(ri, col)) {
                Some(&i) => &anchors[i],
                None => {
                    col += 1.0;
                    continue;
                }
            };
            let formatting = covering.cell.get("formatting").nullish_or(Js::new_obj()).spread();
            if covering.colspan > 1.0 {
                formatting.set("gridSpan", covering.colspan);
            } else {
                formatting.delete("gridSpan");
            }
            formatting.set("vMerge", "continue");
            cells.push(covering.cell.with(&[("content", Js::new_arr(Vec::new())), ("formatting", formatting)]));
            stories.push(match &covering.story {
                None => Js::Undefined,
                Some(story) => obj! { "covering" => story.as_str(), "rowOffset" => ri - covering.row },
            });
            col += covering.colspan;
        }
        cell_stories.push(Js::new_arr(stories));

        let attrs = table_row_attr_defaults_b();
        attrs.assign(&row_payload.get("trPr").nullish_or(Js::new_obj()));
        let row = obj! { "type" => "tableRow" };
        if attrs.get("_extraAttributes").truthy() {
            row.set("extraAttributes", attrs.get("_extraAttributes"));
        }
        row.set("formatting", table_row_attrs_to_formatting(&attrs)?);
        row.set("cells", Js::new_arr(cells));
        let ins = tracked_info(&attrs.get("trIns"), true, None)?;
        let del = tracked_info(&attrs.get("trDel"), true, None)?;
        if ins.truthy() {
            row.set("structuralChange", obj! { "type" => "tableRowInsertion", "info" => ins });
        } else if del.truthy() {
            row.set("structuralChange", obj! { "type" => "tableRowDeletion", "info" => del });
        }
        let tr_pr_change = attrs.get("trPrChange");
        if tr_pr_change.is_arr() && tr_pr_change.len() > 0 {
            row.set("propertyChanges", tr_pr_change);
        }
        rows.push(row);
    }
    let rows = Js::new_arr(rows);
    let cell_stories = Js::new_arr(cell_stories);

    normalize_v_merge_runs(&rows)?;
    let grid_raw = payload.get("grid");
    let grid: Vec<Js> = if grid_raw.is_arr() { grid_raw.items().into_iter().filter(|w| w.is_num()).collect() } else { Vec::new() };
    let attrs = table_attr_defaults_b();
    attrs.assign(&payload.get("tblPr").nullish_or(Js::new_obj()));
    attrs.set("columnWidths", if grid.is_empty() { Js::Undefined } else { Js::new_arr(grid) });
    context.apply_table_baseline(payload, &attrs)?;
    let mut formatting = table_attrs_to_formatting(&attrs)?;
    if !formatting.get("borders").truthy() && !attrs.get("_originalFormatting").truthy() {
        let borders = infer_table_borders(&rows)?;
        if borders.truthy() {
            formatting = formatting.nullish_or(Js::new_obj()).with(&[("borders", borders)]);
        }
    }
    let column_widths = attrs.get("columnWidths");
    let table = obj! {
        "type" => "table",
        "columnWidths" => if column_widths.truthy() { column_widths } else { Js::Undefined },
        "formatting" => formatting,
        "rows" => rows,
    };
    let tbl_pr_change = attrs.get("tblPrChange");
    if tbl_pr_change.is_arr() && tbl_pr_change.len() > 0 {
        table.set("propertyChanges", tbl_pr_change);
    }
    let mut prefix_hint: Option<String> = None;
    'outer: for row in &row_payloads {
        let cells = row.get("cells");
        for cell in if cells.is_arr() { cells.items() } else { Vec::new() } {
            let story = match as_string(&cell.path(&["tcPr", "_sourceCellStory"]))? {
                Some(s) => Some(s),
                None => as_string(&cell.get("story"))?,
            };
            if let Some((prefix, _, _)) = story.as_deref().and_then(|s| match_cell_story(s, true)) {
                prefix_hint = Some(prefix);
                break 'outer;
            }
        }
    }
    restore_table_bookmarks(context, &table, &cell_stories, prefix_hint.as_deref())?;
    projected_table_cells_set(&table, cell_stories);
    Ok(table)
}

/// `tableCellKey` (src/yrs/yrsToDocument.ts:2652).
pub(crate) fn table_cell_key(entry: &Js) -> Result<Option<String>> {
    Ok(if let Some(s) = entry.as_str() {
        Some(s.to_string())
    } else if entry.truthy() {
        Some(format!("{}#{}", js::to_js_string(&entry.get("covering")), js::to_js_string(&entry.get("rowOffset"))))
    } else {
        None
    })
}

/// `baseTableCells` (src/yrs/yrsToDocument.ts:2657). A `Js::Map` of `{ row, cell }`.
pub(crate) fn base_table_cells(table: &Js, prefix: &str) -> Result<Js> {
    let cells = Js::new_map();
    if let Some(recorded) = projected_table_cells_get(table) {
        for (row_index, row) in recorded.items().iter().enumerate() {
            for (cell_index, entry) in row.items().iter().enumerate() {
                if let Some(key) = table_cell_key(entry)? {
                    if !cells.map_has(&key) {
                        cells.map_set(&key, obj! { "row" => row_index as f64, "cell" => cell_index as f64 });
                    }
                }
            }
        }
        return Ok(cells);
    }
    let seeded = seeded_cell_indices(table)?;
    let mut spans: HashMap<String, (String, usize)> = HashMap::new();
    for (row_index, row) in table.get("rows").items().iter().enumerate() {
        let starts = grid_starts(row)?;
        let seeded_row = seeded.at(row_index);
        for (cell_index, cell) in row.get("cells").items().iter().enumerate() {
            let ordinal = seeded_row
                .items()
                .iter()
                .position(|v| v.strict_eq(&Js::from(cell_index as f64)));
            let column = js::num_to_string(starts[cell_index]);
            if let Some(ordinal) = ordinal {
                let story = format!("{}:r{}c{}", prefix, row_index, ordinal);
                cells.map_set(&story, obj! { "row" => row_index as f64, "cell" => cell_index as f64 });
                if cell.path(&["formatting", "vMerge"]).as_str() == Some("restart") {
                    spans.insert(column, (story, row_index));
                } else {
                    spans.remove(&column);
                }
                continue;
            }
            if let Some((story, anchor_row)) = spans.get(&column) {
                cells.map_set(
                    &format!("{}#{}", story, row_index - anchor_row),
                    obj! { "row" => row_index as f64, "cell" => cell_index as f64 },
                );
            }
        }
    }
    Ok(cells)
}

/// `gridStarts` (src/yrs/yrsToDocument.ts:2688).
pub(crate) fn grid_starts(row: &Js) -> Result<Vec<f64>> {
    let mut column = num(&row.path(&["formatting", "gridBefore"]).nullish_or(Js::from(0.0)));
    Ok(row
        .get("cells")
        .items()
        .iter()
        .map(|cell| {
            let start = column;
            column += num(&cell.path(&["formatting", "gridSpan"]).nullish_or(Js::from(1.0)));
            start
        })
        .collect())
}

/// `rowCellBookmarks` (src/yrs/yrsToDocument.ts:2699).
pub(crate) fn row_cell_bookmarks(row: &Js) -> Result<Vec<String>> {
    let mut out = Vec::new();
    for cell in row.get("cells").items() {
        out.extend(str_items(&cell.get("leadingBookmarks")));
    }
    out.extend(str_items(&row.get("trailingBookmarks")));
    Ok(out)
}

/// `tableHasBookmarks` (src/yrs/yrsToDocument.ts:2703).
pub(crate) fn table_has_bookmarks(table: &Js) -> Result<bool> {
    if table.get("trailingBookmarks").len() > 0 {
        return Ok(true);
    }
    for row in table.get("rows").items() {
        if row.get("leadingBookmarks").len() > 0 || !row_cell_bookmarks(&row)?.is_empty() {
            return Ok(true);
        }
    }
    Ok(false)
}

/// `bookmarkMarker` (src/yrs/yrsToDocument.ts:2709).
pub(crate) fn bookmark_marker(key: &str, scan: &Js) -> Result<Js> {
    let marker = obj! { "type" => "paragraph", "paraId" => key, "content" => Js::new_arr(Vec::new()) };
    if scan.truthy() {
        marker.set("scan", scan.clone());
    }
    Ok(marker)
}

/// `distributeBookmarks` (src/yrs/yrsToDocument.ts:2713).
pub(crate) fn distribute_bookmarks(restored: &Js, items: &Js, leading: &dyn Fn(&Js, &str) -> Result<()>, trailing: &dyn Fn(&str) -> Result<()>) -> Result<()> {
    let mut index = 0usize;
    let len = items.len();
    for block in restored.items() {
        if is_raw_xml(&block)? {
            let xml = js::to_js_string(&block.get("xml"));
            if index < len {
                leading(&items.at(index), &xml)?;
            } else {
                trailing(&xml)?;
            }
        } else {
            index += 1;
        }
    }
    Ok(())
}

#[derive(Clone, Copy)]
enum Resolved {
    Direct { row: f64, cell: f64 },
    Cover { row: f64, column: f64 },
}

fn push_bookmark(target: &Js, key: &str, xml: &str) {
    if target.get(key).is_nullish() {
        target.set(key, Js::new_arr(Vec::new()));
    }
    target.get(key).push(xml);
}

fn index_of_f(values: Option<&Vec<f64>>, needle: f64) -> f64 {
    match values {
        None => -1.0,
        Some(v) => v.iter().position(|x| *x == needle).map_or(-1.0, |i| i as f64),
    }
}

fn row_starts(starts: &[Vec<f64>], row: f64) -> Option<&Vec<f64>> {
    if row >= 0.0 && row.fract() == 0.0 { starts.get(row as usize) } else { None }
}

fn norm_bits(x: f64) -> u64 {
    (x + 0.0).to_bits()
}

/// `restoreTableBookmarks` (src/yrs/yrsToDocument.ts:2731).
pub(crate) fn restore_table_bookmarks(context: &crate::yrs_to_document::SaveContext, table: &Js, cell_stories: &Js, prefix_hint: Option<&str>) -> Result<()> {
    let parse = |story: &str| match_cell_story(story, false).map(|(prefix, row, ordinal)| (prefix, row, ordinal));
    let prefix: Option<String> = match prefix_hint {
        Some(p) => Some(p.to_string()),
        None => cell_stories
            .items()
            .iter()
            .flat_map(|row| row.items())
            .find_map(|entry| {
                if let Some(s) = entry.as_str() {
                    parse(s).map(|p| p.0)
                } else if entry.truthy() {
                    parse(&js::to_js_string(&entry.get("covering"))).map(|p| p.0)
                } else {
                    None
                }
            }),
    };
    let prefix = match prefix {
        Some(p) => p,
        None => return Ok(()),
    };
    let source = context.source_table(&prefix)?;
    if !source.truthy() || !table_has_bookmarks(&source)? {
        return Ok(());
    }
    let located = base_table_cells(&source, &prefix)?;
    let source_rows = source.get("rows");
    let mut starts: Vec<Vec<f64>> = Vec::new();
    for row in source_rows.items() {
        starts.push(grid_starts(&row)?);
    }
    let continues = |row: f64, cell: f64| -> bool {
        at_f(&at_f(&source_rows, row).get("cells"), cell).path(&["formatting", "vMerge"]).as_str() == Some("continue")
    };
    let resolved: Vec<Vec<Option<Resolved>>> = cell_stories
        .items()
        .iter()
        .map(|row| {
            row.items()
                .iter()
                .map(|entry| {
                    if entry.is_undefined() {
                        return None;
                    }
                    if let Some(s) = entry.as_str() {
                        let at = located.map_get(s);
                        return if at.truthy() {
                            Some(Resolved::Direct { row: num(&at.get("row")), cell: num(&at.get("cell")) })
                        } else {
                            None
                        };
                    }
                    let anchor = located.map_get(&js::to_js_string(&entry.get("covering")));
                    if !anchor.truthy() {
                        return None;
                    }
                    let (ar, ac) = (num(&anchor.get("row")), num(&anchor.get("cell")));
                    let column = row_starts(&starts, ar)
                        .and_then(|s| if ac >= 0.0 { s.get(ac as usize).copied() } else { None })
                        .unwrap_or(f64::NAN);
                    Some(Resolved::Cover { row: ar + num(&entry.get("rowOffset")), column })
                })
                .collect()
        })
        .collect();
    let mut seen: HashSet<u64> = HashSet::new();
    let mut row_keys: Vec<Option<f64>> = Vec::new();
    for row in &resolved {
        let mut rows: indexmap::IndexSet<u64> = indexmap::IndexSet::new();
        for cell in row.iter().flatten() {
            if let Resolved::Direct { row, .. } = cell {
                rows.insert(norm_bits(*row));
            }
        }
        let source_row = if rows.len() == 1 { Some(f64::from_bits(rows[0])) } else { None };
        match source_row {
            Some(r) if !seen.contains(&norm_bits(r)) => {
                seen.insert(norm_bits(r));
                row_keys.push(Some(r));
            }
            _ => row_keys.push(None),
        }
    }
    for (index, row) in resolved.iter().enumerate() {
        if row_keys[index].is_some() || row.iter().any(|c| matches!(c, Some(Resolved::Direct { .. }))) {
            continue;
        }
        let mut rows: indexmap::IndexSet<u64> = indexmap::IndexSet::new();
        for cell in row.iter().flatten() {
            if let Resolved::Cover { row, column } = cell {
                if continues(*row, index_of_f(row_starts(&starts, *row), *column)) {
                    rows.insert(norm_bits(*row));
                }
            }
        }
        let source_row = if rows.len() == 1 { Some(f64::from_bits(rows[0])) } else { None };
        if let Some(r) = source_row {
            if !seen.contains(&norm_bits(r)) {
                seen.insert(norm_bits(r));
                row_keys[index] = Some(r);
            }
        }
    }
    let sources: Vec<Vec<Option<(f64, f64)>>> = resolved
        .iter()
        .enumerate()
        .map(|(index, row)| {
            row.iter()
                .map(|cell| match cell {
                    None => None,
                    Some(Resolved::Direct { row, cell }) => Some((*row, *cell)),
                    Some(Resolved::Cover { column, .. }) => {
                        let source_row = row_keys[index];
                        let covered = match source_row {
                            None => -1.0,
                            Some(r) => index_of_f(row_starts(&starts, r), *column),
                        };
                        if covered >= 0.0 && continues(source_row.unwrap(), covered) {
                            Some((source_row.unwrap(), covered))
                        } else {
                            None
                        }
                    }
                })
                .collect()
        })
        .collect();
    let content_of = |rows: &Js| -> Result<Js> {
        let mut all = Vec::new();
        for row in rows.items() {
            for cell in row.get("cells").items() {
                all.extend(cell.get("content").items());
            }
        }
        projected_bookmark_positions(&Js::new_arr(all))
    };
    let source_content = content_of(&source_rows)?;
    let table_rows = table.get("rows");
    let projected_content = content_of(&table_rows)?;
    let orphaned = |xml: &str| -> Result<bool> {
        let marker = raw_block_bookmark(xml)?;
        let id = marker.get("id");
        if id.is_undefined() {
            return Ok(false);
        }
        let id = js::to_js_string(&id);
        let partner = if marker.get("kind").as_str() == Some("start") { "end" } else { "start" };
        Ok(!source_content.map_get(&id).get(partner).is_undefined()
            && projected_content.map_get(&id).get(partner).is_undefined())
    };
    let raw = |xml: &str| -> Result<Vec<Js>> {
        Ok(if orphaned(xml)? { Vec::new() } else { vec![obj! { "type" => "rawXml", "xml" => xml }] })
    };
    let mut projected = Vec::new();
    for (index, row) in table_rows.items().iter().enumerate() {
        let key = match row_keys[index] {
            None => format!("N{}", index),
            Some(r) => format!("R{}", js::num_to_string(r)),
        };
        let mut scan = Vec::new();
        for cell in row.get("cells").items() {
            scan.extend(cell.get("content").items());
        }
        projected.push(bookmark_marker(&key, &Js::new_arr(scan))?);
    }
    let projected = Js::new_arr(projected);
    let placed = projected_bookmark_positions(&projected)?;
    let mut promoted: HashSet<String> = HashSet::new();
    for (index, source_index) in row_keys.iter().enumerate() {
        let source_index = match source_index {
            Some(s) => *s,
            None => continue,
        };
        for xml in row_cell_bookmarks(&at_f(&source_rows, source_index))? {
            let marker = raw_block_bookmark(&xml)?;
            let id = marker.get("id");
            if id.is_undefined() {
                continue;
            }
            let is_start = marker.get("kind").as_str() == Some("start");
            let partner = placed.map_get(&js::to_js_string(&id)).get(if is_start { "end" } else { "start" });
            if !partner.is_undefined() {
                let p = num(&partner);
                let mid = index as f64 + 0.5;
                if if is_start { p < mid } else { p > mid } {
                    promoted.insert(xml);
                }
            }
        }
    }
    let mut base = Vec::new();
    for (index, row) in source_rows.items().iter().enumerate() {
        for xml in str_items(&row.get("leadingBookmarks")) {
            base.extend(raw(&xml)?);
        }
        for xml in row_cell_bookmarks(row)? {
            if !seen.contains(&norm_bits(index as f64)) || promoted.contains(&xml) {
                base.extend(raw(&xml)?);
            }
        }
        base.push(bookmark_marker(&format!("R{}", index), &Js::Undefined)?);
    }
    for xml in str_items(&source.get("trailingBookmarks")) {
        base.extend(raw(&xml)?);
    }
    table.delete("trailingBookmarks");
    for row in table_rows.items() {
        row.delete("leadingBookmarks");
        row.delete("trailingBookmarks");
        for cell in row.get("cells").items() {
            cell.delete("leadingBookmarks");
        }
    }
    distribute_bookmarks(
        &restore_block_bookmarks(&projected, &Js::new_arr(base), &format!("{}:rows", prefix))?,
        &table_rows,
        &|row, xml| {
            push_bookmark(row, "leadingBookmarks", xml);
            Ok(())
        },
        &|xml| {
            push_bookmark(table, "trailingBookmarks", xml);
            Ok(())
        },
    )?;
    for (index, row) in table_rows.items().iter().enumerate() {
        let source_index = match row_keys[index] {
            Some(s) => s,
            None => continue,
        };
        let source_row = at_f(&source_rows, source_index);
        if row_cell_bookmarks(&source_row)?.is_empty() {
            continue;
        }
        let mut cell_base = Vec::new();
        let kept = |xml: &str| -> Result<Vec<Js>> { if promoted.contains(xml) { Ok(Vec::new()) } else { raw(xml) } };
        for (cell_index, cell) in source_row.get("cells").items().iter().enumerate() {
            for xml in str_items(&cell.get("leadingBookmarks")) {
                cell_base.extend(kept(&xml)?);
            }
            cell_base.push(bookmark_marker(&format!("C{}", cell_index), &Js::Undefined)?);
        }
        for xml in str_items(&source_row.get("trailingBookmarks")) {
            cell_base.extend(kept(&xml)?);
        }
        let mut cell_projected = Vec::new();
        for (cell_index, cell) in row.get("cells").items().iter().enumerate() {
            let key = sources[index].get(cell_index).copied().flatten();
            let name = match key {
                Some((r, c)) if r == source_index => format!("C{}", js::num_to_string(c)),
                _ => format!("N{}", cell_index),
            };
            cell_projected.push(bookmark_marker(&name, &cell.get("content"))?);
        }
        distribute_bookmarks(
            &restore_block_bookmarks(
                &Js::new_arr(cell_projected),
                &Js::new_arr(cell_base),
                &format!("{}:r{}", prefix, js::num_to_string(source_index)),
            )?,
            &row.get("cells"),
            &|cell, xml| {
                push_bookmark(cell, "leadingBookmarks", xml);
                Ok(())
            },
            &|xml| {
                push_bookmark(row, "trailingBookmarks", xml);
                Ok(())
            },
        )?;
    }
    Ok(())
}

/// `pageBreakParagraph` (src/yrs/yrsToDocument.ts:2851).
pub(crate) fn page_break_paragraph() -> Result<Js> {
    Ok(obj! {
        "type" => "paragraph",
        "content" => arr![obj! { "type" => "run", "content" => arr![obj! { "type" => "break", "breakType" => "page" }] }],
    })
}

/// `isNoteMark` (src/yrs/yrsToDocument.ts:2859).
pub(crate) fn is_note_mark(content: &Js) -> Result<bool> {
    let t = ty(content);
    let children = content.get("content");
    if matches!(t.as_str(), "insertion" | "deletion" | "moveFrom" | "moveTo") {
        if children.len() == 0 {
            return Ok(false);
        }
        for child in children.items() {
            if !is_note_mark(&child)? {
                return Ok(false);
            }
        }
        return Ok(true);
    }
    Ok(t == "run"
        && children.len() > 0
        && children.items().iter().all(|e| type_is(e, "footnoteRefMark") || type_is(e, "endnoteRefMark")))
}

/// `hasNoteMark` (src/yrs/yrsToDocument.ts:2877).
pub(crate) fn has_note_mark(blocks: &Js) -> Result<bool> {
    for block in blocks.items() {
        if type_is(&block, "paragraph") {
            for child in block.get("content").items() {
                if is_note_mark(&child)? {
                    return Ok(true);
                }
            }
        }
    }
    Ok(false)
}

/// `restoreNoteMarks` (src/yrs/yrsToDocument.ts:2886).
pub(crate) fn restore_note_marks(projected: &Js, base: &Js) -> Result<Js> {
    if has_note_mark(projected)? {
        return Ok(projected.clone());
    }
    let opening = base
        .items()
        .into_iter()
        .find(|b| type_is(b, "paragraph"))
        .map(|b| b.get("content"))
        .unwrap_or_default()
        .items();
    let mut end = None;
    for (i, child) in opening.iter().enumerate() {
        if !is_note_mark(child)? {
            end = Some(i);
            break;
        }
    }
    let marks: Vec<Js> = opening[..end.unwrap_or(opening.len())].to_vec();
    if marks.is_empty() {
        return Ok(projected.clone());
    }
    let index = match projected.items().iter().position(|b| type_is(b, "paragraph")) {
        Some(i) => i,
        None => return Ok(projected.clone()),
    };
    let target = projected.at(index);
    let restored = projected.copy_arr();
    let mut merged = marks;
    merged.extend(target.get("content").items());
    restored.set_at(index, clone_projected_paragraph(&target, &obj! { "content" => Js::new_arr(merged) }, &target)?);
    Ok(restored)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn same(actual: &Js, expected: &str) {
        assert_eq!(actual.stringify(), Js::parse(expected).unwrap().stringify());
    }

    fn sample() -> Js {
        Js::parse(r#"[{"type":"run","content":[{"type":"text","text":"ab😀c"}]},{"type":"hyperlink","href":"x","children":[{"type":"run","content":[{"type":"text","text":"def"}]}]},{"type":"insertion","info":{"id":1},"content":[{"type":"run","content":[{"type":"text","text":"gh"}]}]},{"type":"simpleField","content":[]},{"type":"run","content":[{"type":"text","text":"ij"}]}]"#).unwrap()
    }

    // Expectations from running the TS functions with bun.
    #[test]
    fn content_lengths() {
        let lengths: Vec<f64> = sample().items().iter().map(|c| paragraph_content_length(c).unwrap()).collect();
        assert_eq!(lengths, vec![5.0, 3.0, 2.0, 1.0, 2.0]);
    }

    #[test]
    fn split_content_halves() {
        let c = sample();
        same(&split_content(&c.at(1), 1.0).unwrap(), r#"[{"type":"hyperlink","href":"x","children":[{"type":"run","content":[{"type":"text","text":"d"}]}]},{"type":"hyperlink","href":"x","children":[{"type":"run","content":[{"type":"text","text":"ef"}]}]}]"#);
        same(&split_content(&c.at(2), 1.0).unwrap(), r#"[{"type":"insertion","info":{"id":1},"content":[{"type":"run","content":[{"type":"text","text":"g"}]}]},{"type":"insertion","info":{"id":1},"content":[{"type":"run","content":[{"type":"text","text":"h"}]}]}]"#);
        assert!(split_content(&c.at(3), 0.0).unwrap().is_null());
        same(&split_content(&c.at(0), 4.0).unwrap(), r#"[{"type":"run","content":[{"type":"text","text":"ab😀"}]},{"type":"run","content":[{"type":"text","text":"c"}]}]"#);
    }

    #[test]
    fn insert_boundaries_orders_markers() {
        let b = Js::parse(r#"[{"id":1,"kind":"start","offset":1},{"id":1,"kind":"end","offset":7},{"id":2,"kind":"start","offset":5},{"id":2,"kind":"end","offset":5},{"id":3,"kind":"end","offset":10},{"id":4,"kind":"start","offset":10},{"id":0,"kind":"start","offset":10}]"#).unwrap();
        let out = insert_boundaries(&sample(), &b, &comment_marker_default).unwrap();
        same(&out, r#"[{"type":"run","content":[{"type":"text","text":"a"}]},{"type":"commentRangeStart","id":1},{"type":"run","content":[{"type":"text","text":"b😀c"}]},{"type":"commentRangeStart","id":2},{"type":"commentRangeEnd","id":2},{"type":"hyperlink","href":"x","children":[{"type":"run","content":[{"type":"text","text":"de"}]}]},{"type":"commentRangeEnd","id":1},{"type":"hyperlink","href":"x","children":[{"type":"run","content":[{"type":"text","text":"f"}]}]},{"type":"insertion","info":{"id":1},"content":[{"type":"run","content":[{"type":"text","text":"gh"}]}]},{"type":"commentRangeEnd","id":3},{"type":"commentRangeStart","id":0},{"type":"commentRangeStart","id":4},{"type":"simpleField","content":[]},{"type":"run","content":[{"type":"text","text":"ij"}]}]"#);
        let empty = Js::new_arr(Vec::new());
        let c = sample();
        assert!(insert_boundaries(&c, &empty, &comment_marker_default).unwrap().strict_eq(&c));
    }

    #[test]
    fn edit_midpoints() {
        same(&edit_midpoint("abcdef", "azced").unwrap(), "[4,5]");
        same(&edit_midpoint("kitten", "sitting").unwrap(), "[4,5]");
        same(&edit_midpoint("a😀b", "ab").unwrap(), "[2,1]");
        same(&edit_midpoint("xyz", "q").unwrap(), "[2,0]");
    }

    #[test]
    fn text_matches() {
        same(&source_text_matches("hello world", "hello brave world").unwrap(), r#"[{"original":0,"current":0,"length":6},{"original":6,"current":12,"length":5}]"#);
        same(&source_text_matches("the quick brown fox", "a quick red fox!").unwrap(), r#"[{"original":3,"current":1,"length":7},{"original":11,"current":8,"length":1},{"original":15,"current":11,"length":4}]"#);
        same(&source_text_matches("a😀b", "ab😀").unwrap(), r#"[{"original":0,"current":0,"length":1},{"original":1,"current":2,"length":2}]"#);
    }

    #[test]
    fn cell_story_pattern() {
        assert_eq!(match_cell_story("body:t0:r1c2", false), Some(("body:t0".into(), "1".into(), "2".into())));
        assert_eq!(match_cell_story("a:t1:r0c0:t2:r3c4~5", true).map(|m| m.0), Some("a:t1:r0c0:t2".into()));
        assert_eq!(match_cell_story("a:t1:r0c0~5", false), None);
        assert_eq!(match_cell_story("t1:r0c0", true), None);
    }
}
