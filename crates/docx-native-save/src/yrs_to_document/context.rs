//! Port of `packages/docx/src/yrs/yrsToDocument.ts` (part).
#![allow(unused_variables, unused_imports, dead_code, clippy::all)]

use std::cell::RefCell;
use std::collections::{HashMap, HashSet};
use std::rc::Rc;
use indexmap::IndexMap;
use crate::js::{self, Js, Marks, Result, SaveError};
use crate::session::{Inheritance, Session};
use crate::{arr, obj, throw};
use super::*;

use crate::content_control_values::apply_content_control_value;
use crate::font_family_merge::merge_font_family;
use crate::header_footer_alias_projection::{header_footer_alias_of, mark_header_footer_alias};
use crate::sdt_attributes::sdt_attrs_to_props;
use crate::sections::section_properties_equal;
use crate::style_resolver::{create_style_resolver, StyleResolver};
use crate::table_row_spans::seeded_cell_indices;

// ----- small JS helpers -----

/// `asObject(value)`: a non-null, non-array object.
fn as_obj_c(value: &Js) -> Option<Js> {
    match value {
        Js::Obj(_) | Js::Map(_) | Js::Bytes(_) => Some(value.clone()),
        _ => None,
    }
}

fn as_str_c(value: &Js) -> Option<String> {
    value.as_str().map(str::to_owned)
}

fn is_raw_xml_c(block: &Js) -> bool {
    block.get("type").as_str() == Some("rawXml")
}

fn type_is(value: &Js, ty: &str) -> bool {
    value.get("type").as_str() == Some(ty)
}

/// JS `\s`.
fn is_js_space(c: char) -> bool {
    matches!(
        c,
        '\t' | '\n' | '\u{0B}' | '\u{0C}' | '\r' | ' ' | '\u{A0}' | '\u{1680}' | '\u{2000}'..='\u{200A}'
            | '\u{2028}' | '\u{2029}' | '\u{202F}' | '\u{205F}' | '\u{3000}' | '\u{FEFF}'
    )
}

fn js_trim_start(s: &str) -> &str {
    s.trim_start_matches(is_js_space)
}

fn is_line_terminator(c: char) -> bool {
    matches!(c, '\n' | '\r' | '\u{2028}' | '\u{2029}')
}

/// Consumes `\d+` at the start of `s`; the rest, or `None` without a digit.
fn digits(s: &str) -> Option<&str> {
    let n = s.bytes().take_while(u8::is_ascii_digit).count();
    (n > 0).then(|| &s[n..])
}

/// `s` matches `:r\d+c\d+(?:~\d+)?$`.
fn cell_suffix(s: &str) -> bool {
    let Some(rest) = s.strip_prefix(":r").and_then(digits) else { return false };
    let Some(rest) = rest.strip_prefix('c').and_then(digits) else { return false };
    if rest.is_empty() {
        return true;
    }
    rest.strip_prefix('~').and_then(digits).map_or(false, str::is_empty)
}

/// `NESTED_STORY_ID.exec(story)?.[1]`: `/^(.*?)(?::t\d+:r\d+c\d+(?:~\d+)?|:sdt\d+)$/`.
fn nested_story_parent(story: &str) -> Option<String> {
    for (i, c) in story.char_indices().chain(std::iter::once((story.len(), '\0'))) {
        let rest = &story[i..];
        let table = rest
            .strip_prefix(":t")
            .and_then(digits)
            .map_or(false, cell_suffix);
        let sdt = rest
            .strip_prefix(":sdt")
            .and_then(digits)
            .map_or(false, str::is_empty);
        if table || sdt {
            return Some(story[..i].to_owned());
        }
        if i < story.len() && is_line_terminator(c) {
            return None;
        }
    }
    None
}

/// `story?.match(/^(.*:t\d+):r\d+c\d+(?:~\d+)?$/)?.[1]`.
fn table_prefix_of(story: &str) -> Option<String> {
    let mut ends: Vec<usize> = story.char_indices().map(|(i, _)| i).collect();
    ends.push(story.len());
    for &i in ends.iter().rev() {
        if !cell_suffix(&story[i..]) {
            continue;
        }
        let head = &story[..i];
        if head.chars().any(is_line_terminator) {
            continue;
        }
        // `head` must end with `:t\d+`.
        let trimmed = head.trim_end_matches(|c: char| c.is_ascii_digit());
        if trimmed.len() < head.len() && trimmed.ends_with(":t") {
            return Some(head.to_owned());
        }
    }
    None
}

/// `/^<w:bookmark(Start|End)(?=[\s/>]|$)/.exec(xml)?.[1]` over an already trimmed string.
fn bookmark_kind(xml: &str) -> Option<&'static str> {
    for (name, kind) in [("<w:bookmarkStart", "Start"), ("<w:bookmarkEnd", "End")] {
        if let Some(rest) = xml.strip_prefix(name) {
            match rest.chars().next() {
                None => return Some(kind),
                Some(c) if is_js_space(c) || c == '/' || c == '>' => return Some(kind),
                _ => {}
            }
        }
    }
    None
}

fn js_slice(arr: &Js, start: f64, end: Option<f64>) -> Js {
    let items = arr.items();
    let len = items.len() as f64;
    let norm = |v: f64| if v < 0.0 { (len + v).max(0.0) } else { v.min(len) };
    let s = norm(start);
    let e = norm(end.unwrap_or(len));
    if e <= s {
        return Js::new_arr(Vec::new());
    }
    Js::new_arr(items[s as usize..e as usize].to_vec())
}

fn js_min(values: &[f64]) -> f64 {
    let mut out = f64::INFINITY;
    for &v in values {
        if v.is_nan() {
            return f64::NAN;
        }
        if v < out {
            out = v;
        }
    }
    out
}

fn js_max(values: &[f64]) -> f64 {
    let mut out = f64::NEG_INFINITY;
    for &v in values {
        if v.is_nan() {
            return f64::NAN;
        }
        if v > out {
            out = v;
        }
    }
    out
}

/// A `Map` key for a JS value (numbers by their string form).
fn key_of(value: &Js) -> String {
    js::to_js_string(value)
}

fn map_delete(map: &Js, key: &str) {
    if let Some(m) = map.as_map() {
        m.0.borrow_mut().shift_remove(key);
    }
}

fn map_clear(map: &Js) {
    if let Some(m) = map.as_map() {
        m.0.borrow_mut().clear();
    }
}

fn str_arr(values: &[String]) -> Js {
    Js::new_arr(values.iter().map(|v| Js::from(v.as_str())).collect())
}

fn arr_strings(value: &Js) -> Vec<String> {
    value.items().iter().map(|v| v.str_or_empty().to_owned()).collect()
}

/// `TABLE_CELL_ATTR_DEFAULTS` (src/yrs/yrsToDocument.ts:205).
fn table_cell_attr_defaults_c() -> Js {
    obj! {
        "colspan" => 1,
        "rowspan" => 1,
        "colwidth" => Js::Null,
        "width" => Js::Null,
        "widthType" => Js::Null,
        "verticalAlign" => Js::Null,
        "backgroundColor" => Js::Null,
        "borders" => Js::Null,
        "margins" => Js::Null,
        "textDirection" => Js::Null,
        "noWrap" => false,
        "_originalFormatting" => Js::Null,
        "_originalResolvedFill" => Js::Null,
        "cellMarker" => Js::Null,
        "tcPrChange" => Js::Null,
    }
}

const WORD_NAMESPACE_C: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

// ----- projection memo -----

/// `ProjectedBlockMemo` (src/yrs/yrsToDocument.ts:3306).
#[derive(Clone, Default)]
pub(crate) struct ProjectedBlockMemo {
    pub(crate) table_baseline: Option<Js>,
    pub(crate) table_source_prefix: Option<String>,
    pub(crate) key: Option<String>,
    pub(crate) inputs: Option<Js>,
    pub(crate) children: Option<Vec<Js>>,
    pub(crate) session_key: Option<String>,
    pub(crate) trailing: Option<Vec<String>>,
    /// `undefined` when absent.
    pub(crate) source_runs: Js,
}

thread_local! {
    /// `projectedBlocks` (src/yrs/yrsToDocument.ts:3339).
    static PROJECTED_BLOCKS: Marks<ProjectedBlockMemo> = Marks::new();
    /// `sessionProjectionMemos`, keyed by the session's address for the duration of a projection.
    static SESSION_MEMOS: RefCell<HashMap<usize, Js>> = RefCell::new(HashMap::new());
    /// `EMPTY_BLOCKS` (one array, compared by identity).
    static EMPTY_BLOCKS: Js = Js::new_arr(Vec::new());
}

fn empty_blocks() -> Js {
    EMPTY_BLOCKS.with(Js::clone)
}

/// `projectedBlocks.get(block)`.
pub(crate) fn projected_blocks_get(block: &Js) -> Option<ProjectedBlockMemo> {
    PROJECTED_BLOCKS.with(|m| m.get(block))
}

/// `projectedBlocks.set(block, memo)`.
pub(crate) fn projected_blocks_set(block: &Js, memo: ProjectedBlockMemo) {
    PROJECTED_BLOCKS.with(|m| m.set(block, memo))
}

/// `projectedBlocks.get(paragraph)?.sourceRuns` (read by `restoreOriginalRuns`).
pub(crate) fn projected_block_source_runs(paragraph: &Js) -> Js {
    projected_blocks_get(paragraph).map(|m| m.source_runs).unwrap_or_default()
}

fn session_key_of(session: &Session) -> usize {
    session as *const Session as *const () as usize
}

fn clear_session_projection_memo(session: &Session) {
    SESSION_MEMOS.with(|m| {
        m.borrow_mut().remove(&session_key_of(session));
    });
}

// ----- chart recovery -----

/// `ChartRecovery.drawingFor` of a recovery built by [`SaveContext::recovery`].
pub(crate) fn recovery_drawing_for(recovery: &Js, r_id: Option<&str>, path: Option<&str>) -> Option<String> {
    let scoped = recovery.get("scoped");
    let by_rid = match r_id {
        Some(r_id) if scoped.is_map() => scoped.map_get(&format!("rId:{r_id}")),
        _ => Js::Undefined,
    };
    let found = if by_rid.is_nullish() {
        match path {
            Some(path) if scoped.is_map() => scoped.map_get(&format!("path:{path}")),
            _ => Js::Undefined,
        }
    } else {
        by_rid
    };
    found.as_str().map(str::to_owned)
}

/// `ChartRecovery.unrecoverable`.
pub(crate) fn recovery_unrecoverable(recovery: &Js, r_id: Option<&str>, path: Option<&str>) {
    let identity = match (r_id, path) {
        (Some(r_id), _) => format!(" (rId {r_id})"),
        (None, Some(path)) => format!(" (path {path})"),
        _ => String::new(),
    };
    recovery.get("warnings").push(format!(
        "chart run carries no drawing to replay{identity}; keeping the run out of the output"
    ));
}

/// `ChartRecovery.warn`.
pub(crate) fn recovery_warn(recovery: &Js, message: &str) {
    recovery.get("warnings").push(message);
}

// ----- base collection -----

fn collect_base_paragraphs_map(stories: &IndexMap<String, Js>) -> IndexMap<String, Js> {
    let mut paragraphs = IndexMap::new();
    for (story_id, blocks) in stories {
        let mut index = 0;
        for block in blocks.items() {
            if !type_is(&block, "paragraph") {
                continue;
            }
            let key = projected_blocks_get(&block)
                .and_then(|m| m.session_key)
                .unwrap_or_else(|| {
                    let pid = if !block.get("repeatedParaId").truthy() {
                        block.get("paraId")
                    } else {
                        Js::Bool(false)
                    };
                    if pid.truthy() {
                        key_of(&pid)
                    } else {
                        format!("{story_id}:p{index}")
                    }
                });
            index += 1;
            if !paragraphs.contains_key(&key) {
                paragraphs.insert(key, block);
            }
        }
    }
    paragraphs
}

/// `collectBaseParagraphs` (src/yrs/yrsToDocument.ts:2908).
pub(crate) fn collect_base_paragraphs(stories: &Js) -> Result<Js> {
    let map: IndexMap<String, Js> = stories.map_entries().into_iter().collect();
    Ok(Js::map_from(collect_base_paragraphs_map(&map)))
}

fn collect_base_stories_map(session: &Session, document: &Js) -> Result<IndexMap<String, Js>> {
    fn visit(stories: &mut IndexMap<String, Js>, story_id: String, blocks: Js) -> Result<()> {
        if stories.contains_key(&story_id) {
            return Ok(());
        }
        stories.insert(story_id.clone(), blocks.clone());
        let mut table_index = 0;
        let mut sdt_index = 0;
        for block in blocks.items() {
            if type_is(&block, "blockSdt") {
                let id = format!("{story_id}:sdt{sdt_index}");
                sdt_index += 1;
                visit(stories, id, block.get("content"))?;
                continue;
            }
            if !type_is(&block, "table") {
                continue;
            }
            let current_table_index = table_index;
            table_index += 1;
            if let Some(recorded) = projected_table_cells_get(&block) {
                for (row_index, row) in recorded.items().iter().enumerate() {
                    for (cell_index, entry) in row.items().iter().enumerate() {
                        if let Some(entry) = entry.as_str() {
                            let content = block
                                .get("rows")
                                .at(row_index)
                                .get("cells")
                                .at(cell_index)
                                .get("content");
                            visit(stories, entry.to_owned(), content)?;
                        }
                    }
                }
                continue;
            }
            for (row_index, cells) in seeded_cell_indices(&block)?.items().iter().enumerate() {
                for (ordinal, cell_index) in cells.items().iter().enumerate() {
                    let cell_index = js::to_number(cell_index) as usize;
                    let content = block
                        .get("rows")
                        .at(row_index)
                        .get("cells")
                        .at(cell_index)
                        .get("content");
                    visit(
                        stories,
                        format!("{story_id}:t{current_table_index}:r{row_index}c{ordinal}"),
                        content,
                    )?;
                }
            }
        }
        Ok(())
    }

    let mut stories = IndexMap::new();
    let package = document.get("package");
    visit(&mut stories, "body".into(), package.path(&["document", "content"]))?;
    for (r_id, part) in package.get("headers").map_entries() {
        visit(&mut stories, session.header_footer_story(&r_id), part.get("content"))?;
    }
    for (r_id, part) in package.get("footers").map_entries() {
        visit(&mut stories, session.header_footer_story(&r_id), part.get("content"))?;
    }
    for note in package.get("footnotes").items() {
        visit(&mut stories, format!("fn:{}", key_of(&note.get("id"))), note.get("content"))?;
    }
    for note in package.get("endnotes").items() {
        visit(&mut stories, format!("en:{}", key_of(&note.get("id"))), note.get("content"))?;
    }
    Ok(stories)
}

/// `collectBaseStories` (src/yrs/yrsToDocument.ts:2926).
pub(crate) fn collect_base_stories(session: &Session, document: &Js) -> Result<Js> {
    Ok(Js::map_from(collect_base_stories_map(session, document)?))
}

fn comment_ranges_map(session: &Session, comments: &Js, session_ids: &Js) -> Result<IndexMap<String, Js>> {
    let mut by_story: IndexMap<String, Js> = IndexMap::new();
    for comment in comments.items() {
        let id = comment.get("id");
        let mapped = if session_ids.is_map() {
            session_ids.map_get(&key_of(&id))
        } else {
            Js::Undefined
        };
        let lookup = if mapped.is_nullish() { key_of(&id) } else { key_of(&mapped) };
        let anchors = match session.resolve_comment(&lookup) {
            Ok(anchors) => anchors,
            Err(_) => continue,
        };
        let mut story_groups: IndexMap<String, Vec<Js>> = IndexMap::new();
        for anchor in anchors.items() {
            story_groups.entry(key_of(&anchor.get("story"))).or_default().push(anchor);
        }
        for (story, group) in story_groups {
            if group.is_empty() {
                continue;
            }
            let starts: Vec<f64> = group.iter().map(|a| js::to_number(&a.get("start"))).collect();
            let ends: Vec<f64> = group.iter().map(|a| js::to_number(&a.get("end"))).collect();
            let start = js_min(&starts);
            let end = js_max(&ends);
            let ranges = by_story
                .entry(story)
                .or_insert_with(|| Js::new_arr(Vec::new()))
                .clone();
            ranges.push(obj! { "id" => id.clone(), "start" => start, "end" => end });
        }
    }
    Ok(by_story)
}

/// `commentRanges` (src/yrs/yrsToDocument.ts:2970).
pub(crate) fn comment_ranges(session: &Session, comments: &Js, session_ids: &Js) -> Result<Js> {
    Ok(Js::map_from(comment_ranges_map(session, comments, session_ids)?))
}

/// `collectChartDrawings` (src/yrs/yrsToDocument.ts:3008). `into` is a `Js::Map`.
pub(crate) fn collect_chart_drawings(blocks: &Js, into: &Js) -> Result<()> {
    for block in blocks.items() {
        if type_is(&block, "paragraph") {
            for run in paragraph_chart_runs(&block.get("content"))?.items() {
                for content in run.get("content").items() {
                    let chart = content.get("chart");
                    if !type_is(&content, "chart") || !chart.get("drawingXml").truthy() {
                        continue;
                    }
                    if !chart.get("rId").is_undefined() {
                        let key = format!("rId:{}", key_of(&chart.get("rId")));
                        if !into.map_has(&key) {
                            into.map_set(&key, chart.get("drawingXml"));
                        }
                    }
                    if !chart.get("path").is_undefined() {
                        let key = format!("path:{}", key_of(&chart.get("path")));
                        if !into.map_has(&key) {
                            into.map_set(&key, chart.get("drawingXml"));
                        }
                    }
                }
            }
        } else if type_is(&block, "table") {
            for row in block.get("rows").items() {
                for cell in row.get("cells").items() {
                    collect_chart_drawings(&cell.get("content"), into)?;
                }
            }
        } else if type_is(&block, "blockSdt") {
            collect_chart_drawings(&block.get("content"), into)?;
        }
    }
    Ok(())
}

/// `paragraphChartRuns` (src/yrs/yrsToDocument.ts:3032).
pub(crate) fn paragraph_chart_runs(content: &Js) -> Result<Js> {
    fn visit_inline(child: &Js, runs: &Js) {
        let ty = child.get("type");
        match ty.as_str() {
            Some("run") => runs.push(child.clone()),
            Some("hyperlink") => {
                for inline in child.get("structuredChildren").nullish_or(child.get("children")).items() {
                    visit_inline(&inline, runs);
                }
            }
            Some("inlineSdt") => {
                for inline in child.get("content").items() {
                    visit_inline(&inline, runs);
                }
            }
            Some("simpleField") => {
                for run in child.get("content").items() {
                    if type_is(&run, "run") {
                        runs.push(run);
                    }
                }
            }
            Some("complexField") => {
                for run in child.get("fieldCode").items().into_iter().chain(child.get("fieldResult").items()) {
                    runs.push(run);
                }
            }
            Some("insertion") | Some("deletion") | Some("moveFrom") | Some("moveTo") => {
                for inline in child.get("content").items() {
                    visit_inline(&inline, runs);
                }
            }
            _ => {}
        }
    }
    let runs = Js::new_arr(Vec::new());
    for child in content.items() {
        visit_inline(&child, &runs);
    }
    Ok(runs)
}

// ----- block bookmarks -----

/// `isBlockBookmark` (src/yrs/yrsToDocument.ts:3055).
pub(crate) fn is_block_bookmark(block: &Js) -> Result<bool> {
    Ok(is_raw_xml_c(block) && bookmark_kind(js_trim_start(block.get("xml").str_or_empty())).is_some())
}

/// `containsBookmark` (src/yrs/yrsToDocument.ts:3059).
pub(crate) fn contains_bookmark(blocks: &Js) -> Result<bool> {
    for block in blocks.items() {
        if is_block_bookmark(&block)? {
            return Ok(true);
        }
        if type_is(&block, "blockSdt") {
            if contains_bookmark(&block.get("content"))? {
                return Ok(true);
            }
            continue;
        }
        if type_is(&block, "table") {
            if table_has_bookmarks(&block)? {
                return Ok(true);
            }
            for row in block.get("rows").items() {
                for cell in row.get("cells").items() {
                    if contains_bookmark(&cell.get("content"))? {
                        return Ok(true);
                    }
                }
            }
        }
    }
    Ok(false)
}

/// `assertCommentBookmarkReplacement` (src/yrs/yrsToDocument.ts:3068).
pub(crate) fn assert_comment_bookmark_replacement(blocks: &Js) -> Result<()> {
    for block in blocks.items() {
        let hit = if type_is(&block, "blockSdt") {
            contains_bookmark(&block.get("content"))?
        } else {
            type_is(&block, "table") && contains_bookmark(&arr![block.clone()])?
        };
        if hit {
            throw!("Cannot replace comment body with bookmarks in nested blocks");
        }
    }
    Ok(())
}

/// The entity decoding of `rawBlockBookmark`.
fn decode_entities(value: &str) -> String {
    let chars: Vec<char> = value.chars().collect();
    let mut out = String::new();
    let mut i = 0;
    'outer: while i < chars.len() {
        if chars[i] == '&' {
            // `&#x([0-9a-fA-F]+);` | `&#([0-9]+);`
            if chars.get(i + 1) == Some(&'#') {
                let (hex, start) = if chars.get(i + 2) == Some(&'x') { (true, i + 3) } else { (false, i + 2) };
                let mut j = start;
                while j < chars.len()
                    && (if hex { chars[j].is_ascii_hexdigit() } else { chars[j].is_ascii_digit() })
                {
                    j += 1;
                }
                if j > start && chars.get(j) == Some(&';') {
                    let digits: String = chars[start..j].iter().collect();
                    let radix = if hex { 16.0 } else { 10.0 };
                    let mut code = 0f64;
                    for d in digits.chars() {
                        code = code * radix + d.to_digit(16).unwrap() as f64;
                    }
                    if code <= 0x10ffff as f64 {
                        // A lone surrogate cannot live in a Rust string: U+FFFD stands in.
                        out.push(char::from_u32(code as u32).unwrap_or('\u{FFFD}'));
                    } else {
                        out.extend(&chars[i..=j]);
                    }
                    i = j + 1;
                    continue;
                }
            }
            for (name, ch) in [("amp", '&'), ("lt", '<'), ("gt", '>'), ("quot", '"'), ("apos", '\'')] {
                let n: Vec<char> = name.chars().collect();
                let end = i + 1 + n.len();
                if end < chars.len() && chars[i + 1..end] == n[..] && chars[end] == ';' {
                    out.push(ch);
                    i = end + 1;
                    continue 'outer;
                }
            }
        }
        out.push(chars[i]);
        i += 1;
    }
    out
}

/// `/^<[^<>"']*(?:"[^"]*"[^<>"']*|'[^']*'[^<>"']*)*>/.exec(xml)?.[0]`.
fn leading_tag(xml: &str) -> Option<&str> {
    let mut chars = xml.char_indices();
    match chars.next() {
        Some((_, '<')) => {}
        _ => return None,
    }
    let mut quote: Option<char> = None;
    for (i, c) in chars {
        match quote {
            Some(q) => {
                if c == q {
                    quote = None;
                }
            }
            None => match c {
                '>' => return Some(&xml[..i + 1]),
                '<' => return None,
                '"' | '\'' => quote = Some(c),
                _ => {}
            },
        }
    }
    None
}

/// `tag.matchAll(/([^\s=<>/]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g)`: name and raw value.
fn tag_attributes(tag: &str) -> Vec<(String, String)> {
    let chars: Vec<char> = tag.chars().collect();
    let is_name = |c: char| !(is_js_space(c) || matches!(c, '=' | '<' | '>' | '/'));
    let mut out = Vec::new();
    let mut i = 0;
    while i < chars.len() {
        let try_at = |i: usize| -> Option<(String, String, usize)> {
            let mut j = i;
            while j < chars.len() && is_name(chars[j]) {
                j += 1;
            }
            if j == i {
                return None;
            }
            let name: String = chars[i..j].iter().collect();
            while j < chars.len() && is_js_space(chars[j]) {
                j += 1;
            }
            if chars.get(j) != Some(&'=') {
                return None;
            }
            j += 1;
            while j < chars.len() && is_js_space(chars[j]) {
                j += 1;
            }
            let q = *chars.get(j)?;
            if q != '"' && q != '\'' {
                return None;
            }
            let start = j + 1;
            let mut k = start;
            while k < chars.len() && chars[k] != q {
                k += 1;
            }
            if k >= chars.len() {
                return None;
            }
            Some((name, chars[start..k].iter().collect(), k + 1))
        };
        match try_at(i) {
            Some((name, value, next)) => {
                out.push((name, value));
                i = next;
            }
            None => i += 1,
        }
    }
    out
}

fn raw_block_bookmark_parts(xml: &str) -> Result<Option<(&'static str, Option<String>)>> {
    let xml = js_trim_start(xml);
    let Some(kind) = bookmark_kind(xml) else { return Ok(None) };
    let Some(tag) = leading_tag(xml) else { throw!("Invalid block bookmark") };
    let attributes: Vec<(String, String)> = tag_attributes(tag)
        .into_iter()
        .map(|(name, value)| (name, decode_entities(&value)))
        .collect();
    let mut namespaces: IndexMap<String, String> = IndexMap::new();
    namespaces.insert("w".into(), WORD_NAMESPACE_C.into());
    namespaces.insert("xml".into(), "http://www.w3.org/XML/1998/namespace".into());
    for (name, value) in &attributes {
        if let Some(prefix) = name.strip_prefix("xmlns:") {
            namespaces.insert(prefix.to_owned(), value.clone());
        }
    }
    let mut value: Option<String> = None;
    for (name, attribute) in &attributes {
        let mut parts = name.split(':');
        let prefix = parts.next().unwrap_or("");
        let local = parts.next();
        if prefix == "xmlns" || (local != Some("id") && local != Some("name")) {
            continue;
        }
        let Some(namespace) = namespaces.get(prefix) else {
            throw!("Unbound block bookmark attribute {name}");
        };
        if namespace != WORD_NAMESPACE_C || local != Some("id") {
            continue;
        }
        if value.is_some() {
            throw!("Duplicate block bookmark id");
        }
        value = Some(attribute.clone());
    }
    let id = value.map(|value| {
        let n = js::to_number(&Js::from(value.as_str()));
        if n.is_finite() && n.fract() == 0.0 && n.abs() <= 9007199254740991.0 {
            js::num_to_string(n)
        } else {
            value
        }
    });
    Ok(Some((if kind == "Start" { "start" } else { "end" }, id)))
}

/// `rawBlockBookmark` (src/yrs/yrsToDocument.ts:3078). `{ kind, id? }` or `undefined`.
pub(crate) fn raw_block_bookmark(xml: &str) -> Result<Js> {
    Ok(match raw_block_bookmark_parts(xml)? {
        None => Js::Undefined,
        Some((kind, id)) => obj! { "kind" => kind, "id" => Js::from(id) },
    })
}

fn story_block_keys_vec(blocks: &Js, story_id: &str) -> Vec<Option<String>> {
    let mut paragraph = 0;
    let mut table = 0;
    let mut sdt = 0;
    blocks
        .items()
        .iter()
        .map(|block| {
            let memo = projected_blocks_get(block);
            if type_is(block, "paragraph") {
                let key = memo.and_then(|m| m.session_key).unwrap_or_else(|| {
                    let pid = if !block.get("repeatedParaId").truthy() { block.get("paraId") } else { Js::Bool(false) };
                    if pid.truthy() { key_of(&pid) } else { format!("{story_id}:p{paragraph}") }
                });
                paragraph += 1;
                return Some(format!("P{key}"));
            }
            if type_is(block, "table") {
                let key = memo
                    .and_then(|m| m.table_source_prefix)
                    .unwrap_or_else(|| format!("{story_id}:t{table}"));
                table += 1;
                return Some(format!("T{key}"));
            }
            if type_is(block, "blockSdt") {
                let key = memo.and_then(|m| m.key).unwrap_or_else(|| format!("S{story_id}:sdt{sdt}"));
                sdt += 1;
                return Some(key);
            }
            None
        })
        .collect()
}

/// `storyBlockKeys` (src/yrs/yrsToDocument.ts:3110).
pub(crate) fn story_block_keys(blocks: &Js, story_id: &str) -> Result<Js> {
    Ok(Js::new_arr(
        story_block_keys_vec(blocks, story_id).into_iter().map(Js::from).collect(),
    ))
}

type Pair = (Option<f64>, Option<f64>);

fn projected_bookmark_positions_map(blocks: &Js) -> Result<IndexMap<String, Pair>> {
    fn record(positions: &mut IndexMap<String, Pair>, kind: &str, id: String, position: f64) {
        let pair = positions.entry(id).or_insert((None, None));
        let slot = if kind == "start" { &mut pair.0 } else { &mut pair.1 };
        if slot.is_none() {
            *slot = Some(position);
        }
    }
    fn visit(value: &Js, position: f64, visited: &mut HashSet<usize>, positions: &mut IndexMap<String, Pair>) -> Result<()> {
        if !value.is_object_like() {
            return Ok(());
        }
        let p = value.ptr().unwrap();
        if !visited.insert(p) {
            return Ok(());
        }
        if value.is_arr() {
            for child in value.items() {
                visit(&child, position, visited, positions)?;
            }
            return Ok(());
        }
        let ty = value.get("type");
        let marker: Option<(&str, Option<String>)> = if ty.as_str() == Some("rawXml") && value.get("xml").is_str() {
            raw_block_bookmark_parts(value.get("xml").str_or_empty())?
        } else if ty.as_str() == Some("bookmarkStart") || ty.as_str() == Some("bookmarkEnd") {
            Some((
                if ty.as_str() == Some("bookmarkStart") { "start" } else { "end" },
                Some(key_of(&value.get("id"))),
            ))
        } else {
            None
        };
        if let Some((kind, Some(id))) = marker {
            record(positions, kind, id, position);
            return Ok(());
        }
        for key in ["leadingBookmarks", "trailingBookmarks"] {
            let fragments = value.get(key);
            if !fragments.is_arr() {
                continue;
            }
            for xml in fragments.items() {
                let boundary = match xml.as_str() {
                    Some(xml) => raw_block_bookmark_parts(xml)?,
                    None => None,
                };
                if let Some((kind, Some(id))) = boundary {
                    record(positions, kind, id, position);
                }
            }
        }
        for (_, child) in value.entries() {
            visit(&child, position, visited, positions)?;
        }
        Ok(())
    }
    let mut positions = IndexMap::new();
    let mut visited = HashSet::new();
    for (index, block) in blocks.items().iter().enumerate() {
        visit(block, index as f64 + 0.5, &mut visited, &mut positions)?;
    }
    Ok(positions)
}

/// `projectedBookmarkPositions` (src/yrs/yrsToDocument.ts:3135). A `Js::Map` of id to `{ start?, end? }`.
pub(crate) fn projected_bookmark_positions(blocks: &Js) -> Result<Js> {
    let map = projected_bookmark_positions_map(blocks)?;
    Ok(Js::map_from(map.into_iter().map(|(id, (start, end))| {
        let pair = Js::new_obj();
        if let Some(s) = start {
            pair.set("start", s);
        }
        if let Some(e) = end {
            pair.set("end", e);
        }
        (id, pair)
    })))
}

struct Placement {
    block: Js,
    index: usize,
    position: f64,
    omit: bool,
    following: Option<Js>,
}

/// `restoreBlockBookmarks` (src/yrs/yrsToDocument.ts:3182).
pub(crate) fn restore_block_bookmarks(projected: &Js, base: &Js, story_id: &str) -> Result<Js> {
    let base_items = base.items();
    let mut any = false;
    for block in &base_items {
        if is_block_bookmark(block)? {
            any = true;
            break;
        }
    }
    if !any {
        return Ok(projected.clone());
    }
    let keys = story_block_keys_vec(base, story_id);
    let mut live: HashMap<String, usize> = HashMap::new();
    for (index, key) in story_block_keys_vec(projected, story_id).into_iter().enumerate() {
        if let Some(key) = key {
            live.insert(key, index);
        }
    }
    let mut next_positions: Vec<Option<usize>> = vec![None; base_items.len()];
    let mut next: Option<usize> = None;
    for index in (0..base_items.len()).rev() {
        next_positions[index] = next;
        if let Some(position) = keys[index].as_ref().and_then(|k| live.get(k)) {
            next = Some(*position);
        }
    }
    let present = projected_bookmark_positions_map(projected)?;
    let mut pairs: IndexMap<String, (Option<usize>, Option<usize>)> = IndexMap::new();
    let mut boundaries: Vec<Placement> = Vec::new();
    let mut previous: Option<usize> = None;
    for index in 0..base_items.len() {
        let block = &base_items[index];
        let marker = if is_raw_xml_c(block) {
            raw_block_bookmark_parts(block.get("xml").str_or_empty())?
        } else {
            None
        };
        if let Some((kind, id)) = marker {
            let id = id.unwrap_or_else(|| format!("\0{index}"));
            let pair = pairs.entry(id.clone()).or_insert((None, None));
            let slot = if kind == "start" { &mut pair.0 } else { &mut pair.1 };
            if slot.is_some() {
                throw!("Duplicate block bookmark {kind} for id {id}");
            }
            let existing = present.get(&id);
            let omit = match existing {
                Some((s, e)) => if kind == "start" { s.is_some() } else { e.is_some() },
                None => false,
            };
            let position = match next_positions[index] {
                Some(p) => p as f64,
                None => match previous {
                    None => 0.0,
                    Some(p) => p as f64 + 1.0,
                },
            };
            *slot = Some(boundaries.len());
            boundaries.push(Placement { block: block.clone(), index, position, omit, following: None });
        }
        if let Some(position) = keys[index].as_ref().and_then(|k| live.get(k)) {
            previous = Some(*position);
        }
    }
    for (id, (start, end)) in &pairs {
        let existing = present.get(id).copied().unwrap_or((None, None));
        if let Some(s) = start {
            if !boundaries[*s].omit {
                if let Some(e) = existing.1 {
                    boundaries[*s].position = boundaries[*s].position.min(e.floor());
                }
            }
        }
        if let Some(e) = end {
            if !boundaries[*e].omit {
                if let Some(s) = existing.0 {
                    boundaries[*e].position = boundaries[*e].position.max(s.ceil());
                }
            }
        }
        if let (Some(s), Some(e)) = (start, end) {
            let (sb, eb) = (&boundaries[*s], &boundaries[*e]);
            if !sb.omit
                && !eb.omit
                && (eb.position < sb.position || (eb.position == sb.position && eb.index < sb.index))
            {
                let end_block = eb.block.clone();
                boundaries[*s].following = Some(end_block);
                boundaries[*e].omit = true;
            }
        }
    }
    let mut insertions: HashMap<i64, Vec<Js>> = HashMap::new();
    for boundary in &boundaries {
        if boundary.omit {
            continue;
        }
        let bucket = insertions.entry(boundary.position as i64).or_default();
        bucket.push(boundary.block.clone());
        if let Some(following) = &boundary.following {
            bucket.push(following.clone());
        }
    }
    let projected_items = projected.items();
    let mut restored = Vec::new();
    for index in 0..=projected_items.len() {
        if let Some(bucket) = insertions.get(&(index as i64)) {
            restored.extend(bucket.iter().cloned());
        }
        if index < projected_items.len() {
            restored.push(projected_items[index].clone());
        }
    }
    Ok(Js::new_arr(restored))
}

/// `restoreTextBoxBookmarks` (src/yrs/yrsToDocument.ts:3257).
pub(crate) fn restore_text_box_bookmarks(content: &Js, base: &Js) -> Result<()> {
    fn bodies(value: &Js) -> IndexMap<String, Vec<Js>> {
        fn visit(value: &Js, found: &mut IndexMap<String, Vec<Js>>) {
            if value.is_arr() {
                for child in value.items() {
                    visit(&child, found);
                }
            } else if value.is_object_like() {
                if let Some(body) = as_obj_c(&value.get("textBody")) {
                    if body.get("content").is_arr() {
                        let id = match value.get("id").as_str() {
                            Some(id) => id.to_owned(),
                            None => found.len().to_string(),
                        };
                        found.entry(id).or_default().push(body);
                    }
                }
                for (key, child) in value.entries() {
                    if key != "drawingModel" {
                        visit(&child, found);
                    }
                }
            }
        }
        let mut found = IndexMap::new();
        visit(value, &mut found);
        found
    }
    let sources = bodies(base);
    for (id, copies) in bodies(content) {
        for (index, body) in copies.iter().enumerate() {
            if let Some(source) = sources.get(&id).and_then(|s| s.get(index)) {
                let restored = restore_block_bookmarks(&body.get("content"), &source.get("content"), &format!("txbx:{id}"))?;
                body.set("content", restored);
            }
        }
    }
    Ok(())
}

/// `restoreRawBlocks` (src/yrs/yrsToDocument.ts:3287).
pub(crate) fn restore_raw_blocks(projected: &Js, base: &Js, story_id: &str) -> Result<Js> {
    let projected = restore_block_bookmarks(projected, base, story_id)?;
    let base_items = base.items();
    let mut offset: f64 = 0.0;
    for index in 0..base_items.len() {
        let block = &base_items[index];
        if !is_raw_xml_c(block) || is_block_bookmark(block)? {
            continue;
        }
        let current = projected.items();
        let following = base_items[index + 1..].iter().find(|candidate| {
            type_is(candidate, "paragraph")
                && candidate.get("paraId").truthy()
                && current.iter().any(|entry| {
                    type_is(entry, "paragraph") && entry.get("paraId").strict_eq(&candidate.get("paraId"))
                })
        });
        let anchor: i64 = match following {
            Some(following) if type_is(following, "paragraph") => current
                .iter()
                .position(|entry| {
                    type_is(entry, "paragraph") && entry.get("paraId").strict_eq(&following.get("paraId"))
                })
                .map_or(-1, |p| p as i64),
            _ => -1,
        };
        let position = if anchor >= 0 {
            anchor as f64
        } else {
            (index as f64 + offset).min(current.len() as f64)
        };
        if let Some(mut items) = projected.items_mut() {
            let at = (position as usize).min(items.len());
            items.insert(at, block.clone());
        }
        offset = (position - index as f64).max(0.0);
    }
    Ok(projected)
}

// ----- projected paragraphs -----

/// `cloneProjectedParagraph` (src/yrs/yrsToDocument.ts:3341). `changes`/`identity` may be undefined.
pub(crate) fn clone_projected_paragraph(paragraph: &Js, changes: &Js, identity: &Js) -> Result<Js> {
    let clone = paragraph.spread();
    clone.assign(changes);
    let identity = if identity.is_undefined() { paragraph } else { identity };
    if let Some(memo) = projected_blocks_get(identity) {
        if memo.session_key.as_deref().map_or(false, |k| !k.is_empty()) {
            projected_blocks_set(
                &clone,
                ProjectedBlockMemo {
                    session_key: memo.session_key,
                    source_runs: memo.source_runs,
                    ..Default::default()
                },
            );
        }
    }
    Ok(clone)
}

/// `projectedSessionKey` (src/yrs/yrsToDocument.ts:3355).
pub(crate) fn projected_session_key(paragraph: &Js) -> Result<Option<String>> {
    Ok(projected_blocks_get(paragraph).and_then(|m| m.session_key))
}

/// `carryParagraphIdentities` (src/yrs/yrsToDocument.ts:3363).
pub(crate) fn carry_paragraph_identities(replacement: &Js, replaced: &Js) -> Result<Js> {
    let previous: Vec<Js> = replaced.items().into_iter().filter(|b| type_is(b, "paragraph")).collect();
    let mut index = 0;
    let mut out = Vec::new();
    for block in replacement.items() {
        if !type_is(&block, "paragraph") {
            out.push(block);
            continue;
        }
        let source = previous.get(index).cloned();
        index += 1;
        match source {
            Some(source) if source.get("paraId").truthy() => {
                out.push(clone_projected_paragraph(&block, &obj! { "paraId" => source.get("paraId") }, &source)?);
            }
            _ => out.push(block),
        }
    }
    Ok(Js::new_arr(out))
}

/// `ownProjectedParagraphs` (src/yrs/yrsToDocument.ts:3382).
pub(crate) fn own_projected_paragraphs(blocks: &Js) -> Result<Js> {
    let mut out = Vec::new();
    for block in blocks.items() {
        if type_is(&block, "paragraph") {
            out.push(clone_projected_paragraph(&block, &Js::Undefined, &Js::Undefined)?);
            continue;
        }
        if !type_is(&block, "table") {
            out.push(block);
            continue;
        }
        let mut rows = Vec::new();
        for row in block.get("rows").items() {
            let mut cells = Vec::new();
            for cell in row.get("cells").items() {
                cells.push(cell.with(&[("content", own_projected_paragraphs(&cell.get("content"))?)]));
            }
            rows.push(row.with(&[("cells", Js::new_arr(cells))]));
        }
        let own = block.with(&[("rows", Js::new_arr(rows))]);
        if let Some(prior) = projected_blocks_get(&block) {
            projected_blocks_set(
                &own,
                ProjectedBlockMemo {
                    table_baseline: prior.table_baseline,
                    table_source_prefix: prior.table_source_prefix,
                    ..Default::default()
                },
            );
        }
        out.push(own);
    }
    Ok(Js::new_arr(out))
}

/// `sessionProjectionMemo` (src/yrs/yrsToDocument.ts:3407): `{ wholesale, clean, dirty, stories,
/// parents }`, the sets as `Js::Map`s of `true`.
pub(crate) fn session_projection_memo(session: &Session) -> Result<Js> {
    Ok(SESSION_MEMOS.with(|m| {
        m.borrow_mut()
            .entry(session_key_of(session))
            .or_insert_with(|| {
                obj! {
                    "wholesale" => true,
                    "clean" => Js::new_map(),
                    "dirty" => Js::new_map(),
                    "stories" => Js::new_map(),
                    "parents" => Js::new_map(),
                }
            })
            .clone()
    }))
}

/// `sameJson` (src/yrs/yrsToDocument.ts:3423).
pub(crate) fn same_json(a: &Js, b: &Js) -> Result<bool> {
    Ok(same_json_c(a, b))
}

fn same_json_c(a: &Js, b: &Js) -> bool {
    if a.strict_eq(b) {
        return true;
    }
    if a.is_arr() && b.is_arr() {
        let (x, y) = (a.items(), b.items());
        return x.len() == y.len() && x.iter().zip(y.iter()).all(|(l, r)| same_json_c(l, r));
    }
    let (Some(left), Some(right)) = (as_obj_c(a), as_obj_c(b)) else { return false };
    let keys = left.keys();
    keys.len() == right.keys().len()
        && keys.iter().all(|key| right.has(key) && same_json_c(&left.get(key), &right.get(key)))
}

// ----- block breaks -----

fn block_break_kind_c(embed_kind: &Js) -> Option<&'static str> {
    match embed_kind.as_str() {
        Some("pageBreak") => Some("page"),
        Some("columnBreak") => Some("column"),
        _ => None,
    }
}

fn is_block_break_c(item: &Js) -> bool {
    item.get("kind").as_str() == Some("embed") && block_break_kind_c(&item.get("embedKind")).is_some()
}

/// `isBlockBreak` (src/yrs/yrsToDocument.ts:3441).
pub(crate) fn is_block_break(item: &Js) -> Result<bool> {
    Ok(is_block_break_c(item))
}

/// `blockBreakKind` (src/yrs/yrsToDocument.ts:3446).
pub(crate) fn block_break_kind(embed_kind: &str) -> Result<Js> {
    Ok(match block_break_kind_c(&Js::from(embed_kind)) {
        Some(kind) => Js::from(kind),
        None => Js::Null,
    })
}

/// `runFlowTokens` (src/yrs/yrsToDocument.ts:3452).
pub(crate) fn run_flow_tokens(run: &Js, tokens: &Js) -> Result<()> {
    for content in run.get("content").items() {
        let break_type = content.get("breakType");
        if type_is(&content, "break") && matches!(break_type.as_str(), Some("page") | Some("column")) {
            tokens.push(break_type);
        } else if !type_is(&content, "text") || content.get("text").as_str().map_or(false, |t| js::len16(t) > 0) {
            tokens.push("visible");
        }
    }
    Ok(())
}

/// `flowTokens` (src/yrs/yrsToDocument.ts:3460). `tokens` may be undefined.
pub(crate) fn flow_tokens(content: &Js, tokens: &Js) -> Result<Js> {
    let tokens = if tokens.is_undefined() { Js::new_arr(Vec::new()) } else { tokens.clone() };
    for item in content.items() {
        match item.get("type").as_str() {
            Some("run") => run_flow_tokens(&item, &tokens)?,
            Some("hyperlink") => {
                for child in item.get("children").items() {
                    if type_is(&child, "run") {
                        run_flow_tokens(&child, &tokens)?;
                    }
                }
            }
            Some("simpleField") => {
                for child in item.get("content").items() {
                    if type_is(&child, "run") {
                        run_flow_tokens(&child, &tokens)?;
                    }
                }
            }
            Some("complexField") => {
                for child in item.get("fieldCode").items().into_iter().chain(item.get("fieldResult").items()) {
                    run_flow_tokens(&child, &tokens)?;
                }
            }
            Some("inlineSdt") => {
                flow_tokens(&item.get("content"), &tokens)?;
            }
            Some("insertion") | Some("deletion") | Some("moveFrom") | Some("moveTo") => {
                for child in item.get("content").items() {
                    if type_is(&child, "run") {
                        run_flow_tokens(&child, &tokens)?;
                    }
                }
            }
            Some("mathEquation") => tokens.push("visible"),
            _ => {}
        }
    }
    Ok(tokens)
}

fn flow_token_vec(paragraph: &Js) -> Result<Vec<String>> {
    Ok(arr_strings(&flow_tokens(&paragraph.get("content"), &Js::Undefined)?))
}

fn leading_block_breaks_vec(paragraph: &Js) -> Result<Vec<String>> {
    let tokens = flow_token_vec(paragraph)?;
    if !tokens.iter().any(|t| t == "visible") {
        let last = tokens.iter().rposition(|t| t == "column").map_or(0, |p| p + 1);
        return Ok(tokens[..last].to_vec());
    }
    Ok(if tokens.first().map(String::as_str) == Some("column") { vec!["column".into()] } else { vec![] })
}

fn trailing_block_breaks_vec(paragraph: &Js) -> Result<Vec<String>> {
    let tokens = flow_token_vec(paragraph)?;
    if !tokens.iter().any(|t| t == "visible") {
        let from = tokens.iter().rposition(|t| t == "column").map_or(0, |p| p + 1);
        return Ok(tokens[from..].to_vec());
    }
    let mut trailing = Vec::new();
    let mut leading = false;
    let mut visible = false;
    for token in tokens {
        if token == "visible" {
            visible = true;
        } else if visible || leading {
            trailing.push(token);
        } else {
            leading = true;
        }
    }
    Ok(trailing)
}

/// `leadingBlockBreaks` (src/yrs/yrsToDocument.ts:3477).
pub(crate) fn leading_block_breaks(paragraph: &Js) -> Result<Js> {
    Ok(str_arr(&leading_block_breaks_vec(paragraph)?))
}

/// `trailingBlockBreaks` (src/yrs/yrsToDocument.ts:3484).
pub(crate) fn trailing_block_breaks(paragraph: &Js) -> Result<Js> {
    Ok(str_arr(&trailing_block_breaks_vec(paragraph)?))
}

/// `noteYrsStoriesDirty` (src/yrs/yrsToDocument.ts:3499). `stories` is `"all"`, a story id, or
/// an array of story ids.
pub(crate) fn note_yrs_stories_dirty(session: &Session, stories: &Js) -> Result<()> {
    let memo = session_projection_memo(session)?;
    if stories.as_str() == Some("all") {
        memo.set("wholesale", true);
        map_clear(&memo.get("clean"));
        map_clear(&memo.get("dirty"));
        map_clear(&memo.get("stories"));
        map_clear(&memo.get("parents"));
        return Ok(());
    }
    let mut queue: Vec<String> = match stories.as_str() {
        Some(s) => vec![s.to_owned()],
        None if stories.is_map() => stories.map_keys(),
        None => stories.items().iter().map(key_of).collect(),
    };
    let mut queued: HashSet<String> = queue.iter().cloned().collect();
    let mut index = 0;
    while index < queue.len() {
        let story = queue[index].clone();
        memo.get("dirty").map_set(&story, true);
        map_delete(&memo.get("clean"), &story);
        map_delete(&memo.get("stories"), &story);
        let recorded = memo.get("parents").map_get(&story);
        let parent = if recorded.is_nullish() {
            nested_story_parent(&story)
        } else {
            Some(key_of(&recorded))
        };
        if let Some(parent) = parent {
            if !parent.is_empty() && !queued.contains(&parent) {
                queued.insert(parent.clone());
                queue.push(parent);
            }
        }
        index += 1;
    }
    Ok(())
}

// ----- projection -----

fn options_story_ids(options: &Js) -> Option<Js> {
    let ids = options.get("storyIds");
    if ids.is_undefined() { None } else { Some(ids) }
}

fn story_set_has(set: &Js, story: &str) -> bool {
    if set.is_map() {
        set.map_has(story)
    } else if set.is_arr() {
        set.items().iter().any(|v| v.as_str() == Some(story))
    } else {
        set.has(story)
    }
}

/// `yrsToDocument` (src/yrs/yrsToDocument.ts:4243).
pub(crate) fn yrs_to_document(session: &Session, base: &Js, options: &Js) -> Result<Js> {
    let options = if options.is_nullish() { Js::new_obj() } else { options.clone() };
    project_document(session, base, &options, &Js::Undefined)
}

/// `yrsBodyToDocumentWithRevisionIds` (src/yrs/yrsToDocument.ts:4255).
pub(crate) fn yrs_body_to_document_with_revision_ids(session: &Session, base: &Js, revision_ids: &Js) -> Result<Js> {
    project_document(session, base, &obj! { "storyIds" => arr!["body"] }, revision_ids)
}

/// `projectDocument` (src/yrs/yrsToDocument.ts:4263).
pub(crate) fn project_document(session: &Session, base: &Js, options: &Js, revision_ids: &Js) -> Result<Js> {
    let _media = project_media_with(session);
    let result = (|| {
        let context = SaveContext::new(session, base, &options.get("commentIds"), revision_ids)?;
        let result = project_stories(&context, session, base, options);
        context.dispose()?;
        result
    })();
    // Each save opens a fresh session: its memo lives only as long as the projection.
    clear_session_projection_memo(session);
    result
}

/// `withHostSectionChanges` (src/yrs/yrsToDocument.ts:4281).
pub(crate) fn with_host_section_changes(authored: &Js, source: &Js, host: &Js) -> Result<Js> {
    let merged = authored.spread();
    let mut keys: IndexMap<String, ()> = IndexMap::new();
    for key in source.keys().into_iter().chain(host.keys()) {
        keys.insert(key, ());
    }
    for key in keys.keys() {
        if section_properties_equal(&host.get(key), &source.get(key))?
            || !section_properties_equal(&authored.get(key), &source.get(key))?
        {
            continue;
        }
        if host.get(key).is_undefined() {
            merged.delete(key);
        } else {
            merged.set(key, host.get(key));
        }
    }
    Ok(merged)
}

/// `projectStories` (src/yrs/yrsToDocument.ts:4292).
pub(crate) fn project_stories(context: &crate::yrs_to_document::SaveContext, session: &Session, base: &Js, options: &Js) -> Result<Js> {
    let story_ids = options_story_ids(options);
    let should_project = |story_id: &str| -> bool {
        match &story_ids {
            None => true,
            Some(set) => story_set_has(set, story_id),
        }
    };
    let package = base.get("package");
    let body_content = if context.story_ids.contains("body") && should_project("body") {
        context.story_to_blocks("body")?
    } else {
        package.path(&["document", "content"])
    };

    let mut selected_header_footer_stories: HashSet<String> = HashSet::new();
    for r_id in package.get("headers").map_keys().into_iter().chain(package.get("footers").map_keys()) {
        if should_project(&format!("hf:{r_id}")) {
            selected_header_footer_stories.insert(session.header_footer_story(&r_id));
        }
    }
    let should_project_header_footer =
        |story_id: &str| -> bool { should_project(story_id) || selected_header_footer_stories.contains(story_id) };

    let header_footer_contents: RefCell<HashMap<String, Js>> = RefCell::new(HashMap::new());
    let project_header_footer = |story_id: &str| -> Result<Js> {
        if let Some(content) = header_footer_contents.borrow().get(story_id) {
            if content.truthy() {
                return Ok(content.clone());
            }
        }
        let content = context.story_to_blocks(story_id)?;
        header_footer_contents.borrow_mut().insert(story_id.to_owned(), content.clone());
        Ok(content)
    };
    let canonical_for = |r_id: &str, story_id: &str| -> Option<String> {
        if story_id == format!("hf:{r_id}") {
            None
        } else {
            Some(js::slice16(story_id, 3.0, None))
        }
    };
    let project_header_footers = |entries: Js| -> Result<Js> {
        if !entries.truthy() {
            return Ok(entries);
        }
        let routes: Vec<(String, Js, String)> = entries
            .map_entries()
            .into_iter()
            .map(|(r_id, part)| {
                let story = session.header_footer_story(&r_id);
                (r_id, part, story)
            })
            .collect();
        let project = story_ids.is_none() || routes.iter().any(|(_, _, s)| should_project_header_footer(s));
        let mut remark = false;
        for (r_id, part, story_id) in &routes {
            if header_footer_alias_of(part)? != canonical_for(r_id, story_id) {
                remark = true;
                break;
            }
        }
        if !project && !remark {
            return Ok(entries);
        }
        let mut out = Vec::new();
        for (r_id, part, story_id) in routes {
            let canonical = canonical_for(&r_id, &story_id);
            let next = if context.story_ids.contains(&story_id) && should_project_header_footer(&story_id) {
                part.with(&[("content", project_header_footer(&story_id)?)])
            } else if header_footer_alias_of(&part)? == canonical {
                part.clone()
            } else {
                part.spread()
            };
            if let Some(canonical) = &canonical {
                mark_header_footer_alias(&next, canonical)?;
            }
            out.push((r_id, next));
        }
        Ok(Js::map_from(out))
    };
    let headers = project_header_footers(package.get("headers"))?;
    let footers = project_header_footers(package.get("footers"))?;

    let project_notes = |notes: Js, prefix: &str| -> Result<Js> {
        let should_project_notes = story_ids.is_none()
            || (notes.is_arr()
                && notes
                    .items()
                    .iter()
                    .any(|note| should_project(&format!("{prefix}{}", key_of(&note.get("id"))))));
        if !should_project_notes || !notes.is_arr() {
            return Ok(notes);
        }
        let mut out = Vec::new();
        for note in notes.items() {
            let story_id = format!("{prefix}{}", key_of(&note.get("id")));
            if context.story_ids.contains(&story_id) && should_project(&story_id) {
                let content = restore_note_marks(&context.story_to_blocks(&story_id)?, &note.get("content"))?;
                out.push(note.with(&[("content", content), ("verbatimXml", Js::Undefined)]));
            } else {
                out.push(note);
            }
        }
        Ok(Js::new_arr(out))
    };
    let footnotes = project_notes(package.get("footnotes"), "fn:")?;
    let endnotes = project_notes(package.get("endnotes"), "en:")?;

    let base_comments = package.path(&["document", "comments"]);
    let mut has_bookmarked = false;
    for comment in base_comments.items() {
        if contains_bookmark(&comment.get("blockContent").nullish_or(Js::new_arr(Vec::new())))? {
            has_bookmarked = true;
            break;
        }
    }
    let comment_bodies: Option<HashMap<String, Js>> = if has_bookmarked {
        let mut bodies = HashMap::new();
        for comment in session.list_comments()?.items() {
            // A `Map` keyed by the session id, looked up by `String(comment.id)`.
            if let Some(id) = comment.get("id").as_str() {
                bodies.insert(id.to_owned(), comment.get("body"));
            }
        }
        Some(bodies)
    } else {
        None
    };
    let comments = match &comment_bodies {
        Some(comment_bodies) if base_comments.is_arr() => {
            let mut out = Vec::new();
            for comment in base_comments.items() {
                if !contains_bookmark(&comment.get("blockContent").nullish_or(Js::new_arr(Vec::new())))? {
                    out.push(comment);
                    continue;
                }
                let body = comment_bodies.get(&key_of(&comment.get("id"))).cloned().unwrap_or_default();
                let original = comment.get("blockContent").nullish_or(comment.get("content"));
                let valid = body.is_arr()
                    && body.items().iter().all(|block| {
                        block.truthy()
                            && matches!(
                                block.get("type").as_str(),
                                Some("paragraph") | Some("table") | Some("blockSdt") | Some("rawXml")
                            )
                    });
                if !valid || same_json_c(&body, &original) {
                    out.push(comment);
                    continue;
                }
                assert_comment_bookmark_replacement(&original)?;
                let block_content =
                    restore_block_bookmarks(&body, &original, &format!("comment:{}", key_of(&comment.get("id"))))?;
                let content = Js::new_arr(
                    block_content.items().into_iter().filter(|b| type_is(b, "paragraph")).collect(),
                );
                out.push(comment.with(&[("content", content), ("blockContent", block_content)]));
            }
            Js::new_arr(out)
        }
        _ => Js::Undefined,
    };

    let source_package = session.source_package();
    let live_sections = if source_package.is_some() && should_project("body") {
        session.list_sections()?.nullish_or(Js::new_arr(Vec::new()))
    } else {
        Js::new_arr(Vec::new())
    };
    let source_final = match source_package {
        Some(source) => {
            let metadata = source.keys.get("metadata");
            Js::parse(metadata.str_or_empty())?.get("finalSectionProperties")
        }
        None => Js::Undefined,
    };
    let sections = live_sections.items();
    let source_final_or_empty = if source_final.is_nullish() { Js::new_obj() } else { source_final.clone() };
    let final_section_changes = !sections.is_empty()
        && !section_properties_equal(&source_final_or_empty, &sections.last().unwrap().get("authored"))?;
    let mut section_changes = final_section_changes;
    if !section_changes && !sections.is_empty() {
        for section in &sections[..sections.len() - 1] {
            let id = key_of(&section.get("closingParagraphId"));
            if context.section_changed(&id, &section.get("authored"))? {
                section_changes = true;
                break;
            }
        }
    }
    let projected_sections = if section_changes {
        let mut start: f64 = 0.0;
        let body_items = body_content.items();
        let mut out = Vec::new();
        for section in &sections {
            let closing = section.get("closingParagraphId");
            let end = if closing.as_str() == Some("final") {
                body_items.len() as f64
            } else {
                let mut found: i64 = -1;
                for (i, block) in body_items.iter().enumerate() {
                    if type_is(block, "paragraph") {
                        let key = projected_session_key(block)?;
                        if Js::from(key).strict_eq(&closing) {
                            found = i as i64;
                            break;
                        }
                    }
                }
                (found + 1) as f64
            };
            let content = js_slice(&body_content, start, Some(end));
            start = end;
            out.push(obj! {
                "id" => closing.clone(),
                "properties" => section.get("resolved"),
                "content" => content,
            });
        }
        Js::new_arr(out)
    } else {
        Js::Undefined
    };

    let result = base.spread();
    let warnings = context.warnings.clone();
    if warnings.len() > 0 {
        let mut all = base.get("warnings").nullish_or(Js::new_arr(Vec::new())).items();
        all.extend(warnings.items());
        result.set("warnings", Js::new_arr(all));
    }
    let document = package.get("document").spread();
    document.set("content", body_content.clone());
    if comments.truthy() {
        document.set("comments", comments);
    }
    if section_changes {
        document.set("sections", projected_sections);
    }
    if final_section_changes {
        let host_final = package.path(&["document", "finalSectionProperties"]);
        let host_final = if host_final.is_nullish() { Js::new_obj() } else { host_final };
        document.set(
            "finalSectionProperties",
            with_host_section_changes(&sections.last().unwrap().get("authored"), &source_final_or_empty, &host_final)?,
        );
    }
    let out_package = package.spread();
    out_package.set("document", document);
    if headers.truthy() {
        out_package.set("headers", headers);
    }
    if footers.truthy() {
        out_package.set("footers", footers);
    }
    if footnotes.truthy() {
        out_package.set("footnotes", footnotes);
    }
    if endnotes.truthy() {
        out_package.set("endnotes", endnotes);
    }
    result.set("package", out_package);
    Ok(result)
}

// ----- SaveContext -----

/// `class SaveContext`.
pub(crate) struct SaveContext<'a> {
    session: &'a Session<'a>,
    pub(crate) story_ids: HashSet<String>,
    /// `warnings`: a shared `Js` array of strings (recoveries push onto it).
    pub(crate) warnings: Js,
    base_paragraphs: IndexMap<String, Js>,
    reference_runs: HashMap<String, Js>,
    comment_parents: HashMap<String, Js>,
    has_comment_reference_style: bool,
    base_stories: IndexMap<String, Js>,
    /// Story id to a `Js::Map` of drawings.
    base_chart_drawings: HashMap<String, Js>,
    comments: IndexMap<String, Js>,
    subtree_comments: IndexMap<String, IndexMap<String, Js>>,
    memo: Js,
    styles: StyleResolver,
    synthetic_defaults: bool,
    inherited: RefCell<HashMap<String, Js>>,
    style_definitions: Js,
    run_inheritance: RefCell<Option<Rc<Inheritance>>>,
    inherited_runs: RefCell<HashMap<String, Js>>,
    source_tables: HashMap<String, Js>,
    /// Keyed by table identity; the table kept alive alongside.
    table_baselines: RefCell<HashMap<usize, (Js, Js)>>,
    payload_baselines: Marks<Js>,
    source_cells: RefCell<HashMap<String, Js>>,
    cell_paragraphs: RefCell<HashMap<String, Js>>,
    cell_runs: RefCell<HashMap<String, Js>>,
    font_scheme: Js,
    revision_ids: Js,
}

fn memo_parents_owner(memo: &Js, owner: &str) -> Option<String> {
    let recorded = memo.get("parents").map_get(owner);
    if recorded.is_nullish() {
        nested_story_parent(owner)
    } else {
        Some(key_of(&recorded))
    }
}

impl<'a> SaveContext<'a> {
    /// `SaveContext.sourceTable` (src/yrs/yrsToDocument.ts:3550).
    pub(crate) fn source_table(&self, prefix: &str) -> Result<Js> {
        Ok(self.source_tables.get(prefix).cloned().unwrap_or_default())
    }

    /// `SaveContext.constructor` (src/yrs/yrsToDocument.ts:3559).
    pub(crate) fn new(session: &'a Session<'a>, base: &Js, comment_ids: &Js, revision_ids: &Js) -> Result<Self> {
        let story_ids: HashSet<String> = session.story_ids().into_iter().collect();
        let base_stories = collect_base_stories_map(session, base)?;
        let mut source_tables = HashMap::new();
        let mut base_chart_drawings = HashMap::new();
        for (story_id, blocks) in &base_stories {
            let mut table_index = 0;
            for block in blocks.items() {
                if !type_is(&block, "table") {
                    continue;
                }
                let prefix = projected_blocks_get(&block)
                    .and_then(|m| m.table_source_prefix)
                    .unwrap_or_else(|| format!("{story_id}:t{table_index}"));
                source_tables.insert(prefix, block);
                table_index += 1;
            }
            let scoped = Js::new_map();
            collect_chart_drawings(blocks, &scoped)?;
            if !scoped.map_keys().is_empty() {
                base_chart_drawings.insert(story_id.clone(), scoped);
            }
        }
        let base_paragraphs = collect_base_paragraphs_map(&base_stories);
        let mut reference_runs = HashMap::new();
        for paragraph in base_paragraphs.values() {
            for run in paragraph_chart_runs(&paragraph.get("content"))?.items() {
                for entry in run.get("content").items() {
                    if type_is(&entry, "commentReference") && !entry.get("id").is_undefined() {
                        reference_runs.insert(key_of(&entry.get("id")), run.clone());
                    }
                }
            }
        }
        let package = base.get("package");
        let mut comment_parents = HashMap::new();
        for comment in package.path(&["document", "comments"]).items() {
            if !comment.get("parentId").is_undefined() {
                comment_parents.insert(key_of(&comment.get("id")), comment.get("parentId"));
            }
        }
        let style_list = package.path(&["styles", "styles"]).items();
        let has_comment_reference_style = style_list.iter().any(|style| {
            style.get("styleId").as_str() == Some("CommentReference") && style.get("type").as_str() == Some("character")
        });
        let comments = comment_ranges_map(session, &package.path(&["document", "comments"]), comment_ids)?;
        let memo = session_projection_memo(session)?;
        let style_definitions = package.get("styles");
        let font_scheme = package.path(&["theme", "fontScheme"]);
        let styles = create_style_resolver(&package.get("styles"))?;
        let default_style = styles.get_default_paragraph_style()?;
        let synthetic_defaults =
            !default_style.is_undefined() && !style_list.iter().any(|s| s.strict_eq(&default_style));
        let mut subtree_comments: IndexMap<String, IndexMap<String, Js>> = IndexMap::new();
        for (story, ranges) in &comments {
            let mut seen: HashSet<String> = HashSet::new();
            let mut owner: Option<String> = Some(story.clone());
            while let Some(current) = owner {
                if seen.contains(&current) {
                    break;
                }
                seen.insert(current.clone());
                subtree_comments
                    .entry(current.clone())
                    .or_default()
                    .insert(story.clone(), ranges.clone());
                owner = memo_parents_owner(&memo, &current);
            }
        }
        Ok(Self {
            session,
            story_ids,
            warnings: Js::new_arr(Vec::new()),
            base_paragraphs,
            reference_runs,
            comment_parents,
            has_comment_reference_style,
            base_stories,
            base_chart_drawings,
            comments,
            subtree_comments,
            memo,
            styles,
            synthetic_defaults,
            inherited: RefCell::new(HashMap::new()),
            style_definitions,
            run_inheritance: RefCell::new(None),
            inherited_runs: RefCell::new(HashMap::new()),
            source_tables,
            table_baselines: RefCell::new(HashMap::new()),
            payload_baselines: Marks::new(),
            source_cells: RefCell::new(HashMap::new()),
            cell_paragraphs: RefCell::new(HashMap::new()),
            cell_runs: RefCell::new(HashMap::new()),
            font_scheme,
            revision_ids: revision_ids.clone(),
        })
    }

    /// `SaveContext.storyIsClean` (src/yrs/yrsToDocument.ts:3617).
    pub(crate) fn story_is_clean(&self, story_id: &str) -> Result<bool> {
        Ok(if self.memo.get("wholesale").truthy() {
            self.memo.get("clean").map_has(story_id)
        } else {
            !self.memo.get("dirty").map_has(story_id)
        })
    }

    /// `SaveContext.sectionChanged` (src/yrs/yrsToDocument.ts:3623).
    pub(crate) fn section_changed(&self, para_id: &str, properties: &Js) -> Result<bool> {
        let base = self
            .base_paragraphs
            .get(para_id)
            .map(|p| p.get("sectionProperties"))
            .unwrap_or_default();
        Ok(!section_properties_equal(&base, properties)?)
    }

    /// `SaveContext.inheritance` (src/yrs/yrsToDocument.ts:3627).
    pub(crate) fn inheritance(&self) -> Result<Rc<Inheritance>> {
        let mut slot = self.run_inheritance.borrow_mut();
        if slot.is_none() {
            *slot = Some(Rc::new(self.session.run_inheritance(&self.style_definitions)));
        }
        Ok(slot.as_ref().unwrap().clone())
    }

    /// `SaveContext.prepareTable` (src/yrs/yrsToDocument.ts:3631).
    pub(crate) fn prepare_table(&self, payload: &Js, parent: &str) -> Result<()> {
        let rows = payload.get("rows");
        let cells: Vec<Js> = if rows.is_arr() {
            rows.items()
                .iter()
                .flat_map(|row| {
                    let cells = row.get("cells");
                    if cells.is_arr() { cells.items() } else { Vec::new() }
                })
                .collect()
        } else {
            Vec::new()
        };
        let mut source: Option<(String, Js)> = None;
        if payload.path(&["tblPr", "_originalFormatting"]).truthy() {
            for cell in &cells {
                let story = as_str_c(&cell.path(&["tcPr", "_sourceCellStory"])).or_else(|| as_str_c(&cell.get("story")));
                let prefix = story.as_deref().and_then(table_prefix_of);
                if let Some(prefix) = prefix.filter(|p| !p.is_empty()) {
                    if let Some(table) = self.source_tables.get(&prefix) {
                        if table.truthy() {
                            source = Some((prefix, table.clone()));
                            break;
                        }
                    }
                }
            }
        }
        if let Some((prefix, table)) = source {
            let ptr = table.ptr().unwrap_or(0);
            let existing = self.table_baselines.borrow().get(&ptr).map(|(_, b)| b.clone());
            let baseline = match existing {
                Some(b) => b,
                None => {
                    let baseline = match projected_blocks_get(&table).and_then(|m| m.table_baseline) {
                        Some(b) if !b.is_nullish() => b,
                        _ => self.inheritance()?.table_attributes(&table),
                    };
                    self.table_baselines.borrow_mut().insert(ptr, (table.clone(), baseline.clone()));
                    baseline
                }
            };
            self.payload_baselines.set(payload, baseline.clone());
            let has_memo_baseline = projected_blocks_get(&table)
                .and_then(|m| m.table_baseline)
                .map_or(false, |b| b.truthy());
            let recorded = if has_memo_baseline { None } else { projected_table_cells_get(&table) };
            if let Some(recorded) = recorded {
                let seeded = seeded_cell_indices(&table)?;
                for (row_index, row) in recorded.items().iter().enumerate() {
                    for (cell_index, entry) in row.items().iter().enumerate() {
                        let seeded_row = seeded.at(row_index);
                        let ordinal: i64 = if seeded_row.is_arr() {
                            seeded_row
                                .items()
                                .iter()
                                .position(|v| v.strict_eq(&Js::Num(cell_index as f64)))
                                .map_or(-1, |p| p as i64)
                        } else {
                            -1
                        };
                        let cell = if ordinal >= 0 {
                            baseline.get("rows").at(row_index).at(ordinal as usize)
                        } else {
                            Js::Undefined
                        };
                        if let Some(entry) = entry.as_str() {
                            if cell.truthy() {
                                self.source_cells.borrow_mut().insert(entry.to_owned(), cell);
                            }
                        }
                    }
                }
            } else {
                for (row_index, row) in baseline.get("rows").items().iter().enumerate() {
                    for (cell_index, cell) in row.items().iter().enumerate() {
                        self.source_cells
                            .borrow_mut()
                            .insert(format!("{prefix}:r{row_index}c{cell_index}"), cell.clone());
                    }
                }
            }
        }
        for cell in &cells {
            let Some(story) = cell.get("story").as_str().map(str::to_owned) else { continue };
            self.memo.get("parents").map_set(&story, parent);
            let source = as_str_c(&cell.path(&["tcPr", "_sourceCellStory"])).unwrap_or_else(|| story.clone());
            let source_cell = self.source_cells.borrow().get(&source).cloned().unwrap_or_default();
            let pf = source_cell.get("paragraphFormatting");
            let rf = source_cell.get("runFormatting");
            self.cell_paragraphs
                .borrow_mut()
                .insert(story.clone(), if pf.is_nullish() { Js::Undefined } else { pf });
            self.cell_runs
                .borrow_mut()
                .insert(story, if rf.is_nullish() { Js::Undefined } else { rf });
        }
        Ok(())
    }

    /// `SaveContext.applyTableBaseline` (src/yrs/yrsToDocument.ts:3680).
    pub(crate) fn apply_table_baseline(&self, payload: &Js, attrs: &Js) -> Result<()> {
        if let Some(baseline) = self.payload_baselines.get(payload) {
            if baseline.truthy() {
                attrs.set("_originalResolvedMargins", baseline.path(&["attributes", "cellMargins"]));
            }
        }
        Ok(())
    }

    /// `SaveContext.cellAttrs` (src/yrs/yrsToDocument.ts:3687).
    pub(crate) fn cell_attrs(&self, payload: &Js) -> Result<Js> {
        let attrs = table_cell_attr_defaults_c();
        let tc_pr = payload.get("tcPr");
        if !tc_pr.is_nullish() {
            attrs.assign(&tc_pr);
        }
        let origin = as_str_c(&payload.path(&["tcPr", "_sourceCellStory"]));
        let source = origin.clone().or_else(|| as_str_c(&payload.get("story")));
        let baseline = match &source {
            Some(source) => self.source_cells.borrow().get(source).cloned().unwrap_or_default(),
            None => Js::Undefined,
        };
        if baseline.truthy() {
            attrs.set("_originalResolvedBorders", baseline.path(&["attributes", "borders"]));
            attrs.set("_originalResolvedMargins", baseline.path(&["attributes", "margins"]));
            let original = as_obj_c(&baseline.path(&["attributes", "_originalFormatting"]));
            if let (Some(origin), Some(original)) = (&origin, original) {
                if !origin.is_empty() && !attrs.get("_originalFormatting").truthy() {
                    let formatting = original.spread();
                    formatting.delete("gridSpan");
                    formatting.delete("vMerge");
                    attrs.set("_originalFormatting", formatting);
                }
            }
        }
        Ok(attrs)
    }

    /// `SaveContext.cellContents` (src/yrs/yrsToDocument.ts:3709).
    pub(crate) fn cell_contents(&self, payload: &Js, parent: &str) -> Result<Js> {
        let mut contents = Vec::new();
        let rows = payload.get("rows");
        for row in if rows.is_arr() { rows.items() } else { Vec::new() } {
            let cells = row.get("cells");
            for cell in if cells.is_arr() { cells.items() } else { Vec::new() } {
                let story = cell.get("story");
                if !story.is_undefined() {
                    self.memo.get("parents").map_set(&key_of(&story), parent);
                }
                match story.as_str() {
                    Some(story) if self.story_ids.contains(story) => contents.push(self.story_to_blocks(story)?),
                    _ => contents.push(empty_blocks()),
                }
            }
        }
        Ok(Js::new_arr(contents))
    }

    /// `SaveContext.restoreCommentReferences` (src/yrs/yrsToDocument.ts:3724).
    pub(crate) fn restore_comment_references(&self, paragraph: &Js) -> Result<()> {
        let mut held: HashSet<String> = HashSet::new();
        for run in paragraph_chart_runs(&paragraph.get("content"))?.items() {
            for entry in run.get("content").items() {
                if !type_is(&entry, "commentReference") || entry.get("id").is_undefined() {
                    continue;
                }
                let id = key_of(&entry.get("id"));
                held.insert(id.clone());
                if let Some(source) = self.reference_runs.get(&id) {
                    run.set("formatting", source.get("formatting"));
                    run.set("propertyChanges", source.get("propertyChanges"));
                }
            }
        }
        let mut content = Vec::new();
        for child in paragraph.get("content").items() {
            let child_id = key_of(&child.get("id"));
            if !type_is(&child, "commentRangeEnd") || held.contains(&child_id) {
                content.push(child);
                continue;
            }
            let Some(parent) = self.comment_parents.get(&child_id).cloned() else {
                content.push(child);
                continue;
            };
            let mut seen: HashSet<String> = HashSet::new();
            seen.insert(child_id.clone());
            let mut id: Option<String> = Some(key_of(&parent));
            let mut source: Option<Js> = None;
            while let Some(current) = id.clone() {
                if seen.contains(&current) {
                    break;
                }
                seen.insert(current.clone());
                source = self.reference_runs.get(&current).cloned();
                if source.is_some() {
                    break;
                }
                id = self.comment_parents.get(&current).map(key_of);
            }
            let source_formatting = source.as_ref().map(|s| s.get("formatting")).unwrap_or_default();
            let formatting = if self.has_comment_reference_style {
                source_formatting.with(&[("styleId", Js::from("CommentReference"))])
            } else {
                source_formatting
            };
            held.insert(child_id);
            let id_value = child.get("id");
            content.push(child);
            content.push(obj! {
                "type" => "run",
                "formatting" => formatting,
                "content" => arr![obj! { "type" => "commentReference", "id" => id_value }],
            });
        }
        paragraph.set("content", Js::new_arr(content));
        Ok(())
    }

    /// `SaveContext.inheritedFormatting` (src/yrs/yrsToDocument.ts:3762).
    pub(crate) fn inherited_formatting(&self, story_id: &str, properties: &Js) -> Result<Js> {
        let cell = self.cell_paragraphs.borrow().contains_key(story_id);
        let nested = !cell && nested_story_parent(story_id).is_some();
        if self.synthetic_defaults && !cell && !nested {
            return Ok(Js::Undefined);
        }
        let numbered = !properties.get("numPr").is_nullish();
        let style = properties.get("pStyle").as_str().unwrap_or("").to_owned();
        let table_formatting = self.cell_paragraphs.borrow().get(story_id).cloned().unwrap_or_default();
        let key = if cell {
            arr![style.as_str(), table_formatting.clone(), numbered].stringify().unwrap_or_default()
        } else {
            arr![style.as_str(), numbered, nested].stringify().unwrap_or_default()
        };
        if !self.inherited.borrow().contains_key(&key) {
            let style_opt = if style.is_empty() { None } else { Some(style.as_str()) };
            let resolved = if cell {
                self.inheritance()?.paragraph_formatting(style_opt, &table_formatting)
            } else {
                self.styles
                    .resolve_paragraph_style(style_opt, &Js::Undefined, &Js::Undefined)?
                    .get("paragraphFormatting")
            };
            let value = if resolved.is_undefined() || (!numbered && !resolved.get("numPr").is_nullish()) {
                Js::Undefined
            } else if numbered || nested {
                let formatting = Js::new_obj();
                if !resolved.get("autoSpaceDE").is_undefined() {
                    formatting.set("autoSpaceDE", resolved.get("autoSpaceDE"));
                }
                if !resolved.get("autoSpaceDN").is_undefined() {
                    formatting.set("autoSpaceDN", resolved.get("autoSpaceDN"));
                }
                if !resolved.get("adjustRightInd").is_undefined() {
                    formatting.set("adjustRightInd", resolved.get("adjustRightInd"));
                }
                formatting
            } else {
                let formatting = resolved.spread();
                formatting.delete("indentLeft");
                formatting.delete("indentRight");
                formatting
            };
            self.inherited.borrow_mut().insert(key.clone(), value);
        }
        Ok(self.inherited.borrow().get(&key).cloned().unwrap_or_default())
    }

    fn fonts_key(&self, fonts: &Js) -> Result<Js> {
        let mut out = Vec::new();
        for (name, r#ref) in [("ascii", "asciiTheme"), ("hAnsi", "hAnsiTheme"), ("eastAsia", "eastAsiaTheme"), ("cs", "csTheme")] {
            out.push(canonical_font_pair(fonts, name, r#ref, &self.font_scheme)?);
        }
        Ok(Js::new_arr(out))
    }

    fn formatting_key(&self, formatting: &Js) -> Result<String> {
        let copy = formatting.spread();
        copy.delete("pageBreakBefore");
        copy.delete("frame");
        let run_properties = formatting.get("runProperties");
        if run_properties.truthy() {
            copy.set(
                "runProperties",
                run_properties.with(&[("fontFamily", self.fonts_key(&run_properties.get("fontFamily"))?)]),
            );
        }
        let round_trip = Js::parse(&copy.stringify().unwrap_or_else(|| "null".into()))?;
        stable_stringify(&round_trip)
    }

    /// `SaveContext.copiedParagraphProperties` (src/yrs/yrsToDocument.ts:3798).
    pub(crate) fn copied_paragraph_properties(&self, story: &str, properties: &Js, items: &Js, preceding: &Js) -> Result<Js> {
        let original = as_obj_c(&properties.get("_originalFormatting"));
        let original_font = original
            .as_ref()
            .and_then(|o| as_obj_c(&o.get("runProperties")))
            .and_then(|rp| as_obj_c(&rp.get("fontFamily")));
        if !properties.get("numPr").truthy()
            || original_font.is_none()
            || properties.get("_originalRunBoundaries").is_arr()
        {
            return Ok(properties.clone());
        }
        let original = original.unwrap();
        struct Possible {
            paragraph: Js,
            run: Js,
            mark: Js,
            matches: bool,
        }
        let mut possible: Vec<Possible> = Vec::new();
        let p_style = properties.get("pStyle");
        let paragraph_style = p_style.as_str();
        let blocks = self.base_stories.get(story).cloned().unwrap_or_else(|| Js::new_arr(Vec::new()));
        for paragraph in blocks.items() {
            if !type_is(&paragraph, "paragraph") {
                continue;
            }
            let formatting = paragraph.get("formatting");
            let formatting = if formatting.is_nullish() { Js::new_obj() } else { formatting };
            if self.formatting_key(&formatting)? != self.formatting_key(&original)? {
                continue;
            }
            let runs: Vec<Js> = source_runs(&paragraph.get("content"), true)?
                .items()
                .into_iter()
                .filter(|run| js::len16(run.get("text").str_or_empty()) > 0)
                .collect();
            let Some(run) = runs.last().cloned() else { continue };
            let run_formatting = run.get("formatting");
            let run_style = run_formatting.get("styleId");
            let table_run = self.cell_runs.borrow().get(story).cloned().unwrap_or_default();
            let inherited = self.inherited_run(paragraph_style, run_style.as_str(), &table_run)?;
            let expected = if run_formatting.get("fontFamily").truthy() {
                merge_font_family(&inherited.get("fontFamily"), &run_formatting.get("fontFamily"))?
            } else {
                inherited.get("fontFamily")
            };
            let expected_key = self.fonts_key(&expected)?;
            let mut matches = false;
            for item in items.items() {
                let attributes = item.get("attributes");
                let formatting = direct_formatting(&attributes)?;
                let live = as_obj_c(&attributes.get("language"));
                if !same_json_c(&self.fonts_key(&formatting.get("fontFamily"))?, &expected_key) {
                    continue;
                }
                let ok = match &live {
                    None => true,
                    Some(live) => {
                        let language = run_formatting.get("language");
                        let language = if language.is_nullish() { Js::new_obj() } else { language };
                        language.entries().iter().all(|(slot, value)| live.get(slot).strict_eq(value))
                    }
                };
                if ok {
                    matches = true;
                    break;
                }
            }
            possible.push(Possible {
                paragraph: paragraph.clone(),
                run: obj! {
                    "fontFamily" => run_formatting.get("fontFamily"),
                    "language" => run_formatting.get("language"),
                },
                mark: paragraph.path(&["formatting", "runProperties"]),
                matches,
            });
        }
        let matching: Vec<&Possible> = possible.iter().filter(|c| c.matches).collect();
        let candidates: Vec<&Possible> = if !matching.is_empty() { matching } else { possible.iter().collect() };
        let preceding = preceding.items();
        let mut candidate: Option<&Possible> = None;
        for index in (0..preceding.len()).rev() {
            candidate = candidates.iter().copied().find(|c| c.paragraph.strict_eq(&preceding[index]));
            if candidate.is_some() {
                break;
            }
        }
        let candidate = match candidate {
            Some(c) => c,
            None => {
                let Some(first) = candidates.first().copied() else { return Ok(properties.clone()) };
                if candidates
                    .iter()
                    .any(|other| !same_json_c(&other.run, &first.run) || !same_json_c(&other.mark, &first.mark))
                {
                    return Ok(properties.clone());
                }
                first
            }
        };
        let formatting = candidate.run.clone();
        let marks = as_obj_c(&original.get("runProperties"));
        let source_mark = candidate.mark.clone();
        let run_properties = match marks {
            Some(marks) => marks.with(&[
                ("fontFamily", source_mark.get("fontFamily")),
                ("language", source_mark.get("language")),
            ]),
            None => Js::Undefined,
        };
        let text: String = items
            .items()
            .iter()
            .map(|item| {
                if item.get("kind").as_str() == Some("text") {
                    js::to_js_string(&item.get("text"))
                } else {
                    "\u{fffc}".to_owned()
                }
            })
            .collect();
        Ok(properties.with(&[
            ("_originalFormatting", original.with(&[("runProperties", run_properties)])),
            ("_originalRunBoundaries", arr![obj! { "text" => text, "formatting" => formatting }]),
        ]))
    }

    /// `SaveContext.runFormatter` (src/yrs/yrsToDocument.ts:3871). Returns the formatter closure.
    pub(crate) fn run_formatter<'s>(&'s self, story: &str, properties: &Js) -> impl Fn(&Js) -> Result<Js> + 's {
        let paragraph_style = properties.get("pStyle").as_str().map(str::to_owned);
        let story = story.to_owned();
        move |attributes: &Js| {
            let run_style = as_obj_c(&attributes.get("runStyle"))
                .map(|r| r.get("styleId"))
                .and_then(|s| as_str_c(&s))
                .filter(|s| !s.is_empty());
            let table_run = self.cell_runs.borrow().get(&story).cloned().unwrap_or_default();
            let inherited = self.inherited_run(paragraph_style.as_deref(), run_style.as_deref(), &table_run)?;
            relative_formatting(attributes, &inherited, &self.font_scheme)
        }
    }

    /// `SaveContext.inheritedRun` (src/yrs/yrsToDocument.ts:3881).
    pub(crate) fn inherited_run(&self, paragraph_style: Option<&str>, run_style: Option<&str>, table_run: &Js) -> Result<Js> {
        let key = arr![Js::from(paragraph_style), Js::from(run_style), table_run.clone()]
            .stringify()
            .unwrap_or_default();
        if let Some(inherited) = self.inherited_runs.borrow().get(&key) {
            return Ok(inherited.clone());
        }
        let inheritance = self.inheritance()?;
        let table_run = if table_run.is_undefined() { None } else { Some(table_run) };
        let inherited = attrs_to_text_formatting(&inheritance.attributes(paragraph_style, run_style, table_run))?;
        self.inherited_runs.borrow_mut().insert(key, inherited.clone());
        Ok(inherited)
    }

    /// `SaveContext.dispose` (src/yrs/yrsToDocument.ts:3892).
    pub(crate) fn dispose(&self) -> Result<()> {
        *self.run_inheritance.borrow_mut() = None;
        Ok(())
    }

    /// `SaveContext.recovery` (src/yrs/yrsToDocument.ts:3897): `{ scoped, warnings }`, read through
    /// [`recovery_drawing_for`], [`recovery_unrecoverable`] and [`recovery_warn`].
    pub(crate) fn recovery(&self, story_id: &str) -> Result<Js> {
        Ok(obj! {
            "scoped" => self.base_chart_drawings.get(story_id).cloned().unwrap_or_default(),
            "warnings" => self.warnings.clone(),
        })
    }

    /// `SaveContext.storyToBlocks` (src/yrs/yrsToDocument.ts:3915).
    pub(crate) fn story_to_blocks(&self, story_id: &str) -> Result<Js> {
        let base_blocks: Option<Js> = self.base_stories.get(story_id).cloned();
        let story_comments = self.comments.get(story_id).cloned().unwrap_or_else(|| Js::new_arr(Vec::new()));
        let comments_key = match self.subtree_comments.get(story_id) {
            Some(nested) => stable_stringify(&Js::obj_from(nested.clone()))?,
            None => String::new(),
        };
        let prior_story = self.memo.get("stories").map_get(story_id);
        if !prior_story.is_undefined()
            && prior_story.get("commentsKey").as_str() == Some(comments_key.as_str())
            && self.story_is_clean(story_id)?
            && (nested_story_parent(story_id).is_some()
                || base_blocks.is_none()
                || base_blocks.as_ref().unwrap().strict_eq(&prior_story.get("blocks")))
        {
            return Ok(prior_story.get("blocks"));
        }
        let mut blocks: Vec<Js> = Vec::new();
        let paragraph_ids = self.session.story_paragraph_ids(story_id)?;
        let segments = self.session.story_segments(story_id)?.items();
        let mut trailing_by_paragraph: HashMap<String, Js> = HashMap::new();
        let mut closing: Option<String> = None;
        for segment in &segments {
            if segment.get("kind").as_str() == Some("pilcrow") {
                let id = key_of(&segment.get("paraId"));
                trailing_by_paragraph.insert(id.clone(), Js::new_arr(Vec::new()));
                closing = Some(id);
            } else {
                let kind = if segment.get("kind").as_str() == Some("embed") {
                    block_break_kind_c(&segment.get("embedKind"))
                } else {
                    None
                };
                match (kind, &closing) {
                    (Some(kind), Some(c)) => trailing_by_paragraph[c].push(kind),
                    _ => closing = None,
                }
            }
        }
        let mut items = Js::new_arr(Vec::new());
        let mut paragraph_start: f64 = 0.0;
        let mut inline_start: f64 = 0.0;
        let mut paragraph_index = 0usize;
        let preceding_base_paragraphs = Js::new_arr(Vec::new());
        let mut story_offset: f64 = 0.0;
        let mut candidates_by_key: Option<IndexMap<String, Vec<Js>>> = None;

        let candidates_for = |key: &str, cache: &mut Option<IndexMap<String, Vec<Js>>>| -> Vec<Js> {
            if cache.is_none() {
                let mut map: IndexMap<String, Vec<Js>> = IndexMap::new();
                for block in base_blocks.as_ref().map(|b| b.items()).unwrap_or_default() {
                    let Some(block_key) = projected_blocks_get(&block).and_then(|m| m.key) else { continue };
                    map.entry(block_key).or_default().push(block);
                }
                *cache = Some(map);
            }
            cache.as_ref().unwrap().get(key).cloned().unwrap_or_default()
        };
        let reuse_container = |candidates: Vec<Js>, inputs: &Js, children: &[Js]| -> Option<Js> {
            candidates.into_iter().find(|block| {
                let Some(memo) = projected_blocks_get(block) else { return false };
                match (&memo.inputs, &memo.children) {
                    (Some(memo_inputs), Some(memo_children)) => {
                        same_json_c(memo_inputs, inputs)
                            && memo_children.len() == children.len()
                            && memo_children.iter().zip(children.iter()).all(|(a, b)| a.strict_eq(b))
                    }
                    _ => false,
                }
            })
        };
        let paragraph_comment_boundaries = |end: f64, paragraph_start: f64, inline_start: f64| -> Js {
            let boundaries = Js::new_arr(Vec::new());
            for range in story_comments.items() {
                let start = js::to_number(&range.get("start"));
                let range_end = js::to_number(&range.get("end"));
                if start >= paragraph_start && start <= end {
                    let offset = (start - inline_start).max(0.0);
                    boundaries.push(obj! { "id" => range.get("id"), "kind" => "start", "offset" => offset });
                }
                if range_end >= paragraph_start && range_end <= end {
                    boundaries.push(obj! {
                        "id" => range.get("id"),
                        "kind" => "end",
                        "offset" => (range_end - inline_start).max(0.0),
                    });
                }
            }
            boundaries
        };
        let push_text = |items: &Js, text: &str, attributes: &Js| {
            let units = js::utf16(text);
            let mut cursor = 0;
            for index in 0..units.len() {
                if units[index] != b'\t' as u16 {
                    continue;
                }
                if index > cursor {
                    items.push(obj! {
                        "kind" => "text",
                        "text" => js::from_utf16(&units[cursor..index]),
                        "attributes" => attributes.clone(),
                    });
                }
                items.push(obj! {
                    "kind" => "embed",
                    "embedKind" => "tab",
                    "payload" => Js::new_obj(),
                    "attributes" => attributes.clone(),
                });
                cursor = index + 1;
            }
            if cursor < units.len() {
                items.push(obj! {
                    "kind" => "text",
                    "text" => js::from_utf16(&units[cursor..]),
                    "attributes" => attributes.clone(),
                });
            }
        };

        let mut witnesses: Vec<String> = Vec::new();
        let mut breaks = Js::new_arr(Vec::new());
        for segment in &segments {
            let kind_str = segment.get("kind");
            if kind_str.as_str() == Some("text") {
                let text = js::to_js_string(&segment.get("text"));
                push_text(&items, &text, &segment.get("attributes"));
                story_offset += js::len16(&text) as f64;
                witnesses = Vec::new();
                continue;
            }
            if kind_str.as_str() == Some("pilcrow") {
                let para_id = key_of(&segment.get("paraId"));
                let live_trailing = trailing_by_paragraph.get(&para_id).cloned().unwrap_or_default();
                let saved = paragraph_ids.at(paragraph_index);
                let saved_para_id = if saved.is_nullish() { String::new() } else { key_of(&saved) };
                let base_paragraph = self.base_paragraphs.get(&para_id).cloned();
                if let Some(bp) = &base_paragraph {
                    preceding_base_paragraphs.push(bp.clone());
                }
                let base_paragraph_js = base_paragraph.clone().unwrap_or_default();
                let boundaries = paragraph_comment_boundaries(story_offset, paragraph_start, inline_start);
                let source = source_break_boundaries(&base_paragraph_js)?;
                let source_runs_value = if source.items().iter().any(|b| b.path(&["breaks", "length"]).truthy()) {
                    source
                } else {
                    Js::Undefined
                };
                let properties = segment.get("properties");
                let inputs = arr![
                    segment.get("paraId"),
                    saved_para_id.as_str(),
                    properties.clone(),
                    segment.get("attributes"),
                    items.clone(),
                    boundaries.clone(),
                    breaks.clone(),
                    live_trailing.clone(),
                    source_runs_value.clone(),
                ];
                let prior_memo = base_paragraph.as_ref().and_then(projected_blocks_get);
                let mut paragraph: Js;
                let reuse = match (&base_paragraph, &prior_memo) {
                    (Some(_), Some(memo)) => memo.inputs.as_ref().map_or(false, |i| same_json_c(i, &inputs)),
                    _ => false,
                };
                if reuse {
                    let bp = base_paragraph.clone().unwrap();
                    witnesses = match prior_memo.and_then(|m| m.trailing) {
                        Some(t) => t,
                        None => trailing_block_breaks_vec(&bp)?,
                    };
                    paragraph = bp;
                } else {
                    // Projection mutates `item.payload`; pin the pre-mutation inputs.
                    let snapshot = Js::new_arr(
                        inputs
                            .items()
                            .iter()
                            .enumerate()
                            .map(|(index, input)| {
                                if index == 4 {
                                    Js::new_arr(input.items().iter().map(Js::spread).collect())
                                } else {
                                    input.clone()
                                }
                            })
                            .collect(),
                    );
                    let paragraph_properties = if base_paragraph.is_some() {
                        properties.clone()
                    } else {
                        self.copied_paragraph_properties(story_id, &properties, &items, &preceding_base_paragraphs)?
                    };
                    let live_breaks = obj! {
                        "leading" => breaks.clone(),
                        "trailing" => live_trailing.clone(),
                        "leadingPage" => properties.get("pageBreakBeforeRun").is_true(),
                    };
                    let recovery = self.recovery(story_id)?;
                    let formatter = self.run_formatter(story_id, &properties);
                    let inherited = self.inherited_formatting(story_id, &properties)?;
                    paragraph = paragraph_from_story(
                        &saved_para_id,
                        &paragraph_properties,
                        &items,
                        &boundaries,
                        &base_paragraph_js,
                        &live_breaks,
                        &recovery,
                        &formatter,
                        &self.revision_ids,
                        &inherited,
                    )?;
                    self.restore_comment_references(&paragraph)?;
                    let trailing = trailing_block_breaks_vec(&paragraph)?;
                    let leading = leading_block_breaks_vec(&paragraph)?;
                    let break_list = arr_strings(&breaks);
                    let consumed = if leading.iter().enumerate().all(|(i, k)| break_list.get(i) == Some(k)) {
                        leading.len()
                    } else {
                        0
                    };
                    let pending: Vec<String> = break_list.get(consumed..).map(<[String]>::to_vec).unwrap_or_default();
                    if !pending.is_empty() {
                        let mut content: Vec<Js> = pending
                            .iter()
                            .map(|break_type| {
                                obj! {
                                    "type" => "run",
                                    "content" => arr![obj! { "type" => "break", "breakType" => break_type.as_str() }],
                                }
                            })
                            .collect();
                        content.extend(paragraph.get("content").items());
                        paragraph = clone_projected_paragraph(
                            &paragraph,
                            &obj! { "content" => Js::new_arr(content) },
                            &Js::Undefined,
                        )?;
                    }
                    projected_blocks_set(
                        &paragraph,
                        ProjectedBlockMemo {
                            inputs: Some(snapshot),
                            session_key: Some(para_id.clone()),
                            trailing: Some(trailing.clone()),
                            source_runs: source_runs_value.clone(),
                            ..Default::default()
                        },
                    );
                    witnesses = trailing;
                }
                blocks.push(paragraph);
                items = Js::new_arr(Vec::new());
                breaks = Js::new_arr(Vec::new());
                paragraph_index += 1;
                story_offset += 1.0;
                paragraph_start = story_offset;
                inline_start = story_offset;
                continue;
            }

            let embed_kind = segment.get("embedKind");
            let payload = segment.get("payload");
            if embed_kind.as_str() == Some("table") {
                self.prepare_table(&payload, story_id)?;
                let inputs = arr![segment.get("attributes"), payload.clone()];
                let rows = payload.get("rows");
                let first_cell = if rows.is_arr() {
                    let cells = rows.at(0).get("cells");
                    if cells.is_arr() { cells.at(0) } else { cells.get("0") }
                } else {
                    Js::Undefined
                };
                let key = format!("T{}", js::to_js_string(&first_cell.get("story").nullish_or(Js::from(""))));
                let contents = if !candidates_for(&key, &mut candidates_by_key).is_empty() {
                    Some(self.cell_contents(&payload, story_id)?)
                } else {
                    None
                };
                let reused = match &contents {
                    Some(contents) => {
                        reuse_container(candidates_for(&key, &mut candidates_by_key), &inputs, &contents.items())
                    }
                    None => None,
                };
                let table = match reused {
                    Some(table) => table,
                    None => {
                        let table = table_from_payload(self, &payload)?;
                        let children = match contents {
                            Some(c) => c,
                            None => self.cell_contents(&payload, story_id)?,
                        };
                        projected_blocks_set(
                            &table,
                            ProjectedBlockMemo {
                                key: Some(key.clone()),
                                inputs: Some(inputs.clone()),
                                table_baseline: self.payload_baselines.get(&payload),
                                table_source_prefix: first_cell.get("story").as_str().and_then(table_prefix_of),
                                children: Some(children.items()),
                                ..Default::default()
                            },
                        );
                        table
                    }
                };
                blocks.push(table);
            } else if embed_kind.as_str() == Some("blockSdt") {
                let child_story = as_str_c(&payload.get("story"));
                if let Some(child) = child_story.as_ref().filter(|s| !s.is_empty()) {
                    self.memo.get("parents").map_set(child, story_id);
                }
                let child_content = match child_story.as_ref() {
                    Some(child) if !child.is_empty() && self.story_ids.contains(child) => self.story_to_blocks(child)?,
                    _ => empty_blocks(),
                };
                let inputs = arr![segment.get("attributes"), payload.clone()];
                let key = format!("S{}", child_story.clone().unwrap_or_default());
                let reused = reuse_container(
                    candidates_for(&key, &mut candidates_by_key),
                    &inputs,
                    &[child_content.clone()],
                );
                if let Some(reused) = reused {
                    blocks.push(reused);
                } else {
                    let mut properties = sdt_attrs_to_props(&payload)?;
                    let mut content = child_content.clone();
                    let authored_value = content_control_value(&payload.get("value"))?;
                    if authored_value.truthy() {
                        if let Ok(applied) = apply_content_control_value(&properties, &authored_value) {
                            properties = applied.get("properties");
                            if let Ok(carried) = carry_paragraph_identities(&applied.get("content"), &child_content) {
                                content = carried;
                            }
                        }
                        // Retain the child story if the authored value is invalid.
                    }
                    let block_content =
                        if content.strict_eq(&empty_blocks()) { Js::new_arr(Vec::new()) } else { content.clone() };
                    let block = obj! {
                        "type" => "blockSdt",
                        "properties" => properties,
                        "content" => block_content.clone(),
                    };
                    // Rebind the child's entry to the embedded array for the next projection.
                    if !content.strict_eq(&child_content) {
                        if let Some(child) = &child_story {
                            let child_entry = self.memo.get("stories").map_get(child);
                            if child_entry.truthy() {
                                child_entry.set("blocks", block_content.clone());
                            }
                        }
                    }
                    projected_blocks_set(
                        &block,
                        ProjectedBlockMemo {
                            key: Some(key.clone()),
                            inputs: Some(inputs.clone()),
                            children: Some(vec![block_content.clone()]),
                            ..Default::default()
                        },
                    );
                    blocks.push(block);
                }
            } else if embed_kind.as_str() == Some("opaque") {
                let blob = as_obj_c(&payload.get("blob"));
                let blob_type = blob.as_ref().map(|b| b.get("type")).unwrap_or_default();
                if blob_type.as_str() == Some("pageBreak") {
                    blocks.push(page_break_paragraph()?);
                } else {
                    // Carry a same-position base block while block SDTs stay opaque.
                    let base_block = base_blocks.as_ref().map(|b| b.at(blocks.len())).unwrap_or_default();
                    if blob_type.as_str() == Some("blockSdt") && type_is(&base_block, "blockSdt") {
                        blocks.push(base_block);
                    }
                    // Standalone text-box blobs do not have a one-to-one base block (the base
                    // model stores them inside paragraph runs), so they remain the documented
                    // opaque carry gap.
                }
            } else {
                items.push(segment.clone());
            }
            match block_break_kind_c(&embed_kind) {
                None => witnesses = Vec::new(),
                Some(kind) => {
                    if let Some(pos) = witnesses.iter().position(|w| w == kind) {
                        witnesses.remove(pos);
                    } else {
                        breaks.push(kind);
                    }
                }
            }
            story_offset += 1.0;
            if items.items().iter().all(is_block_break_c) {
                inline_start = story_offset;
            }
        }

        let pending_breaks: Vec<Js> = breaks
            .items()
            .iter()
            .map(|break_type| {
                obj! {
                    "type" => "run",
                    "content" => arr![obj! { "type" => "break", "breakType" => break_type.clone() }],
                }
            })
            .collect();
        // Defensive recovery for malformed/legacy stories without a final pilcrow.
        if items.len() > 0 {
            let trailing = build_paragraph_content(
                &items,
                &self.recovery(story_id)?,
                &|a: &Js| direct_formatting(a),
                &self.revision_ids,
                false,
            )?;
            if trailing.len() > 0 {
                let content = if breaks.len() > 0 {
                    let mut c = pending_breaks.clone();
                    c.extend(trailing.items());
                    Js::new_arr(c)
                } else {
                    trailing
                };
                blocks.push(obj! { "type" => "paragraph", "content" => content });
                breaks = Js::new_arr(Vec::new());
            }
        }
        if breaks.len() > 0 {
            let index = blocks.iter().rposition(|b| type_is(b, "paragraph"));
            match index {
                Some(index) => {
                    let previous = blocks[index].clone();
                    let mut content = previous.get("content").items();
                    content.extend(pending_breaks.iter().cloned());
                    blocks[index] = clone_projected_paragraph(
                        &previous,
                        &obj! { "content" => Js::new_arr(content) },
                        &Js::Undefined,
                    )?;
                }
                None => blocks.push(obj! { "type" => "paragraph", "content" => Js::new_arr(pending_breaks) }),
            }
        }
        let base_arr = base_blocks.clone().unwrap_or_else(|| Js::new_arr(Vec::new()));
        let projected = restore_raw_blocks(&Js::new_arr(blocks), &base_arr, story_id)?;
        self.memo
            .get("stories")
            .map_set(story_id, obj! { "commentsKey" => comments_key, "blocks" => projected.clone() });
        map_delete(&self.memo.get("dirty"), story_id);
        self.memo.get("clean").map_set(story_id, true);
        Ok(projected)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn nested_story_ids() {
        assert_eq!(nested_story_parent("body:t0:r1c2"), Some("body".into()));
        assert_eq!(nested_story_parent("body:t0:r1c2~3"), Some("body".into()));
        assert_eq!(nested_story_parent("body:sdt0:t1:r0c0"), Some("body:sdt0".into()));
        assert_eq!(nested_story_parent("body:sdt4"), Some("body".into()));
        assert_eq!(nested_story_parent(":sdt4"), Some("".into()));
        assert_eq!(nested_story_parent("body"), None);
        assert_eq!(nested_story_parent("body:t0:r1"), None);
    }

    #[test]
    fn table_prefixes() {
        assert_eq!(table_prefix_of("body:t0:r1c2"), Some("body:t0".into()));
        assert_eq!(table_prefix_of("body:t0:r1c2:t3:r0c0~1"), Some("body:t0:r1c2:t3".into()));
        assert_eq!(table_prefix_of("body:sdt0"), None);
        assert_eq!(table_prefix_of(":t:r1c2"), None);
    }

    #[test]
    fn block_bookmarks() {
        let b = raw_block_bookmark("  <w:bookmarkStart w:id=\"007\" w:name=\"a&amp;b\"/>").unwrap();
        assert_eq!(b.stringify().unwrap(), r#"{"kind":"start","id":"7"}"#);
        let b = raw_block_bookmark("<w:bookmarkEnd w:id='x'/>").unwrap();
        assert_eq!(b.stringify().unwrap(), r#"{"kind":"end","id":"x"}"#);
        assert!(raw_block_bookmark("<w:bookmarkStartX/>").unwrap().is_undefined());
        assert!(raw_block_bookmark("<w:bookmarkStart foo:id=\"1\"/>").is_err());
        assert!(raw_block_bookmark("<w:bookmarkStart w:id=\"1\" w:id=\"2\"/>").is_err());
        let b = raw_block_bookmark("<w:bookmarkEnd xmlns:v=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" v:id=\"&#x31;\"/>").unwrap();
        assert_eq!(b.get("id").as_str(), Some("1"));
        assert_eq!(decode_entities("&lt;&#65;&bogus;&#x110000;"), "<A&bogus;&#x110000;");
    }

    #[test]
    fn same_json_shapes() {
        let a = Js::parse(r#"{"a":[1,{"b":null}]}"#).unwrap();
        let b = Js::parse(r#"{"a":[1,{"b":null}]}"#).unwrap();
        assert!(same_json_c(&a, &b));
        assert!(!same_json_c(&a, &Js::parse(r#"{"a":[1,{"b":0}]}"#).unwrap()));
    }

    #[test]
    fn break_tokens() {
        let p = Js::parse(r#"{"content":[{"type":"run","content":[{"type":"break","breakType":"page"},{"type":"text","text":"x"},{"type":"break","breakType":"column"}]}]}"#).unwrap();
        assert_eq!(trailing_block_breaks_vec(&p).unwrap(), vec!["column".to_string()]);
        assert_eq!(leading_block_breaks_vec(&p).unwrap(), Vec::<String>::new());
    }
}
