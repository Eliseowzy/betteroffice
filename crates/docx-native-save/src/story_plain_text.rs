//! Port of `packages/docx/src/yrs/storyPlainText.ts`.
#![allow(unused_variables, unused_imports, dead_code, clippy::all)]

use crate::js::{Js, Result, SaveError};
use crate::session::{Inheritance, Session};
use crate::{arr, obj, throw};
use indexmap::IndexMap;
use std::collections::{HashMap, HashSet};
use std::rc::Rc;

pub(crate) const MAX_TABLE_COLUMNS: f64 = 16_384.0;
const PLAIN_TEXT_SLOT_LIMIT: f64 = (1u32 << 20) as f64;

fn js_max(a: f64, b: f64) -> f64 {
    if a.is_nan() || b.is_nan() {
        f64::NAN
    } else {
        a.max(b)
    }
}
fn js_min(a: f64, b: f64) -> f64 {
    if a.is_nan() || b.is_nan() {
        f64::NAN
    } else {
        a.min(b)
    }
}
fn num(v: &Js) -> f64 {
    crate::js::to_number(v)
}

/// `s.replace(/\n$/, '')`.
fn strip_final_newline(s: String) -> String {
    match s.strip_suffix('\n') {
        Some(t) => t.to_owned(),
        None => s,
    }
}

/// `positiveSpan` (src/yrs/storyPlainText.ts:43).
pub(crate) fn positive_span(value: &Js) -> Result<f64> {
    Ok(match value {
        Js::Num(n) if n.fract() == 0.0 && n.is_finite() && *n > 0.0 => *n,
        _ => 1.0,
    })
}

/// `objectValue` (src/yrs/storyPlainText.ts:47). `Js::Null` for a non-object.
pub(crate) fn object_value(value: &Js) -> Result<Js> {
    Ok(
        if !value.is_nullish() && value.is_object_like() && !value.is_arr() {
            value.clone()
        } else {
            Js::Null
        },
    )
}

/// `embedPlainText` (src/yrs/storyPlainText.ts:54).
pub(crate) fn embed_plain_text(reader: &Session, kind: &str, payload: &Js) -> Result<String> {
    let s = |k: &str| {
        payload
            .get(k)
            .as_str()
            .map(str::to_owned)
            .unwrap_or_default()
    };
    Ok(match kind {
        "tab" => "\t".into(),
        "break" => "\n".into(),
        "field" => s("displayText"),
        "math" => s("plainText"),
        "sdt" => {
            let content = payload.get("content");
            if !content.is_arr() {
                return Ok(String::new());
            }
            let mut out = String::new();
            for item in content.items() {
                if item.get("kind").as_str() == Some("text") && item.get("text").is_str() {
                    out.push_str(item.get("text").str_or_empty());
                } else {
                    let p = object_value(&item.get("payload"))?;
                    let p = if p.is_null() { Js::new_obj() } else { p };
                    out.push_str(&embed_plain_text(
                        reader,
                        &crate::js::to_js_string(&item.get("kind")),
                        &p,
                    )?);
                }
            }
            out
        }
        "blockSdt" => match payload.get("story").as_str() {
            Some(story) => format!(
                "{}\n",
                strip_final_newline(story_plain_text(reader, story, 0.0, f64::INFINITY)?)
            ),
            None => String::new(),
        },
        "table" => {
            let rows = payload.get("rows");
            if !rows.is_arr() {
                return Ok(String::new());
            }
            table_plain_text(reader, &obj! { "rows" => rows }, &Js::Undefined)?
                .into_iter()
                .map(|row| format!("{row}\n"))
                .collect()
        }
        _ => String::new(),
    })
}

/// `cellPlainText` (src/yrs/storyPlainText.ts:90).
pub(crate) fn cell_plain_text(reader: &Session, story: &str) -> Result<String> {
    let text = strip_final_newline(story_plain_text(reader, story, 0.0, f64::INFINITY)?);
    Ok(if text.contains(['\t', '\n', '"']) {
        format!("\"{}\"", text.replace('"', "\"\""))
    } else {
        text
    })
}

/// `tablePlainText` (src/yrs/storyPlainText.ts:96). `range` is `Js::Undefined` when absent.
pub(crate) fn table_plain_text(reader: &Session, payload: &Js, range: &Js) -> Result<Vec<String>> {
    let ta = table_anchors(payload)?;
    let anchors = ta.get("anchors").items();
    let columns = num(&ta.get("columns"));
    let top = num(&range.get("top").nullish_or(Js::Num(0.0)));
    let bottom = num(&range
        .get("bottom")
        .nullish_or(Js::Num(payload.get("rows").len() as f64 - 1.0)));
    let left = num(&range.get("left").nullish_or(Js::Num(0.0)));
    let right = num(&range.get("right").nullish_or(Js::Num(columns - 1.0)));
    let mut lines = Vec::new();
    let cell_text = |story: &Js| -> Result<String> {
        match story.as_str() {
            Some(s) if !s.is_empty() => cell_plain_text(reader, s),
            _ => Ok(String::new()),
        }
    };
    if (bottom - top + 1.0) * (right - left + 1.0) > PLAIN_TEXT_SLOT_LIMIT {
        let mut index = 0;
        let mut row = top;
        while row <= bottom {
            let mut texts = Vec::new();
            while index < anchors.len() && num(&anchors[index].get("row")) < row {
                index += 1;
            }
            while index < anchors.len() && num(&anchors[index].get("row")) == row {
                let cell = &anchors[index];
                index += 1;
                let c = num(&cell.get("column"));
                if c >= left && c <= right {
                    texts.push(cell_text(&cell.get("story"))?);
                }
            }
            lines.push(texts.join("\t"));
            row += 1.0;
        }
        return Ok(lines);
    }
    let mut by_grid: HashMap<String, Js> = HashMap::new();
    for cell in &anchors {
        by_grid.insert(
            format!(
                "{}:{}",
                crate::js::to_js_string(&cell.get("row")),
                crate::js::to_js_string(&cell.get("column"))
            ),
            cell.get("story"),
        );
    }
    let mut row = top;
    while row <= bottom {
        let mut texts = Vec::new();
        let mut column = left;
        while column <= right {
            let story = by_grid
                .get(&format!(
                    "{}:{}",
                    crate::js::num_to_string(row),
                    crate::js::num_to_string(column)
                ))
                .cloned()
                .unwrap_or_default();
            texts.push(cell_text(&story)?);
            column += 1.0;
        }
        lines.push(texts.join("\t"));
        row += 1.0;
    }
    Ok(lines)
}

/// `storyPlainText` (src/yrs/storyPlainText.ts:135). Callers pass `0.0, f64::INFINITY` for the defaults.
pub(crate) fn story_plain_text(
    reader: &Session,
    story: &str,
    from: f64,
    to: f64,
) -> Result<String> {
    let mut text = String::new();
    let mut offset = 0.0;
    for segment in reader.story_segments(story)?.items() {
        let start = offset;
        let kind = segment.get("kind");
        let seg_text = segment.get("text");
        offset += if kind.as_str() == Some("text") {
            seg_text.len() as f64
        } else {
            1.0
        };
        if offset <= from {
            continue;
        }
        if start >= to {
            break;
        }
        match kind.as_str() {
            Some("text") => text.push_str(&crate::js::slice16(
                seg_text.str_or_empty(),
                js_max(from, start) - start,
                Some(js_min(to, offset) - start),
            )),
            Some("pilcrow") => text.push('\n'),
            _ => text.push_str(&embed_plain_text(
                reader,
                &crate::js::to_js_string(&segment.get("embedKind")),
                &segment.get("payload"),
            )?),
        }
    }
    Ok(text)
}

/// `tableRowGrid` (src/yrs/storyPlainText.ts:154).
pub(crate) fn table_row_grid(tr_pr: &Js) -> Result<Js> {
    let original = tr_pr.get("_originalFormatting");
    let count = |key: &str| {
        let value = if tr_pr.truthy() && tr_pr.has(key) {
            tr_pr.get(key)
        } else {
            original.get(key)
        };
        match value {
            Js::Num(n) if n.is_finite() && n.fract() == 0.0 && n >= 0.0 => n.min(MAX_TABLE_COLUMNS),
            _ => 0.0,
        }
    };
    Ok(obj! { "before" => count("gridBefore"), "after" => count("gridAfter") })
}

/// `tableAnchors` (src/yrs/storyPlainText.ts:165).
pub(crate) fn table_anchors(payload: &Js) -> Result<Js> {
    // (start, end, lastRow)
    let mut active_merges: Vec<(f64, f64, f64)> = Vec::new();
    let anchors = Js::new_arr(Vec::new());
    let grid = payload.get("grid");
    let grid_len = if grid.is_nullish() {
        0.0
    } else if grid.is_arr() || grid.is_str() {
        grid.len() as f64
    } else {
        f64::NAN
    };
    let mut columns = js_min(grid_len, MAX_TABLE_COLUMNS);
    let rows = payload.get("rows").items();
    let row_count = rows.len() as f64;
    for (row_index, row) in rows.iter().enumerate() {
        let row_index = row_index as f64;
        active_merges.retain(|m| m.2 >= row_index);
        let mut next_merges: Vec<(f64, f64, f64)> = Vec::new();
        let mut merge_index = 0;
        let grid_counts = table_row_grid(&row.get("trPr"))?;
        let (before, after) = (
            num(&grid_counts.get("before")),
            num(&grid_counts.get("after")),
        );
        let mut column = before;
        for cell in row.get("cells").items() {
            while merge_index < active_merges.len() && active_merges[merge_index].0 <= column {
                let merge = active_merges[merge_index];
                merge_index += 1;
                column = js_max(column, merge.1);
                next_merges.push(merge);
            }
            let tc_pr = cell.get("tcPr");
            let rowspan = js_max(
                1.0,
                js_min(positive_span(&tc_pr.get("rowspan"))?, row_count - row_index),
            );
            let colspan = js_max(
                1.0,
                js_min(
                    positive_span(&tc_pr.get("colspan"))?,
                    MAX_TABLE_COLUMNS - column,
                ),
            );
            anchors.push(obj! { "row" => row_index, "column" => column, "rowspan" => rowspan, "colspan" => colspan, "story" => cell.get("story") });
            if rowspan > 1.0 {
                next_merges.push((column, column + colspan, row_index + rowspan - 1.0));
            }
            column += colspan;
            columns = js_max(columns, column);
        }
        while merge_index < active_merges.len() {
            next_merges.push(active_merges[merge_index]);
            merge_index += 1;
        }
        let row_end = next_merges.iter().fold(column, |end, m| js_max(end, m.1));
        columns = js_max(columns, js_min(MAX_TABLE_COLUMNS, row_end + after));
        active_merges = next_merges;
    }
    Ok(obj! { "anchors" => anchors, "columns" => columns })
}

/// `storyOffsetForLoc` (src/yrs/storyPlainText.ts:207). The reader's `locateParagraph` is a
/// live session method; here `reader` must carry it as a `locateParagraph` object of
/// `{ [story]: { [paraId]: { start } } }` spans.
pub(crate) fn story_offset_for_loc(reader: &Js, loc: &Js) -> Result<f64> {
    let story = crate::js::to_js_string(&loc.get("story"));
    let para_id = crate::js::to_js_string(&loc.get("paraId"));
    let span = reader.path(&["locateParagraph", &story, &para_id]);
    if span.is_nullish() {
        throw!("paragraph {para_id} is not in story {story}");
    }
    Ok(num(&span.get("start")) + num(&loc.get("offset")))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn table_anchors_match_ts() {
        let payload = Js::parse(
            r#"{"grid":[1,2],"rows":[
{"trPr":{"gridBefore":1},"cells":[{"story":"a","tcPr":{"rowspan":2,"colspan":2}},{"story":"b"}]},
{"trPr":{"_originalFormatting":{"gridAfter":2}},"cells":[{"story":"c"}]}]}"#,
        )
        .unwrap();
        assert_eq!(
            table_anchors(&payload).unwrap().stringify().unwrap(),
            r#"{"anchors":[{"row":0,"column":1,"rowspan":2,"colspan":2,"story":"a"},{"row":0,"column":3,"rowspan":1,"colspan":1,"story":"b"},{"row":1,"column":0,"rowspan":1,"colspan":1,"story":"c"}],"columns":5}"#
        );
    }
}
