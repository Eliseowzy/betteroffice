//! Port of `packages/docx/src/yrs/tableRowSpans.ts`.
#![allow(unused_variables, unused_imports, dead_code, clippy::all)]

use crate::js::{Js, Result, SaveError};
use crate::session::{Inheritance, Session};
use crate::{arr, obj, throw};
use indexmap::IndexMap;
use std::collections::{HashMap, HashSet};
use std::rc::Rc;

fn num_key(v: f64) -> String {
    crate::js::num_to_string(v)
}

/// `calculateRowSpans` (src/yrs/tableRowSpans.ts:7). A `Js::Map` keyed by `row-column` of
/// `{ rowSpan, skip }` objects.
pub(crate) fn calculate_row_spans(table: &Js) -> Result<Js> {
    let result = Js::new_map();
    // `active`: column -> start row (keyed by the number's string form).
    let mut active: IndexMap<String, f64> = IndexMap::new();
    for (row_index, row) in table.get("rows").items().into_iter().enumerate() {
        let row_index = row_index as f64;
        let mut column = row
            .path(&["formatting", "gridBefore"])
            .nullish_or(Js::Num(0.0));
        let mut cells: Vec<(String, Js, String)> = Vec::new();
        for cell in row.get("cells").items() {
            let current = column.clone();
            column = add(
                &column,
                &cell
                    .path(&["formatting", "gridSpan"])
                    .nullish_or(Js::Num(1.0)),
            );
            let col = crate::js::to_js_string(&current);
            cells.push((
                col.clone(),
                cell.path(&["formatting", "vMerge"]),
                format!("{}-{}", num_key(row_index), col),
            ));
        }
        let empty = !cells.is_empty()
            && cells
                .iter()
                .all(|(col, vm, _)| vm.as_str() == Some("continue") && active.contains_key(col));
        if empty {
            for (col, _, key) in &cells {
                active.shift_remove(col);
                result.map_set(key, obj! { "rowSpan" => 1, "skip" => false });
            }
            continue;
        }
        for (col, vm, key) in &cells {
            match vm.as_str() {
                Some("restart") => {
                    active.insert(col.clone(), row_index);
                    result.map_set(key, obj! { "rowSpan" => 1, "skip" => false });
                }
                Some("continue") => match active.get(col) {
                    None => result.map_set(key, obj! { "rowSpan" => 1, "skip" => false }),
                    Some(start) => {
                        let owner = result.map_get(&format!("{}-{}", num_key(*start), col));
                        if owner.truthy() {
                            owner.set("rowSpan", crate::js::to_number(&owner.get("rowSpan")) + 1.0);
                        }
                        result.map_set(key, obj! { "rowSpan" => 1, "skip" => true });
                    }
                },
                _ => {
                    active.shift_remove(col);
                    result.map_set(key, obj! { "rowSpan" => 1, "skip" => false });
                }
            }
        }
    }
    Ok(result)
}

/// `column += x` (numbers; string concatenation if either is a string).
fn add(a: &Js, b: &Js) -> Js {
    if a.is_str() || b.is_str() {
        Js::from(format!(
            "{}{}",
            crate::js::to_js_string(a),
            crate::js::to_js_string(b)
        ))
    } else {
        Js::Num(crate::js::to_number(a) + crate::js::to_number(b))
    }
}

/// `seededCellIndices` (src/yrs/tableRowSpans.ts:49).
pub(crate) fn seeded_cell_indices(table: &Js) -> Result<Js> {
    let spans = calculate_row_spans(table)?;
    let out = Js::new_arr(Vec::new());
    for (row_index, row) in table.get("rows").items().into_iter().enumerate() {
        let mut column = row
            .path(&["formatting", "gridBefore"])
            .nullish_or(Js::Num(0.0));
        let kept = Js::new_arr(Vec::new());
        for (index, cell) in row.get("cells").items().into_iter().enumerate() {
            let start = column.clone();
            column = add(
                &column,
                &cell
                    .path(&["formatting", "gridSpan"])
                    .nullish_or(Js::Num(1.0)),
            );
            let key = format!("{}-{}", row_index, crate::js::to_js_string(&start));
            if !spans.map_get(&key).get("skip").truthy() {
                kept.push(index);
            }
        }
        out.push(kept);
    }
    Ok(out)
}
