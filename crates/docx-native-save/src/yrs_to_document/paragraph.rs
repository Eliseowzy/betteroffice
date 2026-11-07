//! Port of `packages/docx/src/yrs/yrsToDocument.ts` (part).
#![allow(unused_variables, unused_imports, dead_code, clippy::all)]

use std::collections::{HashMap, HashSet};
use std::rc::Rc;
use indexmap::IndexMap;
use crate::js::{Js, Result, SaveError};
use crate::session::{Inheritance, Session};
use crate::{arr, obj, throw};
use super::*;

/// `buildParagraphContent` (src/yrs/yrsToDocument.ts:1455).
pub(crate) fn build_paragraph_content(items: &Js, recovery: &Js, format: &dyn Fn(&Js) -> Result<Js>, revision_ids: &Js, in_control: bool) -> Result<Js> {
    todo!("buildParagraphContent")
}

/// `marksKeyToYrsAttrs` (src/yrs/yrsToDocument.ts:1554).
pub(crate) fn marks_key_to_yrs_attrs(marks_key: Option<&str>) -> Result<Js> {
    todo!("marksKeyToYrsAttrs")
}

/// `noteMarkContent` (src/yrs/yrsToDocument.ts:1587).
pub(crate) fn note_mark_content(boundary: &Js) -> Result<Js> {
    todo!("noteMarkContent")
}

/// `boundaryContent` (src/yrs/yrsToDocument.ts:1598).
pub(crate) fn boundary_content(boundary: &Js, formatting: &Js) -> Result<Js> {
    todo!("boundaryContent")
}

/// `sourceBreakBoundaries` (src/yrs/yrsToDocument.ts:1618).
pub(crate) fn source_break_boundaries(paragraph: &Js) -> Result<Js> {
    todo!("sourceBreakBoundaries")
}

/// `liveRunBoundaries` (src/yrs/yrsToDocument.ts:1644).
pub(crate) fn live_run_boundaries(boundaries: &Js, live: &Js) -> Result<Js> {
    todo!("liveRunBoundaries")
}

/// `paragraphBreaksAreLive` (src/yrs/yrsToDocument.ts:1677).
pub(crate) fn paragraph_breaks_are_live(content: &Js, live: &Js) -> Result<bool> {
    todo!("paragraphBreaksAreLive")
}

/// `restoreRecordedBreaks` (src/yrs/yrsToDocument.ts:1685).
pub(crate) fn restore_recorded_breaks(content: &Js, items: &Js, boundaries: &Js, live_breaks: &Js) -> Result<Js> {
    todo!("restoreRecordedBreaks")
}

/// `restoreOriginalRuns` (src/yrs/yrsToDocument.ts:1747).
pub(crate) fn restore_original_runs(content: &Js, items: &Js, boundaries: &Js, base_paragraph: &Js, live_breaks: &Js) -> Result<Js> {
    todo!("restoreOriginalRuns")
}

/// `runContentLength` (src/yrs/yrsToDocument.ts:1819).
pub(crate) fn run_content_length(content: &Js) -> Result<f64> {
    todo!("runContentLength")
}

/// `runTextLength` (src/yrs/yrsToDocument.ts:1842).
pub(crate) fn run_text_length(run: &Js) -> Result<f64> {
    todo!("runTextLength")
}

/// `linkChildLength` (src/yrs/yrsToDocument.ts:1847).
pub(crate) fn link_child_length(child: &Js) -> Result<f64> {
    todo!("linkChildLength")
}

/// `paragraphContentLength` (src/yrs/yrsToDocument.ts:1862).
pub(crate) fn paragraph_content_length(content: &Js) -> Result<f64> {
    todo!("paragraphContentLength")
}

/// `splitItems` (src/yrs/yrsToDocument.ts:1895).
pub(crate) fn split_items(items: &Js, at: f64, length: &dyn Fn(&Js) -> Result<f64>, split: &dyn Fn(&Js, f64) -> Result<Js>) -> Result<Js> {
    todo!("splitItems")
}

/// `splitRunContent` (src/yrs/yrsToDocument.ts:1919).
pub(crate) fn split_run_content(content: &Js, at: f64) -> Result<Js> {
    todo!("splitRunContent")
}

/// `linkPart` (src/yrs/yrsToDocument.ts:1928).
pub(crate) fn link_part(link: &Js, children: &Js) -> Result<Js> {
    todo!("linkPart")
}

/// `splitContent` (src/yrs/yrsToDocument.ts:1944).
pub(crate) fn split_content(content: &Js, at: f64) -> Result<Js> {
    todo!("splitContent")
}

/// `insertBoundaries` (src/yrs/yrsToDocument.ts:1980).
pub(crate) fn insert_boundaries(content: &Js, boundaries: &Js, make_marker: &dyn Fn(&Js) -> Result<Js>) -> Result<Js> {
    todo!("insertBoundaries")
}

/// `bookmarkBoundaries` (src/yrs/yrsToDocument.ts:2047).
pub(crate) fn bookmark_boundaries(properties: &Js) -> Result<Js> {
    todo!("bookmarkBoundaries")
}

/// `nestedParagraphs` (src/yrs/yrsToDocument.ts:2071).
pub(crate) fn nested_paragraphs(value: &Js, found: &Js) -> Result<Js> {
    todo!("nestedParagraphs")
}

/// `restoreNestedOrdinals` (src/yrs/yrsToDocument.ts:2087).
pub(crate) fn restore_nested_ordinals(content: &Js, base: &Js) -> Result<()> {
    todo!("restoreNestedOrdinals")
}

/// `restoreRawInlines` (src/yrs/yrsToDocument.ts:2098).
pub(crate) fn restore_raw_inlines(content: &Js, base: &Js) -> Result<Js> {
    todo!("restoreRawInlines")
}

/// `paragraphAttrs` (src/yrs/yrsToDocument.ts:2110).
pub(crate) fn paragraph_attrs(properties: &Js) -> Result<Js> {
    todo!("paragraphAttrs")
}

/// `sourceRuns` (src/yrs/yrsToDocument.ts:2119).
pub(crate) fn source_runs(content: &Js, into_fields: bool) -> Result<Js> {
    todo!("sourceRuns")
}

/// `editMidpoint` (src/yrs/yrsToDocument.ts:2143).
pub(crate) fn edit_midpoint(original: &str, current: &str) -> Result<Js> {
    todo!("editMidpoint")
}

/// `sourceTextMatches` (src/yrs/yrsToDocument.ts:2203).
pub(crate) fn source_text_matches(original: &str, current: &str) -> Result<Js> {
    todo!("sourceTextMatches")
}

/// `preserveSourceRunFormatting` (src/yrs/yrsToDocument.ts:2232).
pub(crate) fn preserve_source_run_formatting(items: &Js, sources: &Js) -> Result<Js> {
    todo!("preserveSourceRunFormatting")
}

/// `restoreLineBreaks` (src/yrs/yrsToDocument.ts:2278).
pub(crate) fn restore_line_breaks(content: &Js, base: &Js) -> Result<()> {
    todo!("restoreLineBreaks")
}

/// `paragraphFromStory` (src/yrs/yrsToDocument.ts:2305).
pub(crate) fn paragraph_from_story(para_id: &str, properties: &Js, items: &Js, comment_boundaries: &Js, base_paragraph: &Js, live_breaks: &Js, recovery: &Js, format: &dyn Fn(&Js) -> Result<Js>, revision_ids: &Js, inherited: &Js) -> Result<Js> {
    todo!("paragraphFromStory")
}

/// `inferTableBorders` (src/yrs/yrsToDocument.ts:2414).
pub(crate) fn infer_table_borders(rows: &Js) -> Result<Js> {
    todo!("inferTableBorders")
}

/// `normalizeVMergeRuns` (src/yrs/yrsToDocument.ts:2431).
pub(crate) fn normalize_v_merge_runs(rows: &Js) -> Result<()> {
    todo!("normalizeVMergeRuns")
}

/// `tableCellFromPayload` (src/yrs/yrsToDocument.ts:2482).
pub(crate) fn table_cell_from_payload(context: &crate::yrs_to_document::SaveContext, payload: &Js) -> Result<Js> {
    todo!("tableCellFromPayload")
}

/// `tableFromPayload` (src/yrs/yrsToDocument.ts:2513).
pub(crate) fn table_from_payload(context: &crate::yrs_to_document::SaveContext, payload: &Js) -> Result<Js> {
    todo!("tableFromPayload")
}

/// `tableCellKey` (src/yrs/yrsToDocument.ts:2652).
pub(crate) fn table_cell_key(entry: &Js) -> Result<Option<String>> {
    todo!("tableCellKey")
}

/// `baseTableCells` (src/yrs/yrsToDocument.ts:2657).
pub(crate) fn base_table_cells(table: &Js, prefix: &str) -> Result<Js> {
    todo!("baseTableCells")
}

/// `gridStarts` (src/yrs/yrsToDocument.ts:2688).
pub(crate) fn grid_starts(row: &Js) -> Result<Vec<f64>> {
    todo!("gridStarts")
}

/// `rowCellBookmarks` (src/yrs/yrsToDocument.ts:2699).
pub(crate) fn row_cell_bookmarks(row: &Js) -> Result<Vec<String>> {
    todo!("rowCellBookmarks")
}

/// `tableHasBookmarks` (src/yrs/yrsToDocument.ts:2703).
pub(crate) fn table_has_bookmarks(table: &Js) -> Result<bool> {
    todo!("tableHasBookmarks")
}

/// `bookmarkMarker` (src/yrs/yrsToDocument.ts:2709).
pub(crate) fn bookmark_marker(key: &str, scan: &Js) -> Result<Js> {
    todo!("bookmarkMarker")
}

/// `distributeBookmarks` (src/yrs/yrsToDocument.ts:2713).
pub(crate) fn distribute_bookmarks(restored: &Js, items: &Js, leading: &dyn Fn(&Js, &str) -> Result<()>, trailing: &dyn Fn(&str) -> Result<()>) -> Result<()> {
    todo!("distributeBookmarks")
}

/// `restoreTableBookmarks` (src/yrs/yrsToDocument.ts:2731).
pub(crate) fn restore_table_bookmarks(context: &crate::yrs_to_document::SaveContext, table: &Js, cell_stories: &Js, prefix_hint: Option<&str>) -> Result<()> {
    todo!("restoreTableBookmarks")
}

/// `pageBreakParagraph` (src/yrs/yrsToDocument.ts:2851).
pub(crate) fn page_break_paragraph() -> Result<Js> {
    todo!("pageBreakParagraph")
}

/// `isNoteMark` (src/yrs/yrsToDocument.ts:2859).
pub(crate) fn is_note_mark(content: &Js) -> Result<bool> {
    todo!("isNoteMark")
}

/// `hasNoteMark` (src/yrs/yrsToDocument.ts:2877).
pub(crate) fn has_note_mark(blocks: &Js) -> Result<bool> {
    todo!("hasNoteMark")
}

/// `restoreNoteMarks` (src/yrs/yrsToDocument.ts:2886).
pub(crate) fn restore_note_marks(projected: &Js, base: &Js) -> Result<Js> {
    todo!("restoreNoteMarks")
}
