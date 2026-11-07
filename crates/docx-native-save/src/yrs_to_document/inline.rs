//! Port of `packages/docx/src/yrs/yrsToDocument.ts` (part).
#![allow(unused_variables, unused_imports, dead_code, clippy::all)]

use std::collections::{HashMap, HashSet};
use std::rc::Rc;
use indexmap::IndexMap;
use crate::js::{Js, Result, SaveError};
use crate::session::{Inheritance, Session};
use crate::{arr, obj, throw};
use super::*;

/// `asObject` (src/yrs/yrsToDocument.ts:282).
pub(crate) fn as_object(value: &Js) -> Result<Js> {
    todo!("asObject")
}

/// `asFiniteNumber` (src/yrs/yrsToDocument.ts:288).
pub(crate) fn as_finite_number(value: &Js) -> Result<Option<f64>> {
    todo!("asFiniteNumber")
}

/// `asString` (src/yrs/yrsToDocument.ts:292).
pub(crate) fn as_string(value: &Js) -> Result<Option<String>> {
    todo!("asString")
}

/// `dropNulls` (src/yrs/yrsToDocument.ts:296).
pub(crate) fn drop_nulls(value: &Js) -> Result<Js> {
    todo!("dropNulls")
}

/// `stableStringify` (src/yrs/yrsToDocument.ts:312).
pub(crate) fn stable_stringify(value: &Js) -> Result<String> {
    todo!("stableStringify")
}

/// `stableStringifyObject` (src/yrs/yrsToDocument.ts:331).
pub(crate) fn stable_stringify_object(value: &Js, omitted: Option<&HashSet<String>>) -> Result<String> {
    todo!("stableStringifyObject")
}

/// `stableStringifyWithout` (src/yrs/yrsToDocument.ts:344).
pub(crate) fn stable_stringify_without(value: &Js, omitted: &HashSet<String>) -> Result<String> {
    todo!("stableStringifyWithout")
}

/// `formattingAttrs` (src/yrs/yrsToDocument.ts:348).
pub(crate) fn formatting_attrs(attributes: &Js) -> Result<Js> {
    todo!("formattingAttrs")
}

/// `attrsToTextFormatting` (src/yrs/yrsToDocument.ts:357).
pub(crate) fn attrs_to_text_formatting(attributes: &Js) -> Result<Js> {
    todo!("attrsToTextFormatting")
}

/// `directFormatting` (src/yrs/yrsToDocument.ts:473).
pub(crate) fn direct_formatting(attributes: &Js) -> Result<Js> {
    todo!("directFormatting")
}

/// `lentByLinkStyle` (src/yrs/yrsToDocument.ts:483).
pub(crate) fn lent_by_link_style(mark: &Js) -> Result<bool> {
    todo!("lentByLinkStyle")
}

/// `canonicalFontPair` (src/yrs/yrsToDocument.ts:485).
pub(crate) fn canonical_font_pair(fonts: &Js, name: &str, r#ref: &str, theme: &Js) -> Result<Js> {
    todo!("canonicalFontPair")
}

/// `relativeFontFamily` (src/yrs/yrsToDocument.ts:493).
pub(crate) fn relative_font_family(live: &Js, inherited: &Js, source: &Js, theme: &Js) -> Result<Js> {
    todo!("relativeFontFamily")
}

/// `relativeFormatting` (src/yrs/yrsToDocument.ts:512).
pub(crate) fn relative_formatting(attributes: &Js, inherited: &Js, theme: &Js) -> Result<Js> {
    todo!("relativeFormatting")
}

/// `runContentForText` (src/yrs/yrsToDocument.ts:548).
pub(crate) fn run_content_for_text(text: &str, formatting: &Js) -> Result<Js> {
    todo!("runContentForText")
}

/// `createTextRun` (src/yrs/yrsToDocument.ts:597).
pub(crate) fn create_text_run(text: &str, attributes: &Js, format: &dyn Fn(&Js) -> Result<Js>) -> Result<Js> {
    todo!("createTextRun")
}

/// `runPropertyChanges` (src/yrs/yrsToDocument.ts:607).
pub(crate) fn run_property_changes(attributes: &Js, format: &dyn Fn(&Js) -> Result<Js>) -> Result<Js> {
    todo!("runPropertyChanges")
}

/// `appendTextRun` (src/yrs/yrsToDocument.ts:626).
pub(crate) fn append_text_run(target: &Js, source: &Js) -> Result<()> {
    todo!("appendTextRun")
}

/// `fnv53` (src/yrs/yrsToDocument.ts:634).
pub(crate) fn fnv53(value: &str) -> Result<f64> {
    todo!("fnv53")
}

/// `revisionId` (src/yrs/yrsToDocument.ts:643).
pub(crate) fn revision_id(value: &Js) -> Result<f64> {
    todo!("revisionId")
}

/// `trackedInfo` (src/yrs/yrsToDocument.ts:653).
pub(crate) fn tracked_info(raw: &Js, _pm_shape: bool, id: Option<f64>) -> Result<Js> {
    todo!("trackedInfo")
}

/// `propertyChanges` (src/yrs/yrsToDocument.ts:666).
pub(crate) fn property_changes(changes: &Js) -> Result<Js> {
    todo!("propertyChanges")
}

/// `mappedRevisionId` (src/yrs/yrsToDocument.ts:684).
pub(crate) fn mapped_revision_id(raw: &Js, kind: &Js, revision_ids: &Js) -> Result<Option<f64>> {
    todo!("mappedRevisionId")
}

/// `createHyperlink` (src/yrs/yrsToDocument.ts:694).
pub(crate) fn create_hyperlink(attributes: &Js) -> Result<Js> {
    todo!("createHyperlink")
}

/// `hyperlinkKey` (src/yrs/yrsToDocument.ts:717).
pub(crate) fn hyperlink_key(attributes: &Js) -> Result<Option<String>> {
    todo!("hyperlinkKey")
}

/// `fieldFromPayload` (src/yrs/yrsToDocument.ts:722).
pub(crate) fn field_from_payload(payload: &Js, attributes: &Js, format: &dyn Fn(&Js) -> Result<Js>) -> Result<Js> {
    todo!("fieldFromPayload")
}

/// `mathFromPayload` (src/yrs/yrsToDocument.ts:781).
pub(crate) fn math_from_payload(payload: &Js) -> Result<Js> {
    todo!("mathFromPayload")
}

/// `horizontalRuleRun` (src/yrs/yrsToDocument.ts:790).
pub(crate) fn horizontal_rule_run(payload: &Js, attributes: &Js, format: &dyn Fn(&Js) -> Result<Js>) -> Result<Js> {
    todo!("horizontalRuleRun")
}

/// `mediaSrc` (src/yrs/yrsToDocument.ts:820).
pub(crate) fn media_src(src: &str) -> Result<String> {
    todo!("mediaSrc")
}

/// `resolveMediaSrcs` (src/yrs/yrsToDocument.ts:824).
pub(crate) fn resolve_media_srcs(value: &Js) -> Result<()> {
    todo!("resolveMediaSrcs")
}

/// `imageFromPayload` (src/yrs/yrsToDocument.ts:837).
pub(crate) fn image_from_payload(payload: &Js) -> Result<Js> {
    todo!("imageFromPayload")
}

/// `imageRunFromPayload` (src/yrs/yrsToDocument.ts:936).
pub(crate) fn image_run_from_payload(payload: &Js) -> Result<Js> {
    todo!("imageRunFromPayload")
}

/// `storedShape` (src/yrs/yrsToDocument.ts:941).
pub(crate) fn stored_shape(value: &Js) -> Result<Js> {
    todo!("storedShape")
}

/// `chartRunFromPayload` (src/yrs/yrsToDocument.ts:955).
pub(crate) fn chart_run_from_payload(payload: &Js, recovery: &Js) -> Result<Js> {
    todo!("chartRunFromPayload")
}

/// `replayedRunProperties` (src/yrs/yrsToDocument.ts:983).
pub(crate) fn replayed_run_properties(payload: &Js, attributes: &Js, format: &dyn Fn(&Js) -> Result<Js>) -> Result<Js> {
    todo!("replayedRunProperties")
}

/// `opaqueDrawingRunFromPayload` (src/yrs/yrsToDocument.ts:1002).
pub(crate) fn opaque_drawing_run_from_payload(payload: &Js, attributes: &Js, format: &dyn Fn(&Js) -> Result<Js>) -> Result<Js> {
    todo!("opaqueDrawingRunFromPayload")
}

/// `shapeRunFromPayload` (src/yrs/yrsToDocument.ts:1027).
pub(crate) fn shape_run_from_payload(payload: &Js, attributes: &Js, format: &dyn Fn(&Js) -> Result<Js>) -> Result<Js> {
    todo!("shapeRunFromPayload")
}

/// `inlineSdtFromPayload` (src/yrs/yrsToDocument.ts:1112).
pub(crate) fn inline_sdt_from_payload(payload: &Js, recovery: &Js, format: &dyn Fn(&Js) -> Result<Js>) -> Result<Js> {
    todo!("inlineSdtFromPayload")
}

/// `inlineSdtContent` (src/yrs/yrsToDocument.ts:1164).
pub(crate) fn inline_sdt_content(content: &Js) -> Result<Js> {
    todo!("inlineSdtContent")
}

/// `fieldInlineContent` (src/yrs/yrsToDocument.ts:1180).
pub(crate) fn field_inline_content(content: &Js) -> Result<Js> {
    todo!("fieldInlineContent")
}

/// `contentControlValue` (src/yrs/yrsToDocument.ts:1192).
pub(crate) fn content_control_value(value: &Js) -> Result<Js> {
    todo!("contentControlValue")
}

/// `commentReferenceFromPayload` (src/yrs/yrsToDocument.ts:1206).
pub(crate) fn comment_reference_from_payload(payload: &Js) -> Result<Js> {
    todo!("commentReferenceFromPayload")
}

/// `ordinaryContentForItem` (src/yrs/yrsToDocument.ts:1215).
pub(crate) fn ordinary_content_for_item(item: &Js, recovery: &Js, format: &dyn Fn(&Js) -> Result<Js>) -> Result<Js> {
    todo!("ordinaryContentForItem")
}

/// `inlineContentForItem` (src/yrs/yrsToDocument.ts:1224).
pub(crate) fn inline_content_for_item(item: &Js, recovery: &Js, format: &dyn Fn(&Js) -> Result<Js>) -> Result<Js> {
    todo!("inlineContentForItem")
}

/// `tabRun` (src/yrs/yrsToDocument.ts:1271).
pub(crate) fn tab_run(attributes: &Js, format: &dyn Fn(&Js) -> Result<Js>) -> Result<Js> {
    todo!("tabRun")
}

/// `trackedContentForItem` (src/yrs/yrsToDocument.ts:1281).
pub(crate) fn tracked_content_for_item(item: &Js, info: &Js, in_control: bool, recovery: &Js, format: &dyn Fn(&Js) -> Result<Js>) -> Result<Js> {
    todo!("trackedContentForItem")
}

/// `addToHyperlink` (src/yrs/yrsToDocument.ts:1328).
pub(crate) fn add_to_hyperlink(hyperlink: &Js, item: &Js, recovery: &Js, format: &dyn Fn(&Js) -> Result<Js>) -> Result<()> {
    todo!("addToHyperlink")
}

/// `projectionSignature` (src/yrs/yrsToDocument.ts:1361).
pub(crate) fn projection_signature(items: &Js) -> Result<String> {
    todo!("projectionSignature")
}

/// `restoreProjectedFieldResults` (src/yrs/yrsToDocument.ts:1378).
pub(crate) fn restore_projected_field_results(items: &Js, recovery: &Js, format: &dyn Fn(&Js) -> Result<Js>) -> Result<Js> {
    todo!("restoreProjectedFieldResults")
}

/// `isTrackedWrapper` (src/yrs/yrsToDocument.ts:1446).
pub(crate) fn is_tracked_wrapper(content: &Js) -> Result<bool> {
    todo!("isTrackedWrapper")
}
