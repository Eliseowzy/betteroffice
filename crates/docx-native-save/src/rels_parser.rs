//! Port of `packages/docx/src/docx/relsParser.ts`.
#![allow(dead_code)]

use crate::js::{Js, Result};

/// `RELATIONSHIP_TYPES.header`.
pub(crate) const RELATIONSHIP_TYPE_HEADER: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";
/// `RELATIONSHIP_TYPES.footer`.
pub(crate) const RELATIONSHIP_TYPE_FOOTER: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";

/// `resolveTarget` (src/docx/relsParser.ts:37).
pub(crate) fn resolve_target(map: &Js, r_id: &str) -> Result<Option<String>> {
    Ok(map.map_get(r_id).get("target").as_str().map(str::to_owned))
}
