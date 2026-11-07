//! Port of `packages/docx/src/docx/headerFooterAliasProjection.ts`.
#![allow(unused_variables, unused_imports, dead_code, clippy::all)]

use std::collections::{HashMap, HashSet};
use std::rc::Rc;
use indexmap::IndexMap;
use crate::js::{Js, Result, SaveError};
use crate::session::{Inheritance, Session};
use crate::{arr, obj, throw};

/// `markHeaderFooterAlias` (src/docx/headerFooterAliasProjection.ts:6).
pub(crate) fn mark_header_footer_alias(entry: &Js, canonical_r_id: &str) -> Result<()> {
    todo!("markHeaderFooterAlias")
}

/// `headerFooterAliasOf` (src/docx/headerFooterAliasProjection.ts:11).
pub(crate) fn header_footer_alias_of(entry: &Js) -> Result<Option<String>> {
    todo!("headerFooterAliasOf")
}

/// `isWrittenByCanonical` (src/docx/headerFooterAliasProjection.ts:16).
pub(crate) fn is_written_by_canonical(entry: &Js, entries: &Js) -> Result<bool> {
    todo!("isWrittenByCanonical")
}
