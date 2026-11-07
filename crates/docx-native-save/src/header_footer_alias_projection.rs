//! Port of `packages/docx/src/docx/headerFooterAliasProjection.ts`.
#![allow(unused_variables, unused_imports, dead_code, clippy::all)]

use crate::js::{Js, Result, SaveError};
use crate::session::{Inheritance, Session};
use crate::{arr, obj, throw};
use indexmap::IndexMap;
use std::collections::{HashMap, HashSet};
use std::rc::Rc;

thread_local! {
    static ALIASES: crate::js::Marks<String> = crate::js::Marks::new();
}

/// `markHeaderFooterAlias` (src/docx/headerFooterAliasProjection.ts:6).
pub(crate) fn mark_header_footer_alias(entry: &Js, canonical_r_id: &str) -> Result<()> {
    ALIASES.with(|a| a.set(entry, canonical_r_id.to_owned()));
    Ok(())
}

/// `headerFooterAliasOf` (src/docx/headerFooterAliasProjection.ts:11).
pub(crate) fn header_footer_alias_of(entry: &Js) -> Result<Option<String>> {
    Ok(ALIASES.with(|a| a.get(entry)))
}

/// `isWrittenByCanonical` (src/docx/headerFooterAliasProjection.ts:16).
pub(crate) fn is_written_by_canonical(entry: &Js, entries: &Js) -> Result<bool> {
    Ok(match header_footer_alias_of(entry)? {
        Some(id) => entries.map_has(&id),
        None => false,
    })
}

/// Drops the `ALIASES` entries of objects that no longer exist.
pub(crate) fn prune_aliases() {
    ALIASES.with(|marks| marks.clear());
}
