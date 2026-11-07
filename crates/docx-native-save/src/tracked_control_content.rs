//! Port of `packages/docx/src/utils/trackedControlContent.ts`.
#![allow(unused_variables, unused_imports, dead_code, clippy::all)]

use crate::js::{Js, Result, SaveError};
use crate::session::{Inheritance, Session};
use crate::{arr, obj, throw};
use indexmap::IndexMap;
use std::collections::{HashMap, HashSet};
use std::rc::Rc;

fn is_revision(t: &str) -> bool {
    matches!(t, "insertion" | "deletion" | "moveFrom" | "moveTo")
}

/// `visitTrackedControlContent` (src/utils/trackedControlContent.ts:3). Callers pass
/// `false, false` for the TS defaults.
pub(crate) fn visit_tracked_control_content(
    node: &Js,
    visit: &dyn Fn(&Js) -> Result<()>,
    control: bool,
    revision: bool,
) -> Result<()> {
    let ty = node.get("type");
    let ty = ty.str_or_empty();
    let control = control || ty == "inlineSdt";
    let revision = revision || is_revision(ty);
    if control && revision {
        visit(node)?;
    }
    let children = if ty == "inlineSdt" || is_revision(ty) {
        node.get("content")
    } else if ty == "hyperlink" {
        node.get("structuredChildren")
            .nullish_or(node.get("children"))
    } else if ty == "simpleField" {
        node.path(&["structuredResult", "inline"])
            .nullish_or(node.get("content"))
    } else if ty == "complexField" {
        for child in node
            .path(&["structuredCode", "inline"])
            .nullish_or(node.get("fieldCode"))
            .items()
        {
            if child.get("type").as_str() == Some("run") && !(control && revision) {
                continue;
            }
            visit_tracked_control_content(&child, visit, control, revision)?;
        }
        node.path(&["structuredResult", "inline"])
            .nullish_or(node.get("fieldResult"))
    } else {
        return Ok(());
    };
    for child in children.items() {
        if child.get("type").as_str() == Some("run") && !(control && revision) {
            continue;
        }
        visit_tracked_control_content(&child, visit, control, revision)?;
    }
    Ok(())
}

/// `hasTrackedControlContent` (src/utils/trackedControlContent.ts:35).
pub(crate) fn has_tracked_control_content(node: &Js) -> Result<bool> {
    let found = std::cell::Cell::new(false);
    visit_tracked_control_content(
        node,
        &|_| {
            found.set(true);
            Ok(())
        },
        false,
        false,
    )?;
    Ok(found.get())
}
