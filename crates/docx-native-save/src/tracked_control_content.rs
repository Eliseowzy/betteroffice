//! Port of `packages/docx/src/utils/trackedControlContent.ts`.
#![allow(unused_variables, unused_imports, dead_code, clippy::all)]

use std::collections::{HashMap, HashSet};
use std::rc::Rc;
use indexmap::IndexMap;
use crate::js::{Js, Result, SaveError};
use crate::session::{Inheritance, Session};
use crate::{arr, obj, throw};

/// `visitTrackedControlContent` (src/utils/trackedControlContent.ts:3).
pub(crate) fn visit_tracked_control_content(node: &Js, visit: &dyn Fn(&Js) -> Result<()>, control: bool, revision: bool) -> Result<()> {
    todo!("visitTrackedControlContent")
}

/// `hasTrackedControlContent` (src/utils/trackedControlContent.ts:35).
pub(crate) fn has_tracked_control_content(node: &Js) -> Result<bool> {
    todo!("hasTrackedControlContent")
}
