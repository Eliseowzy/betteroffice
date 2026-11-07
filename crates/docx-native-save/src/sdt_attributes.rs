//! Port of `packages/docx/src/types/sdtAttributes.ts`.
#![allow(unused_variables, unused_imports, dead_code, clippy::all)]

use std::collections::{HashMap, HashSet};
use std::rc::Rc;
use indexmap::IndexMap;
use crate::js::{Js, Result, SaveError};
use crate::session::{Inheritance, Session};
use crate::{arr, obj, throw};

/// `sdtPropsToAttrs` (src/types/sdtAttributes.ts:4).
pub(crate) fn sdt_props_to_attrs(props: &Js) -> Result<Js> {
    todo!("sdtPropsToAttrs")
}

/// `sdtAttrsToProps` (src/types/sdtAttributes.ts:24).
pub(crate) fn sdt_attrs_to_props(attrs: &Js) -> Result<Js> {
    todo!("sdtAttrsToProps")
}
