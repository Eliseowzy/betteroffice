//! Port of `packages/docx/src/yrs/yrsToDocument.ts`, in three parts by source line:
//! `inline` (to line 1454), `paragraph` (1455–2907) and `context` (2908 on).
#![allow(unused_imports)]

mod context;
mod inline;
mod paragraph;

pub(crate) use context::*;
pub(crate) use inline::*;
pub(crate) use paragraph::*;
