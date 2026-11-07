//! Port of `packages/docx/src/docx/tiffMedia.ts`.
#![allow(unused_variables, unused_imports, dead_code, clippy::all)]

use std::collections::{HashMap, HashSet};
use std::rc::Rc;
use indexmap::IndexMap;
use crate::js::{Js, Result, SaveError};
use crate::session::{Inheritance, Session};
use crate::{arr, obj, throw};

/// `maybeTranscodeTiffMedia` (src/docx/tiffMedia.ts:15).
pub(crate) fn maybe_transcode_tiff_media(bytes: &Rc<Vec<u8>>, mime_type: &str, data_url: &str, transcode: &dyn Fn(&Rc<Vec<u8>>) -> Result<Vec<u8>>) -> Result<Js> {
    todo!("maybeTranscodeTiffMedia")
}

/// `encodeBase64` (src/docx/tiffMedia.ts:34).
pub(crate) fn encode_base64(bytes: &Rc<Vec<u8>>) -> Result<String> {
    todo!("encodeBase64")
}
