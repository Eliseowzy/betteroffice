//! Port of `packages/docx/src/docx/tiffMedia.ts`.

use std::rc::Rc;

use base64::Engine as _;

use crate::js::{Js, Result};
use crate::obj;

/// `MAX_TIFF_BYTES` (shared/media.ts:2).
const MAX_TIFF_BYTES: usize = 32 * 1024 * 1024;

/// `isTiff` (shared/media.ts:5).
fn is_tiff(bytes: &[u8]) -> bool {
    bytes.len() >= 4
        && ((bytes[0] == 0x49 && bytes[1] == 0x49 && bytes[2] == 0x2a && bytes[3] == 0)
            || (bytes[0] == 0x4d && bytes[1] == 0x4d && bytes[2] == 0 && bytes[3] == 0x2a))
}

/// `maybeTranscodeTiffMedia` (src/docx/tiffMedia.ts:15): `{ bytes, mimeType, dataUrl }`.
pub(crate) fn maybe_transcode_tiff_media(
    bytes: &Rc<Vec<u8>>,
    mime_type: &str,
    data_url: &str,
    transcode: &dyn Fn(&Rc<Vec<u8>>) -> Result<Vec<u8>>,
) -> Result<Js> {
    let untouched = || obj! { "bytes" => Js::Bytes(bytes.clone()), "mimeType" => mime_type, "dataUrl" => data_url };
    if !is_tiff(bytes) || bytes.len() > MAX_TIFF_BYTES {
        return Ok(untouched());
    }
    match transcode(bytes) {
        Ok(png) => {
            let png = Rc::new(png);
            let url = format!("data:image/png;base64,{}", encode_base64(&png)?);
            Ok(obj! { "bytes" => Js::Bytes(png), "mimeType" => "image/png", "dataUrl" => url })
        }
        Err(_) => Ok(untouched()),
    }
}

/// `encodeBase64` (src/docx/tiffMedia.ts:34).
pub(crate) fn encode_base64(bytes: &Rc<Vec<u8>>) -> Result<String> {
    Ok(base64::engine::general_purpose::STANDARD.encode(bytes.as_slice()))
}
