//! Port of `packages/docx/src/docx/rezip/parts.ts`.
#![allow(dead_code)]

use std::rc::Rc;

use crate::header_footer_alias_projection::is_written_by_canonical;
use crate::js::{Js, Result, SaveError};
use crate::obj;
use crate::rels_parser::{RELATIONSHIP_TYPE_FOOTER, RELATIONSHIP_TYPE_HEADER};

/// `toBytes` (src/docx/rezip/parts.ts:27): a string as UTF-8, bytes as they are.
pub(crate) fn to_bytes(content: &Js) -> Result<Vec<u8>> {
    match content {
        Js::Str(s) => Ok(s.as_bytes().to_vec()),
        Js::Bytes(b) => Ok((**b).clone()),
        _ => Ok(Vec::new()),
    }
}

/// `partText` (src/docx/rezip/parts.ts:34): UTF-8 keeping a leading BOM.
pub(crate) fn part_text(bytes: Option<&Rc<Vec<u8>>>) -> Result<Option<String>> {
    Ok(bytes.map(|b| String::from_utf8_lossy(b).into_owned()))
}

/// `rezipPartsToArrayBuffer` (src/docx/rezip/parts.ts:40): `pkg` a `Map<path, bytes>`.
pub(crate) fn rezip_parts_to_array_buffer(pkg: &Js) -> Result<Vec<u8>> {
    let mut entries = Vec::new();
    for (path, content) in pkg.map_entries() {
        entries.push((path, to_bytes(&content)?));
    }
    ooxml_opc::rezip_parts(&entries).map_err(SaveError::new)
}

const EMPTY_RELS_XML: &str = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"></Relationships>";

/// `headerFooterFilename` (src/docx/rezip/parts.ts:65).
pub(crate) fn header_footer_filename(target: &str) -> Result<String> {
    Ok(match target.strip_prefix('/') {
        Some(rest) => rest.to_owned(),
        None => format!("word/{target}"),
    })
}

/// `collectParts` (src/docx/rezip/parts.ts:75): `[{ relsPath, blocks }]`.
pub(crate) fn collect_parts(doc: &Js) -> Result<Js> {
    let pkg = doc.get("package");
    let parts = Js::new_arr(vec![obj! {
        "relsPath" => "word/_rels/document.xml.rels",
        "blocks" => pkg.path(&["document", "content"]),
    }]);
    let note_blocks = |notes: &Js| -> Vec<Js> {
        notes
            .items()
            .iter()
            .flat_map(|note| flat(&note.get("content")))
            .collect()
    };
    let mut footnote_blocks = note_blocks(&pkg.get("footnoteSeparators"));
    footnote_blocks.extend(note_blocks(&pkg.get("footnotes")));
    if !footnote_blocks.is_empty() {
        parts.push(obj! { "relsPath" => "word/_rels/footnotes.xml.rels", "blocks" => Js::new_arr(footnote_blocks) });
    }
    let mut endnote_blocks = note_blocks(&pkg.get("endnoteSeparators"));
    endnote_blocks.extend(note_blocks(&pkg.get("endnotes")));
    if !endnote_blocks.is_empty() {
        parts.push(obj! { "relsPath" => "word/_rels/endnotes.xml.rels", "blocks" => Js::new_arr(endnote_blocks) });
    }
    let rels = pkg.get("relationships");
    if !rels.truthy() {
        return Ok(parts);
    }
    let add = |map: &Js, kind: &str| -> Result<()> {
        if !map.truthy() {
            return Ok(());
        }
        for (r_id, hf) in map.map_entries() {
            if is_written_by_canonical(&hf, map)? {
                continue;
            }
            let rel = rels.map_get(&r_id);
            if !rel.truthy()
                || rel.get("type").as_str() != Some(kind)
                || !rel.get("target").truthy()
            {
                continue;
            }
            let filename = header_footer_filename(rel.get("target").str_or_empty())?;
            let basename = filename.strip_prefix("word/").unwrap_or(&filename);
            parts.push(obj! { "relsPath" => format!("word/_rels/{basename}.rels"), "blocks" => hf.get("content") });
        }
        Ok(())
    };
    add(&pkg.get("headers"), RELATIONSHIP_TYPE_HEADER)?;
    add(&pkg.get("footers"), RELATIONSHIP_TYPE_FOOTER)?;
    Ok(parts)
}

/// `Array.prototype.flatMap`'s flattening of one returned value.
fn flat(value: &Js) -> Vec<Js> {
    if value.is_arr() {
        value.items()
    } else {
        vec![value.clone()]
    }
}

/// `readRelsOrStub` (src/docx/rezip/parts.ts:125).
pub(crate) fn read_rels_or_stub(pkg: &Js, rels_path: &str) -> Result<String> {
    let xml = match pkg.map_get(rels_path) {
        Js::Bytes(b) => String::from_utf8_lossy(&b).into_owned(),
        _ => EMPTY_RELS_XML.to_owned(),
    };
    // xml.replace(/<Relationships([^>]*)\/>/, '<Relationships$1></Relationships>')
    let open = "<Relationships";
    let mut from = 0;
    while let Some(found) = xml[from..].find(open) {
        let start = from + found;
        let body = start + open.len();
        if let Some(gt) = xml[body..].find('>') {
            let gt = body + gt;
            if gt > body && xml.as_bytes()[gt - 1] == b'/' {
                let attrs = &xml[body..gt - 1];
                return Ok(format!(
                    "{}<Relationships{attrs}></Relationships>{}",
                    &xml[..start],
                    &xml[gt + 1..]
                ));
            }
        }
        from = start + 1;
    }
    Ok(xml)
}

/// `findMaxRId` (src/docx/rezip/parts.ts:133).
pub(crate) fn find_max_r_id(rels_xml: &str) -> Result<f64> {
    let mut max_id = 0.0;
    let pin = "Id=\"rId";
    let bytes = rels_xml.as_bytes();
    let mut from = 0;
    while let Some(found) = rels_xml[from..].find(pin) {
        let digits = from + found + pin.len();
        let mut end = digits;
        while end < bytes.len() && bytes[end].is_ascii_digit() {
            end += 1;
        }
        if end > digits && end < bytes.len() && bytes[end] == b'"' {
            let id = crate::js::parse_int(&rels_xml[digits..end]);
            if id > max_id {
                max_id = id;
            }
            from = end + 1;
        } else {
            from = from + found + 1;
        }
    }
    Ok(max_id)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn rels_helpers() {
        let pkg = Js::new_map();
        pkg.map_set("r", Js::bytes(b"<x/><Relationships a=\"1\"/>".to_vec()));
        assert_eq!(
            read_rels_or_stub(&pkg, "r").unwrap(),
            "<x/><Relationships a=\"1\"></Relationships>"
        );
        assert_eq!(
            find_max_r_id("Id=\"rId3\" Id=\"rId12\" Id=\"rIdx\"").unwrap(),
            12.0
        );
    }
}
