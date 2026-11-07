//! Port of `packages/docx/src/docx/rustParseFacade.ts`.
#![allow(dead_code)]

use std::collections::HashSet;
use std::rc::Rc;

use crate::js::{self, Js, Result, SaveError};
use crate::tiff_media::maybe_transcode_tiff_media;
use crate::{obj, throw};

const FORBIDDEN_KEYS: [&str; 3] = ["__proto__", "constructor", "prototype"];

fn quote(s: &str) -> String {
    Js::from(s).stringify().unwrap_or_default()
}

fn strings(names: &[&str]) -> Vec<String> {
    names.iter().map(|s| (*s).to_owned()).collect()
}

/// `parseDocumentWithRust` (src/docx/rustParseFacade.ts:36): `parse_docx_s9` natively.
pub(crate) fn parse_document_with_rust(data: &Rc<Vec<u8>>, options: &Js) -> Result<Js> {
    let options_json = options.stringify().unwrap_or_default();
    let json = docx_parse::parse_docx_s9_json(data, &options_json).map_err(SaveError::new)?;
    decode_s9_envelope(&json, data)
}

/// `decodeS9Envelope` (src/docx/rustParseFacade.ts:43).
pub(crate) fn decode_s9_envelope(json: &str, original_buffer: &Rc<Vec<u8>>) -> Result<Js> {
    decode_s9_envelope_value(&Js::parse(json)?, original_buffer)
}

/// `decodeS9EnvelopeValue` (src/docx/rustParseFacade.ts:47).
pub(crate) fn decode_s9_envelope_value(value: &Js, original_buffer: &Rc<Vec<u8>>) -> Result<Js> {
    assert_safe_tree(value, "wire", &Js::Undefined)?;
    let envelope = object_at(value, "wire")?;
    exact_keys(
        &envelope,
        &strings(&[
            "wireVersion",
            "document",
            "embeddedFontParts",
            "fontTableRelationshipsXml",
            "canonicalBase64",
            "canonicalSha256",
        ]),
        "wire",
        &strings(&[
            "fontTableRelationshipsXml",
            "canonicalBase64",
            "canonicalSha256",
        ]),
    )?;
    if !envelope.get("wireVersion").strict_eq(&Js::Num(1.0)) {
        throw!("wire.wireVersion must be 1");
    }
    let document = decode_s9_document(&envelope.get("document"), original_buffer)?;
    let embedded_fonts = decode_binary_parts(&envelope.get("embeddedFontParts"))?;
    let base64 = envelope.get("canonicalBase64");
    let sha_value = envelope.get("canonicalSha256");
    let has_canonical = !base64.is_undefined() || !sha_value.is_undefined();
    if has_canonical && (base64.is_undefined() || sha_value.is_undefined()) {
        throw!("wire canonical fields must be present together");
    }
    let sha = if sha_value.is_undefined() {
        None
    } else {
        Some(string_at(&sha_value, "wire.canonicalSha256")?)
    };
    if let Some(sha) = &sha
        && !(sha.len() == 64
            && sha
                .bytes()
                .all(|b| b.is_ascii_digit() || (b'a'..=b'f').contains(&b)))
        {
            throw!("wire.canonicalSha256 is invalid");
        }
    let result =
        obj! { "wireVersion" => 1.0, "document" => document, "embeddedFonts" => embedded_fonts };
    let font_rels = envelope.get("fontTableRelationshipsXml");
    if !font_rels.is_undefined() {
        result.set(
            "fontTableRelationshipsXml",
            string_at(&font_rels, "wire.fontTableRelationshipsXml")?,
        );
    }
    if !base64.is_undefined() {
        let bytes = decode_base64(&string_at(&base64, "wire.canonicalBase64")?)?;
        result.set("rustCanonicalBytes", Js::bytes(bytes));
        result.set("rustCanonicalSha256", Js::from(sha));
    }
    Ok(result)
}

/// `decodeS9Document` (src/docx/rustParseFacade.ts:104).
pub(crate) fn decode_s9_document(value: &Js, original_buffer: &Rc<Vec<u8>>) -> Result<Js> {
    let wire_document = object_at(value, "wire.document")?;
    exact_keys(
        &wire_document,
        &strings(&["package", "templateVariables", "warnings"]),
        "wire.document",
        &strings(&["templateVariables", "warnings"]),
    )?;
    let pkg = decode_s9_package(&wire_document.get("package"))?;
    let document =
        obj! { "package" => pkg, "originalBuffer" => Js::Bytes(original_buffer.clone()) };
    let warnings = wire_document.get("warnings");
    if !warnings.is_undefined() {
        let list = string_array_at(&warnings, "wire.document.warnings")?;
        document.set(
            "warnings",
            Js::new_arr(list.into_iter().map(Js::from).collect()),
        );
    }
    Ok(document)
}

/// `decodeS9Package` (src/docx/rustParseFacade.ts:118).
pub(crate) fn decode_s9_package(value: &Js) -> Result<Js> {
    let p = "wire.document.package";
    let wire = object_at(value, p)?;
    let optional = strings(&[
        "styles",
        "headerEntries",
        "footerEntries",
        "footnotes",
        "endnotes",
        "footnoteSeparators",
        "endnoteSeparators",
    ]);
    exact_keys(
        &wire,
        &strings(&[
            "document",
            "styles",
            "theme",
            "numbering",
            "settings",
            "fontTable",
            "headerEntries",
            "footerEntries",
            "footnotes",
            "endnotes",
            "footnoteSeparators",
            "endnoteSeparators",
            "relationshipEntries",
            "mediaEntries",
            "chartEntries",
        ]),
        p,
        &optional,
    )?;
    let document_body = decode_s9_document_body(&wire.get("document"))?;
    let pkg = obj! {
        "document" => document_body,
        "theme" => object_at(&wire.get("theme"), "wire.document.package.theme")?,
        "numbering" => object_at(&wire.get("numbering"), "wire.document.package.numbering")?,
        "settings" => object_at(&wire.get("settings"), "wire.document.package.settings")?,
        "fontTable" => object_at(&wire.get("fontTable"), "wire.document.package.fontTable")?,
        "relationships" => decode_relationship_entries(&wire.get("relationshipEntries"), "wire.document.package.relationshipEntries")?,
        "media" => decode_media_entries(&wire.get("mediaEntries"))?,
        "charts" => decode_map_entries(&wire.get("chartEntries"), "wire.document.package.chartEntries", &|entry, path, _| object_at(entry, path))?,
    };
    if !wire.get("styles").is_undefined() {
        pkg.set(
            "styles",
            object_at(&wire.get("styles"), "wire.document.package.styles")?,
        );
    }
    for (wire_name, public_name) in [("headerEntries", "headers"), ("footerEntries", "footers")] {
        let entries = wire.get(wire_name);
        if !entries.is_undefined() {
            let path = format!("wire.document.package.{wire_name}");
            pkg.set(
                public_name,
                decode_map_entries(&entries, &path, &|entry, path, _| object_at(entry, path))?,
            );
        }
    }
    for name in [
        "footnotes",
        "endnotes",
        "footnoteSeparators",
        "endnoteSeparators",
    ] {
        let notes = wire.get(name);
        if !notes.is_undefined() {
            pkg.set(
                name,
                decode_object_array(&notes, &format!("wire.document.package.{name}"))?,
            );
        }
    }
    Ok(pkg)
}

/// `decodeS9DocumentBody` (src/docx/rustParseFacade.ts:224).
pub(crate) fn decode_s9_document_body(value: &Js) -> Result<Js> {
    let path = "wire.document.package.document";
    let body = object_at(value, path)?;
    exact_keys(
        &body,
        &strings(&[
            "content",
            "sections",
            "finalSectionProperties",
            "customRootBindings",
            "comments",
        ]),
        path,
        &strings(&[
            "sections",
            "finalSectionProperties",
            "customRootBindings",
            "comments",
        ]),
    )?;
    let content = body.get("content");
    if !content.is_arr() {
        throw!("{path}.content must be an array");
    }
    for (index, block) in content.items().iter().enumerate() {
        object_at(block, &format!("{path}.content[{index}]"))?;
    }
    let document_body = obj! { "content" => content.clone() };
    let sections = body.get("sections");
    if !sections.is_undefined() {
        if !sections.is_arr() {
            throw!("{path}.sections must be an array");
        }
        let mut out = Vec::new();
        for (index, value) in sections.items().iter().enumerate() {
            let section_path = format!("{path}.sections[{index}]");
            let section = object_at(value, &section_path)?;
            exact_keys(
                &section,
                &strings(&["id", "properties", "contentStart", "contentEnd"]),
                &section_path,
                &strings(&["id"]),
            )?;
            let start = integer_at(
                &section.get("contentStart"),
                &format!("{section_path}.contentStart"),
            )?;
            let end = integer_at(
                &section.get("contentEnd"),
                &format!("{section_path}.contentEnd"),
            )?;
            if start > end || end > content.len() as f64 {
                throw!("{section_path} has an invalid content range");
            }
            let decoded = Js::new_obj();
            if !section.get("id").is_undefined() {
                decoded.set(
                    "id",
                    string_at(&section.get("id"), &format!("{section_path}.id"))?,
                );
            }
            decoded.set(
                "properties",
                object_at(
                    &section.get("properties"),
                    &format!("{section_path}.properties"),
                )?,
            );
            let items = content.items();
            decoded.set(
                "content",
                Js::new_arr(items[start as usize..end as usize].to_vec()),
            );
            out.push(decoded);
        }
        document_body.set("sections", Js::new_arr(out));
    }
    let final_props = body.get("finalSectionProperties");
    if !final_props.is_undefined() {
        document_body.set(
            "finalSectionProperties",
            object_at(&final_props, &format!("{path}.finalSectionProperties"))?,
        );
    }
    let bindings = body.get("customRootBindings");
    if !bindings.is_undefined() {
        document_body.set(
            "customRootBindings",
            decode_object_array(&bindings, &format!("{path}.customRootBindings"))?,
        );
    }
    let comments = body.get("comments");
    if !comments.is_undefined() {
        document_body.set(
            "comments",
            decode_object_array(&comments, &format!("{path}.comments"))?,
        );
    }
    Ok(document_body)
}

/// `decodeRelationshipEntries` (src/docx/rustParseFacade.ts:279).
pub(crate) fn decode_relationship_entries(value: &Js, path: &str) -> Result<Js> {
    decode_map_entries(value, path, &|entry, entry_path, key| {
        decode_relationship(entry, entry_path, key)
    })
}

/// `decodeMediaEntries` (src/docx/rustParseFacade.ts:285).
pub(crate) fn decode_media_entries(value: &Js) -> Result<Js> {
    let path = "wire.document.package.mediaEntries";
    if !value.is_arr() {
        throw!("{path} must be an array");
    }
    let media = Js::new_map();
    let mut by_source_path: std::collections::HashMap<String, Js> =
        std::collections::HashMap::new();
    for (index, entry) in value.items().iter().enumerate() {
        let entry_path = format!("{path}[{index}]");
        if !entry.is_arr() || entry.len() != 2 {
            throw!("{entry_path} must be a key/value pair");
        }
        let alias = string_at(&entry.at(0), &format!("{entry_path}[0]"))?;
        if media.map_has(&alias) {
            throw!("{path} contains duplicate key {}", quote(&alias));
        }
        let file_path = format!("{entry_path}[1]");
        let wire_file = object_at(&entry.at(1), &file_path)?;
        exact_keys(
            &wire_file,
            &strings(&["path", "filename", "mimeType", "base64", "dataUrl"]),
            &file_path,
            &strings(&["filename"]),
        )?;
        let source_path = string_at(&wire_file.get("path"), &format!("{file_path}.path"))?;
        let file = match by_source_path.get(&source_path) {
            Some(file) => file.clone(),
            None => {
                let bytes = Rc::new(decode_base64(&string_at(
                    &wire_file.get("base64"),
                    &format!("{file_path}.base64"),
                )?)?);
                let mime = string_at(&wire_file.get("mimeType"), &format!("{file_path}.mimeType"))?;
                let data_url =
                    string_at(&wire_file.get("dataUrl"), &format!("{file_path}.dataUrl"))?;
                let transcoded =
                    maybe_transcode_tiff_media(&bytes, &mime, &data_url, &decode_tiff_image)?;
                let created = obj! { "path" => source_path.clone() };
                if !wire_file.get("filename").is_undefined() {
                    created.set(
                        "filename",
                        string_at(&wire_file.get("filename"), &format!("{file_path}.filename"))?,
                    );
                }
                created.set("mimeType", transcoded.get("mimeType"));
                let data = match transcoded.get("bytes") {
                    Js::Bytes(b) => exact_array_buffer(&b)?,
                    _ => Vec::new(),
                };
                created.set("data", Js::bytes(data));
                created.set("dataUrl", transcoded.get("dataUrl"));
                by_source_path.insert(source_path, created.clone());
                created
            }
        };
        media.map_set(&alias, file);
    }
    Ok(media)
}

/// `decodeTiffImage` of the parse wasm. This crate builds `docx-parse` without its `tiff`
/// feature, so no transcoder is available: it fails, and the media keeps its source bytes.
fn decode_tiff_image(bytes: &Rc<Vec<u8>>) -> Result<Vec<u8>> {
    docx_parse::decode_tiff_png_bytes(bytes).map_err(SaveError::new)
}

/// `decodeBinaryParts` (src/docx/rustParseFacade.ts:331).
pub(crate) fn decode_binary_parts(value: &Js) -> Result<Js> {
    let path = "wire.embeddedFontParts";
    if !value.is_arr() {
        throw!("{path} must be an array");
    }
    let parts = Js::new_map();
    for (index, entry) in value.items().iter().enumerate() {
        let entry_path = format!("{path}[{index}]");
        let part = object_at(entry, &entry_path)?;
        exact_keys(&part, &strings(&["path", "base64"]), &entry_path, &[])?;
        let part_path = string_at(&part.get("path"), &format!("{entry_path}.path"))?;
        if parts.map_has(&part_path) {
            throw!("{path} contains duplicate key {}", quote(&part_path));
        }
        let bytes = Rc::new(decode_base64(&string_at(
            &part.get("base64"),
            &format!("{entry_path}.base64"),
        )?)?);
        parts.map_set(&part_path, Js::bytes(exact_array_buffer(&bytes)?));
    }
    Ok(parts)
}

/// `decodeMapEntries` (src/docx/rustParseFacade.ts:351).
pub(crate) fn decode_map_entries(
    value: &Js,
    path: &str,
    decode_value: &dyn Fn(&Js, &str, &str) -> Result<Js>,
) -> Result<Js> {
    if !value.is_arr() {
        throw!("{path} must be an array");
    }
    let map = Js::new_map();
    for (index, entry) in value.items().iter().enumerate() {
        let entry_path = format!("{path}[{index}]");
        if !entry.is_arr() || entry.len() != 2 {
            throw!("{entry_path} must be a key/value pair");
        }
        let key = string_at(&entry.at(0), &format!("{entry_path}[0]"))?;
        if map.map_has(&key) {
            throw!("{path} contains duplicate key {}", quote(&key));
        }
        let decoded = decode_value(&entry.at(1), &format!("{entry_path}[1]"), &key)?;
        map.map_set(&key, decoded);
    }
    Ok(map)
}

/// `stringArrayAt` (src/docx/rustParseFacade.ts:370).
pub(crate) fn string_array_at(value: &Js, path: &str) -> Result<Vec<String>> {
    if !value.is_arr() {
        throw!("{path} must be an array");
    }
    value
        .items()
        .iter()
        .enumerate()
        .map(|(index, entry)| string_at(entry, &format!("{path}[{index}]")))
        .collect()
}

/// `decodeObjectArray` (src/docx/rustParseFacade.ts:375).
pub(crate) fn decode_object_array(value: &Js, path: &str) -> Result<Js> {
    if !value.is_arr() {
        throw!("{path} must be an array");
    }
    for (index, entry) in value.items().iter().enumerate() {
        object_at(entry, &format!("{path}[{index}]"))?;
    }
    Ok(value.clone())
}

/// `exactArrayBuffer` (src/docx/rustParseFacade.ts:381).
pub(crate) fn exact_array_buffer(bytes: &Rc<Vec<u8>>) -> Result<Vec<u8>> {
    Ok((**bytes).clone())
}

/// `integerAt` (src/docx/rustParseFacade.ts:385).
pub(crate) fn integer_at(value: &Js, path: &str) -> Result<f64> {
    let number = number_at(value, path)?;
    if !(js::is_integer(&Js::Num(number)) && number.abs() <= 9007199254740991.0) || number < 0.0 {
        throw!("{path} must be a non-negative safe integer");
    }
    Ok(number)
}

/// `parseRelationshipsXmlWithRust` (src/docx/rustParseFacade.ts:394).
pub(crate) fn parse_relationships_xml_with_rust(xml: &str, part_path: &str) -> Result<Js> {
    let json = docx_parse::parse_relationships_xml_json(xml.as_bytes(), part_path)
        .map_err(SaveError::new)?;
    let value = Js::parse(&json)?;
    assert_safe_tree(&value, "relationshipPart", &Js::Undefined)?;
    let part = object_at(&value, "relationshipPart")?;
    exact_keys(
        &part,
        &strings(&[
            "path",
            "relationships",
            "canonicalBase64",
            "canonicalSha256",
        ]),
        "relationshipPart",
        &[],
    )?;
    string_at(&part.get("path"), "relationshipPart.path")?;
    decode_relationship_entries(&part.get("relationships"), "relationshipPart.relationships")
}

/// `decodeRelationship` (src/docx/rustParseFacade.ts:412).
pub(crate) fn decode_relationship(value: &Js, path: &str, map_id: &str) -> Result<Js> {
    let relationship = object_at(value, path)?;
    exact_keys(
        &relationship,
        &strings(&["id", "type", "target", "targetMode"]),
        path,
        &strings(&["targetMode"]),
    )?;
    let id = string_at(&relationship.get("id"), &format!("{path}.id"))?;
    if id != map_id {
        throw!("{path}.id does not match its map key");
    }
    let decoded = obj! {
        "id" => id,
        "type" => string_at(&relationship.get("type"), &format!("{path}.type"))?,
        "target" => string_at(&relationship.get("target"), &format!("{path}.target"))?,
    };
    let mode = relationship.get("targetMode");
    if !mode.is_undefined() {
        if !matches!(mode.as_str(), Some("External" | "Internal")) {
            throw!("{path}.targetMode is invalid");
        }
        decoded.set("targetMode", mode);
    }
    Ok(decoded)
}

/// `assertSafeTree` (src/docx/rustParseFacade.ts:432). `active` is the caller's `WeakSet`;
/// this port always starts from an empty one.
pub(crate) fn assert_safe_tree(value: &Js, path: &str, _active: &Js) -> Result<()> {
    fn current(path: &str, segments: &[String]) -> String {
        if segments.is_empty() {
            path.to_owned()
        } else {
            format!("{path}.{}", segments.join("."))
        }
    }
    fn visit(
        value: &Js,
        path: &str,
        segments: &mut Vec<String>,
        active: &mut HashSet<usize>,
    ) -> Result<()> {
        if !matches!(value, Js::Obj(_) | Js::Arr(_)) {
            return Ok(());
        }
        let p = value.ptr().unwrap();
        if active.contains(&p) {
            throw!("{} contains a cycle", current(path, segments));
        }
        active.insert(p);
        let result = (|| {
            for (key, child) in value.entries() {
                if FORBIDDEN_KEYS.contains(&key.as_str()) {
                    throw!("{} contains forbidden key {key}", current(path, segments));
                }
                segments.push(key);
                visit(&child, path, segments, active)?;
                segments.pop();
            }
            Ok(())
        })();
        active.remove(&p);
        result
    }
    visit(value, path, &mut Vec::new(), &mut HashSet::new())
}

/// `objectAt` (src/docx/rustParseFacade.ts:453).
pub(crate) fn object_at(value: &Js, path: &str) -> Result<Js> {
    if !value.is_object_like() || value.is_arr() {
        throw!("{path} must be an object");
    }
    Ok(value.clone())
}

/// `exactKeys` (src/docx/rustParseFacade.ts:460).
pub(crate) fn exact_keys(
    object: &Js,
    allowed: &[String],
    path: &str,
    optional: &[String],
) -> Result<()> {
    for key in object.keys() {
        if !allowed.contains(&key) {
            throw!("{path} contains unknown key {key}");
        }
    }
    for key in allowed {
        if !optional.contains(key) && !object.has(key) {
            throw!("{path} is missing {key}");
        }
    }
    Ok(())
}

/// `stringAt` (src/docx/rustParseFacade.ts:478).
pub(crate) fn string_at(value: &Js, path: &str) -> Result<String> {
    match value.as_str() {
        Some(s) => Ok(s.to_owned()),
        None => throw!("{path} must be a string"),
    }
}

/// `numberAt` (src/docx/rustParseFacade.ts:483).
pub(crate) fn number_at(value: &Js, path: &str) -> Result<f64> {
    match value {
        Js::Num(n) if n.is_finite() => Ok(*n),
        _ => throw!("{path} must be a finite number"),
    }
}

/// `decodeBase64` (src/docx/rustParseFacade.ts:490): `atob`, forgiving as the HTML spec is.
pub(crate) fn decode_base64(value: &str) -> Result<Vec<u8>> {
    let invalid = || SaveError::new("wire canonicalBase64 is invalid");
    let mut data: Vec<u8> = value
        .bytes()
        .filter(|b| !matches!(b, b'\t' | b'\n' | 0x0c | b'\r' | b' '))
        .collect();
    if data.len().is_multiple_of(4) {
        if data.ends_with(b"==") {
            data.truncate(data.len() - 2);
        } else if data.ends_with(b"=") {
            data.truncate(data.len() - 1);
        }
    }
    if data.len() % 4 == 1 {
        return Err(invalid());
    }
    let mut out = Vec::with_capacity(data.len() * 3 / 4);
    let mut buffer: u32 = 0;
    let mut bits = 0;
    for b in data {
        let v = match b {
            b'A'..=b'Z' => b - b'A',
            b'a'..=b'z' => b - b'a' + 26,
            b'0'..=b'9' => b - b'0' + 52,
            b'+' => 62,
            b'/' => 63,
            _ => return Err(invalid()),
        };
        buffer = (buffer << 6) | u32::from(v);
        bits += 6;
        if bits >= 8 {
            bits -= 8;
            out.push((buffer >> bits) as u8);
            buffer &= (1 << bits) - 1;
        }
    }
    Ok(out)
}

/// `decodeDocxHost` (src/yrs/yrsSessionFacade.ts:156): the host metadata `open_docx` returned
/// for the package `source`, as `{ document, referencedFonts, unusedScriptFonts, wholeBody?,
/// embeddedFonts, fontTableRelationshipsXml? }`.
pub(crate) fn decode_docx_host(json: &str, source: &Rc<Vec<u8>>) -> Result<Js> {
    let value = Js::parse(json)?;
    if !value.is_obj() {
        throw!("DOCX host metadata must be an object");
    }
    let referenced = value.get("referencedFonts");
    if !referenced.is_arr() || !referenced.items().iter().all(Js::is_str) {
        throw!("DOCX host referencedFonts must be a string array");
    }
    let unused = value
        .get("unusedScriptFonts")
        .nullish_or(Js::new_arr(Vec::new()));
    if !unused.is_arr() || !unused.items().iter().all(Js::is_str) {
        throw!("DOCX host unusedScriptFonts must be a string array");
    }
    let result = decode_s9_envelope_value(&value.get("envelope"), source)?;
    let host = obj! {
        "document" => result.get("document"),
        "referencedFonts" => referenced,
        "unusedScriptFonts" => unused,
    };
    if value.get("wholeBody").is_true() {
        host.set("wholeBody", true);
    }
    host.set("embeddedFonts", result.get("embeddedFonts"));
    if !result.get("fontTableRelationshipsXml").is_undefined() {
        host.set(
            "fontTableRelationshipsXml",
            result.get("fontTableRelationshipsXml"),
        );
    }
    Ok(host)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn base64_like_atob() {
        assert_eq!(decode_base64("aGk=").unwrap(), b"hi");
        assert_eq!(decode_base64("aGk").unwrap(), b"hi");
        assert_eq!(decode_base64(" aG\nk= ").unwrap(), b"hi");
        assert!(decode_base64("a").is_err());
        assert!(decode_base64("a*==").is_err());
    }

    #[test]
    fn decodes_a_tiny_envelope() {
        // Expectations from running decodeS9EnvelopeValue with bun on the same envelope.
        let json = r#"{"wireVersion":1,"document":{"package":{"document":{"content":[{"type":"paragraph","content":[]}],"sections":[{"properties":{"a":1},"contentStart":0,"contentEnd":1}]},"theme":{},"numbering":{},"settings":{},"fontTable":{},"relationshipEntries":[["rId1",{"id":"rId1","type":"t","target":"x.xml","targetMode":"Internal"}]],"mediaEntries":[["a",{"path":"word/media/a.png","mimeType":"image/png","base64":"aGk=","dataUrl":"data:x"}],["b",{"path":"word/media/a.png","mimeType":"image/png","base64":"aGk=","dataUrl":"data:x"}]],"chartEntries":[]},"warnings":["w"]},"embeddedFontParts":[{"path":"f","base64":"AA=="}]}"#;
        let source = Rc::new(vec![1u8, 2]);
        let out = decode_s9_envelope(json, &source).unwrap();
        let doc = out.get("document");
        assert_eq!(doc.keys(), vec!["package", "originalBuffer", "warnings"]);
        let pkg = doc.get("package");
        assert_eq!(
            pkg.keys(),
            vec![
                "document",
                "theme",
                "numbering",
                "settings",
                "fontTable",
                "relationships",
                "media",
                "charts"
            ]
        );
        assert_eq!(
            pkg.get("relationships")
                .map_get("rId1")
                .stringify()
                .unwrap(),
            r#"{"id":"rId1","type":"t","target":"x.xml","targetMode":"Internal"}"#
        );
        let media = pkg.get("media");
        assert!(media.map_get("a").strict_eq(&media.map_get("b")));
        assert_eq!(
            media.map_get("a").keys(),
            vec!["path", "mimeType", "data", "dataUrl"]
        );
        let body = pkg.get("document");
        assert!(
            body.get("sections")
                .at(0)
                .get("content")
                .at(0)
                .strict_eq(&body.get("content").at(0))
        );
        assert_eq!(out.keys(), vec!["wireVersion", "document", "embeddedFonts"]);
        assert_eq!(
            out.get("embeddedFonts")
                .map_get("f")
                .as_bytes()
                .unwrap()
                .as_slice(),
            &[0u8]
        );
        let bad = json.replacen("\"wireVersion\":1", "\"wireVersion\":1,\"__proto__\":{}", 1);
        assert_eq!(
            decode_s9_envelope(&bad, &source).unwrap_err().message,
            "wire contains forbidden key __proto__"
        );
    }
}
