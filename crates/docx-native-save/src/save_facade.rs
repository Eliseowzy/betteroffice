//! Port of `packages/docx/src/docx/rustSaveFacade.ts`.
#![allow(dead_code)]
//!
//! The native save threads two things the TS facade takes from its environment through
//! [`SaveContext`]: the determinism clock (`new Date().toISOString()`, fixed by
//! `SaveOptions::now`) and the capture of the request JSON handed to the writer.

use std::cell::RefCell;
use std::collections::HashSet;
use std::rc::Rc;
use std::time::{SystemTime, UNIX_EPOCH};

use indexmap::IndexMap;
use sha2::{Digest, Sha256};

use crate::header_footer_alias_projection::is_written_by_canonical;
use crate::js::{Js, Result, SaveError};
use crate::rezip_parts::{collect_parts, header_footer_filename, part_text};
use crate::session_revision_ids::with_session_revision_ids;
use crate::tracked_control_content::visit_tracked_control_content;
use crate::{obj, throw};

/// `PARAGRAPH_PATCH_WRITER_VERSION` (src/wasm/parse.ts:50).
const PARAGRAPH_PATCH_WRITER_VERSION: u32 = 1;

/// What a native save sets around `saveYrsDocx` and reads back after it.
#[derive(Default)]
pub(crate) struct SaveContext {
    /// The clock `writeDocumentWithRust` stamps; the current time when `None`.
    pub now: Option<String>,
    /// Keep each request JSON handed to the writer in `wire`.
    pub capture_wire: bool,
    /// The last request JSON handed to the writer.
    pub wire: Option<String>,
    /// The warnings the last write returned (document warnings, then the writer's).
    pub warnings: Vec<String>,
    /// Leave `session.recordSavedParagraphIds` uncalled (`SaveOptions::record_saved_paragraph_ids` false).
    pub skip_record_saved_paragraph_ids: bool,
}

thread_local! {
    static CONTEXT: RefCell<Option<SaveContext>> = const { RefCell::new(None) };
}

/// Installs `context` for the saves on this thread until [`take_save_context`].
pub(crate) fn set_save_context(context: SaveContext) {
    CONTEXT.with(|c| *c.borrow_mut() = Some(context));
}

/// Removes the installed context, with what the saves recorded in it.
pub(crate) fn take_save_context() -> Option<SaveContext> {
    CONTEXT.with(|c| c.borrow_mut().take())
}

/// Whether the installed context asks the save not to record the paragraph IDs it wrote.
pub(crate) fn skips_recording_saved_paragraph_ids() -> bool {
    CONTEXT.with(|c| {
        c.borrow()
            .as_ref()
            .is_some_and(|c| c.skip_record_saved_paragraph_ids)
    })
}

/// `new Date().toISOString()` for `millis` since the epoch.
pub(crate) fn iso_string(millis: i64) -> String {
    let days = millis.div_euclid(86_400_000);
    let ms_of_day = millis.rem_euclid(86_400_000);
    // civil_from_days (Howard Hinnant).
    let z = days + 719_468;
    let era = z.div_euclid(146_097);
    let doe = z.rem_euclid(146_097);
    let yoe = (doe - doe / 1460 + doe / 36_524 - doe / 146_096) / 365;
    let mut year = yoe + era * 400;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let day = doy - (153 * mp + 2) / 5 + 1;
    let month = if mp < 10 { mp + 3 } else { mp - 9 };
    if month <= 2 {
        year += 1;
    }
    let (h, m, s, ms) = (
        ms_of_day / 3_600_000,
        ms_of_day / 60_000 % 60,
        ms_of_day / 1000 % 60,
        ms_of_day % 1000,
    );
    let year = if (0..=9999).contains(&year) {
        format!("{year:04}")
    } else if year < 0 {
        format!("-{:06}", -year)
    } else {
        format!("+{year:06}")
    };
    format!("{year}-{month:02}-{day:02}T{h:02}:{m:02}:{s:02}.{ms:03}Z")
}

fn now_iso() -> String {
    if let Some(now) = CONTEXT.with(|c| c.borrow().as_ref().and_then(|c| c.now.clone())) {
        return now;
    }
    let millis = match SystemTime::now().duration_since(UNIX_EPOCH) {
        Ok(d) => d.as_millis() as i64,
        Err(e) => -(e.duration().as_millis() as i64),
    };
    iso_string(millis)
}

/// `headerFooterSaveEntries` (src/docx/rustSaveFacade.ts:83).
pub(crate) fn header_footer_save_entries(entries: &Js) -> Result<Js> {
    if !entries.truthy() {
        return Ok(Js::new_arr(Vec::new()));
    }
    let mut out = Vec::new();
    for (key, part) in entries.map_entries() {
        if !is_written_by_canonical(&part, entries)? {
            out.push(Js::new_arr(vec![Js::from(key), part]));
        }
    }
    Ok(Js::new_arr(out))
}

/// `sha256Hex` (src/docx/rustSaveFacade.ts:91).
pub(crate) fn sha256_hex(bytes: &Rc<Vec<u8>>) -> Result<String> {
    Ok(hex(&Sha256::digest(bytes.as_slice())))
}

pub(crate) fn hex(bytes: &[u8]) -> String {
    bytes.iter().map(|b| format!("{b:02x}")).collect()
}

/// `writeDocumentWithRust` (src/docx/rustSaveFacade.ts:97): `{ buffer, determinism, warnings }`.
#[allow(clippy::too_many_arguments)]
pub(crate) fn write_document_with_rust(
    document: &Js,
    original_buffer: &Rc<Vec<u8>>,
    options: &Js,
    selective: &Js,
    determinism: &Js,
    paragraph_ids: &Js,
    skip_mutations: bool,
    _yield_save: Option<&dyn Fn(&Js) -> Result<Js>>,
) -> Result<Js> {
    for part in paragraph_ids.get("splicedParts").items() {
        if !part.get("changed").is_empty()
            && docx_parse::paragraph_patch_writer_version_native() != PARAGRAPH_PATCH_WRITER_VERSION
        {
            throw!(
                "DOCX_MINIMAL_PATCH [INVALID_PLAN] {} p#?: paragraphPatchWriterVersion must report PARAGRAPH_PATCH_WRITER_VERSION=1",
                crate::js::to_js_string(&part.get("part"))
            );
        }
    }
    let fixed = if determinism.is_nullish() {
        obj! { "seed" => sha256_hex(original_buffer)?, "now" => now_iso() }
    } else {
        determinism.clone()
    };
    validate_determinism(&fixed)?;

    let pkg = document.get("package");
    let request = obj! {
        "determinism" => fixed.clone(),
        "document" => pkg.get("document"),
        "headerEntries" => header_footer_save_entries(&pkg.get("headers"))?,
        "footerEntries" => header_footer_save_entries(&pkg.get("footers"))?,
        "footnotes" => pkg.get("footnotes").nullish_or(Js::new_arr(Vec::new())),
        "endnotes" => pkg.get("endnotes").nullish_or(Js::new_arr(Vec::new())),
        "footnoteSeparators" => pkg.get("footnoteSeparators").nullish_or(Js::new_arr(Vec::new())),
        "endnoteSeparators" => pkg.get("endnoteSeparators").nullish_or(Js::new_arr(Vec::new())),
        "relationshipEntries" => Js::new_arr(
            pkg.get("relationships").map_entries().into_iter().map(|(k, v)| Js::new_arr(vec![Js::from(k), v])).collect()
        ),
    };
    if !pkg.get("numbering").is_undefined() {
        request.set("numbering", pkg.get("numbering"));
    }
    let write_options = obj! { "updateModifiedDate" => options.get("updateModifiedDate").nullish_or(Js::Bool(true)) };
    if !options.get("modifiedBy").is_undefined() {
        write_options.set("modifiedBy", options.get("modifiedBy"));
    }
    request.set("options", write_options);
    if !selective.is_undefined() {
        let changed = selective.get("changedParaIds");
        let ids = if changed.is_map() {
            Js::new_arr(changed.map_keys().into_iter().map(Js::from).collect())
        } else {
            changed.copy_arr()
        };
        let sel = obj! { "changedParaIds" => ids };
        if !selective.get("sourceParagraphs").is_undefined() {
            sel.set("sourceParagraphs", selective.get("sourceParagraphs"));
        }
        request.set("selective", sel);
    }
    if !paragraph_ids.is_undefined() {
        request.set("paragraphIds", paragraph_ids.clone());
    }
    assert_safe_save_tree(&request, "save", &Js::Undefined)?;
    let request_for_write = request.clone();
    let wire =
        with_session_revision_ids(&request, &|| Ok(Js::from(request_for_write.stringify())))?;
    let wire = wire.as_str().unwrap_or_default().to_owned();
    CONTEXT.with(|c| {
        if let Some(context) = c.borrow_mut().as_mut() {
            if context.capture_wire {
                context.wire = Some(wire.clone());
            }
        }
    });
    let framed = docx_parse::write_docx_s13_with_warnings_json(&wire, original_buffer)
        .map_err(SaveError::new)?;
    if framed.len() < 8 {
        throw!("Rust save response is truncated");
    }
    let warnings_length = u64::from_le_bytes(framed[..8].try_into().unwrap());
    if warnings_length > 9007199254740991 || 8 + warnings_length > framed.len() as u64 {
        throw!("Rust save response carries a corrupt warnings frame");
    }
    let warnings_end = 8 + warnings_length as usize;
    let warnings = Js::parse(&String::from_utf8_lossy(&framed[8..warnings_end]))?;
    if !warnings.is_arr() || !warnings.items().iter().all(Js::is_str) {
        throw!("Rust save response carries malformed warnings");
    }
    let buffer = Rc::new(framed[warnings_end..].to_vec());
    if selective.is_undefined() && !skip_mutations {
        apply_rust_save_mutations(document, original_buffer, &buffer)?;
    }
    let mut all: Vec<Js> = document.get("warnings").items();
    all.extend(warnings.items());
    CONTEXT.with(|c| {
        if let Some(context) = c.borrow_mut().as_mut() {
            context.warnings = all.iter().map(|w| w.str_or_empty().to_owned()).collect();
        }
    });
    Ok(
        obj! { "buffer" => Js::Bytes(buffer), "determinism" => fixed, "warnings" => Js::new_arr(all) },
    )
}

/// `unzipContainer`: the package's parts by path.
pub(crate) fn unzip(bytes: &[u8]) -> Result<IndexMap<String, Rc<Vec<u8>>>> {
    Ok(ooxml_opc::unzip_parts(bytes)
        .map_err(SaveError::new)?
        .into_iter()
        .map(|(k, v)| (k, Rc::new(v)))
        .collect())
}

/// `rezipContainer`.
pub(crate) fn rezip(parts: &IndexMap<String, Rc<Vec<u8>>>) -> Result<Vec<u8>> {
    let entries: Vec<(String, &[u8])> = parts
        .iter()
        .map(|(k, v)| (k.clone(), v.as_slice()))
        .collect();
    ooxml_opc::rezip_parts_borrowed(&entries).map_err(SaveError::new)
}

/// `applyRustSaveMutations` (src/docx/rustSaveFacade.ts:169).
pub(crate) fn apply_rust_save_mutations(
    document: &Js,
    original_buffer: &Rc<Vec<u8>>,
    saved_buffer: &Rc<Vec<u8>>,
) -> Result<()> {
    let original_parts = unzip(original_buffer)?;
    let saved_parts = unzip(saved_buffer)?;
    let attr = |tag: &str, name: &str| relationship_attribute(tag, name).unwrap_or(None);

    for part in collect_parts(document)?.items() {
        let rels_path = part.get("relsPath").str_or_empty().to_owned();
        let original_relationships = part_text(original_parts.get(&rels_path))?.unwrap_or_default();
        let saved_relationships = part_text(saved_parts.get(&rels_path))?.unwrap_or_default();
        let original_ids: HashSet<Option<String>> = relationship_tags(&original_relationships)?
            .iter()
            .map(|tag| attr(tag, "Id"))
            .collect();
        let saved_tags = relationship_tags(&saved_relationships)?;
        let new_image_ids: Vec<String> = saved_tags
            .iter()
            .filter(|tag| attr(tag, "Type").is_some_and(|t| t.ends_with("/image")))
            .map(|tag| attr(tag, "Id"))
            .filter_map(|id| {
                id.filter(|id| !id.is_empty() && !original_ids.contains(&Some(id.clone())))
            })
            .collect();
        let images = collect_new_images(&part.get("blocks"))?;
        if images.len() > new_image_ids.len() {
            throw!(
                "Rust save image mutation mismatch in {rels_path}: {} model images, {} relationships",
                images.len(),
                new_image_ids.len()
            );
        }
        for (index, image) in images.items().iter().enumerate() {
            image.set("rId", new_image_ids[index].as_str());
        }

        let saved_hyperlinks: Vec<&String> = saved_tags
            .iter()
            .filter(|tag| {
                attr(tag, "Type").is_some_and(|t| t.ends_with("/hyperlink"))
                    && attr(tag, "TargetMode").as_deref() == Some("External")
            })
            .collect();
        for hyperlink in collect_external_hyperlinks(&part.get("blocks"))?.items() {
            if !hyperlink.get("href").truthy() {
                let r_id = hyperlink.get("rId");
                if relationship_for_id(&saved_relationships, r_id.as_str())?.is_none() {
                    hyperlink.set("rId", Js::Undefined);
                }
                continue;
            }
            let href = hyperlink.get("href");
            let mut found = None;
            for tag in &saved_hyperlinks {
                if href.as_str()
                    == Some(decode_xml_entities(&attr(tag, "Target").unwrap_or_default())?.as_str())
                {
                    found = Some(*tag);
                    break;
                }
            }
            if let Some(id) = found.and_then(|tag| attr(tag, "Id")) {
                if !id.is_empty() {
                    hyperlink.set("rId", id);
                }
            }
        }
    }

    let pkg = document.get("package");
    let headers = pkg.get("headers");
    let relationships = pkg.get("relationships");
    if !headers.truthy() || !relationships.truthy() {
        return Ok(());
    }
    for (owner_id, header) in headers.map_entries() {
        let watermark = header.get("watermark");
        if !watermark.truthy() || watermark.get("kind").as_str() != Some("picture") {
            continue;
        }
        let owner = relationships.map_get(&owner_id);
        if !owner.get("target").truthy() {
            continue;
        }
        let filename = header_footer_filename(owner.get("target").str_or_empty())?;
        let filename = filename.strip_prefix("word/").unwrap_or(&filename);
        let rels_path = format!("word/_rels/{filename}.rels");
        let saved_relationships = part_text(saved_parts.get(&rels_path))?.unwrap_or_default();
        if relationship_for_id(&saved_relationships, watermark.get("relId").as_str())?.is_some() {
            continue;
        }
        let media_path = watermark.get("mediaPath");
        let media_filename = media_path
            .as_str()
            .map(|p| p.rsplit('/').next().unwrap_or("").to_owned());
        let matched = relationship_tags(&saved_relationships)?
            .into_iter()
            .find(|tag| {
                if !attr(tag, "Type").is_some_and(|t| t.ends_with("/image")) {
                    return false;
                }
                let target =
                    attr(tag, "Target").map(|t| t.rsplit('/').next().unwrap_or("").to_owned());
                match &media_filename {
                    Some(name) if !name.is_empty() => target.as_deref() == Some(name.as_str()),
                    _ => true,
                }
            });
        if let Some(id) = matched.and_then(|tag| attr(&tag, "Id")) {
            if !id.is_empty() {
                watermark.set("relId", id);
            }
        }
    }
    Ok(())
}

fn is_tracked_wrapper_type(t: &str) -> bool {
    matches!(t, "insertion" | "deletion" | "moveFrom" | "moveTo")
}

/// `collectNewImages` (src/docx/rustSaveFacade.ts:239).
pub(crate) fn collect_new_images(blocks: &Js) -> Result<Js> {
    let images = Js::new_arr(Vec::new());
    let visit_run = |run: &Js| {
        for content in run.get("content").items() {
            let image = content.get("image");
            if content.get("type").as_str() == Some("drawing")
                && image
                    .get("src")
                    .as_str()
                    .is_some_and(|s| s.starts_with("data:"))
                && !image.get("rId").truthy()
            {
                images.push(image);
            }
        }
    };
    let visit_node = |node: &Js| -> Result<()> {
        if node.get("type").as_str() == Some("run") {
            visit_run(node);
        }
        Ok(())
    };
    for block in blocks.items() {
        match block.get("type").as_str() {
            Some("paragraph") => {
                for content in block.get("content").items() {
                    let t = content.get("type");
                    if t.as_str() == Some("run") {
                        visit_run(&content);
                    } else if t.as_str().is_some_and(is_tracked_wrapper_type) {
                        for inline in content.get("content").items() {
                            if inline.get("type").as_str() == Some("run") {
                                visit_run(&inline);
                            } else {
                                visit_tracked_control_content(&inline, &visit_node, false, true)?;
                            }
                        }
                    } else {
                        visit_tracked_control_content(&content, &visit_node, false, false)?;
                    }
                }
            }
            Some("table") => {
                for row in block.get("rows").items() {
                    for cell in row.get("cells").items() {
                        for image in collect_new_images(&cell.get("content"))?.items() {
                            images.push(image);
                        }
                    }
                }
            }
            _ => {}
        }
    }
    Ok(images)
}

/// `collectExternalHyperlinks` (src/docx/rustSaveFacade.ts:283).
pub(crate) fn collect_external_hyperlinks(blocks: &Js) -> Result<Js> {
    let hyperlinks = Js::new_arr(Vec::new());
    let wanted = |node: &Js| {
        node.get("type").as_str() == Some("hyperlink")
            && (node.get("href").truthy() || node.get("rId").truthy())
            && !node.get("anchor").truthy()
    };
    for block in blocks.items() {
        match block.get("type").as_str() {
            Some("paragraph") => {
                for content in block.get("content").items() {
                    if wanted(&content) {
                        hyperlinks.push(content.clone());
                    }
                    if content.get("type").as_str() != Some("run") {
                        visit_tracked_control_content(
                            &content,
                            &|node| {
                                if wanted(node) {
                                    hyperlinks.push(node.clone());
                                }
                                Ok(())
                            },
                            false,
                            false,
                        )?;
                    }
                }
            }
            Some("table") => {
                for row in block.get("rows").items() {
                    for cell in row.get("cells").items() {
                        for h in collect_external_hyperlinks(&cell.get("content"))?.items() {
                            hyperlinks.push(h);
                        }
                    }
                }
            }
            Some("blockSdt") => {
                for h in collect_external_hyperlinks(&block.get("content"))?.items() {
                    hyperlinks.push(h);
                }
            }
            _ => {}
        }
    }
    Ok(hyperlinks)
}

fn is_word_byte(b: u8) -> bool {
    b.is_ascii_alphanumeric() || b == b'_'
}

/// `relationshipTags` (src/docx/rustSaveFacade.ts:310): `/<Relationship\b[^>]*\/?>/g`.
pub(crate) fn relationship_tags(xml: &str) -> Result<Vec<String>> {
    let bytes = xml.as_bytes();
    let pin = "<Relationship";
    let mut tags = Vec::new();
    let mut from = 0;
    while let Some(found) = xml[from..].find(pin) {
        let start = from + found;
        let after = start + pin.len();
        if bytes.get(after).is_some_and(|b| is_word_byte(*b)) {
            from = start + 1;
            continue;
        }
        match xml[after..].find('>') {
            Some(gt) => {
                let end = after + gt + 1;
                tags.push(xml[start..end].to_owned());
                from = end;
            }
            None => from = start + 1,
        }
    }
    Ok(tags)
}

/// `relationshipAttribute` (src/docx/rustSaveFacade.ts:314): `` /\b${name}="([^"]*)"/ ``.
pub(crate) fn relationship_attribute(tag: &str, name: &str) -> Result<Option<String>> {
    let bytes = tag.as_bytes();
    let pin = format!("{name}=\"");
    let mut from = 0;
    while let Some(found) = tag[from..].find(&pin) {
        let start = from + found;
        from = start + 1;
        let boundary = start == 0 || !is_word_byte(bytes[start - 1]);
        if !boundary && name.as_bytes().first().is_some_and(|b| is_word_byte(*b)) {
            continue;
        }
        let value = start + pin.len();
        if let Some(close) = tag[value..].find('"') {
            return Ok(Some(tag[value..value + close].to_owned()));
        }
    }
    Ok(None)
}

/// `relationshipForId` (src/docx/rustSaveFacade.ts:318).
pub(crate) fn relationship_for_id(xml: &str, id: Option<&str>) -> Result<Option<String>> {
    let Some(id) = id.filter(|id| !id.is_empty()) else {
        return Ok(None);
    };
    for tag in relationship_tags(xml)? {
        if relationship_attribute(&tag, "Id")?.as_deref() == Some(id) {
            return Ok(Some(tag));
        }
    }
    Ok(None)
}

/// `String.fromCodePoint(n)`; a lone surrogate becomes U+FFFD (Rust strings cannot hold one).
fn from_code_point(n: f64) -> Result<String> {
    if !(n.fract() == 0.0 && (0.0..=1_114_111.0).contains(&n)) {
        return Err(SaveError::new(format!(
            "Invalid code point {}",
            crate::js::num_to_string(n)
        )));
    }
    Ok(char::from_u32(n as u32).unwrap_or('\u{fffd}').to_string())
}

/// `value.replace(/&#(\d+);/g, …)` / `/&#x([0-9A-Fa-f]+);/g`.
fn replace_numeric(value: &str, hex: bool) -> Result<String> {
    let pin = if hex { "&#x" } else { "&#" };
    let mut out = String::new();
    let mut rest = value;
    while let Some(found) = rest.find(pin) {
        let digits_start = found + pin.len();
        let digits_len = rest[digits_start..]
            .bytes()
            .take_while(|b| {
                if hex {
                    b.is_ascii_hexdigit()
                } else {
                    b.is_ascii_digit()
                }
            })
            .count();
        let end = digits_start + digits_len;
        if digits_len > 0 && rest.as_bytes().get(end) == Some(&b';') {
            let digits = &rest[digits_start..end];
            let n = if hex {
                digits.chars().fold(0.0f64, |acc, c| {
                    acc * 16.0 + f64::from(c.to_digit(16).unwrap())
                })
            } else {
                crate::js::to_number(&Js::from(digits))
            };
            out.push_str(&rest[..found]);
            out.push_str(&from_code_point(n)?);
            rest = &rest[end + 1..];
        } else {
            out.push_str(&rest[..found + 1]);
            rest = &rest[found + 1..];
        }
    }
    out.push_str(rest);
    Ok(out)
}

/// `decodeXmlEntities` (src/docx/rustSaveFacade.ts:323).
pub(crate) fn decode_xml_entities(value: &str) -> Result<String> {
    let v = value
        .replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&quot;", "\"")
        .replace("&apos;", "'");
    let v = replace_numeric(&v, false)?;
    let v = replace_numeric(&v, true)?;
    Ok(v.replace("&amp;", "&"))
}

fn is_iso_millis(now: &str) -> bool {
    let b = now.as_bytes();
    let shape = b"dddd-dd-ddTdd:dd:dd.dddZ";
    b.len() == shape.len()
        && b.iter().zip(shape.iter()).all(|(c, s)| {
            if *s == b'd' {
                c.is_ascii_digit()
            } else {
                c == s
            }
        })
}

/// `validateDeterminism` (src/docx/rustSaveFacade.ts:336).
pub(crate) fn validate_determinism(value: &Js) -> Result<()> {
    let seed = crate::js::to_js_string(&value.get("seed"));
    if !(seed.len() == 64 && seed.bytes().all(|b| b.is_ascii_hexdigit())) {
        throw!("Rust save determinism seed must be a SHA-256 hex digest");
    }
    if !is_iso_millis(&crate::js::to_js_string(&value.get("now"))) {
        throw!("Rust save clock must be a UTC ISO-8601 millisecond timestamp");
    }
    Ok(())
}

const FORBIDDEN_KEYS: [&str; 3] = ["__proto__", "constructor", "prototype"];

/// `assertSafeSaveTree` (src/docx/rustSaveFacade.ts:348). `active` is the caller's `WeakSet`;
/// this port always starts from an empty one.
pub(crate) fn assert_safe_save_tree(value: &Js, path: &str, _active: &Js) -> Result<()> {
    enum Segment {
        Index(usize),
        Key(String),
    }
    fn current(path: &str, segments: &[Segment]) -> String {
        let mut result = path.to_owned();
        for segment in segments {
            match segment {
                Segment::Index(i) => result.push_str(&format!("[{i}]")),
                Segment::Key(k) => {
                    result.push('.');
                    result.push_str(k);
                }
            }
        }
        result
    }
    fn visit(
        value: &Js,
        path: &str,
        segments: &mut Vec<Segment>,
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
        if value.is_arr() {
            for (index, item) in value.items().iter().enumerate() {
                segments.push(Segment::Index(index));
                visit(item, path, segments, active)?;
                segments.pop();
            }
        } else {
            for (key, child) in value.entries() {
                if FORBIDDEN_KEYS.contains(&key.as_str()) {
                    throw!("{} contains forbidden key {key}", current(path, segments));
                }
                segments.push(Segment::Key(key));
                visit(&child, path, segments, active)?;
                segments.pop();
            }
        }
        active.remove(&p);
        Ok(())
    }
    visit(value, path, &mut Vec::new(), &mut HashSet::new())
}

/// `exactArrayBuffer` (src/docx/rustSaveFacade.ts:382).
pub(crate) fn exact_array_buffer(bytes: &Rc<Vec<u8>>) -> Result<Vec<u8>> {
    Ok((**bytes).clone())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn iso_strings_match_to_iso_string() {
        // bun: new Date(n).toISOString()
        assert_eq!(iso_string(0), "1970-01-01T00:00:00.000Z");
        assert_eq!(iso_string(1_700_000_000_123), "2023-11-14T22:13:20.123Z");
        assert_eq!(iso_string(951_782_400_000), "2000-02-29T00:00:00.000Z");
        assert_eq!(iso_string(-1), "1969-12-31T23:59:59.999Z");
        assert_eq!(iso_string(4_102_444_800_000), "2100-01-01T00:00:00.000Z");
        assert!(is_iso_millis(&now_iso()));
    }

    #[test]
    fn relationship_scanning() {
        let xml = r#"<Relationships><Relationship Id="rId1" Type="a/image" Target="m/x.png"/><RelationshipX Id="no"/><Relationship TargetMode="External" Id="rId2" Target="a&amp;b&#65;&#x42;"></Relationship></Relationships>"#;
        let tags = relationship_tags(xml).unwrap();
        assert_eq!(tags.len(), 2);
        assert_eq!(
            relationship_attribute(&tags[1], "Id").unwrap().as_deref(),
            Some("rId2")
        );
        assert_eq!(relationship_attribute(&tags[1], "Mode").unwrap(), None);
        assert_eq!(
            decode_xml_entities(&relationship_attribute(&tags[1], "Target").unwrap().unwrap())
                .unwrap(),
            "a&bAB"
        );
        assert_eq!(decode_xml_entities("&amp;lt;").unwrap(), "&lt;");
        assert!(relationship_for_id(xml, Some("rId1")).unwrap().is_some());
    }
}
