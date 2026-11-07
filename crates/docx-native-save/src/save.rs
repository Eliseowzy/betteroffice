//! Port of `packages/docx/src/yrs/saveYrsDocx.ts`.

use std::cell::{Cell, RefCell};
use std::collections::{HashMap, HashSet};
use std::rc::Rc;

use indexmap::{IndexMap, IndexSet};
use sha2::{Digest, Sha256};

use crate::editor_save_keys::{editor_save_keys, note_separator_keys};
use crate::header_footer_alias_projection::{
    header_footer_alias_of, is_written_by_canonical, mark_header_footer_alias,
};
use crate::host_media::resolve_comment_media;
use crate::js::{self, Js, Result, SaveError};
use crate::parse_facade::parse_document_with_rust;
use crate::resident_comment_projection::saved_comment_ids;
use crate::save_facade::{self, SaveContext, hex, rezip, unzip, write_document_with_rust};
use crate::sections::section_properties_equal;
use crate::session::Session;
use crate::yrs_to_document::{
    assert_comment_bookmark_replacement, projected_session_key, restore_block_bookmarks,
    yrs_to_document,
};
use crate::{arr, obj, throw};

/// A JS `Map`/`Set` key for a primitive or object, with SameValueZero semantics.
fn mkey(value: &Js) -> String {
    match value.ptr() {
        Some(p) => format!("object:{p}"),
        None => match value {
            Js::Num(n) if *n == 0.0 => "number:0".into(),
            other => format!("{}:{}", other.type_of(), js::to_js_string(other)),
        },
    }
}

fn empty_arr() -> Js {
    Js::new_arr(Vec::new())
}

/// `x?.length` truthiness.
fn has_length(value: &Js) -> bool {
    !value.is_empty()
}

/// `Array.prototype.includes` (SameValueZero) on an array, `undefined` for a non-array.
fn includes(list: &Js, value: &Js) -> Option<bool> {
    if !list.is_arr() {
        return None;
    }
    Some(list.items().iter().any(|item| match (item, value) {
        (Js::Num(a), Js::Num(b)) => a == b || (a.is_nan() && b.is_nan()),
        (a, b) => a.strict_eq(b),
    }))
}

/// `part.replace(/^\//, '')`.
fn strip_slash(s: &str) -> &str {
    s.strip_prefix('/').unwrap_or(s)
}

/// `storyBlocks` (src/yrs/saveYrsDocx.ts:78).
pub(crate) fn story_blocks(document: &Js) -> Result<Js> {
    let pkg = document.get("package");
    let mut out = vec![pkg.path(&["document", "content"])];
    for entries in [pkg.get("headers"), pkg.get("footers")] {
        if !entries.truthy() {
            continue;
        }
        for (_, part) in entries.map_entries() {
            if !is_written_by_canonical(&part, &entries)? {
                out.push(part.get("content"));
            }
        }
    }
    for note in pkg.get("footnotes").items() {
        out.push(note.get("content"));
    }
    for note in pkg.get("endnotes").items() {
        out.push(note.get("content"));
    }
    Ok(Js::new_arr(out))
}

const COMMENTS_PART: &str = "word/comments.xml";
const COMMENT_SOURCE_PARTS: [&str; 8] = [
    COMMENTS_PART,
    "word/commentsExtended.xml",
    "word/commentsIds.xml",
    "word/commentsExtensible.xml",
    "word/_rels/comments.xml.rels",
    "[Content_Types].xml",
    "word/_rels/document.xml.rels",
    "docProps/core.xml",
];

/// `commentSourceBuffer` (src/yrs/saveYrsDocx.ts:98).
pub(crate) fn comment_source_buffer(
    original: &Rc<Vec<u8>>,
    baseline: &Rc<Vec<u8>>,
) -> Result<Vec<u8>> {
    let mut parts = unzip(original)?;
    let saved = unzip(baseline)?;
    for path in COMMENT_SOURCE_PARTS {
        match saved.get(path) {
            Some(bytes) => {
                parts.insert(path.to_owned(), bytes.clone());
            }
            None => {
                parts.shift_remove(path);
            }
        }
    }
    rezip(&parts)
}

/// `commentBlocks` (src/yrs/saveYrsDocx.ts:108): `body` itself, or `undefined`.
pub(crate) fn comment_blocks(body: &Js) -> Result<Js> {
    let ok = body.is_arr()
        && body.items().iter().all(|entry| {
            entry.truthy()
                && matches!(
                    entry.get("type").as_str(),
                    Some("paragraph" | "table" | "blockSdt" | "rawXml")
                )
        });
    Ok(if ok { body.clone() } else { Js::Undefined })
}

/// `commentContent` (src/yrs/saveYrsDocx.ts:114).
pub(crate) fn comment_content(body: &Js) -> Result<Js> {
    let blocks = comment_blocks(body)?;
    if blocks.truthy() && !blocks.is_empty() {
        return Ok(Js::new_arr(
            blocks
                .items()
                .into_iter()
                .filter(|b| b.get("type").as_str() == Some("paragraph"))
                .collect(),
        ));
    }
    let text = body.as_str().unwrap_or("");
    let content = if text.is_empty() {
        empty_arr()
    } else {
        arr![
            obj! { "type" => "run", "formatting" => obj!{}, "content" => arr![obj! { "type" => "text", "text" => text }] }
        ]
    };
    Ok(arr![
        obj! { "type" => "paragraph", "formatting" => obj!{}, "content" => content }
    ])
}

/// Whether `key` is an array index, which JS orders first among an object's keys.
fn array_index(key: &str) -> Option<u32> {
    if key == "0" {
        return Some(0);
    }
    if key.is_empty() || key.starts_with('0') || !key.bytes().all(|b| b.is_ascii_digit()) {
        return None;
    }
    key.parse::<u32>().ok().filter(|n| *n != u32::MAX)
}

/// `commentBodyKey` (src/yrs/saveYrsDocx.ts:129): JSON with every object's keys sorted.
pub(crate) fn comment_body_key(body: &Js) -> Result<Option<String>> {
    fn sorted(value: &Js) -> Js {
        match value {
            Js::Arr(_) => Js::new_arr(value.items().iter().map(sorted).collect()),
            Js::Obj(_) => {
                let mut entries = value.entries();
                entries.sort_by(|(a, _), (b, _)| js::cmp16(a, b));
                // Object.fromEntries puts array-index keys first, in ascending order.
                let (mut indices, rest): (Vec<_>, Vec<_>) = entries
                    .into_iter()
                    .partition(|(k, _)| array_index(k).is_some());
                indices.sort_by_key(|(k, _)| array_index(k));
                let out = Js::new_obj();
                for (k, v) in indices.into_iter().chain(rest) {
                    out.set(&k, sorted(&v));
                }
                out
            }
            Js::Map(_) | Js::Bytes(_) => Js::new_obj(),
            other => other.clone(),
        }
    }
    Ok(sorted(body).stringify())
}

/// `savedComments` (src/yrs/saveYrsDocx.ts:136): `{ comments, sessionIds, changed }`.
/// `identities` is a `Map` of session comment id to `{ id, paraId?, durableId? }` (or
/// undefined), `baseline` the baseline's comment list (or undefined). `sessionIds`, a
/// `Map<number, string>` in TS, is a `Map` keyed by `String(number)`.
pub(crate) fn saved_comments(
    session: &Session,
    base: &Js,
    identities: &Js,
    baseline: &Js,
) -> Result<Js> {
    let source = base
        .path(&["package", "document", "comments"])
        .nullish_or(empty_arr());
    let listed = session.list_comments()?;
    let data_url = |token: &str| -> Result<Option<String>> { Ok(session.media_data_url(token)) };
    resolve_comment_media(&listed, &data_url)?;
    let source_ids: Vec<f64> = source
        .items()
        .iter()
        .map(|c| js::to_number(&c.get("id")))
        .collect();
    let previous_ids = if identities.truthy() {
        Js::map_from(
            identities
                .map_entries()
                .into_iter()
                .map(|(id, identity)| (id, identity.get("id"))),
        )
    } else {
        Js::Undefined
    };
    let ids = saved_comment_ids(&source_ids, &listed, &previous_ids)?;
    let mut held: IndexMap<String, Js> = IndexMap::new();
    for comment in listed.items() {
        held.insert(mkey(&comment.get("id")), comment);
    }
    let prior = Js::new_arr(baseline.items()).deep_clone();
    resolve_comment_media(&prior, &data_url)?;
    let mut previous: HashMap<String, Js> = HashMap::new();
    for comment in prior.items() {
        previous.insert(mkey(&comment.get("id")), comment);
    }
    let mut aliases: HashMap<String, String> = HashMap::new();
    if identities.truthy() {
        for (id, identity) in identities.map_entries() {
            aliases.insert(mkey(&identity.get("id")), id);
        }
    }
    let session_id = |comment: &Js| -> String {
        aliases
            .get(&mkey(&comment.get("id")))
            .cloned()
            .unwrap_or_else(|| js::to_js_string(&comment.get("id")))
    };
    let mut retained: Option<HashSet<String>> = None;
    if source
        .items()
        .iter()
        .any(|c| !held.contains_key(&mkey(&Js::from(session_id(c)))))
    {
        let read = session.export_structured(&obj! {
            "revisionView" => "markup", "stories" => arr!["comments"], "maxBlocks" => 1_000_000.0, "maxBytes" => 67_108_864.0,
        })?;
        if !read.get("ok").truthy() || read.path(&["content", "truncated"]).truthy() {
            throw!("The remaining source comments could not be read for saving");
        }
        let mut set = HashSet::new();
        for story in read.path(&["content", "stories"]).items() {
            if story.get("comment").truthy() {
                set.insert(mkey(&story.path(&["comment", "id"])));
            }
        }
        retained = Some(set);
    }
    let changed = Cell::new(false);
    let parent_of = |parent_id: &Js| -> Js {
        if parent_id.is_null() {
            return Js::Undefined;
        }
        let id = match parent_id.as_str() {
            Some(s) => ids.map_get(s),
            None => Js::Undefined,
        };
        if !id.is_undefined() {
            return id;
        }
        let numeric = js::to_number(parent_id);
        if numeric.is_finite() && Js::from(js::num_to_string(numeric)).strict_eq(parent_id) {
            Js::Num(numeric)
        } else {
            Js::Undefined
        }
    };
    let mut comments: Vec<Js> = Vec::new();
    let kept: Vec<Js> = source
        .items()
        .into_iter()
        .filter(|comment| match &retained {
            None => true,
            Some(set) if set.contains(&mkey(&Js::from(session_id(comment)))) => true,
            Some(_) => {
                changed.set(true);
                false
            }
        })
        .collect();
    for comment in kept {
        let key = mkey(&Js::from(session_id(&comment)));
        let entry = held.get(&key).cloned();
        let before = previous.get(&key).cloned();
        held.shift_remove(&key);
        let Some(entry) = entry else {
            comments.push(comment);
            continue;
        };
        if !entry.get("author").truthy()
            && !entry.get("date").truthy()
            && entry.get("body").is_nullish()
            && !entry.get("done").truthy()
            && entry.get("parentId").is_null()
        {
            comments.push(comment);
            continue;
        }
        let comment_body = comment
            .get("blockContent")
            .nullish_or(comment.get("content"));
        let prior_body = match &before {
            Some(b) => b.get("body"),
            None => comment_body.clone(),
        };
        let body_changed = !entry.get("body").is_nullish()
            && comment_body_key(&entry.get("body"))? != comment_body_key(&prior_body)?;
        let author = match &before {
            Some(b) if entry.get("author").strict_eq(&b.get("author")) => comment.get("author"),
            _ => entry.get("author"),
        };
        let before_date = before.as_ref().map(|b| b.get("date")).unwrap_or_default();
        let date = if entry.get("date").strict_eq(
            &before_date
                .nullish_or(comment.get("date"))
                .nullish_or(Js::from("")),
        ) {
            comment.get("date")
        } else {
            entry.get("date").or(Js::Undefined)
        };
        let parent_id = parent_of(&entry.get("parentId"));
        if !body_changed
            && author.strict_eq(&comment.get("author"))
            && date.strict_eq(&comment.get("date"))
            && entry
                .get("done")
                .strict_eq(&comment.get("done").nullish_or(Js::Bool(false)))
            && parent_id.strict_eq(&comment.get("parentId"))
        {
            comments.push(comment);
            continue;
        }
        changed.set(true);
        if body_changed {
            assert_comment_bookmark_replacement(&comment_body)?;
        }
        let content = if body_changed {
            comment_content(&entry.get("body"))?
        } else {
            Js::Undefined
        };
        let blocks = comment_blocks(&entry.get("body"))?
            .nullish_or(content.clone())
            .nullish_or(empty_arr());
        let restored = if content.truthy() {
            restore_block_bookmarks(
                &blocks,
                &comment_body,
                &format!("comment:{}", js::to_js_string(&comment.get("id"))),
            )?
        } else {
            blocks.clone()
        };
        let out = comment.with(&[
            ("author", author),
            ("date", date),
            ("done", entry.get("done")),
            ("parentId", parent_id),
        ]);
        if !comment.get("status").is_undefined() {
            out.set(
                "status",
                if entry.get("done").truthy() {
                    "resolved"
                } else {
                    "active"
                },
            );
        }
        if content.truthy() {
            out.set("content", content);
            out.set(
                "blockContent",
                if restored.strict_eq(&blocks) {
                    Js::Undefined
                } else {
                    restored
                },
            );
        }
        comments.push(out);
    }
    let session_ids = Js::new_map();
    for (id, numeric) in ids.map_entries() {
        let text = js::to_js_string(&numeric);
        if id != text {
            session_ids.map_set(&text, id);
        }
    }
    for entry in held.values() {
        changed.set(true);
        let entry_id = entry.get("id");
        let id = match entry_id.as_str() {
            Some(s) => ids.map_get(s),
            None => Js::Undefined,
        };
        let identity = match entry_id.as_str() {
            Some(s) if identities.truthy() => identities.map_get(s),
            _ => Js::Undefined,
        };
        if !Js::from(js::to_js_string(&id)).strict_eq(&entry_id) {
            session_ids.map_set(&js::to_js_string(&id), entry_id.clone());
        }
        let added = obj! { "id" => id };
        if identity.get("paraId").truthy() {
            added.set("paraId", identity.get("paraId"));
        }
        if identity.get("durableId").truthy() {
            added.set("durableId", identity.get("durableId"));
        }
        added.set("author", entry.get("author"));
        added.set("done", entry.get("done"));
        added.set("parentId", parent_of(&entry.get("parentId")));
        if entry.get("date").truthy() {
            added.set("date", entry.get("date"));
        }
        added.set("content", comment_content(&entry.get("body"))?);
        comments.push(added);
    }
    Ok(
        obj! { "comments" => Js::new_arr(comments), "sessionIds" => session_ids, "changed" => changed.get() },
    )
}

/// `savedParagraphs` (src/yrs/saveYrsDocx.ts:217).
pub(crate) fn saved_paragraphs(document: &Js, identities: &Js) -> Result<Js> {
    let mut by_key: HashMap<String, Js> = HashMap::new();
    for identity in identities.get("paragraphs").items() {
        if identity.get("session").truthy() {
            by_key.insert(mkey(&identity.path(&["session", "paraId"])), identity);
        }
    }
    let saved = empty_arr();
    fn visit(blocks: &Js, by_key: &HashMap<String, Js>, saved: &Js) -> Result<()> {
        for block in blocks.items() {
            match block.get("type").as_str() {
                Some("table") => {
                    for row in block.get("rows").items() {
                        for cell in row.get("cells").items() {
                            visit(&cell.get("content"), by_key, saved)?;
                        }
                    }
                }
                Some("blockSdt") => visit(&block.get("content"), by_key, saved)?,
                Some("paragraph") if block.get("paraId").truthy() => {
                    let key = projected_session_key(&block)?;
                    let identity = key
                        .and_then(|k| by_key.get(&mkey(&Js::from(k))).cloned())
                        .unwrap_or_default();
                    if !identity.get("session").truthy() || !identity.get("persisted").truthy() {
                        continue;
                    }
                    saved.push(obj! {
                        "session" => identity.get("session"),
                        "persisted" => identity.get("persisted").with(&[("paraId", block.get("paraId"))]),
                    });
                }
                _ => {}
            }
        }
        Ok(())
    }
    for blocks in story_blocks(document)?.items() {
        visit(&blocks, &by_key, &saved)?;
    }
    Ok(saved)
}

/// `mapParagraphs` (src/yrs/saveYrsDocx.ts:253): `map` returns the paragraph, a replacement,
/// or `null` to drop it; the same array when nothing changes.
pub(crate) fn map_paragraphs(blocks: &Js, map: &dyn Fn(&Js) -> Result<Js>) -> Result<Js> {
    let mut kept = Vec::new();
    let mut changed = false;
    for block in blocks.items() {
        let mut next = block.clone();
        match block.get("type").as_str() {
            Some("paragraph") => next = map(&block)?,
            Some("table") => {
                let old_rows = block.get("rows").items();
                let mut rows = Vec::with_capacity(old_rows.len());
                for row in &old_rows {
                    let old_cells = row.get("cells").items();
                    let mut cells = Vec::with_capacity(old_cells.len());
                    for cell in &old_cells {
                        let content = map_paragraphs(&cell.get("content"), map)?;
                        cells.push(if content.strict_eq(&cell.get("content")) {
                            cell.clone()
                        } else {
                            cell.with(&[("content", content)])
                        });
                    }
                    let same = cells
                        .iter()
                        .enumerate()
                        .all(|(i, c)| c.strict_eq(&old_cells.get(i).cloned().unwrap_or_default()));
                    rows.push(if same {
                        row.clone()
                    } else {
                        row.with(&[("cells", Js::new_arr(cells))])
                    });
                }
                if rows
                    .iter()
                    .enumerate()
                    .any(|(i, r)| !r.strict_eq(&old_rows.get(i).cloned().unwrap_or_default()))
                {
                    next = block.with(&[("rows", Js::new_arr(rows))]);
                }
            }
            Some("blockSdt") => {
                let content = map_paragraphs(&block.get("content"), map)?;
                if !content.strict_eq(&block.get("content")) {
                    next = block.with(&[("content", content)]);
                }
            }
            _ => {}
        }
        changed = changed || !next.strict_eq(&block);
        if next.truthy() {
            kept.push(next);
        }
    }
    Ok(if changed {
        Js::new_arr(kept)
    } else {
        blocks.clone()
    })
}

/// `mapStoryParagraphs` (src/yrs/saveYrsDocx.ts:283).
pub(crate) fn map_story_paragraphs(document: &Js, map: &dyn Fn(&Js) -> Result<Js>) -> Result<Js> {
    let with_content = |owner: &Js| -> Result<Js> {
        let content = map_paragraphs(&owner.get("content"), map)?;
        Ok(if content.strict_eq(&owner.get("content")) {
            owner.clone()
        } else {
            owner.with(&[("content", content)])
        })
    };
    let pkg = document.get("package");
    let parts = |stories: Js| -> Result<Js> {
        if !stories.truthy() {
            return Ok(stories);
        }
        let out = Js::new_map();
        for (id, part) in stories.map_entries() {
            let projected = with_content(&part)?;
            let canonical = header_footer_alias_of(&part)?;
            if !projected.strict_eq(&part)
                && let Some(canonical) = canonical {
                    mark_header_footer_alias(&projected, &canonical)?;
                }
            out.map_set(&id, projected);
        }
        Ok(out)
    };
    let notes = |notes: Js| -> Result<Js> {
        if notes.is_nullish() {
            return Ok(Js::Undefined);
        }
        Ok(Js::new_arr(
            notes
                .items()
                .iter()
                .map(&with_content)
                .collect::<Result<Vec<_>>>()?,
        ))
    };
    let new_pkg = pkg.with(&[
        ("document", with_content(&pkg.get("document"))?),
        ("headers", parts(pkg.get("headers"))?),
        ("footers", parts(pkg.get("footers"))?),
        ("footnotes", notes(pkg.get("footnotes"))?),
        ("endnotes", notes(pkg.get("endnotes"))?),
    ]);
    Ok(document.with(&[("package", new_pkg)]))
}

/// `withoutSyntheticParagraphs` (src/yrs/saveYrsDocx.ts:317).
pub(crate) fn without_synthetic_paragraphs(document: &Js, identities: &Js) -> Result<Js> {
    let mut synthetic: HashSet<String> = HashSet::new();
    for identity in identities.get("paragraphs").items() {
        if identity.get("origin").as_str() == Some("synthetic") && identity.get("session").truthy()
        {
            synthetic.insert(mkey(&identity.path(&["session", "paraId"])));
        }
    }
    if synthetic.is_empty() {
        return Ok(document.clone());
    }
    map_story_paragraphs(document, &|paragraph| {
        let key = projected_session_key(paragraph)?;
        let empty = paragraph.get("content").items().iter().all(|node| {
            node.get("type").as_str() == Some("run")
                && node.get("content").items().iter().all(|content| {
                    content.get("type").as_str() == Some("break")
                        && matches!(content.get("breakType").as_str(), Some("page" | "column"))
                })
        });
        Ok(
            if empty && key.is_some_and(|k| synthetic.contains(&mkey(&Js::from(k)))) {
                Js::Null
            } else {
                paragraph.clone()
            },
        )
    })
}

/// `withSourceOrdinals` (src/yrs/saveYrsDocx.ts:341).
pub(crate) fn with_source_ordinals(document: &Js, parts: &Js) -> Result<Js> {
    let mut ordinals: HashMap<String, Js> = HashMap::new();
    for part in parts.items() {
        for pair in part.get("paragraphs").items() {
            ordinals.insert(mkey(&pair.at(1)), pair.at(0));
        }
    }
    if ordinals.is_empty() {
        return Ok(document.clone());
    }
    map_story_paragraphs(document, &|paragraph| {
        let key = projected_session_key(paragraph)?;
        let ordinal = key
            .and_then(|k| ordinals.get(&mkey(&Js::from(k))).cloned())
            .unwrap_or_default();
        Ok(
            if ordinal.is_undefined() || paragraph.get("sourceOrdinal").strict_eq(&ordinal) {
                paragraph.clone()
            } else {
                paragraph.with(&[("sourceOrdinal", ordinal)])
            },
        )
    })
}

/// `captureSessionSave` (src/yrs/saveYrsDocx.ts:365); `identities` undefined for the default.
pub(crate) fn capture_session_save(session: &Session, identities: &Js) -> Result<Js> {
    let identities = if identities.is_undefined() {
        session.paragraph_identities()?
    } else {
        identities.clone()
    };
    let plan = session.paragraph_save_plan()?;
    Ok(obj! { "identities" => identities, "plan" => plan })
}

/// `sectionSaveMetadata` (src/yrs/saveYrsDocx.ts:375).
pub(crate) fn section_save_metadata(document: &Js) -> Result<String> {
    let pkg = document.get("package");
    let keys = editor_save_keys(
        &document.with(&[(
            "package",
            pkg.with(&[(
                "document",
                pkg.get("document").with(&[
                    ("sections", Js::Undefined),
                    ("finalSectionProperties", Js::Undefined),
                ]),
            )]),
        )]),
        &Js::Undefined,
    )?;
    let metadata = obj! {
        "keys" => keys,
        "numbering" => pkg.get("numbering"),
        "settings" => pkg.get("settings"),
        "styles" => pkg.get("styles"),
        "theme" => pkg.get("theme"),
        "fontTable" => pkg.get("fontTable"),
        "properties" => pkg.get("properties"),
        "background" => pkg.get("background"),
        "customRootBindings" => pkg.path(&["document", "customRootBindings"]),
        "commentAuthors" => pkg.path(&["document", "commentAuthors"]),
        "footnoteSeparators" => pkg.get("footnoteSeparators"),
        "endnoteSeparators" => pkg.get("endnoteSeparators"),
    };
    Ok(metadata.stringify_maps_as_entries().unwrap_or_default())
}

const MODELED_SECTION_CHILDREN: [&str; 16] = [
    "headerReference",
    "footerReference",
    "footnotePr",
    "endnotePr",
    "type",
    "pgSz",
    "pgMar",
    "paperSrc",
    "pgBorders",
    "lnNumType",
    "pgNumType",
    "cols",
    "vAlign",
    "titlePg",
    "bidi",
    "docGrid",
];

/// One match of `/<(\/?)([\w.-]+(?::[\w.-]+)?)(?:[^>"']|"[^"]*"|'[^']*')*?(\/?)>/` at `at`:
/// `(close, name, empty, end)`.
fn match_tag(b: &[u8], at: usize) -> Option<(bool, &[u8], bool, usize)> {
    let name_char = |c: u8| c.is_ascii_alphanumeric() || c == b'_' || c == b'.' || c == b'-';
    let mut i = at + 1;
    let close = b.get(i) == Some(&b'/');
    if close {
        i += 1;
    }
    let name_start = i;
    while i < b.len() && name_char(b[i]) {
        i += 1;
    }
    if i == name_start {
        return None;
    }
    if b.get(i) == Some(&b':') && b.get(i + 1).is_some_and(|c| name_char(*c)) {
        i += 1;
        while i < b.len() && name_char(b[i]) {
            i += 1;
        }
    }
    let name = &b[name_start..i];
    loop {
        match b.get(i)? {
            b'/' if b.get(i + 1) == Some(&b'>') => return Some((close, name, true, i + 2)),
            b'>' => return Some((close, name, false, i + 1)),
            q @ (b'"' | b'\'') => {
                let q = *q;
                let rel = b[i + 1..].iter().position(|c| *c == q)?;
                i += rel + 2;
            }
            _ => i += 1,
        }
    }
}

/// `sectionsModeled` (src/yrs/saveYrsDocx.ts:392).
pub(crate) fn sections_modeled(xml: &str) -> Result<bool> {
    let b = xml.as_bytes();
    let mut depth = 0i64;
    let mut prefix: &[u8] = b"";
    let mut i = 0;
    while i < b.len() {
        if b[i] != b'<' {
            i += 1;
            continue;
        }
        let Some((close, name, empty, end)) = match_tag(b, i) else {
            i += 1;
            continue;
        };
        i = end;
        let (space, local): (&[u8], &[u8]) = match name.iter().position(|c| *c == b':') {
            Some(colon) => (&name[..colon], &name[colon + 1..]),
            None => (b"", name),
        };
        if depth == 0 {
            if local == b"sectPr" && !close && !empty {
                depth = 1;
                prefix = space;
            }
            continue;
        }
        if close {
            depth -= 1;
            continue;
        }
        if depth == 1
            && (space != prefix
                || !MODELED_SECTION_CHILDREN
                    .iter()
                    .any(|c| c.as_bytes() == local))
        {
            return Ok(false);
        }
        if !empty {
            depth += 1;
        }
    }
    Ok(true)
}

/// Collects the paragraphs of `blocks`, in table cells and block content controls too.
fn visit_paragraphs(blocks: &Js, paragraphs: &mut Vec<Js>) {
    for block in blocks.items() {
        match block.get("type").as_str() {
            Some("paragraph") => paragraphs.push(block),
            Some("table") => {
                for row in block.get("rows").items() {
                    for cell in row.get("cells").items() {
                        visit_paragraphs(&cell.get("content"), paragraphs);
                    }
                }
            }
            Some("blockSdt") => visit_paragraphs(&block.get("content"), paragraphs),
            _ => {}
        }
    }
}

fn index_of(list: &[Js], value: &Js) -> i64 {
    list.iter()
        .position(|p| p.strict_eq(value))
        .map_or(-1, |i| i as i64)
}

/// `list.slice(start, end)` with JS negative indices.
fn slice<T: Clone>(list: &[T], start: i64, end: Option<i64>) -> Vec<T> {
    let len = list.len() as i64;
    let norm = |i: i64| if i < 0 { (len + i).max(0) } else { i.min(len) };
    let (s, e) = (norm(start), end.map_or(len, norm));
    if s >= e {
        Vec::new()
    } else {
        list[s as usize..e as usize].to_vec()
    }
}

/// `sectionSavePlan` (src/yrs/saveYrsDocx.ts:417): `{ plan?, lossy, strict? }` or undefined.
pub(crate) fn section_save_plan(
    session: &Session,
    output: &Js,
    capture: &Js,
    original_buffer: &Rc<Vec<u8>>,
) -> Result<Js> {
    let source = session.materialized_source()?;
    if !source.truthy() {
        return Ok(Js::Undefined);
    }
    let plan = capture.get("plan");
    let identities = capture.path(&["identities", "paragraphs"]).items();
    let mut planned_splices = plan.get("splicedParts").items();
    planned_splices.extend(plan.get("sectionSplicedParts").items());
    let source_content = source.path(&["package", "document", "content"]);
    let output_content = output.path(&["package", "document", "content"]);
    let mut source_paragraphs = Vec::new();
    let mut output_paragraphs = Vec::new();
    visit_paragraphs(&source_content, &mut source_paragraphs);
    visit_paragraphs(&output_content, &mut output_paragraphs);
    let body: HashSet<usize> = source_content
        .items()
        .iter()
        .filter(|b| b.get("type").as_str() == Some("paragraph"))
        .filter_map(Js::ptr)
        .collect();
    let ordinal_of = |p: &Js, index: usize| p.get("sourceOrdinal").nullish_or(Js::from(index));
    let mut by_ordinal: HashMap<String, Js> = HashMap::new();
    for (index, p) in source_paragraphs.iter().enumerate() {
        by_ordinal.insert(mkey(&ordinal_of(p, index)), p.clone());
    }
    // donors: [{ ordinal, properties }]
    let donors: Vec<(Js, Js)> = source_paragraphs
        .iter()
        .enumerate()
        .filter(|(_, p)| {
            p.ptr().is_some_and(|ptr| body.contains(&ptr)) && p.get("sectionProperties").truthy()
        })
        .map(|(index, p)| (ordinal_of(p, index), p.get("sectionProperties")))
        .collect();
    let structural = planned_splices
        .iter()
        .any(|part| has_length(&part.get("inserted")) || has_length(&part.get("removed")));
    let contiguous = source_paragraphs
        .iter()
        .enumerate()
        .all(|(index, p)| ordinal_of(p, index).strict_eq(&Js::from(index)));
    let mut original_by_key: HashMap<String, Js> = HashMap::new();
    for identity in &identities {
        if identity.get("session").truthy()
            && identity.get("source").truthy()
            && identity.path(&["session", "story"]).as_str() == Some("body")
        {
            original_by_key.insert(
                mkey(&identity.path(&["session", "paraId"])),
                identity.path(&["source", "paragraphOrdinal"]),
            );
        }
    }
    let original_of = |key: &Js| original_by_key.get(&mkey(key)).cloned().unwrap_or_default();
    let source_final = source.path(&["package", "document", "finalSectionProperties"]);
    let mut section_changes = false;
    for section in session.list_sections()?.items() {
        let closing = section.get("closingParagraphId");
        let before = if closing.as_str() == Some("final") {
            source_final.clone().nullish_or(obj! {})
        } else {
            let ordinal = original_of(&closing).nullish_or(Js::Num(-1.0));
            by_ordinal
                .get(&mkey(&ordinal))
                .map(|p| p.get("sectionProperties"))
                .unwrap_or_default()
        };
        if !section_properties_equal(&before, &section.get("authored"))? {
            section_changes = true;
            break;
        }
    }
    if !section_changes {
        return Ok(Js::Undefined);
    }
    let mut edits: Vec<Js> = Vec::new();
    let body_paragraphs: Vec<Js> = output_content
        .items()
        .into_iter()
        .filter(|b| b.get("type").as_str() == Some("paragraph"))
        .collect();
    let mut source_positions: HashMap<usize, Js> = HashMap::new();
    let mut legacy = false;
    let added_boundary = Cell::new(false);
    let part_uri = identities
        .iter()
        .find(|i| {
            i.path(&["session", "story"]).as_str() == Some("body") && i.get("source").truthy()
        })
        .map(|i| i.path(&["source", "partUri"]))
        .unwrap_or_default();
    let part = match part_uri.as_str() {
        Some(uri) => strip_slash(uri).to_owned(),
        None => "word/document.xml".to_owned(),
    };
    let bytes = unzip(original_buffer)?.get(&part).cloned();
    let lossy = || -> Result<bool> {
        Ok(added_boundary.get()
            && !(match &bytes {
                Some(b) => sections_modeled(&String::from_utf8_lossy(b))?,
                None => false,
            }))
    };
    let ordinary = |message: &str| -> Result<Js> {
        if lossy()? {
            return Err(SaveError::section_unsupported(message));
        }
        Ok(obj! { "lossy" => false })
    };
    const SUPPORTED: [&str; 11] = [
        "orientation",
        "pageWidth",
        "pageHeight",
        "marginTop",
        "marginBottom",
        "marginLeft",
        "marginRight",
        "headerDistance",
        "footerDistance",
        "gutter",
        "sectionStart",
    ];
    let changed = |before: &Js, after: &Js| -> Result<Option<Js>> {
        let patch = Js::new_obj();
        let mut keys: IndexSet<String> = before.keys().into_iter().collect();
        keys.extend(after.keys());
        for key in keys {
            let a = before.get(&key);
            let b = after.get(&key);
            if section_properties_equal(&a, &b)? {
                continue;
            }
            if !SUPPORTED.contains(&key.as_str()) || b.is_undefined() {
                return Ok(None);
            }
            patch.set(&key, b);
        }
        Ok(Some(patch))
    };
    let session_key_of =
        |p: &Js| -> Result<Js> { Ok(Js::from(projected_session_key(p)?.unwrap_or_default())) };
    for block in output_content.items() {
        if block.get("type").as_str() != Some("paragraph")
            || !block.get("sectionProperties").truthy()
        {
            continue;
        }
        let key = projected_session_key(&block)?;
        let ordinal = match &key {
            Some(k) => original_of(&Js::from(k.as_str())),
            None => Js::Undefined,
        };
        let prior = if ordinal.is_undefined() {
            Js::Undefined
        } else {
            by_ordinal
                .get(&mkey(&ordinal))
                .map(|p| p.get("sectionProperties"))
                .unwrap_or_default()
        };
        if prior.truthy() && section_properties_equal(&prior, &block.get("sectionProperties"))? {
            continue;
        }
        if !prior.truthy() {
            added_boundary.set(true);
        }
        let at = index_of(&output_paragraphs, &block);
        let mut following = Js::Undefined;
        let mapped: Vec<Js> = slice(&output_paragraphs, at, None)
            .iter()
            .map(|p| Ok(original_of(&session_key_of(p)?)))
            .collect::<Result<_>>()?;
        if let Some(o) = mapped.into_iter().find(|o| !o.is_undefined()) {
            following = o;
        }
        let mut preceding = Js::Undefined;
        let mut earlier = slice(&output_paragraphs, 0, Some(at));
        earlier.reverse();
        let mapped: Vec<Js> = earlier
            .iter()
            .map(|p| Ok(original_of(&session_key_of(p)?)))
            .collect::<Result<_>>()?;
        if let Some(o) = mapped.into_iter().find(|o| !o.is_undefined()) {
            preceding = o;
        }
        let mut donor: Option<(Js, Js)> = if prior.truthy() {
            Some((ordinal.clone(), prior.clone()))
        } else {
            donors
                .iter()
                .find(|(o, _)| {
                    !following.is_undefined() && js::to_number(o) >= js::to_number(&following)
                })
                .cloned()
        };
        let base_props = donor
            .as_ref()
            .map(|(_, p)| p.clone())
            .unwrap_or_default()
            .nullish_or(source_final.clone())
            .nullish_or(obj! {});
        let mut properties = changed(&base_props, &block.get("sectionProperties"))?;
        if properties.is_none() && ordinal.is_undefined() && !preceding.is_undefined() {
            let previous = donors
                .iter()
                .find(|(o, _)| o.strict_eq(&preceding))
                .cloned();
            if let Some(previous) = previous
                && let Some(inherited) = changed(&previous.1, &block.get("sectionProperties"))? {
                    donor = Some(previous);
                    properties = Some(inherited);
                }
        }
        let Some(properties) = properties else {
            legacy = true;
            continue;
        };
        let edit = obj! {
            "paragraph" => index_of(&body_paragraphs, &block) as f64,
            "bodyParagraphs" => body_paragraphs.len(),
            "donor" => donor.as_ref().map(|(o, _)| o.clone()).unwrap_or_default().nullish_or(Js::Null),
            "properties" => properties,
        };
        edits.push(edit.clone());
        if !structural {
            source_positions.insert(
                edit.ptr().unwrap(),
                ordinal.clone().nullish_or(Js::from(at as f64)),
            );
        }
    }
    let final_props = output
        .path(&["package", "document", "finalSectionProperties"])
        .nullish_or(obj! {});
    let source_final_or_empty = source_final.clone().nullish_or(obj! {});
    if !section_properties_equal(&final_props, &source_final_or_empty)? {
        match changed(&source_final.clone().nullish_or(obj! {}), &final_props)? {
            None => legacy = true,
            Some(properties) => edits.push(
                obj! { "paragraph" => Js::Null, "donor" => Js::Null, "properties" => properties },
            ),
        }
    }
    if legacy {
        return ordinary("Cannot safely compose section patches with legacy section properties");
    }
    if edits.is_empty() {
        return Ok(Js::Undefined);
    }
    let Some(part_bytes) = bytes.clone() else {
        return ordinary("Section source part is unavailable");
    };
    let digest = hex(&Sha256::digest(part_bytes.as_slice()));
    let mut known_parts: IndexSet<String> = IndexSet::new();
    known_parts.insert(part.clone());
    for identity in &identities {
        if identity.get("source").truthy() {
            known_parts.insert(
                strip_slash(identity.path(&["source", "partUri"]).str_or_empty()).to_owned(),
            );
        }
    }
    for splice in &planned_splices {
        known_parts.insert(js::to_js_string(&splice.get("part")));
    }
    let patched_parts = plan.get("patchedParts").items();
    let only_sections = identities.iter().all(|i| {
        !i.get("session").truthy()
            || !i.get("source").is_null()
            || i.get("origin").as_str() == Some("synthetic")
    }) && section_save_metadata(&source)? == section_save_metadata(output)?
        && plan.get("assignments").is_empty()
        && patched_parts.iter().all(|p| p.get("paraIds").is_empty())
        && known_parts.iter().all(|path| {
            if patched_parts
                .iter()
                .any(|p| p.get("part").as_str() == Some(path.as_str()))
            {
                return true;
            }
            let splice = planned_splices
                .iter()
                .find(|p| p.get("part").as_str() == Some(path.as_str()));
            match splice {
                None => false,
                Some(splice) => {
                    *path == part
                        && !has_length(&splice.get("inserted"))
                        && !has_length(&splice.get("removed"))
                        && !has_length(&splice.get("tablePatches"))
                        && splice
                            .get("changed")
                            .items()
                            .iter()
                            .all(|o| includes(&splice.get("sectionOnly"), o) == Some(true))
                }
            }
        });
    let placed: Vec<Js> = if only_sections {
        edits
            .iter()
            .map(|edit| match source_positions.get(&edit.ptr().unwrap()) {
                None => edit.clone(),
                Some(paragraph) => obj! { "paragraph" => paragraph.clone(), "donor" => edit.get("donor"), "properties" => edit.get("properties") },
            })
            .collect()
    } else {
        edits.clone()
    };
    let strict = structural && !contiguous && edits.iter().any(|e| !e.get("paragraph").is_null());
    Ok(obj! {
        "plan" => obj! { "onlySections" => only_sections, "sha256" => digest, "edits" => Js::new_arr(placed) },
        "lossy" => lossy()?,
        "strict" => strict,
    })
}

/// `writeSessionSave` (src/yrs/saveYrsDocx.ts:524): `{ bytes, paragraphs, conflicts }`.
#[allow(clippy::too_many_arguments)]
pub(crate) fn write_session_save(
    session: &Session,
    document: &Js,
    capture: &Js,
    original_buffer: &Rc<Vec<u8>>,
    options: &Js,
    patches: &dyn Fn(&str) -> Result<bool>,
    skip_mutations: bool,
    splices: &dyn Fn(&str) -> Result<bool>,
    yield_save: Option<&dyn Fn(&Js) -> Result<Js>>,
) -> Result<Js> {
    let identities = capture.get("identities");
    let capture_plan = capture.get("plan");
    let output = without_synthetic_paragraphs(document, &identities)?;
    let source_buffer = session
        .source_package()
        .map(|s| s.bytes.clone())
        .unwrap_or_else(|| original_buffer.clone());
    let body_identity = identities.get("paragraphs").items().into_iter().find(|i| {
        i.path(&["session", "story"]).as_str() == Some("body") && i.get("source").truthy()
    });
    let body_part = match body_identity {
        None => "word/document.xml".to_owned(),
        Some(identity) => match identity.path(&["source", "partUri"]) {
            Js::Undefined | Js::Null => "word/document.xml".to_owned(),
            Js::Str(uri) => strip_slash(&uri).to_owned(),
            _ => throw!("partUri.replace is not a function"),
        },
    };
    let separators = note_separator_keys(document)?;
    let source_separators: RefCell<Option<Js>> = RefCell::new(None);
    let comment_deletion = options.get("commentDeletion").truthy();
    let comment_baseline = options.get("commentBaseline");
    let rewritten = |part: &str| -> Result<bool> {
        if comment_deletion {
            return Ok(false);
        }
        let Some(kind) = ["footnotes", "endnotes"]
            .into_iter()
            .find(|name| strip_slash(part) == format!("word/{name}.xml"))
        else {
            return Ok(false);
        };
        if source_separators.borrow().is_none() {
            let value = session.source_note_separators()?;
            *source_separators.borrow_mut() = Some(value);
        }
        let held = source_separators.borrow().clone().unwrap_or_default();
        Ok(!held.get(kind).strict_eq(&separators.get(kind)))
    };
    // `unspliced` only feeds a debug log; it is evaluated for its effects.
    for entry in capture_plan.get("splicedParts").items() {
        let part = js::to_js_string(&entry.get("part"));
        let spliced = splices(&part)? && !rewritten(&part)?;
        let _ = !spliced
            && (has_length(&entry.get("paragraphPatches"))
                || has_length(&entry.get("legacyParagraphs")));
    }
    let splices_for = |edits: &Js| -> Result<Vec<Js>> {
        let mut all = capture_plan.get("splicedParts").items();
        if edits.truthy() {
            all.extend(capture_plan.get("sectionSplicedParts").items());
        }
        let mut out = Vec::new();
        for entry in all {
            let part = js::to_js_string(&entry.get("part"));
            let keep = (splices(&part)? && !rewritten(&part)?)
                || (edits.truthy()
                    && part == body_part
                    && entry.get("sha256").strict_eq(&edits.get("sha256")));
            if keep {
                out.push(entry);
            }
        }
        Ok(out)
    };
    let section_plan = if comment_deletion {
        Js::Undefined
    } else {
        section_save_plan(session, &output, capture, &source_buffer)?
    };
    let refuse = |message: &str| SaveError::section_unsupported(message);
    let write = |section_edits: &Js| -> Result<(Js, Js)> {
        let model_body = |part: &str| {
            !section_plan.is_undefined() && !section_edits.truthy() && part == body_part
        };
        let spliced: Vec<Js> = splices_for(section_edits)?
            .into_iter()
            .filter(|entry| !model_body(&js::to_js_string(&entry.get("part"))))
            .collect();
        let plan = Js::new_obj();
        if comment_baseline.truthy() {
            let mut ids: IndexSet<String> = IndexSet::new();
            for comment in session
                .materialized_source()?
                .path(&["package", "document", "comments"])
                .items()
            {
                ids.insert(js::to_js_string(&comment.get("id")));
            }
            for (_, identity) in options.get("commentIdentities").map_entries() {
                ids.insert(js::to_js_string(&identity.get("id")));
            }
            let written = output.path(&["package", "document", "comments"]).items();
            let removed: Vec<Js> = ids
                .into_iter()
                .filter(|id| {
                    !written
                        .iter()
                        .any(|c| js::to_js_string(&c.get("id")) == *id)
                })
                .map(Js::from)
                .collect();
            plan.set("removedCommentIds", Js::new_arr(removed));
        }
        if section_edits.truthy() {
            plan.set("sectionEdits", section_edits.clone());
        }
        plan.set("assignments", capture_plan.get("assignments"));
        let mut patched = Vec::new();
        for entry in capture_plan.get("patchedParts").items() {
            let part = js::to_js_string(&entry.get("part"));
            if patches(&part)?
                && !rewritten(&part)?
                && (section_plan.is_undefined() || part != body_part)
            {
                patched.push(entry);
            }
        }
        plan.set("patchedParts", Js::new_arr(patched));
        if !spliced.is_empty() {
            let mut out = Vec::new();
            for entry in &spliced {
                let part = entry.get("part");
                let section_only = entry.get("sectionOnly");
                let filter = section_edits.truthy() && js::to_js_string(&part) == body_part;
                let kept = |entries: &Js, ordinal: &dyn Fn(&Js) -> Js| -> Js {
                    if filter {
                        Js::new_arr(
                            entries
                                .items()
                                .into_iter()
                                .filter(|e| includes(&section_only, &ordinal(e)) != Some(true))
                                .collect(),
                        )
                    } else {
                        entries.clone()
                    }
                };
                let o = obj! {
                    "part" => part.clone(),
                    "sha256" => entry.get("sha256"),
                    "paragraphs" => Js::new_arr(entry.get("paragraphs").items().iter().map(|p| p.at(0)).collect()),
                    "changed" => kept(&entry.get("changed"), &|o| o.clone()),
                };
                let inserted = entry.get("inserted");
                if has_length(&inserted) {
                    o.set("inserted", inserted);
                }
                let removed = entry.get("removed");
                if has_length(&removed) {
                    o.set("removed", removed);
                }
                let patches = entry.get("paragraphPatches");
                if !patches.is_undefined() {
                    o.set("paragraphPatches", kept(&patches, &|p| p.get("ordinal")));
                }
                let legacy = entry.get("legacyParagraphs");
                if !legacy.is_undefined() {
                    o.set("legacyParagraphs", kept(&legacy, &|p| p.get("ordinal")));
                }
                let tables = entry.get("tablePatches");
                if !tables.is_undefined() {
                    o.set("tablePatches", tables);
                }
                out.push(o);
            }
            plan.set("splicedParts", Js::new_arr(out));
        }
        let source: Rc<Vec<u8>> = if comment_deletion && comment_baseline.truthy() {
            Rc::new(
                comment_baseline
                    .get("bytes")
                    .as_bytes()
                    .map(|b| (**b).clone())
                    .unwrap_or_default(),
            )
        } else if comment_baseline.truthy() {
            let baseline = comment_baseline
                .get("bytes")
                .as_bytes()
                .cloned()
                .unwrap_or_default();
            let from = if section_edits.truthy() {
                &source_buffer
            } else {
                original_buffer
            };
            Rc::new(comment_source_buffer(from, &baseline)?)
        } else if section_edits.truthy() {
            source_buffer.clone()
        } else {
            original_buffer.clone()
        };
        let projected = with_source_ordinals(&output, &Js::new_arr(spliced))?;
        let result = write_document_with_rust(
            &projected,
            &source,
            options,
            &Js::Undefined,
            &Js::Undefined,
            &plan,
            skip_mutations,
            yield_save,
        )?;
        Ok((result.get("buffer"), plan))
    };
    let mut section_edits = section_plan.get("plan");
    if section_edits.truthy()
        && !section_edits.get("onlySections").truthy()
        && !splices_for(&section_edits)?
            .iter()
            .any(|e| js::to_js_string(&e.get("part")) == body_part)
    {
        if section_plan.get("lossy").truthy() {
            return Err(refuse(
                "Cannot safely compose sections with this document edit",
            ));
        }
        section_edits = Js::Undefined;
    }
    let (buffer, plan) = match write(&section_edits) {
        Ok(result) => result,
        Err(error) => {
            if !section_edits.truthy() {
                return Err(error);
            }
            if section_plan.get("lossy").truthy() {
                return Err(refuse(&error.message));
            }
            if section_plan.get("strict").truthy() {
                return Err(refuse(
                    "Cannot safely locate section boundaries after this structural edit",
                ));
            }
            write(&Js::Undefined)?
        }
    };
    let paragraphs = saved_paragraphs(&output, &identities)?;
    let bytes = buffer.as_bytes().cloned().unwrap_or_default();
    let held = crate::js::timed("written_paragraph_ids", || {
        session.written_paragraph_ids(&bytes)
    })?;
    let mut held_sets: HashMap<String, HashSet<String>> = HashMap::new();
    for (part, ids) in held.entries() {
        held_sets.insert(
            part,
            ids.items()
                .iter()
                .filter_map(|id| id.as_str().map(str::to_owned))
                .collect(),
        );
    }
    let is_written = |part_uri: &str, para_id: &str| {
        held_sets
            .get(part_uri)
            .is_some_and(|set| set.contains(&para_id.to_uppercase()))
    };
    let written: Vec<Js> = paragraphs
        .items()
        .into_iter()
        .filter(|p| {
            let persisted = p.get("persisted");
            is_written(
                &js::to_js_string(&persisted.path(&["story", "partUri"])),
                &js::to_js_string(&persisted.get("paraId")),
            )
        })
        .collect();
    let mut stale: HashSet<String> = HashSet::new();
    if !save_facade::skips_recording_saved_paragraph_ids() {
        let mut pairs: Vec<Js> = written
            .iter()
            .map(|p| {
                arr![
                    p.path(&["session", "paraId"]),
                    p.path(&["persisted", "paraId"])
                ]
            })
            .collect();
        for assignment in plan.get("assignments").items() {
            let part = js::to_js_string(&assignment.get("part"));
            let para_id = assignment.get("paraId");
            if is_written(&format!("/{part}"), &js::to_js_string(&para_id)) {
                pairs.push(arr![
                    format!("/{part}#{}", js::to_js_string(&assignment.get("ordinal"))),
                    para_id
                ]);
            }
        }
        for pair in session
            .record_saved_paragraph_ids(&Js::new_arr(pairs))?
            .items()
        {
            stale.insert(format!(
                "{}\u{0}{}",
                js::to_js_string(&pair.at(0)),
                js::to_js_string(&pair.at(1))
            ));
        }
    }
    let is_stale = |p: &Js| {
        stale.contains(&format!(
            "{}\u{0}{}",
            js::to_js_string(&p.path(&["session", "paraId"])),
            js::to_js_string(&p.path(&["persisted", "paraId"]))
        ))
    };
    let (conflicts, kept): (Vec<Js>, Vec<Js>) = written.into_iter().partition(|p| is_stale(p));
    Ok(
        obj! { "bytes" => Js::Bytes(bytes), "paragraphs" => Js::new_arr(kept), "conflicts" => Js::new_arr(conflicts) },
    )
}

/// `saveYrsDocx` (src/yrs/saveYrsDocx.ts:677).
pub(crate) fn save_yrs_docx(session: &Session, options: &Js) -> Result<Js> {
    let mut base = crate::js::timed("materialize_docx", || session.materialize_docx())?;
    if !base.get("originalBuffer").truthy() {
        throw!("saveYrsDocx requires a session opened from DOCX bytes");
    }
    let comment_baseline = options.get("commentBaseline");
    let comment_deletion = options.get("commentDeletion").truthy();
    if comment_deletion && !comment_baseline.truthy() {
        throw!("Comment deletion requires a saved comment baseline");
    }
    let original = base
        .get("originalBuffer")
        .as_bytes()
        .cloned()
        .unwrap_or_default();
    let mut capture = crate::js::timed("capture_session_save", || {
        capture_session_save(session, &Js::Undefined)
    })?;
    if comment_baseline.truthy() {
        let baseline_bytes = comment_baseline
            .get("bytes")
            .as_bytes()
            .cloned()
            .unwrap_or_default();
        let source = if comment_deletion {
            Rc::new((*baseline_bytes).clone())
        } else {
            Rc::new(comment_source_buffer(&original, &baseline_bytes)?)
        };
        let saved = parse_document_with_rust(
            &source,
            &obj! { "parseHeadersFooters" => false, "parseNotes" => false, "detectVariables" => false, "includeCanonical" => false },
        )?
        .path(&["document", "package", "document", "comments"]);
        let pkg = base.get("package");
        base = base.with(&[(
            "package",
            pkg.with(&[("document", pkg.get("document").with(&[("comments", saved)]))]),
        )]);
        let plan = capture.get("plan");
        let not_comments = |list: &Js| {
            Js::new_arr(
                list.items()
                    .into_iter()
                    .filter(|e| e.get("part").as_str() != Some(COMMENTS_PART))
                    .collect(),
            )
        };
        let spliced = plan.get("splicedParts");
        let new_plan = plan.with(&[
            ("assignments", not_comments(&plan.get("assignments"))),
            ("patchedParts", not_comments(&plan.get("patchedParts"))),
            (
                "splicedParts",
                if spliced.is_nullish() {
                    Js::Undefined
                } else {
                    not_comments(&spliced)
                },
            ),
        ]);
        capture = capture.with(&[("plan", new_plan)]);
        if comment_deletion {
            let parts: Vec<Js> = unzip(&baseline_bytes)?
                .keys()
                .filter(|part| part.ends_with(".xml") && part.as_str() != COMMENTS_PART)
                .map(|part| obj! { "part" => part.as_str(), "paraIds" => empty_arr() })
                .collect();
            capture.set(
                "plan",
                obj! { "assignments" => empty_arr(), "patchedParts" => Js::new_arr(parts) },
            );
        }
    }
    let comments = saved_comments(
        session,
        &base,
        &options.get("commentIdentities"),
        &comment_baseline.get("comments"),
    )?;
    let changed = comments.get("changed").truthy();
    if comment_baseline.truthy()
        && !changed
        && base.path(&["package", "document", "comments"]).truthy()
    {
        capture
            .path(&["plan", "patchedParts"])
            .push(obj! { "part" => COMMENTS_PART, "paraIds" => empty_arr() });
    }
    let projection_base = if changed {
        let pkg = base.get("package");
        base.with(&[(
            "package",
            pkg.with(&[(
                "document",
                pkg.get("document")
                    .with(&[("comments", comments.get("comments"))]),
            )]),
        )])
    } else {
        base.clone()
    };
    let document = crate::js::timed("yrs_to_document", || {
        yrs_to_document(
            session,
            &projection_base,
            &obj! { "commentIds" => comments.get("sessionIds") },
        )
    })?;
    let patches = |part: &str| -> Result<bool> { Ok(!changed || part != COMMENTS_PART) };
    write_session_save(
        session, &document, &capture, &original, options, &patches, false, &patches, None,
    )
}

/// The native entry point: `saveYrsDocx(session, { updateModifiedDate, modifiedBy })` with the
/// writer's determinism `now` taken from `options.now` when given.
pub(crate) fn save_docx_native(
    session: &Session,
    options: &crate::SaveOptions,
) -> Result<crate::SavedDocx> {
    struct Reset;
    impl Drop for Reset {
        fn drop(&mut self) {
            save_facade::take_save_context();
        }
    }
    save_facade::set_save_context(SaveContext {
        now: options.now.clone(),
        capture_wire: options.capture_wire,
        skip_record_saved_paragraph_ids: !options.record_saved_paragraph_ids,
        ..SaveContext::default()
    });
    let _reset = Reset;
    let save_options = obj! {
        "updateModifiedDate" => options.update_modified_date,
        "modifiedBy" => options.modified_by.clone(),
    };
    let saved = save_yrs_docx(session, &save_options)?;
    let context = save_facade::take_save_context().unwrap_or_default();
    Ok(crate::SavedDocx {
        bytes: saved
            .get("bytes")
            .as_bytes()
            .map(|b| (**b).clone())
            .unwrap_or_default(),
        paragraphs: saved.get("paragraphs"),
        conflicts: saved.get("conflicts"),
        warnings: context.warnings,
        wire: context.wire,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn sections_modeled_scans_tags() {
        // Expectations from sectionsModeled run with bun on the same strings.
        assert!(sections_modeled(r#"<w:body><w:p/><w:sectPr w:rsidR="1"><w:pgSz w:w="1"/><w:pgMar a=">"/><w:cols><w:col/></w:cols></w:sectPr></w:body>"#).unwrap());
        assert!(!sections_modeled(r#"<w:sectPr><w:pgSz/><w:formProt/></w:sectPr>"#).unwrap());
        assert!(!sections_modeled(r#"<w:sectPr><x:pgSz/></w:sectPr>"#).unwrap());
        assert!(sections_modeled(r#"<w:sectPr/><w:formProt/>"#).unwrap());
        assert!(sections_modeled(r#"<w:sectPr><w:pgSz a="unterminated/></w:sectPr>"#).unwrap());
        assert!(!sections_modeled(r#"<sectPr><pgSz/><w:cols/></sectPr>"#).unwrap());
    }

    #[test]
    fn comment_body_keys_sort_keys() {
        let body =
            Js::parse(r#"[{"b":1,"a":{"d":[{"z":1,"y":2}],"c":null},"10":0,"2":0}]"#).unwrap();
        assert_eq!(
            comment_body_key(&body).unwrap().unwrap(),
            r#"[{"2":0,"10":0,"a":{"c":null,"d":[{"y":2,"z":1}]},"b":1}]"#
        );
        assert_eq!(comment_body_key(&Js::Undefined).unwrap(), None);
    }

    #[test]
    fn map_paragraphs_keeps_identity() {
        let blocks = Js::parse(r#"[{"type":"paragraph","id":1},{"type":"table","rows":[{"cells":[{"content":[{"type":"paragraph","id":2}]}]}]}]"#).unwrap();
        let same = map_paragraphs(&blocks, &|p| Ok(p.clone())).unwrap();
        assert!(same.strict_eq(&blocks));
        let dropped = map_paragraphs(&blocks, &|p| {
            Ok(if p.get("id").as_f64() == Some(2.0) {
                Js::Null
            } else {
                p.clone()
            })
        })
        .unwrap();
        assert!(!dropped.strict_eq(&blocks));
        assert!(dropped.at(0).strict_eq(&blocks.at(0)));
        assert_eq!(
            dropped
                .at(1)
                .path(&["rows"])
                .at(0)
                .get("cells")
                .at(0)
                .get("content")
                .len(),
            0
        );
        assert_eq!(
            blocks
                .at(1)
                .get("rows")
                .at(0)
                .get("cells")
                .at(0)
                .get("content")
                .len(),
            1
        );
    }
}
