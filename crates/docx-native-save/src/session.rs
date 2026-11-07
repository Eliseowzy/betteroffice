//! The `YrsSession` facade the TypeScript save reads, over a native [`EditSessionCore`].
//!
//! Each method returns what the facade's method of the same name returns in
//! `packages/docx/src/yrs/yrsSessionFacade.ts`, as a [`Js`] value. The session internals
//! (`sessionInternals.ts`) are the [`Session::source_package`], [`Session::materialized_source`],
//! [`Session::source_note_separators`] and [`Session::header_footer_aliases`] methods.

use std::cell::RefCell;
use std::collections::HashMap;
use std::rc::Rc;

use docx_edit::RunInheritance;
use docx_edit::session::{EditSessionCore, EngineError};
use indexmap::IndexMap;

use crate::editor_save_keys::{editor_save_keys, note_separator_keys};
use crate::js::{Js, Result, SaveError};
use crate::parse_facade::decode_s9_envelope;

fn engine<T>(result: std::result::Result<T, EngineError>) -> Result<T> {
    result.map_err(|e| SaveError::new(e.to_string()))
}

/// What the facade keeps about the package a session was opened from.
#[derive(Clone)]
pub struct OpenedSource {
    /// The source package bytes (`docxSource`).
    pub bytes: Rc<Vec<u8>>,
    /// `editorSaveKeys(host.document)` of the opened host document (`docxSourceKeys`).
    pub keys: Js,
}

impl OpenedSource {
    /// The facade's bookkeeping after `openDocx(bytes, …)` or `bootstrapPeer(…, bytes, host)`:
    /// `open_json` is what `open_docx` returned for `bytes`.
    pub fn after_open(bytes: Vec<u8>, open_json: &str) -> Result<Self> {
        let bytes = Rc::new(bytes);
        let host = crate::parse_facade::decode_docx_host(open_json, &bytes)?;
        let document = host.get("document");
        document.set("originalBuffer", Js::Bytes(bytes.clone()));
        let keys = editor_save_keys(&document, &Js::Undefined)?;
        Ok(Self { bytes, keys })
    }
}

/// A session being saved.
pub struct Session<'a> {
    pub core: &'a EditSessionCore,
    source: Option<OpenedSource>,
    materialized: RefCell<Option<Js>>,
    separators: RefCell<Option<Js>>,
    media: RefCell<HashMap<String, Option<String>>>,
    aliases: RefCell<Option<IndexMap<String, String>>>,
}

impl<'a> Session<'a> {
    pub fn new(core: &'a EditSessionCore, source: Option<OpenedSource>) -> Self {
        Self {
            core,
            source,
            materialized: RefCell::new(None),
            separators: RefCell::new(None),
            media: RefCell::new(HashMap::new()),
            aliases: RefCell::new(None),
        }
    }

    pub fn story_segments(&self, story: &str) -> Result<Js> {
        Js::parse(&engine(self.core.story_segments(story))?)
    }

    pub fn story_paragraph_ids(&self, story: &str) -> Result<Js> {
        Js::parse(&engine(self.core.story_paragraph_ids(story))?)
    }

    pub fn story_ids(&self) -> Vec<String> {
        self.core.story_ids()
    }

    pub fn has_story(&self, story: &str) -> bool {
        self.core.has_story(story)
    }

    pub fn list_comments(&self) -> Result<Js> {
        Js::parse(&engine(self.core.list_comments())?)
    }

    pub fn resolve_comment(&self, comment_id: &str) -> Result<Js> {
        Js::parse(&engine(self.core.resolve_comment(comment_id))?)
    }

    /// `listSections()`: `[]` without a body story.
    pub fn list_sections(&self) -> Result<Js> {
        if !self.has_story("body") {
            return Ok(Js::new_arr(Vec::new()));
        }
        Js::parse(&engine(self.core.list_sections_json())?)
    }

    /// `mediaDataUrl(token)`: the `data:` URL of a `media:{n}` source, cached per token.
    pub fn media_data_url(&self, token: &str) -> Option<String> {
        if !token.starts_with("media:") {
            return None;
        }
        if let Some(url) = self.media.borrow().get(token) {
            return url.clone();
        }
        let url = self.core.media_data_url(token);
        self.media
            .borrow_mut()
            .insert(token.to_owned(), url.clone());
        url
    }

    pub fn paragraph_identities(&self) -> Result<Js> {
        Js::parse(&engine(self.core.paragraph_identities())?)
    }

    /// `paragraphSavePlan()`, refusing a plan without the producer capability.
    pub fn paragraph_save_plan(&self) -> Result<Js> {
        let plan = Js::parse(&engine(self.core.paragraph_save_plan())?)?;
        for part in plan.get("splicedParts").items() {
            if part.get("changed").is_empty() {
                continue;
            }
            if !part.get("paragraphPatches").truthy() || !part.get("legacyParagraphs").truthy() {
                crate::throw!(
                    "DOCX_MINIMAL_PATCH [INVALID_PLAN] {} p#?: missing producer capability",
                    part.get("part").str_or_empty()
                );
            }
        }
        Ok(plan)
    }

    /// `recordSavedParagraphIds(saved)`: the stale `[owner, paraId]` pairs.
    pub fn record_saved_paragraph_ids(&self, saved: &Js) -> Result<Js> {
        let json = saved.stringify().unwrap_or_else(|| "[]".into());
        Js::parse(&engine(self.core.record_saved_paragraph_ids(&json))?)
    }

    pub fn written_paragraph_ids(&self, bytes: &[u8]) -> Result<Js> {
        Js::parse(&engine(self.core.written_paragraph_ids(bytes))?)
    }

    pub fn export_structured(&self, options: &Js) -> Result<Js> {
        let json = options.stringify().unwrap_or_else(|| "{}".into());
        Js::parse(&engine(self.core.export_structured_json(&json))?)
    }

    /// `runInheritance(styles)`.
    pub fn run_inheritance(&self, styles: &Js) -> Inheritance {
        let value = if styles.is_undefined() {
            serde_json::Value::Null
        } else {
            styles.to_serde()
        };
        Inheritance(RefCell::new(RunInheritance::new(Some(&value))))
    }

    /// `materializeDocx()`: a copy of the source document with its `originalBuffer`.
    pub fn materialize_docx(&self) -> Result<Js> {
        let Some(source) = &self.source else {
            return Ok(Js::Null);
        };
        let document = self.materialized_source()?;
        if document.is_nullish() {
            return Ok(Js::Null);
        }
        let copy = document.with(&[("originalBuffer", Js::Undefined)]).deep_clone();
        copy.set("originalBuffer", Js::Bytes(source.bytes.clone()));
        Ok(copy)
    }

    // ----- sessionInternals -----

    /// `sessionMaterializedSource(session)`: the source package parsed whole, cached.
    pub fn materialized_source(&self) -> Result<Js> {
        let Some(source) = &self.source else {
            return Ok(Js::Null);
        };
        if let Some(document) = self.materialized.borrow().as_ref() {
            return Ok(document.clone());
        }
        let Some(json) = engine(self.core.materialize_docx())? else {
            return Ok(Js::Null);
        };
        let document = decode_s9_envelope(&json, &source.bytes)?.get("document");
        *self.materialized.borrow_mut() = Some(document.clone());
        Ok(document)
    }

    /// `sessionSourcePackage(session)`: `{ buffer, keys }` or `None`.
    pub fn source_package(&self) -> Option<&OpenedSource> {
        self.source.as_ref()
    }

    /// `sessionSourceNoteSeparators(session)`.
    pub fn source_note_separators(&self) -> Result<Js> {
        if self.source.is_none() {
            return Ok(Js::Null);
        }
        if let Some(keys) = self.separators.borrow().as_ref() {
            return Ok(keys.clone());
        }
        let full = self.materialized_source()?;
        if full.is_nullish() {
            return Ok(Js::Null);
        }
        let keys = note_separator_keys(&full)?;
        *self.separators.borrow_mut() = Some(keys.clone());
        Ok(keys)
    }

    /// The header and footer alias relationship ids and their canonical ids.
    pub fn header_footer_aliases(&self) -> IndexMap<String, String> {
        if let Some(aliases) = self.aliases.borrow().as_ref() {
            return aliases.clone();
        }
        let mut aliases = IndexMap::new();
        if let Ok(value) = Js::parse(&self.core.header_footer_aliases_json()) {
            let entries = value.entries();
            if value.is_obj() && entries.iter().all(|(_, v)| v.is_str()) {
                for (k, v) in entries {
                    aliases.insert(k, v.str_or_empty().to_owned());
                }
            }
        }
        *self.aliases.borrow_mut() = Some(aliases.clone());
        aliases
    }

    /// `headerFooterStory(session, rId)`.
    pub fn header_footer_story(&self, r_id: &str) -> String {
        match self.header_footer_aliases().get(r_id) {
            Some(canonical) if !canonical.is_empty() => format!("hf:{canonical}"),
            _ => format!("hf:{r_id}"),
        }
    }
}

/// `YrsRunInheritance` of `runInheritance(styles)`.
pub struct Inheritance(RefCell<RunInheritance>);

impl Inheritance {
    /// `attributes(paragraphStyle, runStyle, tableRun)`.
    pub fn attributes(
        &self,
        paragraph_style: Option<&str>,
        run_style: Option<&str>,
        table_run: Option<&Js>,
    ) -> Js {
        let mut inner = self.0.borrow_mut();
        let attrs = match table_run {
            Some(table_run) => inner.table_run_attrs(paragraph_style, run_style, table_run.to_serde()),
            None => inner.run_attrs(paragraph_style, run_style),
        };
        Js::from_serde(&serde_json::to_value(&attrs).unwrap_or(serde_json::Value::Null))
    }

    /// `paragraphFormatting(paragraphStyle, tableFormatting)`; `undefined` for none.
    pub fn paragraph_formatting(&self, paragraph_style: Option<&str>, table_formatting: &Js) -> Js {
        let table_formatting = if table_formatting.is_nullish() {
            None
        } else {
            Some(table_formatting.to_serde())
        };
        let formatting = self.0.borrow_mut().paragraph_formatting(paragraph_style, table_formatting);
        let json = formatting.unwrap_or(serde_json::Value::Null);
        if json.is_null() {
            Js::Undefined
        } else {
            Js::from_serde(&json)
        }
    }

    /// `tableAttributes(table)`.
    pub fn table_attributes(&self, table: &Js) -> Js {
        let reduced = crate::obj! {
            "formatting" => table.get("formatting"),
            "columnWidths" => table.get("columnWidths"),
            "rows" => Js::new_arr(table.get("rows").items().iter().map(|row| crate::obj! {
                "formatting" => row.get("formatting"),
                "cells" => Js::new_arr(row.get("cells").items().iter().map(|cell| crate::obj! {
                    "formatting" => cell.get("formatting"),
                }).collect()),
            }).collect()),
        };
        Js::from_serde(&self.0.borrow().table_attrs(&reduced.to_serde()))
    }
}
