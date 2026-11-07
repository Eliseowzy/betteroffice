//! Native save of a `docx-edit` session: the TypeScript `saveYrsDocx` of `@betteroffice/docx`
//! ported to Rust, so a host without a JavaScript runtime writes the edited `.docx` itself.
//!
//! The port follows the TypeScript module by module (each file names its source) over
//! [`js::Js`], a value type with JavaScript semantics, and writes through the same
//! `docx-parse` package writer the TypeScript save calls through WebAssembly.

pub mod js;
pub mod session;

mod content_control_values;
mod editor_save_keys;
mod font_family_merge;
mod font_resolver;
mod header_footer_alias_projection;
mod host_media;
mod paragraph_formatting_merge;
mod parse_facade;
mod raw_xml;
mod rels_parser;
mod resident_comment_projection;
mod rezip_parts;
mod save;
mod save_facade;
mod save_formatting;
mod sdt_attributes;
mod sections;
mod session_revision_ids;
mod story_plain_text;
mod style_resolver;
mod table_row_spans;
mod text_formatting_merge;
mod tiff_media;
mod tracked_control_content;
mod units;
mod yrs_to_document;

pub use js::SaveError;
pub use session::OpenedSource;

use docx_edit::session::EditSessionCore;

/// How [`save_docx`] writes the package.
#[derive(Debug, Clone)]
pub struct SaveOptions {
    /// Stamp `dcterms:modified` (the TS `updateModifiedDate`, default true).
    pub update_modified_date: bool,
    /// `cp:lastModifiedBy` (the TS `modifiedBy`).
    pub modified_by: Option<String>,
    /// The save's clock as an ISO 8601 UTC string (`new Date().toISOString()`); now when `None`.
    pub now: Option<String>,
    /// Keep the JSON request handed to the package writer in [`SavedDocx::wire`].
    pub capture_wire: bool,
}

impl Default for SaveOptions {
    fn default() -> Self {
        Self {
            update_modified_date: true,
            modified_by: None,
            now: None,
            capture_wire: false,
        }
    }
}

/// What [`save_docx`] wrote: `DocxSavedDocument` of the TS save.
#[derive(Debug, Clone)]
pub struct SavedDocx {
    pub bytes: Vec<u8>,
    /// `paragraphs`: each saved session paragraph's session and persisted anchors.
    pub paragraphs: js::Js,
    /// `conflicts`: saved paragraphs whose Word paragraph ID the live session reassigned.
    pub conflicts: js::Js,
    pub warnings: Vec<String>,
    /// The writer request, when [`SaveOptions::capture_wire`].
    pub wire: Option<String>,
}

/// Saves `core`, opened from `source`, as `saveYrsDocx(session)` does: unchanged parts keep their
/// source bytes, and the Word paragraph IDs the save writes are recorded in the session.
pub fn save_docx(
    core: &EditSessionCore,
    source: &OpenedSource,
    options: &SaveOptions,
) -> Result<SavedDocx, SaveError> {
    let session = session::Session::new(core, Some(source.clone()));
    save::save_docx_native(&session, options)
}
