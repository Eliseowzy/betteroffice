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
