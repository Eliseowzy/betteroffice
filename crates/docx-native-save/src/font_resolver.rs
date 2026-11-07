//! Port of `packages/docx/src/utils/fontResolver.ts`.
#![allow(unused_variables, unused_imports, dead_code, clippy::all)]

use std::collections::{HashMap, HashSet};
use std::rc::Rc;
use indexmap::IndexMap;
use crate::js::{Js, Result, SaveError};
use crate::session::{Inheritance, Session};
use crate::{arr, obj, throw};

/// `detectFontCategory` (src/utils/fontResolver.ts:354).
pub(crate) fn detect_font_category(font_name: &str) -> Result<Js> {
    todo!("detectFontCategory")
}

/// `resolveFontFamily` (src/utils/fontResolver.ts:446).
pub(crate) fn resolve_font_family(docx_font_name: &str) -> Result<Js> {
    todo!("resolveFontFamily")
}

/// `escapeQuotedFontName` (src/utils/fontResolver.ts:484).
pub(crate) fn escape_quoted_font_name(font_name: &str) -> Result<String> {
    todo!("escapeQuotedFontName")
}

/// `quoteFontName` (src/utils/fontResolver.ts:491).
pub(crate) fn quote_font_name(font_name: &str) -> Result<String> {
    todo!("quoteFontName")
}

/// `resolveThemeFont` (src/utils/fontResolver.ts:516).
pub(crate) fn resolve_theme_font(theme_ref: &str, font_scheme: &Js) -> Result<Option<String>> {
    todo!("resolveThemeFont")
}

/// `getGoogleFontsToLoad` (src/utils/fontResolver.ts:550).
pub(crate) fn get_google_fonts_to_load(docx_fonts: &[String]) -> Result<Vec<String>> {
    todo!("getGoogleFontsToLoad")
}

/// `buildFontFamilyString` (src/utils/fontResolver.ts:570).
pub(crate) fn build_font_family_string(fonts: &[String], category: &Js) -> Result<String> {
    todo!("buildFontFamilyString")
}

/// `getGoogleFontEquivalent` (src/utils/fontResolver.ts:591).
pub(crate) fn get_google_font_equivalent(docx_font_name: &str) -> Result<Option<String>> {
    todo!("getGoogleFontEquivalent")
}

/// `hasGoogleFontEquivalent` (src/utils/fontResolver.ts:602).
pub(crate) fn has_google_font_equivalent(docx_font_name: &str) -> Result<bool> {
    todo!("hasGoogleFontEquivalent")
}
