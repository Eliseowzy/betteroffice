//! Port of `packages/docx/src/yrs/yrsToDocument.ts` (part).
#![allow(unused_variables, unused_imports, dead_code, clippy::all)]

use super::*;
use crate::js::{self, Js, Result, SaveError};
use crate::session::{Inheritance, Session};
use crate::{arr, obj, throw};
use indexmap::IndexMap;
use std::cell::Cell;
use std::collections::{HashMap, HashSet};
use std::rc::Rc;

use crate::content_control_values::apply_content_control_value;
use crate::font_family_merge::merge_font_family;
use crate::font_resolver::resolve_theme_font;
use crate::sdt_attributes::sdt_attrs_to_props;
use crate::session_revision_ids::mark_revision_info;
use crate::tracked_control_content::has_tracked_control_content;
use crate::units::pixels_to_emu;

type Format<'a> = &'a dyn Fn(&Js) -> Result<Js>;

// ----- constants -----

fn defaults(entries: &[(&str, Js)]) -> Js {
    let o = Js::new_obj();
    for (k, v) in entries {
        o.set(k, v.clone());
    }
    o
}

const PARAGRAPH_ATTR_KEYS: &[&str] = &[
    "paraId",
    "textId",
    "alignment",
    "spaceBefore",
    "spaceAfter",
    "lineSpacing",
    "lineSpacingRule",
    "spacingExplicit",
    "indentLeft",
    "indentRight",
    "indentFirstLine",
    "hangingIndent",
    "numPr",
    "numPrFromStyle",
    "listNumFmt",
    "listIsBullet",
    "listMarker",
    "listMarkerHidden",
    "listMarkerFontFamily",
    "listMarkerFontSize",
    "listMarkerSuffix",
    "listLevelNumFmts",
    "listAbstractNumId",
    "listStartOverride",
    "styleId",
    "borders",
    "shading",
    "tabs",
    "pageBreakBefore",
    "renderedPageBreakBefore",
    "keepNext",
    "keepLines",
    "widowControl",
    "contextualSpacing",
    "snapToGrid",
    "autoSpaceDE",
    "autoSpaceDN",
    "adjustRightInd",
    "cnfStyle",
    "defaultTextFormatting",
    "sectionBreakType",
    "bidi",
    "outlineLevel",
    "bookmarks",
    "_originalFormatting",
    "_originalRunBoundaries",
    "_sectionProperties",
    "pPrIns",
    "pPrDel",
    "pPrChange",
];

thread_local! {
    static PARAGRAPH_DEFAULTS: Js = {
        let o = Js::new_obj();
        for k in PARAGRAPH_ATTR_KEYS {
            o.set(k, if *k == "hangingIndent" { Js::Bool(false) } else { Js::Null });
        }
        o
    };
    static TABLE_DEFAULTS: Js = defaults(&[
        ("styleId", Js::Null), ("width", Js::Null), ("widthType", Js::Null),
        ("justification", Js::Null), ("columnWidths", Js::Null), ("tableLayout", Js::Null),
        ("floating", Js::Null), ("cellMargins", Js::Null), ("look", Js::Null), ("bidi", Js::Null),
        ("_originalFormatting", Js::Null), ("tblPrChange", Js::Null),
    ]);
    static TABLE_ROW_DEFAULTS: Js = defaults(&[
        ("height", Js::Null), ("heightRule", Js::Null), ("isHeader", Js::Bool(false)),
        ("_originalFormatting", Js::Null), ("trIns", Js::Null), ("trDel", Js::Null),
        ("trPrChange", Js::Null),
    ]);
    static TABLE_CELL_DEFAULTS: Js = defaults(&[
        ("colspan", Js::Num(1.0)), ("rowspan", Js::Num(1.0)), ("colwidth", Js::Null),
        ("width", Js::Null), ("widthType", Js::Null), ("verticalAlign", Js::Null),
        ("backgroundColor", Js::Null), ("borders", Js::Null), ("margins", Js::Null),
        ("textDirection", Js::Null), ("noWrap", Js::Bool(false)), ("_originalFormatting", Js::Null),
        ("_originalResolvedFill", Js::Null), ("cellMarker", Js::Null), ("tcPrChange", Js::Null),
    ]);
}

/// `PARAGRAPH_ATTR_DEFAULTS` (one shared object, as the TS module constant).
pub(crate) fn paragraph_attr_defaults() -> Js {
    PARAGRAPH_DEFAULTS.with(|v| v.clone())
}
/// `TABLE_ATTR_DEFAULTS`.
pub(crate) fn table_attr_defaults() -> Js {
    TABLE_DEFAULTS.with(|v| v.clone())
}
/// `TABLE_ROW_ATTR_DEFAULTS`.
pub(crate) fn table_row_attr_defaults() -> Js {
    TABLE_ROW_DEFAULTS.with(|v| v.clone())
}
/// `TABLE_CELL_ATTR_DEFAULTS`.
pub(crate) fn table_cell_attr_defaults() -> Js {
    TABLE_CELL_DEFAULTS.with(|v| v.clone())
}

/// `BOOLEAN_MARKS`.
pub(crate) const BOOLEAN_MARKS: &[&str] = &[
    "bold",
    "italic",
    "superscript",
    "subscript",
    "allCaps",
    "smallCaps",
    "emboss",
    "imprint",
    "textShadow",
    "textOutline",
    "hidden",
    "rtl",
];

/// `NON_FORMATTING_ATTRS`.
pub(crate) const NON_FORMATTING_ATTRS: &[&str] = &["hyperlink", "ins", "del"];

/// `ON_OFF_RUN_PROPERTIES`.
pub(crate) const ON_OFF_RUN_PROPERTIES: &[&str] = &[
    "bold",
    "italic",
    "allCaps",
    "smallCaps",
    "strike",
    "doubleStrike",
    "outline",
    "shadow",
    "emboss",
    "imprint",
    "hidden",
    "rtl",
    "boldCs",
    "italicCs",
];

/// `SOURCE_RUN_PROPERTIES`.
pub(crate) const SOURCE_RUN_PROPERTIES: &[&str] = &[
    "bold",
    "italic",
    "fontSize",
    "boldCs",
    "italicCs",
    "fontSizeCs",
    "language",
    "fontFamily",
];

/// `SIMPLE_JSON_KEY.test(key)`: `/^[A-Za-z0-9_$]+$/`.
pub(crate) fn is_simple_json_key(key: &str) -> bool {
    !key.is_empty()
        && key
            .bytes()
            .all(|b| b.is_ascii_alphanumeric() || b == b'_' || b == b'$')
}

/// `SPECIAL_RUN_TEXT.test(text)`: `/[­‑-]/`.
fn has_special_run_text(text: &str) -> bool {
    text.chars()
        .any(|c| c == '\u{ad}' || c == '\u{2011}' || ('\u{f000}'..='\u{f8ff}').contains(&c))
}

// ----- projected media (module-level `let projectedMedia`) -----

thread_local! {
    static PROJECTED_MEDIA: Cell<*const ()> = const { Cell::new(std::ptr::null()) };
}

/// Restores the previous `projectedMedia` when dropped.
pub(crate) struct ProjectedMediaGuard(*const ());

impl Drop for ProjectedMediaGuard {
    fn drop(&mut self) {
        PROJECTED_MEDIA.with(|cell| cell.set(self.0));
    }
}

/// `projectedMedia = (token) => session.mediaDataUrl(token)` for the guard's lifetime.
pub(crate) fn project_media_with(session: &Session) -> ProjectedMediaGuard {
    let pointer = session as *const Session as *const ();
    ProjectedMediaGuard(PROJECTED_MEDIA.with(|cell| cell.replace(pointer)))
}

/// `projectedMedia?.(token)`: `None` without a projection under way (`Some(None)` for a token
/// the session cannot resolve).
fn projected_media(token: &str) -> Option<Option<String>> {
    let pointer = PROJECTED_MEDIA.with(|cell| cell.get());
    if pointer.is_null() {
        return None;
    }
    // SAFETY: the pointer is set only by `project_media_with` from a live `&Session`, and the
    // returned guard (which borrows nothing but is dropped before the borrow ends in every
    // caller's scope) restores the previous value; it is never dereferenced after that.
    let session = unsafe { &*(pointer as *const Session) };
    Some(session.media_data_url(token))
}

// ----- chart recovery -----

/// `rId`/`path` as the template strings of `ChartRecovery` see them (`undefined` → `None`).
fn recovery_key(value: &Js) -> Option<String> {
    if value.is_undefined() {
        None
    } else {
        Some(js::to_js_string(value))
    }
}

// ----- small helpers -----

fn as_obj_value(value: &Js) -> Option<Js> {
    if matches!(value, Js::Obj(_) | Js::Map(_) | Js::Bytes(_)) {
        Some(value.clone())
    } else {
        None
    }
}

fn opt_str(value: Option<String>) -> Js {
    value.map_or(Js::Undefined, Js::from)
}

fn type_is(value: &Js, ty: &str) -> bool {
    value.get("type").as_str() == Some(ty)
}

/// `s.replace('#', '')`: the first `#` only.
fn replace_first_hash(s: &str) -> String {
    s.replacen('#', "", 1)
}

/// `str.match(/rotate\(([-\d.]+)deg\)/)?.[1]`.
fn match_rotation(s: &str) -> Option<String> {
    let mut from = 0;
    while let Some(found) = s[from..].find("rotate(") {
        let start = from + found + "rotate(".len();
        let rest = &s[start..];
        let len = rest
            .bytes()
            .take_while(|b| b.is_ascii_digit() || *b == b'-' || *b == b'.')
            .count();
        if len > 0 && rest[len..].starts_with("deg)") {
            return Some(rest[..len].to_owned());
        }
        from = from + found + 1;
    }
    None
}

fn on_off(key: &str) -> bool {
    ON_OFF_RUN_PROPERTIES.contains(&key)
}

// ----- ported functions -----

/// `asObject` (src/yrs/yrsToDocument.ts:282).
pub(crate) fn as_object(value: &Js) -> Result<Js> {
    Ok(as_obj_value(value).unwrap_or(Js::Undefined))
}

/// `asFiniteNumber` (src/yrs/yrsToDocument.ts:288).
pub(crate) fn as_finite_number(value: &Js) -> Result<Option<f64>> {
    Ok(match value {
        Js::Num(n) if n.is_finite() => Some(*n),
        _ => None,
    })
}

/// `asString` (src/yrs/yrsToDocument.ts:292).
pub(crate) fn as_string(value: &Js) -> Result<Option<String>> {
    Ok(value.as_str().map(str::to_owned))
}

/// `dropNulls` (src/yrs/yrsToDocument.ts:296).
pub(crate) fn drop_nulls(value: &Js) -> Result<Js> {
    if value.is_arr() {
        let items = value
            .items()
            .iter()
            .map(drop_nulls)
            .collect::<Result<Vec<_>>>()?;
        return Ok(Js::new_arr(items));
    }
    let Some(object) = as_obj_value(value) else {
        return Ok(value.clone());
    };
    let result = Js::new_obj();
    for (key, entry) in object.entries() {
        if entry.is_nullish() {
            continue;
        }
        result.set(&key, drop_nulls(&entry)?);
    }
    Ok(result)
}

/// `stableStringify` (src/yrs/yrsToDocument.ts:312).
pub(crate) fn stable_stringify(value: &Js) -> Result<String> {
    match value {
        Js::Undefined | Js::Null => Ok("null".into()),
        Js::Arr(_) => {
            let mut result = String::from("[");
            for (index, item) in value.items().iter().enumerate() {
                if index > 0 {
                    result.push(',');
                }
                result.push_str(&stable_stringify(item)?);
            }
            result.push(']');
            Ok(result)
        }
        Js::Obj(_) | Js::Map(_) | Js::Bytes(_) => stable_stringify_object(value, None),
        Js::Num(n) => Ok(if n.is_finite() {
            js::num_to_string(*n)
        } else {
            "null".into()
        }),
        Js::Bool(b) => Ok(b.to_string()),
        Js::Str(_) => Ok(value.stringify().unwrap_or_default()),
    }
}

/// `stableStringifyObject` (src/yrs/yrsToDocument.ts:331).
pub(crate) fn stable_stringify_object(
    value: &Js,
    omitted: Option<&HashSet<String>>,
) -> Result<String> {
    let mut entries: Vec<(String, Js)> = match value {
        Js::Bytes(b) => b
            .iter()
            .enumerate()
            .map(|(i, v)| (i.to_string(), Js::Num(*v as f64)))
            .collect(),
        _ => value.entries(),
    };
    entries.sort_by(|a, b| js::cmp16(&a.0, &b.0));
    let mut result = String::from("{");
    let mut separator = "";
    for (key, entry) in entries {
        if omitted.is_some_and(|o| o.contains(&key)) {
            continue;
        }
        let quoted = if is_simple_json_key(&key) {
            format!("\"{key}\"")
        } else {
            Js::from(key.as_str()).stringify().unwrap_or_default()
        };
        result.push_str(separator);
        result.push_str(&quoted);
        result.push(':');
        result.push_str(&stable_stringify(&entry)?);
        separator = ",";
    }
    result.push('}');
    Ok(result)
}

/// `stableStringifyWithout` (src/yrs/yrsToDocument.ts:344).
pub(crate) fn stable_stringify_without(value: &Js, omitted: &HashSet<String>) -> Result<String> {
    stable_stringify_object(value, Some(omitted))
}

/// `formattingAttrs` (src/yrs/yrsToDocument.ts:348).
pub(crate) fn formatting_attrs(attributes: &Js) -> Result<Js> {
    let result = attributes.spread();
    result.delete("hyperlink");
    result.delete("ins");
    result.delete("del");
    Ok(result)
}

/// `attrsToTextFormatting` (src/yrs/yrsToDocument.ts:357).
pub(crate) fn attrs_to_text_formatting(attributes: &Js) -> Result<Js> {
    let formatting = Js::new_obj();
    let a = |k: &str| attributes.get(k);

    if a("bold").truthy() {
        formatting.set("bold", true);
    }
    if a("italic").truthy() {
        formatting.set("italic", true);
    }
    if a("boldCs").truthy() {
        formatting.set("boldCs", true);
    }
    if a("italicCs").truthy() {
        formatting.set("italicCs", true);
    }

    if let Some(underline) = as_obj_value(&a("underline")) {
        if !underline.get("inheritedHyperlink").is_true() {
            let style = as_string(&underline.get("style"))?
                .filter(|s| !s.is_empty())
                .unwrap_or_else(|| "single".into());
            formatting.set(
                "underline",
                obj! {
                    "style" => style,
                    "color" => as_obj_value(&underline.get("color")).unwrap_or(Js::Null),
                },
            );
        }
    }

    if let Some(strike) = as_obj_value(&a("strike")) {
        if strike.get("double").truthy() {
            formatting.set("doubleStrike", true);
        } else {
            formatting.set("strike", true);
        }
    }

    if let Some(color) = as_obj_value(&a("textColor")) {
        if !color.get("inheritedHyperlink").is_true() {
            let s = |k: &str| -> Result<Js> {
                Ok(as_string(&color.get(k))?.map_or(Js::Null, Js::from))
            };
            formatting.set(
                "color",
                obj! {
                    "rgb" => s("rgb")?,
                    "themeColor" => color.get("themeColor").nullish_or(Js::Null),
                    "themeTint" => s("themeTint")?,
                    "themeShade" => s("themeShade")?,
                },
            );
        }
    }

    let highlight = match as_string(&a("highlight"))? {
        Some(h) => Some(h),
        None => as_string(&as_object(&a("highlight"))?.get("color"))?,
    };
    if let Some(h) = highlight {
        formatting.set("highlight", h);
    }

    if let Some(font_size) = as_obj_value(&a("fontSize")) {
        let size = as_finite_number(&font_size.get("size"))?;
        let size_cs = as_finite_number(&font_size.get("sizeCs"))?;
        if let Some(v) = size {
            formatting.set("fontSize", v);
        }
        if let Some(v) = size_cs {
            formatting.set("fontSizeCs", v);
        }
    }

    if let Some(ff) = as_obj_value(&a("fontFamily")) {
        let s = |k: &str| -> Result<Option<String>> { as_string(&ff.get(k)) };
        formatting.set(
            "fontFamily",
            obj! {
                "ascii" => s("ascii")?.map_or(Js::Null, Js::from),
                "hAnsi" => s("hAnsi")?.map_or(Js::Null, Js::from),
                "eastAsia" => opt_str(s("eastAsia")?),
                "cs" => opt_str(s("cs")?),
                "asciiTheme" => ff.get("asciiTheme").nullish_or(Js::Null),
                "hAnsiTheme" => opt_str(s("hAnsiTheme")?),
                "eastAsiaTheme" => opt_str(s("eastAsiaTheme")?),
                "csTheme" => opt_str(s("csTheme")?),
                "hint" => opt_str(s("hint")?),
            },
        );
    }

    if a("superscript").truthy() {
        formatting.set("vertAlign", "superscript");
    }
    if a("subscript").truthy() {
        formatting.set("vertAlign", "subscript");
    }
    if a("allCaps").truthy() {
        formatting.set("allCaps", true);
    }
    if a("smallCaps").truthy() {
        formatting.set("smallCaps", true);
    }

    if let Some(spacing) = as_obj_value(&a("characterSpacing")) {
        let char_spacing = as_finite_number(&spacing.get("spacing"))?;
        let position = as_finite_number(&spacing.get("position"))?;
        let scale = as_finite_number(&spacing.get("scale"))?;
        let kerning = as_finite_number(&spacing.get("kerning"))?;
        if let Some(v) = char_spacing {
            formatting.set("spacing", v);
        }
        if let Some(v) = position {
            formatting.set("position", v);
        }
        if let Some(v) = scale {
            formatting.set("scale", v);
        }
        if let Some(v) = kerning {
            formatting.set("kerning", v);
        }
    }

    if a("emboss").truthy() {
        formatting.set("emboss", true);
    }
    if a("imprint").truthy() {
        formatting.set("imprint", true);
    }
    if a("textShadow").truthy() {
        formatting.set("shadow", true);
    }

    if let Some(emphasis) = as_obj_value(&a("emphasisMark")) {
        let t = as_string(&emphasis.get("type"))?
            .filter(|s| !s.is_empty())
            .unwrap_or_else(|| "dot".into());
        formatting.set("emphasisMark", t);
    }
    if a("textOutline").truthy() {
        formatting.set("outline", true);
    }
    if a("hidden").truthy() {
        formatting.set("hidden", true);
    }
    if a("rtl").truthy() {
        formatting.set("rtl", true);
    }

    if let Some(effect) = as_obj_value(&a("textEffect")) {
        let e = as_string(&effect.get("effect"))?
            .filter(|s| !s.is_empty())
            .unwrap_or_else(|| "blinkBackground".into());
        formatting.set("effect", e);
    }

    let modern = as_object(&a("modernTextEffects"))?;
    if modern.get("effects").truthy() {
        formatting.set("modernEffects", modern.get("effects"));
    }

    let run_style = as_object(&a("runStyle"))?;
    if let Some(style_id) = as_string(&run_style.get("styleId"))? {
        if !style_id.is_empty() {
            formatting.set("styleId", style_id);
        }
    }

    Ok(formatting)
}

/// `directFormatting` (src/yrs/yrsToDocument.ts:473).
pub(crate) fn direct_formatting(attributes: &Js) -> Result<Js> {
    attrs_to_text_formatting(&formatting_attrs(attributes)?)
}

/// `lentByLinkStyle` (src/yrs/yrsToDocument.ts:483).
pub(crate) fn lent_by_link_style(mark: &Js) -> Result<bool> {
    Ok(as_object(mark)?.get("inheritedHyperlink").is_true())
}

/// `canonicalFontPair` (src/yrs/yrsToDocument.ts:485).
pub(crate) fn canonical_font_pair(fonts: &Js, name: &str, r#ref: &str, theme: &Js) -> Result<Js> {
    let reference = as_string(&fonts.get(r#ref))?;
    let explicit = as_string(&fonts.get(name))?;
    // `reference && resolveThemeFont(reference, theme)`: only read when `reference` is truthy.
    let second = match reference.as_deref() {
        Some(r) if !r.is_empty() => {
            let resolved = resolve_theme_font(r, theme)?;
            if explicit.is_some() && explicit == resolved {
                Js::Null
            } else {
                explicit.clone().map_or(Js::Null, Js::from)
            }
        }
        _ => explicit.clone().map_or(Js::Null, Js::from),
    };
    Ok(arr![reference.map_or(Js::Null, Js::from), second])
}

/// `relativeFontFamily` (src/yrs/yrsToDocument.ts:493).
pub(crate) fn relative_font_family(
    live: &Js,
    inherited: &Js,
    source: &Js,
    theme: &Js,
) -> Result<Js> {
    let expected = if source.truthy() {
        merge_font_family(inherited, source)?
    } else {
        inherited.clone()
    };
    let result = Js::new_obj();
    for (name, r#ref) in [
        ("ascii", "asciiTheme"),
        ("hAnsi", "hAnsiTheme"),
        ("eastAsia", "eastAsiaTheme"),
        ("cs", "csTheme"),
    ] {
        let unchanged = same_json(
            &canonical_font_pair(live, name, r#ref, theme)?,
            &canonical_font_pair(&expected, name, r#ref, theme)?,
        )?;
        let values = if unchanged { source } else { live };
        for key in [name, r#ref] {
            let v = values.get(key);
            if !v.is_nullish() {
                result.set(key, v);
            }
        }
    }
    let hint = source.get("hint").or_else(|| live.get("hint"));
    if !hint.is_undefined() {
        result.set("hint", hint);
    }
    Ok(if result.key_count() > 0 {
        result
    } else {
        Js::Undefined
    })
}

/// `relativeFormatting` (src/yrs/yrsToDocument.ts:512).
pub(crate) fn relative_formatting(attributes: &Js, inherited: &Js, theme: &Js) -> Result<Js> {
    let formatting = direct_formatting(attributes)?;
    let result = Js::new_obj();
    let mut keys: IndexMap<String, ()> = IndexMap::new();
    for k in formatting.keys().into_iter().chain(inherited.keys()) {
        keys.insert(k, ());
    }
    for key in keys.keys() {
        let key = key.as_str();
        if key == "fontFamily" {
            continue;
        }
        if (key == "boldCs" || key == "italicCs") && !attributes.get(key).is_bool() {
            continue;
        }
        let value = formatting.get(key);
        if same_json(&value, &inherited.get(key))? {
            continue;
        }
        if !value.is_undefined() {
            result.set(key, value);
        } else if on_off(key) {
            result.set(key, false);
        } else if key == "color" && !lent_by_link_style(&attributes.get("textColor"))? {
            result.set(key, obj! { "auto" => true });
        } else if key == "underline" && !lent_by_link_style(&attributes.get("underline"))? {
            result.set(key, obj! { "style" => "none" });
        }
    }
    let source = as_object(&attributes.get("_sourceRunFormatting"))?;
    let fonts = relative_font_family(
        &formatting.get("fontFamily"),
        &inherited.get("fontFamily"),
        &source.get("fontFamily"),
        theme,
    )?;
    if fonts.truthy() {
        result.set("fontFamily", fonts);
    }
    for &key in SOURCE_RUN_PROPERTIES {
        if key == "fontFamily" {
            continue;
        }
        let value = source.get(key);
        if value.is_undefined() {
            continue;
        }
        let live = if on_off(key) {
            formatting.get(key).nullish_or(Js::Bool(false))
        } else {
            formatting.get(key)
        };
        if key == "language"
            || same_json(&value, &live)?
            || ((key == "boldCs" || key == "italicCs") && attributes.get(key).is_undefined())
        {
            result.set(key, value);
        }
    }
    Ok(result)
}

/// `runContentForText` (src/yrs/yrsToDocument.ts:548).
pub(crate) fn run_content_for_text(text: &str, formatting: &Js) -> Result<Js> {
    run_content_for_text_lazy(text, &|| Ok(formatting.clone()))
}

/// `runContentForText` with a `() => TextFormatting` argument.
fn run_content_for_text_lazy(text: &str, formatting: &dyn Fn() -> Result<Js>) -> Result<Js> {
    if !has_special_run_text(text) {
        return Ok(if text.is_empty() {
            arr![]
        } else {
            arr![obj! { "type" => "text", "text" => text }]
        });
    }
    let formatting = formatting()?;
    let content = arr![];
    let mut plain = String::new();
    let flush = |plain: &mut String| {
        if plain.is_empty() {
            return;
        }
        content.push(obj! { "type" => "text", "text" => plain.as_str() });
        plain.clear();
    };
    let family = formatting.get("fontFamily");
    let fonts: Vec<Js> = if family.truthy() {
        ["ascii", "hAnsi", "eastAsia", "cs"]
            .iter()
            .map(|k| family.get(k))
            .filter(Js::truthy)
            .collect()
    } else {
        Vec::new()
    };
    let symbol_font = if !fonts.is_empty() && fonts.iter().all(|f| f.strict_eq(&fonts[0])) {
        fonts[0].clone()
    } else {
        Js::Null
    };

    for ch in text.chars() {
        let code = ch as u32;
        if ch == '\u{ad}' {
            flush(&mut plain);
            content.push(obj! { "type" => "softHyphen" });
        } else if ch == '\u{2011}' {
            flush(&mut plain);
            content.push(obj! { "type" => "noBreakHyphen" });
        } else if symbol_font.truthy() && (0xf000..=0xf8ff).contains(&code) {
            flush(&mut plain);
            content.push(obj! {
                "type" => "symbol",
                "font" => symbol_font.clone(),
                "char" => format!("{code:04X}"),
            });
        } else {
            plain.push(ch);
        }
    }
    flush(&mut plain);
    Ok(content)
}

/// `createTextRun` (src/yrs/yrsToDocument.ts:597).
pub(crate) fn create_text_run(
    text: &str,
    attributes: &Js,
    format: &dyn Fn(&Js) -> Result<Js>,
) -> Result<Js> {
    let formatting = format(attributes)?;
    let run = obj! { "type" => "run" };
    run.assign(&run_property_changes(attributes, format)?);
    run.set(
        "formatting",
        if formatting.key_count() > 0 {
            formatting
        } else {
            Js::Undefined
        },
    );
    run.set(
        "content",
        run_content_for_text_lazy(text, &|| direct_formatting(attributes))?,
    );
    Ok(run)
}

/// `runPropertyChanges` (src/yrs/yrsToDocument.ts:607).
pub(crate) fn run_property_changes(
    attributes: &Js,
    format: &dyn Fn(&Js) -> Result<Js>,
) -> Result<Js> {
    let raw_changes = attributes.get("rPrChange");
    if !raw_changes.is_arr() || raw_changes.len() == 0 {
        return Ok(Js::new_obj());
    }
    let mut changes = Vec::new();
    for raw in raw_changes.items() {
        let Some(change) = as_obj_value(&raw) else {
            changes.push(raw);
            continue;
        };
        let previous = as_object(&change.get("previousAttributes"))?;
        let info = as_object(&change.get("info"))?;
        let id = info.get("id");
        let generated = id.is_str() && !js::to_number(&id).is_finite();
        let out = obj! { "type" => "runPropertyChange" };
        out.assign(&change);
        if previous.truthy() && generated {
            let attrs = previous.with(&[(
                "_sourceRunFormatting",
                attributes.get("_sourceRunFormatting"),
            )]);
            out.set("previousFormatting", format(&attrs)?);
        }
        changes.push(out);
    }
    Ok(obj! { "propertyChanges" => property_changes(&Js::new_arr(changes))? })
}

/// `appendTextRun` (src/yrs/yrsToDocument.ts:626).
pub(crate) fn append_text_run(target: &Js, source: &Js) -> Result<()> {
    let target_content = target.get("content");
    for item in source.get("content").items() {
        let previous = target_content.last();
        if type_is(&previous, "text") && type_is(&item, "text") {
            let joined = format!(
                "{}{}",
                js::to_js_string(&previous.get("text")),
                js::to_js_string(&item.get("text"))
            );
            previous.set("text", joined);
        } else {
            target_content.push(item);
        }
    }
    Ok(())
}

/// `fnv53` (src/yrs/yrsToDocument.ts:634).
pub(crate) fn fnv53(value: &str) -> Result<f64> {
    let mut hash: u64 = 0xcbf29ce484222325;
    for byte in value.as_bytes() {
        hash = (hash ^ u64::from(*byte)).wrapping_mul(0x100000001b3);
    }
    Ok((hash & ((1u64 << 53) - 1)) as f64)
}

/// `revisionId` (src/yrs/yrsToDocument.ts:643).
pub(crate) fn revision_id(value: &Js) -> Result<f64> {
    match value {
        Js::Num(n) if n.is_finite() => Ok(*n),
        Js::Str(s) => {
            let parsed = js::to_number(value);
            if parsed.is_finite() {
                Ok(parsed)
            } else {
                fnv53(s)
            }
        }
        _ => Ok(0.0),
    }
}

/// `trackedInfo` (src/yrs/yrsToDocument.ts:653).
pub(crate) fn tracked_info(raw: &Js, _pm_shape: bool, id: Option<f64>) -> Result<Js> {
    let Some(value) = as_obj_value(raw) else {
        return Ok(Js::Null);
    };
    let author = as_string(&value.get("author"))?
        .filter(|s| !s.is_empty())
        .unwrap_or_else(|| "Unknown".into());
    let date = as_string(&value.get("date"))?;
    let source = value.get("revisionId").or_else(|| value.get("id"));
    let info = obj! {
        "id" => match id { Some(i) => i, None => revision_id(&source)? },
        "author" => author,
    };
    if let Some(d) = date.filter(|d| !d.is_empty()) {
        info.set("date", d);
    }
    mark_revision_info(
        &info,
        id.is_none() && source.is_str() && !js::to_number(&source).is_finite(),
    )
}

/// `propertyChanges` (src/yrs/yrsToDocument.ts:666).
pub(crate) fn property_changes(changes: &Js) -> Result<Js> {
    let mut out = Vec::new();
    for change in changes.items() {
        let value = as_obj_value(&change);
        let info = as_object(&value.clone().unwrap_or_default().get("info"))?;
        let id = info.get("revisionId").or_else(|| info.get("id"));
        match value {
            Some(v) if id.is_str() && !js::to_number(&id).is_finite() => {
                out.push(v.with(&[("info", tracked_info(&info, false, None)?)]));
            }
            _ => out.push(change),
        }
    }
    Ok(Js::new_arr(out))
}

/// `mappedRevisionId` (src/yrs/yrsToDocument.ts:684).
pub(crate) fn mapped_revision_id(raw: &Js, kind: &Js, revision_ids: &Js) -> Result<Option<f64>> {
    let value = as_object(raw)?;
    let id = value.get("revisionId").or_else(|| value.get("id"));
    let Some(id) = id.as_str() else {
        return Ok(None);
    };
    if !revision_ids.is_map() {
        return Ok(None);
    }
    let numbers = revision_ids.map_get(id);
    Ok(numbers.get(kind.str_or_empty()).as_f64())
}

/// `createHyperlink` (src/yrs/yrsToDocument.ts:694).
pub(crate) fn create_hyperlink(attributes: &Js) -> Result<Js> {
    let Some(link) = as_obj_value(&attributes.get("hyperlink")) else {
        return Ok(Js::Null);
    };
    let href = as_string(&link.get("href"))?.unwrap_or_default();
    let tooltip = opt_str(as_string(&link.get("tooltip"))?.filter(|s| !s.is_empty()));
    let history = link.get("history");
    if let Some(anchor) = href.strip_prefix('#') {
        let out = obj! { "type" => "hyperlink", "anchor" => anchor, "tooltip" => tooltip };
        if history.is_bool() {
            out.set("history", history);
        }
        out.set("children", arr![]);
        return Ok(out);
    }
    let out = obj! { "type" => "hyperlink", "href" => href.as_str(), "tooltip" => tooltip };
    if history.is_bool() {
        out.set("history", history);
    }
    out.set(
        "rId",
        opt_str(as_string(&link.get("rId"))?.filter(|s| !s.is_empty())),
    );
    out.set("children", arr![]);
    Ok(out)
}

/// `hyperlinkKey` (src/yrs/yrsToDocument.ts:717).
pub(crate) fn hyperlink_key(attributes: &Js) -> Result<Option<String>> {
    let Some(link) = as_obj_value(&attributes.get("hyperlink")) else {
        return Ok(None);
    };
    Ok(Some(as_string(&link.get("href"))?.unwrap_or_default()))
}

/// The `try` body of `fieldFromPayload`'s cached path; `None` where TS falls through.
fn stored_field(field_data: &str, payload: &Js) -> Result<Option<Js>> {
    let Ok(stored) = Js::parse(field_data) else {
        return Ok(None);
    };
    if field_data.contains("\"media:") {
        resolve_media_srcs(&stored)?;
    }
    let ty = stored.get("type");
    let children = if ty.as_str() == Some("simpleField") {
        stored.get("content")
    } else {
        stored.get("fieldResult")
    };
    if matches!(ty.as_str(), Some("simpleField" | "complexField"))
        && stored.get("instruction").is_str()
        && children.is_arr()
    {
        let display_mode = as_string(&payload.get("displayMode"))?;
        let tree = stored.get("fieldTree");
        if tree.truthy() {
            if let Some(mode) = display_mode.filter(|m| !m.is_empty()) {
                // Assigning onto a primitive throws in strict mode: caught, falls through.
                if !tree.is_object_like() {
                    return Ok(None);
                }
                tree.set("displayMode", mode);
            }
        }
        return Ok(Some(stored));
    }
    Ok(None)
}

/// `fieldFromPayload` (src/yrs/yrsToDocument.ts:722).
pub(crate) fn field_from_payload(
    payload: &Js,
    attributes: &Js,
    format: &dyn Fn(&Js) -> Result<Js>,
) -> Result<Js> {
    if let Some(field_data) = as_string(&payload.get("fieldData"))? {
        if !field_data.is_empty() && js::len16(&field_data) <= 2_000_000 {
            if let Some(stored) = stored_field(&field_data, payload)? {
                return Ok(stored);
            }
        }
    }

    let formatting = format(attributes)?;
    let display_text = as_string(&payload.get("displayText"))?.unwrap_or_default();
    let display_run = obj! {
        "type" => "run",
        "content" => arr![obj! { "type" => "text", "text" => display_text }],
    };
    if formatting.key_count() > 0 {
        display_run.set("formatting", formatting);
    }
    let instruction = as_string(&payload.get("instruction"))?.unwrap_or_default();
    let field_type = as_string(&payload.get("fieldType"))?
        .filter(|s| !s.is_empty())
        .unwrap_or_else(|| "UNKNOWN".into());
    let fld_lock = if payload.get("fldLock").is_true() {
        Js::Bool(true)
    } else {
        Js::Undefined
    };
    let dirty = if payload.get("dirty").is_true() {
        Js::Bool(true)
    } else {
        Js::Undefined
    };
    let display_mode = as_string(&payload.get("displayMode"))?
        .filter(|s| !s.is_empty())
        .unwrap_or_else(|| "result".into());
    let tree = obj! { "version" => 1, "displayMode" => display_mode };
    if payload.get("fieldKind").as_str() == Some("complex") {
        return Ok(obj! {
            "type" => "complexField",
            "instruction" => instruction,
            "fieldType" => field_type,
            "fieldCode" => arr![],
            "fieldResult" => arr![display_run],
            "fldLock" => fld_lock,
            "dirty" => dirty,
            "fieldTree" => tree,
        });
    }
    Ok(obj! {
        "type" => "simpleField",
        "instruction" => instruction,
        "fieldType" => field_type,
        "content" => arr![display_run],
        "fldLock" => fld_lock,
        "dirty" => dirty,
        "fieldTree" => tree,
    })
}

/// `mathFromPayload` (src/yrs/yrsToDocument.ts:781).
pub(crate) fn math_from_payload(payload: &Js) -> Result<Js> {
    Ok(obj! {
        "type" => "mathEquation",
        "display" => as_string(&payload.get("display"))?.filter(|s| !s.is_empty()).unwrap_or_else(|| "inline".into()),
        "ommlXml" => as_string(&payload.get("ommlXml"))?.unwrap_or_default(),
        "plainText" => opt_str(as_string(&payload.get("plainText"))?.filter(|s| !s.is_empty())),
    })
}

/// `horizontalRuleRun` (src/yrs/yrsToDocument.ts:790).
pub(crate) fn horizontal_rule_run(
    payload: &Js,
    attributes: &Js,
    format: &dyn Fn(&Js) -> Result<Js>,
) -> Result<Js> {
    let value = as_object(&payload.get("rule"))?;
    let height = as_finite_number(&value.get("height"))?;
    let bad_number = |k: &str| -> Result<bool> {
        Ok(!value.get(k).is_nullish() && as_finite_number(&value.get(k))?.is_none())
    };
    if value.is_undefined()
        || height.is_none()
        || bad_number("width")?
        || bad_number("widthPercent")?
        || !value.get("alignment").is_str()
        || !value.get("noShade").is_bool()
        || !value.get("color").is_str()
        || !value.get("xml").is_str()
    {
        throw!("Malformed horizontalRule embed payload");
    }
    let rule = obj! {
        "width" => as_finite_number(&value.get("width"))?.map_or(Js::Null, Js::from),
        "widthPercent" => as_finite_number(&value.get("widthPercent"))?.map_or(Js::Null, Js::from),
        "height" => height.unwrap(),
        "alignment" => value.get("alignment"),
        "noShade" => value.get("noShade"),
        "color" => value.get("color"),
        "xml" => value.get("xml"),
    };
    let formatting = format(attributes)?;
    let run = obj! {
        "type" => "run",
        "content" => arr![obj! { "type" => "horizontalRule", "rule" => rule }],
    };
    if formatting.key_count() > 0 {
        run.set("formatting", formatting);
    }
    Ok(run)
}

/// `mediaSrc` (src/yrs/yrsToDocument.ts:820).
pub(crate) fn media_src(src: &str) -> Result<String> {
    if !src.starts_with("media:") {
        return Ok(src.to_owned());
    }
    Ok(match projected_media(src) {
        Some(Some(url)) => url,
        _ => src.to_owned(),
    })
}

/// `resolveMediaSrcs` (src/yrs/yrsToDocument.ts:824).
pub(crate) fn resolve_media_srcs(value: &Js) -> Result<()> {
    if value.is_arr() {
        for item in value.items() {
            resolve_media_srcs(&item)?;
        }
        return Ok(());
    }
    let Some(object) = as_obj_value(value) else {
        return Ok(());
    };
    for (key, field) in object.entries() {
        if key == "src" && field.is_str() {
            object.set(&key, media_src(field.str_or_empty())?);
        } else {
            resolve_media_srcs(&field)?;
        }
    }
    Ok(())
}

/// `imageFromPayload` (src/yrs/yrsToDocument.ts:837).
pub(crate) fn image_from_payload(payload: &Js) -> Result<Js> {
    let a = |k: &str| payload.get(k);
    let num = |k: &str| js::to_number(&payload.get(k));
    let wrap = obj! {
        "type" => as_string(&a("wrapType"))?.filter(|s| !s.is_empty()).unwrap_or_else(|| "inline".into()),
    };
    for (key, out) in [
        ("distTop", "distT"),
        ("distBottom", "distB"),
        ("distLeft", "distL"),
        ("distRight", "distR"),
    ] {
        if !a(key).is_undefined() {
            wrap.set(out, pixels_to_emu(num(key))?);
        }
    }
    if a("wrapText").truthy() {
        wrap.set("wrapText", a("wrapText"));
    }

    let or_zero = |n: f64| if n.is_nan() || n == 0.0 { 0.0 } else { n };
    let image = obj! {
        "type" => "image",
        "rId" => as_string(&a("rId"))?.unwrap_or_default(),
        "src" => media_src(&as_string(&a("src"))?.unwrap_or_default())?,
        "alt" => opt_str(as_string(&a("alt"))?.filter(|s| !s.is_empty())),
        "title" => opt_str(as_string(&a("title"))?.filter(|s| !s.is_empty())),
        "shapeType" => opt_str(as_string(&a("shapeType"))?.filter(|s| !s.is_empty())),
        "size" => obj! {
            "width" => pixels_to_emu(or_zero(num("width")))?,
            "height" => pixels_to_emu(or_zero(num("height")))?,
        },
        "wrap" => wrap,
    };

    let transform_attr = a("transform");
    if transform_attr.truthy() {
        let Some(t) = transform_attr.as_str() else {
            throw!("attrs.transform.match is not a function");
        };
        let transform = Js::new_obj();
        if let Some(rotation) = match_rotation(t) {
            transform.set("rotation", js::parse_float(&rotation));
        }
        if t.contains("scaleX(-1)") {
            transform.set("flipH", true);
        }
        if t.contains("scaleY(-1)") {
            transform.set("flipV", true);
        }
        if transform.get("rotation").truthy()
            || transform.get("flipH").truthy()
            || transform.get("flipV").truthy()
        {
            image.set("transform", transform);
        }
    }

    let position = a("position");
    if position.get("horizontal").truthy() && position.get("vertical").truthy() {
        let h = position.get("horizontal");
        let v = position.get("vertical");
        image.set(
            "position",
            obj! {
                "relativeHeight" => position.get("relativeHeight"),
                "horizontal" => obj! {
                    "relativeTo" => h.get("relativeTo").or(Js::from("column")),
                    "alignment" => h.get("align"),
                    "posOffset" => h.get("posOffset"),
                },
                "vertical" => obj! {
                    "relativeTo" => v.get("relativeTo").or(Js::from("paragraph")),
                    "alignment" => v.get("align"),
                    "posOffset" => v.get("posOffset"),
                },
            },
        );
    }

    if a("borderWidth").truthy() && num("borderWidth") > 0.0 {
        let style = a("borderStyle");
        let mapped = match style.as_str() {
            Some("solid") => Some("solid"),
            Some("dotted") => Some("dot"),
            Some("dashed") => Some("dash"),
            Some("double" | "groove" | "ridge" | "inset" | "outset") => Some("solid"),
            _ => None,
        };
        let color = a("borderColor");
        image.set("outline", obj! {
            "width" => pixels_to_emu(num("borderWidth"))?,
            "color" => if color.truthy() { obj! { "rgb" => replace_first_hash(&js::to_js_string(&color)) } } else { Js::Undefined },
            "style" => mapped.filter(|_| style.truthy()).unwrap_or("solid"),
        });
    }
    if a("hlinkHref").truthy() {
        image.set("hlinkHref", a("hlinkHref"));
    }

    let crop = Js::new_obj();
    for (key, out) in [
        ("cropTop", "top"),
        ("cropRight", "right"),
        ("cropBottom", "bottom"),
        ("cropLeft", "left"),
    ] {
        if !a(key).is_nullish() {
            crop.set(out, a(key));
        }
    }
    if crop.key_count() > 0 {
        image.set("crop", crop);
    }
    if !a("opacity").is_nullish() && num("opacity") < 1.0 {
        image.set("opacity", a("opacity"));
    }
    for key in [
        "layoutInCell",
        "noChangeAspect",
        "allowOverlap",
        "horizontalRule",
    ] {
        if !a(key).is_nullish() {
            image.set(key, a(key));
        }
    }

    let padding = Js::new_obj();
    for (key, out) in [
        ("effectExtentTop", "top"),
        ("effectExtentBottom", "bottom"),
        ("effectExtentLeft", "left"),
        ("effectExtentRight", "right"),
    ] {
        if !a(key).is_nullish() {
            padding.set(out, pixels_to_emu(num(key))?);
        }
    }
    if padding.key_count() > 0 {
        image.set("padding", padding);
    }

    Ok(image)
}

/// `imageRunFromPayload` (src/yrs/yrsToDocument.ts:936).
pub(crate) fn image_run_from_payload(payload: &Js) -> Result<Js> {
    Ok(obj! {
        "type" => "run",
        "content" => arr![obj! { "type" => "drawing", "image" => image_from_payload(payload)? }],
    })
}

/// `storedShape` (src/yrs/yrsToDocument.ts:941).
pub(crate) fn stored_shape(value: &Js) -> Result<Js> {
    let Some(json) = as_string(value)?.filter(|s| !s.is_empty()) else {
        return Ok(Js::Undefined);
    };
    let Ok(parsed) = Js::parse(&json) else {
        return Ok(Js::Undefined);
    };
    if parsed.get("type").as_str() != Some("shape") || !parsed.get("shapeType").is_str() {
        return Ok(Js::Undefined);
    }
    if as_obj_value(&parsed.get("size")).is_none() {
        parsed.set("size", obj! { "width" => 0, "height" => 0 });
    }
    if json.contains("\"media:") {
        resolve_media_srcs(&parsed)?;
    }
    Ok(parsed)
}

/// `chartRunFromPayload` (src/yrs/yrsToDocument.ts:955).
pub(crate) fn chart_run_from_payload(payload: &Js, recovery: &Js) -> Result<Js> {
    let Some(json) = as_string(&payload.get("chartJson"))?.filter(|s| !s.is_empty()) else {
        recovery_warn(
            recovery,
            "chart run carries no chart payload; keeping the run out of the output",
        );
        return Ok(Js::Null);
    };
    let Ok(chart) = Js::parse(&json) else {
        recovery_warn(
            recovery,
            "chart run carries an unreadable chart payload; keeping the run out of the output",
        );
        return Ok(Js::Null);
    };
    if json.contains("\"media:") {
        resolve_media_srcs(&chart)?;
    }
    if chart.get("type").as_str() != Some("chart") || !chart.get("chartType").is_str() {
        recovery_warn(
            recovery,
            "chart run carries a malformed chart payload; keeping the run out of the output",
        );
        return Ok(Js::Null);
    }
    if !chart.get("drawingXml").truthy() {
        let r_id = recovery_key(&chart.get("rId"));
        let path = recovery_key(&chart.get("path"));
        let drawing_xml = recovery_drawing_for(recovery, r_id.as_deref(), path.as_deref());
        let Some(drawing_xml) = drawing_xml.filter(|x| !x.is_empty()) else {
            recovery_unrecoverable(recovery, r_id.as_deref(), path.as_deref());
            return Ok(Js::Null);
        };
        chart.set("drawingXml", drawing_xml);
    }
    Ok(obj! { "type" => "run", "content" => arr![obj! { "type" => "chart", "chart" => chart }] })
}

/// `replayedRunProperties` (src/yrs/yrsToDocument.ts:983).
pub(crate) fn replayed_run_properties(
    payload: &Js,
    attributes: &Js,
    format: &dyn Fn(&Js) -> Result<Js>,
) -> Result<Js> {
    let source = as_obj_value(&payload.get("sourceRun"));
    let seeded = source
        .as_ref()
        .and_then(|s| as_obj_value(&s.get("attributes")));
    let (Some(source), Some(seeded)) = (source, seeded) else {
        return Ok(Js::new_obj());
    };
    let before = direct_formatting(&seeded)?;
    let current = direct_formatting(attributes)?;
    let projected = format(attributes)?.spread();
    let formatting = as_object(&source.get("formatting"))?.spread();
    let mut keys: IndexMap<String, ()> = IndexMap::new();
    for k in before.keys().into_iter().chain(current.keys()) {
        keys.insert(k, ());
    }
    for key in keys.keys() {
        if same_json(&before.get(key), &current.get(key))? {
            continue;
        }
        let value = projected.get(key);
        if value.is_undefined() {
            formatting.delete(key);
        } else {
            formatting.set(key, value);
        }
    }
    let out = obj! { "formatting" => formatting };
    out.assign(&run_property_changes(attributes, format)?);
    Ok(out)
}

/// `opaqueDrawingRunFromPayload` (src/yrs/yrsToDocument.ts:1002).
pub(crate) fn opaque_drawing_run_from_payload(
    payload: &Js,
    attributes: &Js,
    format: &dyn Fn(&Js) -> Result<Js>,
) -> Result<Js> {
    let kind = as_string(&payload.get("kind"))?.filter(|s| !s.is_empty());
    let xml = as_string(&payload.get("xml"))?.filter(|s| !s.is_empty());
    let (Some(kind), Some(xml)) = (kind, xml) else {
        return Ok(Js::Null);
    };
    let preview = as_obj_value(&payload.get("preview"));
    let mut previews = Vec::new();
    if payload.get("previews").is_arr() {
        for image in payload.get("previews").items() {
            if let Some(image) = as_obj_value(&image) {
                previews.push(image_from_payload(&image)?);
            }
        }
    }
    let text_box = stored_shape(&payload.get("textBoxJson"))?;
    let content = obj! { "type" => "opaqueDrawing", "kind" => kind.as_str(), "xml" => xml };
    if let Some(p) = preview {
        content.set("preview", image_from_payload(&p)?);
    }
    if !previews.is_empty() {
        content.set("previews", Js::new_arr(previews));
    }
    if text_box.truthy() {
        content.set("textBox", text_box);
    }
    let run = obj! { "type" => "run", "content" => arr![content] };
    if kind == "drawingRun" {
        run.assign(&replayed_run_properties(payload, attributes, format)?);
    } else {
        run.set("formatting", format(attributes)?);
    }
    Ok(run)
}

/// The `try` body of the gradient fill: `None` where the TS throws.
fn gradient_fill(payload: &Js, stops_json: &str) -> Result<Option<Js>> {
    let Ok(stops) = Js::parse(stops_json) else {
        return Ok(None);
    };
    if !stops.is_arr() {
        return Ok(None);
    }
    let mut mapped = Vec::new();
    for stop in stops.items() {
        if stop.is_nullish() {
            return Ok(None);
        }
        let Some(color) = stop.get("color").as_str().map(str::to_owned) else {
            return Ok(None);
        };
        mapped.push(obj! { "position" => stop.get("position"), "color" => obj! { "rgb" => replace_first_hash(&color) } });
    }
    let angle = as_finite_number(&payload.get("gradientAngle"))?.filter(|a| *a != 0.0);
    Ok(Some(obj! {
        "type" => "gradient",
        "gradient" => obj! {
            "type" => as_string(&payload.get("gradientType"))?.filter(|s| !s.is_empty()).unwrap_or_else(|| "linear".into()),
            "angle" => angle,
            "stops" => Js::new_arr(mapped),
        },
    }))
}

/// `shapeRunFromPayload` (src/yrs/yrsToDocument.ts:1027).
pub(crate) fn shape_run_from_payload(
    payload: &Js,
    attributes: &Js,
    format: &dyn Fn(&Js) -> Result<Js>,
) -> Result<Js> {
    let a = |k: &str| payload.get(k);
    let shape = stored_shape(&a("shapeJson"))?.or_else(|| obj! {
        "type" => "shape", "shapeType" => "rect", "size" => obj! { "width" => 0, "height" => 0 },
    });
    if let Some(t) = as_string(&a("shapeType"))?.filter(|s| !s.is_empty()) {
        shape.set("shapeType", t);
    }
    if let Some(id) = as_string(&a("shapeId"))?.filter(|s| !s.is_empty()) {
        shape.set("id", id);
    }
    if a("width").truthy() {
        shape.set(
            "size",
            shape
                .get("size")
                .with(&[("width", Js::Num(pixels_to_emu(js::to_number(&a("width")))?))]),
        );
    }
    if a("height").truthy() {
        shape.set(
            "size",
            shape.get("size").with(&[(
                "height",
                Js::Num(pixels_to_emu(js::to_number(&a("height")))?),
            )]),
        );
    }
    if a("geometryPath").is_arr() && a("geometryPath").len() > 0 {
        shape.set("geometryPath", a("geometryPath"));
    }

    if a("fillType").as_str() == Some("gradient") && a("gradientStops").is_str() {
        match gradient_fill(payload, a("gradientStops").str_or_empty())? {
            Some(fill) => shape.set("fill", fill),
            None => {
                let rgb = as_string(&a("fillColor"))?
                    .filter(|s| !s.is_empty())
                    .unwrap_or_else(|| "000000".into());
                shape.set("fill", obj! { "type" => "solid", "color" => obj! { "rgb" => replace_first_hash(&rgb) } });
            }
        }
    } else if let Some(color) = a("fillColor").as_str() {
        shape.set("fill", obj! {
            "type" => as_string(&a("fillType"))?.filter(|s| !s.is_empty()).unwrap_or_else(|| "solid".into()),
            "color" => obj! { "rgb" => replace_first_hash(color) },
        });
    } else if a("fillType").as_str() == Some("none") {
        shape.set("fill", obj! { "type" => "none" });
    }

    if let Some(width) = a("outlineWidth").as_f64().filter(|w| *w > 0.0) {
        let mapped = match as_string(&a("outlineStyle"))?.as_deref() {
            Some("solid") => "solid",
            Some("dotted") => "dot",
            Some("dashed") => "dash",
            _ => "solid",
        };
        shape.set("outline", obj! {
            "width" => pixels_to_emu(width)?,
            "color" => match a("outlineColor").as_str() { Some(c) => obj! { "rgb" => replace_first_hash(c) }, None => Js::Undefined },
            "style" => mapped,
        });
    }

    let transform = Js::new_obj();
    if a("rotation").is_num() {
        transform.set("rotation", a("rotation"));
    } else if let Some(t) = a("transform").as_str() {
        if let Some(rotation) = match_rotation(t) {
            transform.set("rotation", js::parse_float(&rotation));
        }
    }
    let transform_text = js::to_js_string(&a("transform").or(Js::from("")));
    if a("flipH").truthy() || transform_text.contains("scaleX(-1)") {
        transform.set("flipH", true);
    }
    if a("flipV").truthy() || transform_text.contains("scaleY(-1)") {
        transform.set("flipV", true);
    }
    if transform.get("rotation").truthy()
        || transform.get("flipH").truthy()
        || transform.get("flipV").truthy()
    {
        shape.set("transform", transform);
    }

    let run = obj! { "type" => "run", "content" => arr![obj! { "type" => "shape", "shape" => shape.clone() }] };
    if shape.get("drawingRunXml").truthy() {
        run.assign(&replayed_run_properties(payload, attributes, format)?);
    } else {
        run.set("formatting", format(attributes)?);
    }
    Ok(run)
}

/// `inlineSdtFromPayload` (src/yrs/yrsToDocument.ts:1112).
pub(crate) fn inline_sdt_from_payload(
    payload: &Js,
    recovery: &Js,
    format: &dyn Fn(&Js) -> Result<Js>,
) -> Result<Js> {
    let mut properties = sdt_attrs_to_props(payload)?;
    if let Some(json) = as_string(&payload.get("propertiesJson"))? {
        if !json.is_empty() && js::len16(&json) <= 1_000_000 {
            if let Ok(parsed) = Js::parse(&json) {
                if json.contains("\"media:") {
                    resolve_media_srcs(&parsed)?;
                }
                if parsed.is_object_like() && parsed.get("sdtType").is_str() {
                    properties = parsed;
                }
            }
        }
    }
    let items = arr![];
    let content_items = payload.get("content");
    for raw in if content_items.is_arr() {
        content_items.items()
    } else {
        Vec::new()
    } {
        let entry = as_obj_value(&raw);
        let kind = entry
            .as_ref()
            .and_then(|e| e.get("kind").as_str().map(str::to_owned))
            .filter(|k| !k.is_empty());
        let (Some(entry), Some(kind)) = (entry, kind) else {
            continue;
        };
        let attributes = as_obj_value(&entry.get("attrs")).unwrap_or_else(Js::new_obj);
        if kind == "text" {
            let text = entry.get("text");
            if text.is_str() {
                items.push(obj! { "kind" => "text", "text" => text, "attributes" => attributes });
            }
            continue;
        }
        items.push(obj! {
            "kind" => "embed",
            "embedKind" => kind,
            "payload" => as_obj_value(&entry.get("payload")).unwrap_or_else(Js::new_obj),
            "attributes" => attributes,
        });
    }
    let mut content = inline_sdt_content(&build_paragraph_content(
        &items,
        recovery,
        format,
        &Js::Undefined,
        true,
    )?)?;
    let authored = content_control_value(&payload.get("value"))?;
    if authored.truthy() {
        if let Ok(applied) = apply_content_control_value(&properties, &authored) {
            properties = applied.get("properties");
            let display = applied.get("content").at(0);
            let children = if type_is(&display, "paragraph") {
                display.get("content")
            } else {
                arr![]
            };
            // `.filter` on a non-array throws inside the `try`: content stays.
            if children.is_arr() {
                content = inline_sdt_content(&children)?;
            }
        }
    }
    Ok(obj! { "type" => "inlineSdt", "properties" => properties, "content" => content })
}

fn filter_types(content: &Js, types: &[&str]) -> Js {
    Js::new_arr(
        content
            .items()
            .into_iter()
            .filter(|c| c.get("type").as_str().is_some_and(|t| types.contains(&t)))
            .collect(),
    )
}

/// `inlineSdtContent` (src/yrs/yrsToDocument.ts:1164).
pub(crate) fn inline_sdt_content(content: &Js) -> Result<Js> {
    Ok(filter_types(
        content,
        &[
            "run",
            "hyperlink",
            "simpleField",
            "complexField",
            "inlineSdt",
            "mathEquation",
            "insertion",
            "deletion",
            "moveFrom",
            "moveTo",
        ],
    ))
}

/// `fieldInlineContent` (src/yrs/yrsToDocument.ts:1180).
pub(crate) fn field_inline_content(content: &Js) -> Result<Js> {
    Ok(filter_types(
        content,
        &[
            "run",
            "hyperlink",
            "simpleField",
            "complexField",
            "inlineSdt",
            "mathEquation",
        ],
    ))
}

/// `contentControlValue` (src/yrs/yrsToDocument.ts:1192).
pub(crate) fn content_control_value(value: &Js) -> Result<Js> {
    let authored = as_object(value)?;
    match authored.get("kind").as_str() {
        Some("checkbox") if authored.get("checked").is_bool() => {
            Ok(obj! { "kind" => "checkbox", "checked" => authored.get("checked") })
        }
        Some("dropdown") if authored.get("value").is_str() => {
            Ok(obj! { "kind" => "dropdown", "value" => authored.get("value") })
        }
        Some("date") if authored.get("date").is_str() => {
            Ok(obj! { "kind" => "date", "date" => authored.get("date") })
        }
        _ => Ok(Js::Null),
    }
}

/// `commentReferenceFromPayload` (src/yrs/yrsToDocument.ts:1206).
pub(crate) fn comment_reference_from_payload(payload: &Js) -> Result<Js> {
    if payload.get("modelKind").as_str() != Some("commentReference") {
        return Ok(Js::Null);
    }
    let reference = obj! { "type" => "commentReference" };
    if let Some(id) = as_finite_number(&payload.get("commentId"))? {
        reference.set("id", id);
    }
    Ok(obj! { "type" => "run", "content" => arr![reference] })
}

/// `ordinaryContentForItem` (src/yrs/yrsToDocument.ts:1215).
pub(crate) fn ordinary_content_for_item(
    item: &Js,
    recovery: &Js,
    format: &dyn Fn(&Js) -> Result<Js>,
) -> Result<Js> {
    let content = inline_content_for_item(item, recovery, format)?;
    if type_is(&content, "run") {
        let out = content.spread();
        out.assign(&run_property_changes(&item.get("attributes"), format)?);
        return Ok(out);
    }
    Ok(content)
}

fn item_text(item: &Js) -> String {
    js::to_js_string(&item.get("text"))
}

/// `inlineContentForItem` (src/yrs/yrsToDocument.ts:1224).
pub(crate) fn inline_content_for_item(
    item: &Js,
    recovery: &Js,
    format: &dyn Fn(&Js) -> Result<Js>,
) -> Result<Js> {
    let attributes = item.get("attributes");
    let payload = item.get("payload");
    if item.get("kind").as_str() == Some("text") {
        return create_text_run(&item_text(item), &attributes, format);
    }
    match item.get("embedKind").as_str().unwrap_or("") {
        "break" => Ok(obj! { "type" => "run", "content" => arr![obj! { "type" => "break" }] }),
        "tab" => tab_run(&attributes, format),
        "image" => image_run_from_payload(&payload),
        "horizontalRule" => horizontal_rule_run(&payload, &attributes, format),
        "shape" => shape_run_from_payload(&payload, &attributes, format),
        "chart" => chart_run_from_payload(&payload, recovery),
        "opaqueDrawing" => opaque_drawing_run_from_payload(&payload, &attributes, format),
        "field" => {
            let reference = comment_reference_from_payload(&payload)?;
            if !reference.is_nullish() {
                Ok(reference)
            } else {
                field_from_payload(&payload, &attributes, format)
            }
        }
        "math" => math_from_payload(&payload),
        "sdt" => inline_sdt_from_payload(&payload, recovery, format),
        "noteRef" => {
            let footnote = payload.get("footnoteRefId");
            let endnote = payload.get("endnoteRefId");
            let reference = if !footnote.is_undefined() {
                obj! { "type" => "footnoteRef", "id" => revision_id(&footnote)? }
            } else {
                obj! { "type" => "endnoteRef", "id" => revision_id(&endnote)? }
            };
            Ok(obj! { "type" => "run", "content" => arr![reference] })
        }
        _ => Ok(Js::Null),
    }
}

/// `tabRun` (src/yrs/yrsToDocument.ts:1271).
pub(crate) fn tab_run(attributes: &Js, format: &dyn Fn(&Js) -> Result<Js>) -> Result<Js> {
    let formatting = format(attributes)?;
    let run = obj! { "type" => "run" };
    run.assign(&run_property_changes(attributes, format)?);
    run.set("content", arr![obj! { "type" => "tab" }]);
    if formatting.key_count() > 0 {
        run.set("formatting", formatting);
    }
    Ok(run)
}

fn empty_run() -> Js {
    obj! { "type" => "run", "content" => arr![] }
}

/// `trackedContentForItem` (src/yrs/yrsToDocument.ts:1281).
pub(crate) fn tracked_content_for_item(
    item: &Js,
    info: &Js,
    in_control: bool,
    recovery: &Js,
    format: &dyn Fn(&Js) -> Result<Js>,
) -> Result<Js> {
    let attributes = item.get("attributes");
    let payload = item.get("payload");
    let is_embed = item.get("kind").as_str() == Some("embed");
    let embed_kind = item.get("embedKind");
    let embed = |k: &str| is_embed && embed_kind.as_str() == Some(k);
    let null_to_empty = |v: Js| if v.is_nullish() { empty_run() } else { v };

    let mut child = if in_control {
        let ordinary = ordinary_content_for_item(item, recovery, format)?;
        let list = if ordinary.truthy() {
            arr![ordinary]
        } else {
            arr![]
        };
        null_to_empty(inline_sdt_content(&list)?.at(0))
    } else if embed("sdt") {
        inline_sdt_from_payload(&payload, recovery, format)?
    } else if embed("image") {
        image_run_from_payload(&payload)?
    } else if embed("tab") {
        tab_run(&attributes, format)?
    } else if embed("horizontalRule") {
        horizontal_rule_run(&payload, &attributes, format)?
    } else if embed("shape") {
        shape_run_from_payload(&payload, &attributes, format)?
    } else if embed("chart") {
        null_to_empty(chart_run_from_payload(&payload, recovery)?)
    } else if embed("opaqueDrawing") {
        null_to_empty(opaque_drawing_run_from_payload(
            &payload,
            &attributes,
            format,
        )?)
    } else if item.get("kind").as_str() == Some("text") {
        create_text_run(&item_text(item), &attributes, format)?
    } else {
        empty_run()
    };

    if embed("chart") && type_is(&child, "run") && child.get("content").len() == 0 {
        return Ok(Js::Null);
    }
    if type_is(&child, "run") {
        child.assign(&run_property_changes(&attributes, format)?);
    }
    let hyperlink = create_hyperlink(&attributes)?;
    let ty = child.get("type");
    if hyperlink.truthy()
        && matches!(
            ty.as_str(),
            Some("run" | "mathEquation" | "inlineSdt" | "simpleField" | "complexField")
        )
    {
        if ty.as_str() == Some("run") {
            child.assign(&run_property_changes(&attributes, format)?);
            hyperlink.get("children").push(child.clone());
        } else {
            hyperlink.set("structuredChildren", arr![child.clone()]);
        }
        child = hyperlink;
    }

    let raw = as_obj_value(&attributes.get("ins")).or_else(|| as_obj_value(&attributes.get("del")));
    let is_move_pair = raw.is_some_and(|r| r.get("isMovePair").is_true());
    let ty = if attributes.get("ins").truthy() {
        if is_move_pair { "moveTo" } else { "insertion" }
    } else if is_move_pair {
        "moveFrom"
    } else {
        "deletion"
    };
    Ok(obj! { "type" => ty, "info" => info, "content" => arr![child] })
}

/// `addToHyperlink` (src/yrs/yrsToDocument.ts:1328).
pub(crate) fn add_to_hyperlink(
    hyperlink: &Js,
    item: &Js,
    recovery: &Js,
    format: &dyn Fn(&Js) -> Result<Js>,
) -> Result<()> {
    let attributes = item.get("attributes");
    let payload = item.get("payload");
    let child: Js = if item.get("kind").as_str() == Some("text") {
        create_text_run(&item_text(item), &attributes, format)?
    } else {
        match item.get("embedKind").as_str().unwrap_or("") {
            "break" => obj! { "type" => "run", "content" => arr![obj! { "type" => "break" }] },
            "tab" => obj! { "type" => "run", "content" => arr![obj! { "type" => "tab" }] },
            "horizontalRule" => horizontal_rule_run(&payload, &attributes, format)?,
            "image" => image_run_from_payload(&payload)?,
            "shape" => shape_run_from_payload(&payload, &attributes, format)?,
            "chart" => chart_run_from_payload(&payload, recovery)?,
            "opaqueDrawing" => opaque_drawing_run_from_payload(&payload, &attributes, format)?,
            "field" => {
                let reference = comment_reference_from_payload(&payload)?;
                if !reference.is_nullish() {
                    reference
                } else {
                    field_from_payload(&payload, &attributes, format)?
                }
            }
            "math" => math_from_payload(&payload)?,
            "sdt" => {
                let control = inline_sdt_from_payload(&payload, recovery, format)?;
                if attributes.get("ins").truthy()
                    || attributes.get("del").truthy()
                    || has_tracked_control_content(&control)?
                {
                    control
                } else {
                    Js::Undefined
                }
            }
            _ => Js::Undefined,
        }
    };
    if !child.truthy() {
        return Ok(());
    }
    let is_run = type_is(&child, "run");
    if is_run {
        hyperlink.get("children").push(child.clone());
    }
    if !is_run || hyperlink.get("structuredChildren").truthy() {
        let mut structured = hyperlink.get("structuredChildren");
        if structured.is_nullish() {
            structured = hyperlink.get("children").copy_arr();
            hyperlink.set("structuredChildren", structured.clone());
        }
        structured.push(child);
    }
    Ok(())
}

/// `projectionSignature` (src/yrs/yrsToDocument.ts:1361).
pub(crate) fn projection_signature(items: &Js) -> Result<String> {
    let normalized: Vec<Js> = Vec::new();
    let mut normalized = normalized;
    for source in items.items() {
        let attributes = source.get("attributes").spread();
        attributes.delete("fieldResult");
        attributes.delete("_sourceRunFormatting");
        let item = if source.get("kind").as_str() == Some("embed")
            && source.get("embedKind").as_str() == Some("tab")
        {
            obj! { "kind" => "text", "text" => "\t", "attributes" => attributes.clone() }
        } else {
            source.with(&[("attributes", attributes.clone())])
        };
        let previous = normalized.last().cloned().unwrap_or_default();
        if item.get("kind").as_str() == Some("text")
            && previous.get("kind").as_str() == Some("text")
            && stable_stringify(&previous.get("attributes"))? == stable_stringify(&attributes)?
        {
            previous.set(
                "text",
                format!("{}{}", item_text(&previous), item_text(&item)),
            );
        } else {
            normalized.push(item);
        }
    }
    stable_stringify(&Js::new_arr(normalized))
}

/// `array[index]` for a JS number index.
fn index_at(array: &Js, index: f64) -> Js {
    if index >= 0.0 && index.fract() == 0.0 {
        array.at(index as usize)
    } else {
        Js::Undefined
    }
}

/// `restoreProjectedFieldResults` (src/yrs/yrsToDocument.ts:1378).
pub(crate) fn restore_projected_field_results(
    items: &Js,
    recovery: &Js,
    format: &dyn Fn(&Js) -> Result<Js>,
) -> Result<Js> {
    let all = items.items();
    let mut owners: Vec<(f64, Js, usize)> = Vec::new();
    for (position, item) in all.iter().enumerate() {
        if item.get("kind").as_str() != Some("embed")
            || item.get("embedKind").as_str() != Some("field")
        {
            continue;
        }
        if let Some(id) =
            as_finite_number(&as_object(&item.path(&["payload", "resultProjection"]))?.get("id"))?
        {
            owners.push((id, item.clone(), position));
        }
    }
    if owners.is_empty() {
        return Ok(items.clone());
    }
    // Map<owner, Map<index, items>>, keyed by owner identity, in insertion order.
    let mut groups: Vec<(usize, Vec<(f64, Js)>)> = Vec::new();
    let mut remaining = Vec::new();
    for (position, item) in all.iter().enumerate() {
        let marker = as_object(&item.path(&["attributes", "fieldResult"]))?;
        let id = as_finite_number(&marker.get("id"))?;
        let index = as_finite_number(&marker.get("index"))?;
        let owner = owners
            .iter()
            .find(|(oid, _, p)| Some(*oid) == id && *p > position)
            .map(|(_, o, _)| o.clone());
        let (Some(index), Some(owner)) = (index, owner) else {
            remaining.push(item.clone());
            continue;
        };
        let key = owner.ptr().unwrap();
        let slot = match groups.iter().position(|(k, _)| *k == key) {
            Some(i) => i,
            None => {
                groups.push((key, Vec::new()));
                groups.len() - 1
            }
        };
        let children = &mut groups[slot].1;
        let group = match children.iter().position(|(i, _)| *i == index) {
            Some(i) => children[i].1.clone(),
            None => {
                let g = arr![];
                children.push((index, g.clone()));
                g
            }
        };
        let attributes = item.get("attributes").spread();
        attributes.delete("fieldResult");
        group.push(item.with(&[("attributes", attributes)]));
    }
    for (_, owner, _) in &owners {
        let owner_payload = owner.get("payload");
        let stored = field_from_payload(&owner_payload, &owner.get("attributes"), format)?;
        if !type_is(&stored, "complexField") {
            continue;
        }
        let projection = as_object(&owner_payload.get("resultProjection"))?;
        let originals = if projection.get("children").is_arr() {
            projection.get("children").items()
        } else {
            Vec::new()
        };
        let mut replacements: Vec<(f64, Js)> = Vec::new();
        let owner_key = owner.ptr().unwrap();
        for raw in originals {
            let child = as_object(&raw)?;
            let index = as_finite_number(&child.get("index"))?;
            let Some(index) = index else { continue };
            if !child.get("items").is_arr() {
                continue;
            }
            let current = groups
                .iter()
                .find(|(k, _)| *k == owner_key)
                .and_then(|(_, c)| c.iter().find(|(i, _)| *i == index).map(|(_, g)| g.clone()))
                .unwrap_or_else(|| arr![]);
            if projection_signature(&current)? == projection_signature(&child.get("items"))? {
                continue;
            }
            let rebuilt = field_inline_content(&build_paragraph_content(
                &current,
                recovery,
                format,
                &Js::Undefined,
                false,
            )?)?;
            let original = if index < 0.0 {
                index_at(&stored.path(&["structuredCode", "inline"]), -index - 1.0)
            } else {
                index_at(&stored.path(&["structuredResult", "inline"]), index)
            };
            if type_is(&original, "hyperlink")
                && rebuilt.len() == 1
                && type_is(&rebuilt.at(0), "hyperlink")
            {
                let first = rebuilt.at(0);
                let merged = original.spread();
                merged.assign(&first);
                merged.set("structuredChildren", first.get("structuredChildren"));
                rebuilt.set_at(0, merged);
            }
            match replacements.iter_mut().find(|(i, _)| *i == index) {
                Some(entry) => entry.1 = rebuilt,
                None => replacements.push((index, rebuilt)),
            }
        }
        if replacements.is_empty() {
            continue;
        }
        let replacement = |i: f64| {
            replacements
                .iter()
                .find(|(k, _)| *k == i)
                .map(|(_, v)| v.clone())
        };
        let flat = |list: &Js, key: &dyn Fn(usize) -> f64| -> Js {
            let mut out = Vec::new();
            for (i, child) in list.items().into_iter().enumerate() {
                match replacement(key(i)) {
                    Some(r) => out.extend(r.items()),
                    None => out.push(child),
                }
            }
            Js::new_arr(out)
        };
        let result_inline = stored.path(&["structuredResult", "inline"]);
        let inline = flat(&result_inline.nullish_or(arr![]), &|i| i as f64);
        let code_inline = stored.path(&["structuredCode", "inline"]);
        let code = if code_inline.is_nullish() {
            Js::Undefined
        } else {
            flat(&code_inline, &|i| -(i as f64) - 1.0)
        };
        stored.set(
            "structuredResult",
            stored
                .get("structuredResult")
                .with(&[("inline", inline.clone())]),
        );
        if code.truthy() {
            stored.set(
                "structuredCode",
                stored
                    .get("structuredCode")
                    .with(&[("inline", code.clone())]),
            );
        }
        let tree = stored.get("fieldTree");
        if tree.truthy() {
            tree.set(
                "result",
                tree.get("result").with(&[("inline", inline.clone())]),
            );
            if code.truthy() {
                tree.set("code", tree.get("code").with(&[("inline", code.clone())]));
            }
        }
        let mut field_result = Vec::new();
        for child in inline.items() {
            if type_is(&child, "run") {
                field_result.push(child);
            } else if type_is(&child, "hyperlink") {
                field_result.extend(
                    child
                        .get("children")
                        .items()
                        .into_iter()
                        .filter(|e| type_is(e, "run")),
                );
            }
        }
        stored.set("fieldResult", Js::new_arr(field_result));
        owner.set(
            "payload",
            owner_payload.with(&[(
                "fieldData",
                Js::from(stored.stringify().unwrap_or_default()),
            )]),
        );
    }
    Ok(Js::new_arr(remaining))
}

/// `isTrackedWrapper` (src/yrs/yrsToDocument.ts:1446).
pub(crate) fn is_tracked_wrapper(content: &Js) -> Result<bool> {
    Ok(matches!(
        content.get("type").as_str(),
        Some("insertion" | "deletion" | "moveFrom" | "moveTo")
    ))
}

#[cfg(test)]
mod tests {
    use super::*;

    // Expected values below were produced by running the TS functions with bun.

    #[test]
    fn fnv53_matches_ts() {
        let got: Vec<f64> = ["", "a", "w:rev-1", "中文é😀"]
            .iter()
            .map(|s| fnv53(s).unwrap())
            .collect();
        assert_eq!(
            got,
            vec![
                5239054864098085.0,
                1086646154030220.0,
                5562418442810663.0,
                6307491230161420.0
            ]
        );
    }

    #[test]
    fn revision_id_matches_ts() {
        let inputs = [
            Js::from("12"),
            Js::from("abc"),
            Js::from(""),
            Js::from(" 7 "),
            Js::Num(3.5),
            Js::Null,
            Js::Num(f64::INFINITY),
        ];
        let got: Vec<f64> = inputs.iter().map(|v| revision_id(v).unwrap()).collect();
        assert_eq!(got, vec![12.0, 8903952624080715.0, 0.0, 7.0, 3.5, 0.0, 0.0]);
    }

    #[test]
    fn stable_stringify_matches_ts() {
        let value = obj! {
            "b" => arr![1, Js::Null, Js::Undefined, 1.5e21, 0.1],
            "a" => obj! { "z-k" => "x\"y", "Z" => true, "_" => obj! {} },
            "é" => f64::NAN,
            "" => -0.0,
        };
        assert_eq!(
            stable_stringify(&value).unwrap(),
            r#"{"":0,"a":{"Z":true,"_":{},"z-k":"x\"y"},"b":[1,null,null,1.5e+21,0.1],"é":null}"#
        );
    }

    #[test]
    fn run_content_for_text_matches_ts() {
        let formatting = obj! { "fontFamily" => obj! { "ascii" => "Symbol", "hAnsi" => "Symbol" } };
        let got = run_content_for_text("a\u{ad}b\u{2011}c\u{f041}d", &formatting).unwrap();
        assert_eq!(
            got.stringify().unwrap(),
            r#"[{"type":"text","text":"a"},{"type":"softHyphen"},{"type":"text","text":"b"},{"type":"noBreakHyphen"},{"type":"text","text":"c"},{"type":"symbol","font":"Symbol","char":"F041"},{"type":"text","text":"d"}]"#
        );
        assert_eq!(
            run_content_for_text("plain", &obj! {})
                .unwrap()
                .stringify()
                .unwrap(),
            r#"[{"type":"text","text":"plain"}]"#
        );
        assert_eq!(
            run_content_for_text("", &obj! {})
                .unwrap()
                .stringify()
                .unwrap(),
            "[]"
        );
    }

    #[test]
    fn attrs_to_text_formatting_matches_ts() {
        let attrs = obj! {
            "bold" => 1, "underline" => obj! { "style" => "" },
            "textColor" => obj! { "rgb" => "FF0000", "themeColor" => Js::Undefined },
            "highlight" => obj! { "color" => "yellow" },
            "fontSize" => obj! { "size" => 24, "sizeCs" => "x" },
            "fontFamily" => obj! { "ascii" => "Arial", "cs" => 3 },
            "superscript" => true, "subscript" => true,
            "characterSpacing" => obj! { "spacing" => 20 },
            "emphasisMark" => obj! {}, "runStyle" => obj! { "styleId" => "Strong" },
            "strike" => obj! { "double" => true },
        };
        let got = attrs_to_text_formatting(&attrs).unwrap();
        assert_eq!(
            got.stringify().unwrap(),
            r#"{"bold":true,"underline":{"style":"single","color":null},"doubleStrike":true,"color":{"rgb":"FF0000","themeColor":null,"themeTint":null,"themeShade":null},"highlight":"yellow","fontSize":24,"fontFamily":{"ascii":"Arial","hAnsi":null,"asciiTheme":null},"vertAlign":"subscript","spacing":20,"emphasisMark":"dot","styleId":"Strong"}"#
        );
        assert_eq!(
            got.get("fontFamily").keys(),
            [
                "ascii",
                "hAnsi",
                "eastAsia",
                "cs",
                "asciiTheme",
                "hAnsiTheme",
                "eastAsiaTheme",
                "csTheme",
                "hint"
            ]
        );
        let second = obj! {
            "underline" => obj! { "inheritedHyperlink" => true }, "highlight" => "cyan",
            "textEffect" => obj! { "effect" => "shimmer" }, "modernTextEffects" => obj! { "effects" => arr![1] },
            "textShadow" => 1, "hidden" => 1,
        };
        assert_eq!(
            attrs_to_text_formatting(&second)
                .unwrap()
                .stringify()
                .unwrap(),
            r#"{"highlight":"cyan","shadow":true,"hidden":true,"effect":"shimmer","modernEffects":[1]}"#
        );
    }

    #[test]
    fn rotation_match() {
        assert_eq!(
            match_rotation("scaleX(-1) rotate(-12.5deg)").as_deref(),
            Some("-12.5")
        );
        assert_eq!(
            match_rotation("rotate(1rad) rotate(3deg)").as_deref(),
            Some("3")
        );
        assert_eq!(match_rotation("rotate(deg)"), None);
    }
}
