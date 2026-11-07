//! Port of `packages/docx/src/utils/fontResolver.ts`.
#![allow(unused_variables, unused_imports, dead_code, clippy::all)]

use crate::js::{Js, Result, SaveError};
use crate::session::{Inheritance, Session};
use crate::{arr, obj, throw};
use indexmap::IndexMap;
use std::collections::{HashMap, HashSet};
use std::rc::Rc;

pub(crate) const DEFAULT_SINGLE_LINE_RATIO: f64 = 1.15;

/// `FONT_MAPPINGS`: (key, googleFont, category, fallbackStack, singleLineRatio).
const FONT_MAPPINGS: &[(&str, &str, &str, &[&str], f64)] = &[
    (
        "calibri",
        "Carlito",
        "sans-serif",
        &["Calibri", "Carlito", "Arial", "Helvetica", "sans-serif"],
        1.2207,
    ),
    (
        "cambria",
        "Caladea",
        "serif",
        &["Cambria", "Caladea", "Georgia", "serif"],
        1.2676,
    ),
    (
        "arial",
        "Arimo",
        "sans-serif",
        &["Arial", "Arimo", "Helvetica", "sans-serif"],
        1.1172,
    ),
    (
        "times new roman",
        "Tinos",
        "serif",
        &["Times New Roman", "Tinos", "Times", "serif"],
        1.1074,
    ),
    (
        "courier new",
        "Cousine",
        "monospace",
        &["Courier New", "Cousine", "Courier", "monospace"],
        1.1328,
    ),
    (
        "georgia",
        "Tinos",
        "serif",
        &["Georgia", "Tinos", "Times New Roman", "serif"],
        1.1362,
    ),
    (
        "verdana",
        "Open Sans",
        "sans-serif",
        &["Verdana", "Open Sans", "Arial", "sans-serif"],
        1.2153,
    ),
    (
        "tahoma",
        "Open Sans",
        "sans-serif",
        &["Tahoma", "Open Sans", "Arial", "sans-serif"],
        1.2075,
    ),
    (
        "trebuchet ms",
        "Fira Sans",
        "sans-serif",
        &["Trebuchet MS", "Fira Sans", "Arial", "sans-serif"],
        1.1431,
    ),
    (
        "comic sans ms",
        "Comic Neue",
        "cursive",
        &["Comic Sans MS", "Comic Neue", "cursive"],
        1.3936,
    ),
    (
        "impact",
        "Anton",
        "sans-serif",
        &["Impact", "Anton", "Arial Black", "sans-serif"],
        1.2197,
    ),
    (
        "palatino linotype",
        "EB Garamond",
        "serif",
        &[
            "Palatino Linotype",
            "EB Garamond",
            "Palatino",
            "Georgia",
            "serif",
        ],
        1.0259,
    ),
    (
        "book antiqua",
        "EB Garamond",
        "serif",
        &[
            "Book Antiqua",
            "EB Garamond",
            "Palatino",
            "Georgia",
            "serif",
        ],
        1.0259,
    ),
    (
        "garamond",
        "EB Garamond",
        "serif",
        &["Garamond", "EB Garamond", "Georgia", "serif"],
        1.068,
    ),
    (
        "century gothic",
        "Questrial",
        "sans-serif",
        &["Century Gothic", "Questrial", "Arial", "sans-serif"],
        1.1611,
    ),
    (
        "lucida sans",
        "Open Sans",
        "sans-serif",
        &["Lucida Sans", "Open Sans", "Arial", "sans-serif"],
        1.1655,
    ),
    (
        "lucida console",
        "Cousine",
        "monospace",
        &["Lucida Console", "Cousine", "Courier New", "monospace"],
        1.1387,
    ),
    (
        "consolas",
        "Inconsolata",
        "monospace",
        &[
            "Consolas",
            "Inconsolata",
            "Cousine",
            "Courier New",
            "monospace",
        ],
        1.1626,
    ),
    (
        "ms mincho",
        "Noto Serif JP",
        "serif",
        &["MS Mincho", "Noto Serif JP", "serif"],
        DEFAULT_SINGLE_LINE_RATIO,
    ),
    (
        "ｍｓ 明朝",
        "Noto Serif JP",
        "serif",
        &["MS Mincho", "ＭＳ 明朝", "Noto Serif JP", "serif"],
        DEFAULT_SINGLE_LINE_RATIO,
    ),
    (
        "ｍｓ ｐ明朝",
        "Noto Serif JP",
        "serif",
        &["MS PMincho", "ＭＳ Ｐ明朝", "Noto Serif JP", "serif"],
        DEFAULT_SINGLE_LINE_RATIO,
    ),
    (
        "ms gothic",
        "Noto Sans JP",
        "sans-serif",
        &["MS Gothic", "Noto Sans JP", "sans-serif"],
        DEFAULT_SINGLE_LINE_RATIO,
    ),
    (
        "ｍｓ ゴシック",
        "Noto Sans JP",
        "sans-serif",
        &["MS Gothic", "ＭＳ ゴシック", "Noto Sans JP", "sans-serif"],
        DEFAULT_SINGLE_LINE_RATIO,
    ),
    (
        "ｍｓ ｐゴシック",
        "Noto Sans JP",
        "sans-serif",
        &[
            "MS PGothic",
            "ＭＳ Ｐゴシック",
            "Noto Sans JP",
            "sans-serif",
        ],
        DEFAULT_SINGLE_LINE_RATIO,
    ),
    (
        "simhei",
        "Noto Sans SC",
        "sans-serif",
        &["SimHei", "Noto Sans SC", "sans-serif"],
        DEFAULT_SINGLE_LINE_RATIO,
    ),
    (
        "simsun",
        "Noto Serif SC",
        "serif",
        &["SimSun", "Noto Serif SC", "serif"],
        DEFAULT_SINGLE_LINE_RATIO,
    ),
    (
        "malgun gothic",
        "Noto Sans KR",
        "sans-serif",
        &["Malgun Gothic", "Noto Sans KR", "sans-serif"],
        DEFAULT_SINGLE_LINE_RATIO,
    ),
    (
        "宋体",
        "Noto Serif SC",
        "serif",
        &["SimSun", "宋体", "Noto Serif SC", "serif"],
        DEFAULT_SINGLE_LINE_RATIO,
    ),
    (
        "黑体",
        "Noto Sans SC",
        "sans-serif",
        &["SimHei", "黑体", "Noto Sans SC", "sans-serif"],
        DEFAULT_SINGLE_LINE_RATIO,
    ),
    (
        "微软雅黑",
        "Noto Sans SC",
        "sans-serif",
        &["Microsoft YaHei", "微软雅黑", "Noto Sans SC", "sans-serif"],
        DEFAULT_SINGLE_LINE_RATIO,
    ),
    (
        "等线",
        "Noto Sans SC",
        "sans-serif",
        &["DengXian", "等线", "Noto Sans SC", "sans-serif"],
        DEFAULT_SINGLE_LINE_RATIO,
    ),
    (
        "仿宋",
        "Noto Serif SC",
        "serif",
        &["FangSong", "仿宋", "Noto Serif SC", "serif"],
        DEFAULT_SINGLE_LINE_RATIO,
    ),
    (
        "楷体",
        "Noto Serif SC",
        "serif",
        &["KaiTi", "楷体", "Noto Serif SC", "serif"],
        DEFAULT_SINGLE_LINE_RATIO,
    ),
    (
        "新細明體",
        "Noto Serif TC",
        "serif",
        &["PMingLiU", "新細明體", "Noto Serif TC", "serif"],
        DEFAULT_SINGLE_LINE_RATIO,
    ),
    (
        "細明體",
        "Noto Serif TC",
        "serif",
        &["MingLiU", "細明體", "Noto Serif TC", "serif"],
        DEFAULT_SINGLE_LINE_RATIO,
    ),
    (
        "微軟正黑體",
        "Noto Sans TC",
        "sans-serif",
        &[
            "Microsoft JhengHei",
            "微軟正黑體",
            "Noto Sans TC",
            "sans-serif",
        ],
        DEFAULT_SINGLE_LINE_RATIO,
    ),
    (
        "標楷體",
        "Noto Serif TC",
        "serif",
        &["DFKai-SB", "標楷體", "Noto Serif TC", "serif"],
        DEFAULT_SINGLE_LINE_RATIO,
    ),
    (
        "맑은 고딕",
        "Noto Sans KR",
        "sans-serif",
        &["Malgun Gothic", "맑은 고딕", "Noto Sans KR", "sans-serif"],
        DEFAULT_SINGLE_LINE_RATIO,
    ),
    (
        "굴림",
        "Noto Sans KR",
        "sans-serif",
        &["Gulim", "굴림", "Noto Sans KR", "sans-serif"],
        DEFAULT_SINGLE_LINE_RATIO,
    ),
    (
        "돋움",
        "Noto Sans KR",
        "sans-serif",
        &["Dotum", "돋움", "Noto Sans KR", "sans-serif"],
        DEFAULT_SINGLE_LINE_RATIO,
    ),
    (
        "바탕",
        "Noto Serif KR",
        "serif",
        &["Batang", "바탕", "Noto Serif KR", "serif"],
        DEFAULT_SINGLE_LINE_RATIO,
    ),
    (
        "궁서",
        "Noto Serif KR",
        "serif",
        &["Gungsuh", "궁서", "Noto Serif KR", "serif"],
        DEFAULT_SINGLE_LINE_RATIO,
    ),
];

/// `CJK_FONT_ALIASES`.
const CJK_FONT_ALIASES: &[(&str, &str)] = &[
    ("microsoft yahei", "微软雅黑"),
    ("dengxian", "等线"),
    ("fangsong", "仿宋"),
    ("kaiti", "楷体"),
    ("pmingliu", "新細明體"),
    ("mingliu", "細明體"),
    ("microsoft jhenghei", "微軟正黑體"),
    ("dfkai-sb", "標楷體"),
    ("gulim", "굴림"),
    ("dotum", "돋움"),
    ("batang", "바탕"),
    ("gungsuh", "궁서"),
    ("meiryo", "ms gothic"),
    ("メイリオ", "ms gothic"),
    ("yu gothic", "ms gothic"),
    ("游ゴシック", "ms gothic"),
    ("yu mincho", "ms mincho"),
    ("游明朝", "ms mincho"),
];

/// `DEFAULT_FALLBACKS`.
fn default_fallback(category: &str) -> &'static str {
    match category {
        "sans-serif" => "Arial, Helvetica, sans-serif",
        "serif" => "Times New Roman, Times, serif",
        "monospace" => "Courier New, Courier, monospace",
        "cursive" => "cursive",
        "fantasy" => "fantasy",
        "system-ui" => "system-ui, sans-serif",
        _ => "",
    }
}

/// JS `\s` / `String.prototype.trim` whitespace.
pub(crate) fn is_js_space(c: char) -> bool {
    matches!(
        c,
        '\u{9}' | '\u{a}' | '\u{b}' | '\u{c}' | '\u{d}' | ' ' | '\u{a0}' | '\u{1680}' | '\u{2000}'
            ..='\u{200a}'
                | '\u{2028}'
                | '\u{2029}'
                | '\u{202f}'
                | '\u{205f}'
                | '\u{3000}'
                | '\u{feff}'
    )
}

fn js_trim(s: &str) -> &str {
    s.trim_matches(is_js_space)
}

fn mapping(
    key: &str,
) -> Option<&'static (
    &'static str,
    &'static str,
    &'static str,
    &'static [&'static str],
    f64,
)> {
    FONT_MAPPINGS.iter().find(|m| m.0 == key)
}

/// `/sans[ -]?serif/`.
fn has_sans_serif(lower: &str) -> bool {
    let mut from = 0;
    while let Some(i) = lower[from..].find("sans") {
        let rest = &lower[from + i + 4..];
        if rest.starts_with("serif") || rest.starts_with(" serif") || rest.starts_with("-serif") {
            return true;
        }
        from += i + 4;
    }
    false
}

/// `detectFontCategory` (src/utils/fontResolver.ts:354).
pub(crate) fn detect_font_category(font_name: &str) -> Result<Js> {
    let lower = font_name.to_lowercase();
    let has = |n: &str| lower.contains(n);
    if has("mono")
        || has("courier")
        || has("consolas")
        || has("console")
        || has("code")
        || has("terminal")
    {
        return Ok(Js::from("monospace"));
    }
    if has_sans_serif(&lower) {
        return Ok(Js::from("sans-serif"));
    }
    if [
        "times",
        "georgia",
        "garamond",
        "palatino",
        "baskerville",
        "bodoni",
        "cambria",
        "mincho",
        "明朝",
        "明體",
        "宋",
        "ming",
        "song",
        "serif",
    ]
    .iter()
    .any(|n| has(n))
    {
        return Ok(Js::from("serif"));
    }
    if ["script", "cursive", "comic", "brush", "hand"]
        .iter()
        .any(|n| has(n))
    {
        return Ok(Js::from("cursive"));
    }
    Ok(Js::from("sans-serif"))
}

/// `resolveFontFamily` (src/utils/fontResolver.ts:446).
pub(crate) fn resolve_font_family(docx_font_name: &str) -> Result<Js> {
    let normalized_name = js_trim(docx_font_name).to_lowercase();
    let key = CJK_FONT_ALIASES
        .iter()
        .find(|a| a.0 == normalized_name)
        .map_or(normalized_name.as_str(), |a| a.1);
    if let Some(m) = mapping(key) {
        let stack: Vec<String> =
            m.3.iter()
                .map(|f| quote_font_name(f))
                .collect::<Result<_>>()?;
        return Ok(obj! {
            "googleFont" => m.1,
            "cssFallback" => stack.join(", "),
            "originalFont" => docx_font_name,
            "hasGoogleEquivalent" => true,
            "singleLineRatio" => m.4,
        });
    }
    let category = detect_font_category(docx_font_name)?;
    let default = default_fallback(category.str_or_empty());
    Ok(obj! {
        "googleFont" => Js::Null,
        "cssFallback" => format!("{}, {}", quote_font_name(docx_font_name)?, default),
        "originalFont" => docx_font_name,
        "hasGoogleEquivalent" => false,
        "singleLineRatio" => DEFAULT_SINGLE_LINE_RATIO,
    })
}

/// `escapeQuotedFontName` (src/utils/fontResolver.ts:484).
pub(crate) fn escape_quoted_font_name(font_name: &str) -> Result<String> {
    let mut out = String::with_capacity(font_name.len());
    for c in font_name.chars() {
        match c {
            '"' | '\\' => {
                out.push('\\');
                out.push(c);
            }
            '\n' => out.push_str("\\a "),
            '\r' => out.push_str("\\d "),
            '\u{c}' => out.push_str("\\c "),
            _ => out.push(c),
        }
    }
    Ok(out)
}

/// `quoteFontName` (src/utils/fontResolver.ts:491).
pub(crate) fn quote_font_name(font_name: &str) -> Result<String> {
    let lower = font_name.to_lowercase();
    if [
        "serif",
        "sans-serif",
        "monospace",
        "cursive",
        "fantasy",
        "system-ui",
    ]
    .contains(&lower.as_str())
    {
        return Ok(font_name.to_owned());
    }
    if font_name
        .chars()
        .any(|c| is_js_space(c) || matches!(c, ',' | '\'' | '"' | '(' | ')'))
    {
        return Ok(format!("\"{}\"", escape_quoted_font_name(font_name)?));
    }
    Ok(font_name.to_owned())
}

/// `resolveThemeFont` (src/utils/fontResolver.ts:516).
pub(crate) fn resolve_theme_font(theme_ref: &str, font_scheme: &Js) -> Result<Option<String>> {
    if !font_scheme.truthy() {
        return Ok(None);
    }
    let ref_ = theme_ref.to_lowercase();
    let is_major = ref_.starts_with("major");
    let theme_font = if is_major {
        font_scheme.get("majorFont")
    } else {
        font_scheme.get("minorFont")
    };
    if !theme_font.truthy() {
        return Ok(None);
    }
    let pick = |v: Js| -> Option<String> {
        match v.nullish_or(Js::Null) {
            Js::Null | Js::Undefined => None,
            other => Some(crate::js::to_js_string(&other)),
        }
    };
    if ref_.contains("eastasia") || ref_.contains("ea") {
        return Ok(pick(
            theme_font.get("ea").nullish_or(theme_font.get("latin")),
        ));
    }
    if ref_.contains("cs") || ref_.contains("bidi") {
        return Ok(pick(
            theme_font.get("cs").nullish_or(theme_font.get("latin")),
        ));
    }
    Ok(pick(theme_font.get("latin")))
}

/// `getGoogleFontsToLoad` (src/utils/fontResolver.ts:550).
pub(crate) fn get_google_fonts_to_load(docx_fonts: &[String]) -> Result<Vec<String>> {
    let mut fonts: Vec<String> = Vec::new();
    for font in docx_fonts {
        let resolved = resolve_font_family(font)?;
        if let Some(g) = resolved.get("googleFont").as_str() {
            if !g.is_empty() && !fonts.iter().any(|f| f == g) {
                fonts.push(g.to_owned());
            }
        }
    }
    Ok(fonts)
}

/// `buildFontFamilyString` (src/utils/fontResolver.ts:570). Pass `"sans-serif"` for the TS default.
pub(crate) fn build_font_family_string(fonts: &[String], category: &Js) -> Result<String> {
    let category = if category.is_undefined() {
        "sans-serif".to_owned()
    } else {
        crate::js::to_js_string(category)
    };
    let mut parts: Vec<String> = fonts
        .iter()
        .map(|f| quote_font_name(f))
        .collect::<Result<_>>()?;
    if !parts.iter().any(|p| p.to_lowercase() == category) {
        parts.push(category);
    }
    Ok(parts.join(", "))
}

/// `getGoogleFontEquivalent` (src/utils/fontResolver.ts:591).
pub(crate) fn get_google_font_equivalent(docx_font_name: &str) -> Result<Option<String>> {
    Ok(resolve_font_family(docx_font_name)?
        .get("googleFont")
        .as_str()
        .map(str::to_owned))
}

/// `hasGoogleFontEquivalent` (src/utils/fontResolver.ts:602).
pub(crate) fn has_google_font_equivalent(docx_font_name: &str) -> Result<bool> {
    let normalized_name = js_trim(docx_font_name).to_lowercase();
    Ok(mapping(&normalized_name).is_some())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn resolves_like_ts() {
        let scheme = Js::parse(r#"{"majorFont":{"latin":"Calibri Light","ea":""},"minorFont":{"latin":"Calibri","cs":"Arial"}}"#).unwrap();
        assert_eq!(
            resolve_theme_font("majorEastAsia", &scheme)
                .unwrap()
                .as_deref(),
            Some("")
        );
        assert_eq!(
            resolve_theme_font("minorBidi", &scheme).unwrap().as_deref(),
            Some("Arial")
        );
        assert_eq!(
            resolve_theme_font("minorHAnsi", &scheme)
                .unwrap()
                .as_deref(),
            Some("Calibri")
        );
        assert_eq!(
            resolve_theme_font("majorHAnsi", &Js::Undefined).unwrap(),
            None
        );
        assert_eq!(
            resolve_font_family(" PMingLiU ")
                .unwrap()
                .stringify()
                .unwrap(),
            r#"{"googleFont":"Noto Serif TC","cssFallback":"PMingLiU, 新細明體, \"Noto Serif TC\", serif","originalFont":" PMingLiU ","hasGoogleEquivalent":true,"singleLineRatio":1.15}"#
        );
        assert_eq!(
            resolve_font_family("My \"Sans Serif\"")
                .unwrap()
                .stringify()
                .unwrap(),
            r#"{"googleFont":null,"cssFallback":"\"My \\\"Sans Serif\\\"\", Arial, Helvetica, sans-serif","originalFont":"My \"Sans Serif\"","hasGoogleEquivalent":false,"singleLineRatio":1.15}"#
        );
        assert_eq!(
            detect_font_category("Fira Code").unwrap().as_str(),
            Some("monospace")
        );
        assert_eq!(
            detect_font_category("SongTi").unwrap().as_str(),
            Some("serif")
        );
        assert_eq!(
            build_font_family_string(
                &["Arial".into(), "Times New Roman".into()],
                &Js::from("serif")
            )
            .unwrap(),
            "Arial, \"Times New Roman\", serif"
        );
        assert_eq!(
            get_google_fonts_to_load(&[
                "Calibri".into(),
                "Georgia".into(),
                "Times New Roman".into(),
                "X".into()
            ])
            .unwrap(),
            vec!["Carlito", "Tinos"]
        );
        assert!(!has_google_font_equivalent("pmingliu").unwrap());
    }
}
