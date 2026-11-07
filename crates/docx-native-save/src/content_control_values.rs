//! Port of `packages/docx/src/yrs/contentControlValues.ts`.
#![allow(unused_variables, unused_imports, dead_code, clippy::all)]

use crate::js::{Js, Result, SaveError};
use crate::session::{Inheritance, Session};
use crate::{arr, obj, throw};
use indexmap::IndexMap;
use std::collections::{HashMap, HashSet};
use std::rc::Rc;

fn is_word(b: u8) -> bool {
    b.is_ascii_alphanumeric() || b == b'_'
}

/// `\b` at byte position `i` of `s`.
fn boundary(s: &[u8], i: usize) -> bool {
    let before = i > 0 && is_word(s[i - 1]);
    let after = i < s.len() && is_word(s[i]);
    before != after
}

/// The first match of `<${element}\b[^>]*>` in `xml`, as a byte range.
fn find_open_tag(xml: &str, element: &str) -> Option<(usize, usize)> {
    let b = xml.as_bytes();
    let needle = format!("<{element}");
    let mut from = 0;
    while let Some(i) = xml[from..].find(&needle) {
        let start = from + i;
        let after = start + needle.len();
        if boundary(b, after) {
            if let Some(j) = xml[after..].find('>') {
                return Some((start, after + j + 1));
            }
        }
        from = start + 1;
        while !xml.is_char_boundary(from) {
            from += 1;
        }
    }
    None
}

/// The first match of `\b${attr}="([^"]*)"` in `s`: (match range, value range).
fn find_attr(s: &str, attr: &str) -> Option<((usize, usize), (usize, usize))> {
    let b = s.as_bytes();
    let needle = format!("{attr}=\"");
    let mut from = 0;
    while let Some(i) = s[from..].find(&needle) {
        let start = from + i;
        if boundary(b, start) {
            let vstart = start + needle.len();
            if let Some(j) = s[vstart..].find('"') {
                return Some(((start, vstart + j + 1), (vstart, vstart + j)));
            }
        }
        from = start + 1;
        while !s.is_char_boundary(from) {
            from += 1;
        }
    }
    None
}

/// `readAttr` (src/yrs/contentControlValues.ts:34).
pub(crate) fn read_attr(xml: &str, element: &str, attr: &str) -> Result<Option<String>> {
    let Some((s, e)) = find_open_tag(xml, element) else {
        return Ok(None);
    };
    let el = &xml[s..e];
    Ok(find_attr(el, attr).map(|(_, (vs, ve))| el[vs..ve].to_owned()))
}

/// `setAttr` (src/yrs/contentControlValues.ts:45).
pub(crate) fn set_attr(xml: &str, element: &str, attr: &str, value: &str) -> Result<String> {
    let Some((s, e)) = find_open_tag(xml, element) else {
        return Ok(xml.to_owned());
    };
    let open = &xml[s..e];
    let self_close = open.ends_with("/>") && open.len() >= element.len() + 3;
    let body = &open[1..open.len() - if self_close { 2 } else { 1 }];
    let new_body = match find_attr(body, attr) {
        Some(((ms, me), _)) => format!("{}{attr}=\"{value}\"{}", &body[..ms], &body[me..]),
        None => format!("{body} {attr}=\"{value}\""),
    };
    Ok(format!(
        "{}<{}{}{}",
        &xml[..s],
        new_body,
        if self_close { "/>" } else { ">" },
        &xml[e..]
    ))
}

/// `escapeXmlAttr` (src/yrs/contentControlValues.ts:61).
pub(crate) fn escape_xml_attr(s: &str) -> Result<String> {
    Ok(s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;"))
}

/// `parseInt(s, 16)`.
fn parse_hex(s: &str) -> f64 {
    let t = s.trim_start_matches(crate::font_resolver::is_js_space);
    let (sign, rest) = match t.as_bytes().first() {
        Some(b'-') => (-1.0, &t[1..]),
        Some(b'+') => (1.0, &t[1..]),
        _ => (1.0, t),
    };
    let rest = rest
        .strip_prefix("0x")
        .or_else(|| rest.strip_prefix("0X"))
        .unwrap_or(rest);
    let digits: Vec<u32> = rest.chars().map_while(|c| c.to_digit(16)).collect();
    if digits.is_empty() {
        return f64::NAN;
    }
    sign * digits.iter().fold(0.0, |acc, d| acc * 16.0 + *d as f64)
}

/// `codePointChar` (src/yrs/contentControlValues.ts:70).
pub(crate) fn code_point_char(hex: Option<&str>, fallback: &str) -> Result<String> {
    let Some(hex) = hex.filter(|h| !h.is_empty()) else {
        return Ok(fallback.to_owned());
    };
    let n = parse_hex(hex);
    if n.is_nan() || n < 0.0 || n > 1_114_111.0 || (n >= 55296.0 && n <= 57343.0) {
        return Ok(fallback.to_owned());
    }
    Ok(char::from_u32(n as u32)
        .map(String::from)
        .unwrap_or_else(|| fallback.to_owned()))
}

const MONTHS_LONG: [&str; 12] = [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
];
const MONTHS_SHORT: [&str; 12] = [
    "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
];

/// `formatSdtDate` (src/yrs/contentControlValues.ts:81). Month names are English only
/// (`Intl.DateTimeFormat(language)` is not reproduced for other languages).
pub(crate) fn format_sdt_date(
    iso: &str,
    pattern: Option<&str>,
    language: Option<&str>,
) -> Result<String> {
    let head = crate::js::slice16(iso, 0.0, Some(10.0));
    let parts: Vec<f64> = head.split('-').map(crate::js::parse_int).collect();
    let (y, m, d) = (
        parts.first().copied().unwrap_or(f64::NAN),
        parts.get(1).copied().unwrap_or(f64::NAN),
        parts.get(2).copied().unwrap_or(f64::NAN),
    );
    if !is_valid_iso_date(y, m, d)? {
        return Ok(iso.to_owned());
    }
    let fmt = match pattern {
        Some(p) if !p.trim_matches(crate::font_resolver::is_js_space).is_empty() => p,
        _ => "M/d/yyyy",
    };
    let pad = |n: f64| {
        let s = crate::js::num_to_string(n);
        if crate::js::len16(&s) < 2 {
            format!("0{s}")
        } else {
            s
        }
    };
    let ys = crate::js::num_to_string(y);
    let mi = (m as usize) - 1;
    let tokens: [(&str, String); 8] = [
        ("yyyy", ys.clone()),
        ("yy", crate::js::slice16(&ys, -2.0, None)),
        ("MMMM", MONTHS_LONG[mi].to_owned()),
        ("MMM", MONTHS_SHORT[mi].to_owned()),
        ("MM", pad(m)),
        ("M", crate::js::num_to_string(m)),
        ("dd", pad(d)),
        ("d", crate::js::num_to_string(d)),
    ];
    let mut out = String::new();
    let mut rest = fmt;
    'outer: while !rest.is_empty() {
        for (t, v) in &tokens {
            if let Some(r) = rest.strip_prefix(t) {
                out.push_str(v);
                rest = r;
                continue 'outer;
            }
        }
        let c = rest.chars().next().unwrap();
        out.push(c);
        rest = &rest[c.len_utf8()..];
    }
    Ok(out)
}

fn days_in_month(y: f64, m: f64) -> f64 {
    let leap = (y % 4.0 == 0.0 && y % 100.0 != 0.0) || y % 400.0 == 0.0;
    match m as u32 {
        1 | 3 | 5 | 7 | 8 | 10 | 12 => 31.0,
        4 | 6 | 9 | 11 => 30.0,
        _ => {
            if leap {
                29.0
            } else {
                28.0
            }
        }
    }
}

/// `isValidIsoDate` (src/yrs/contentControlValues.ts:115).
pub(crate) fn is_valid_iso_date(year: f64, month: f64, day: f64) -> Result<bool> {
    let falsy = |n: f64| n == 0.0 || n.is_nan();
    if falsy(year) || falsy(month) || falsy(day) {
        return Ok(false);
    }
    // `Date.UTC` maps years 0..99 to 1900+, and only integer components round-trip.
    if year.fract() != 0.0 || month.fract() != 0.0 || day.fract() != 0.0 {
        return Ok(false);
    }
    if (0.0..=99.0).contains(&year) || year.abs() > 271_821.0 {
        return Ok(false);
    }
    Ok((1.0..=12.0).contains(&month) && day >= 1.0 && day <= days_in_month(year, month))
}

/// `paragraph` (src/yrs/contentControlValues.ts:126).
pub(crate) fn paragraph(text: &str, font: Option<&str>) -> Result<Js> {
    if text.is_empty() {
        return Ok(obj! { "type" => "paragraph", "content" => arr![] });
    }
    let run =
        obj! { "type" => "run", "content" => arr![obj! { "type" => "text", "text" => text }] };
    if let Some(font) = font.filter(|f| !f.is_empty()) {
        run.set("formatting", obj! { "fontFamily" => obj! { "ascii" => font, "hAnsi" => font, "eastAsia" => font, "cs" => font } });
    }
    Ok(obj! { "type" => "paragraph", "content" => arr![run] })
}

/// Removes every match of `<w:showingPlcHdr\b[^>]*\/>` (`paired` false) or
/// `<w:showingPlcHdr\b[^>]*>[\s\S]*?<\/w:showingPlcHdr>` (`paired` true).
fn remove_showing(raw: &str, paired: bool) -> String {
    const OPEN: &str = "<w:showingPlcHdr";
    const CLOSE: &str = "</w:showingPlcHdr>";
    let b = raw.as_bytes();
    let mut out = String::new();
    let mut i = 0;
    let mut copied = 0;
    while let Some(k) = raw[i..].find(OPEN) {
        let start = i + k;
        let after = start + OPEN.len();
        let mut end = None;
        if boundary(b, after) {
            if let Some(j) = raw[after..].find('>') {
                let gt = after + j;
                if !paired {
                    if gt > after && b[gt - 1] == b'/' {
                        end = Some(gt + 1);
                    }
                } else if let Some(c) = raw[gt + 1..].find(CLOSE) {
                    end = Some(gt + 1 + c + CLOSE.len());
                }
            }
        }
        match end {
            Some(e) => {
                out.push_str(&raw[copied..start]);
                copied = e;
                i = e;
            }
            None => i = start + 1,
        }
    }
    out.push_str(&raw[copied..]);
    out
}

/// `clearShowingPlaceholderXml` (src/yrs/contentControlValues.ts:138).
pub(crate) fn clear_showing_placeholder_xml(raw: Option<&str>) -> Result<Option<String>> {
    match raw {
        None => Ok(None),
        Some("") => Ok(Some(String::new())),
        Some(r) => Ok(Some(remove_showing(&remove_showing(r, false), true))),
    }
}

/// `withoutPlaceholder` (src/yrs/contentControlValues.ts:146).
pub(crate) fn without_placeholder(props: &Js, next_raw: &str) -> Result<Js> {
    let cleaned =
        clear_showing_placeholder_xml(Some(next_raw))?.unwrap_or_else(|| next_raw.to_owned());
    Ok(props.with(&[
        ("showingPlaceholder", Js::Bool(false)),
        (
            "rawPropertiesXml",
            if cleaned.is_empty() {
                Js::Undefined
            } else {
                Js::from(cleaned)
            },
        ),
    ]))
}

fn cc_error(message: String) -> SaveError {
    ContentControlValueError { message }.into_error()
}

/// `applyContentControlValue` (src/yrs/contentControlValues.ts:160).
pub(crate) fn apply_content_control_value(props: &Js, value: &Js) -> Result<Js> {
    let raw_js = props.get("rawPropertiesXml").nullish_or(Js::from(""));
    let raw = crate::js::to_js_string(&raw_js);
    let sdt_type = props.get("sdtType");
    let sdt_type_s = crate::js::to_js_string(&sdt_type);
    match value.get("kind").as_str() {
        Some("dropdown") => {
            if sdt_type_s != "dropDownList" && sdt_type_s != "comboBox" || !sdt_type.is_str() {
                return Err(cc_error(format!(
                    "Control is '{sdt_type_s}', not a dropdown/combo box."
                )));
            }
            let wanted = value.get("value");
            let items = props.get("listItems");
            let found = items.items().into_iter().enumerate().find(|(_, it)| {
                it.get("value").strict_eq(&wanted) || it.get("displayText").strict_eq(&wanted)
            });
            if found.is_none() && sdt_type_s == "dropDownList" {
                return Err(cc_error(format!(
                    "'{}' is not one of the control's list items.",
                    crate::js::to_js_string(&wanted)
                )));
            }
            let item = found.as_ref().map(|(_, it)| it.clone()).unwrap_or_default();
            let stored_value = item.get("value").nullish_or(wanted.clone());
            let display_text = item.get("displayText").nullish_or(wanted.clone());
            let selected_index = match &found {
                Some((i, _)) => Js::from(*i),
                None => Js::Undefined,
            };
            let element = if sdt_type_s == "comboBox" {
                "w:comboBox"
            } else {
                "w:dropDownList"
            };
            let next_raw = set_attr(
                &raw,
                element,
                "w:lastValue",
                &escape_xml_attr(&crate::js::to_js_string(&stored_value))?,
            )?;
            let properties = without_placeholder(props, &next_raw)?.spread();
            properties.set("listLastValue", stored_value.clone());
            properties.set(
                "controlState",
                props.get("controlState").with(&[
                    ("value", display_text.clone()),
                    ("selectedValue", stored_value),
                    ("selectedIndex", selected_index),
                    ("placeholder", Js::Bool(false)),
                ]),
            );
            Ok(
                obj! { "properties" => properties, "content" => arr![paragraph(&crate::js::to_js_string(&display_text), None)?] },
            )
        }
        Some("checkbox") => {
            if sdt_type_s != "checkbox" || !sdt_type.is_str() {
                return Err(cc_error(format!(
                    "Control is '{sdt_type_s}', not a checkbox."
                )));
            }
            if read_attr(&raw, "w14:checked", "w14:val")?.is_none() {
                return Err(cc_error(
                    "Checkbox control has no <w14:checked> state to update (not a Word checkbox)."
                        .into(),
                ));
            }
            let checked = value.get("checked").truthy();
            let state_el = if checked {
                "w14:checkedState"
            } else {
                "w14:uncheckedState"
            };
            let ch = code_point_char(
                read_attr(&raw, state_el, "w14:val")?.as_deref(),
                if checked { "☒" } else { "☐" },
            )?;
            let font = read_attr(&raw, state_el, "w14:font")?;
            let next_raw = set_attr(
                &raw,
                "w14:checked",
                "w14:val",
                if checked { "1" } else { "0" },
            )?;
            let properties = without_placeholder(props, &next_raw)?.spread();
            properties.set("checked", value.get("checked"));
            properties.set(
                "controlState",
                props.get("controlState").with(&[
                    ("checked", value.get("checked")),
                    ("value", Js::from(ch.as_str())),
                    ("placeholder", Js::Bool(false)),
                ]),
            );
            Ok(
                obj! { "properties" => properties, "content" => arr![paragraph(&ch, font.as_deref())?] },
            )
        }
        Some("date") => {
            if sdt_type_s != "date" || !sdt_type.is_str() {
                return Err(cc_error(format!(
                    "Control is '{sdt_type_s}', not a date control."
                )));
            }
            let date = crate::js::to_js_string(&value.get("date"));
            let iso = crate::js::slice16(&date, 0.0, Some(10.0));
            let ib = iso.as_bytes();
            let shaped = ib.len() == 10
                && ib.iter().enumerate().all(|(i, c)| {
                    if i == 4 || i == 7 {
                        *c == b'-'
                    } else {
                        c.is_ascii_digit()
                    }
                });
            let (year, month, day) = if shaped {
                (
                    crate::js::parse_float(&iso[0..4]),
                    crate::js::parse_float(&iso[5..7]),
                    crate::js::parse_float(&iso[8..10]),
                )
            } else {
                (f64::NAN, f64::NAN, f64::NAN)
            };
            if !is_valid_iso_date(year, month, day)? {
                return Err(cc_error(format!(
                    "Date must be ISO yyyy-mm-dd, got '{date}'."
                )));
            }
            let full_date = format!("{iso}T00:00:00");
            let next_raw = set_attr(&raw, "w:date", "w:fullDate", &full_date)?;
            let pattern = props
                .path(&["dateState", "format"])
                .nullish_or(props.get("dateFormat"))
                .nullish_or(
                    read_attr(&raw, "w:dateFormat", "w:val")?.map_or(Js::Undefined, Js::from),
                );
            let language = props.path(&["dateState", "language"]);
            let pattern_s = pattern.as_str().map(str::to_owned);
            let language_s = language.as_str().map(str::to_owned);
            let display = format_sdt_date(&iso, pattern_s.as_deref(), language_s.as_deref())?;
            let properties = without_placeholder(props, &next_raw)?.spread();
            properties.set("dateFormat", pattern.clone());
            properties.set(
                "dateState",
                props.get("dateState").with(&[
                    ("fullDate", Js::from(full_date.as_str())),
                    ("format", pattern),
                ]),
            );
            properties.set(
                "controlState",
                props.get("controlState").with(&[
                    ("value", Js::from(iso.as_str())),
                    ("placeholder", Js::Bool(false)),
                ]),
            );
            Ok(obj! { "properties" => properties, "content" => arr![paragraph(&display, None)?] })
        }
        _ => Ok(Js::Undefined),
    }
}

/// `class ContentControlValueError`.
pub(crate) struct ContentControlValueError {
    pub(crate) message: String,
}

impl ContentControlValueError {
    /// `ContentControlValueError.constructor` (src/yrs/contentControlValues.ts:25).
    pub(crate) fn new(message: &str) -> Result<Self> {
        Ok(Self {
            message: message.to_owned(),
        })
    }

    /// The thrown error.
    pub(crate) fn into_error(self) -> SaveError {
        SaveError::new(self.message)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn date_and_attr_helpers_match_ts() {
        assert_eq!(
            format_sdt_date("2024-03-05", Some("dd MMMM yyyy (MMM, M/d/yy)"), None).unwrap(),
            "05 March 2024 (Mar, 3/5/24)"
        );
        assert_eq!(
            format_sdt_date("2023-02-29", None, None).unwrap(),
            "2023-02-29"
        );
        assert_eq!(
            format_sdt_date("0050-01-01", None, None).unwrap(),
            "0050-01-01"
        );
        assert_eq!(
            format_sdt_date("2024-1-9xyz", Some("  "), None).unwrap(),
            "1/9/2024"
        );
        assert_eq!(
            set_attr(r#"<w:sdtPr><w:dropDownList w:lastValue="a"><w:listItem/></w:dropDownList></w:sdtPr>"#, "w:dropDownList", "w:lastValue", "b&amp;").unwrap(),
            r#"<w:sdtPr><w:dropDownList w:lastValue="b&amp;"><w:listItem/></w:dropDownList></w:sdtPr>"#
        );
        assert_eq!(
            set_attr("<w:date/>", "w:date", "w:fullDate", "x").unwrap(),
            r#"<w:date w:fullDate="x"/>"#
        );
        assert_eq!(code_point_char(Some("2612"), "?").unwrap(), "☒");
        assert_eq!(code_point_char(Some("d800"), "?").unwrap(), "?");
        assert_eq!(
            clear_showing_placeholder_xml(Some(
                "<a/><w:showingPlcHdr/><b/><w:showingPlcHdr>x</w:showingPlcHdr>"
            ))
            .unwrap()
            .as_deref(),
            Some("<a/><b/>")
        );
    }

    #[test]
    fn apply_values_match_ts() {
        let props = Js::parse(r#"{"sdtType":"checkbox","rawPropertiesXml":"<w:sdtPr><w:showingPlcHdr/><w14:checkbox><w14:checked w14:val=\"0\"/><w14:checkedState w14:val=\"2612\" w14:font=\"MS Gothic\"/></w14:checkbox></w:sdtPr>","controlState":{"x":1}}"#).unwrap();
        let out = apply_content_control_value(
            &props,
            &Js::parse(r#"{"kind":"checkbox","checked":true}"#).unwrap(),
        )
        .unwrap();
        assert_eq!(
            out.stringify().unwrap(),
            r#"{"properties":{"sdtType":"checkbox","rawPropertiesXml":"<w:sdtPr><w14:checkbox><w14:checked w14:val=\"1\"/><w14:checkedState w14:val=\"2612\" w14:font=\"MS Gothic\"/></w14:checkbox></w:sdtPr>","controlState":{"x":1,"checked":true,"value":"☒","placeholder":false},"showingPlaceholder":false,"checked":true},"content":[{"type":"paragraph","content":[{"type":"run","content":[{"type":"text","text":"☒"}],"formatting":{"fontFamily":{"ascii":"MS Gothic","hAnsi":"MS Gothic","eastAsia":"MS Gothic","cs":"MS Gothic"}}}]}]}"#
        );
        let dd = Js::parse(r#"{"sdtType":"comboBox","listItems":[{"value":"a","displayText":"A"},{"value":"b","displayText":"B"}],"rawPropertiesXml":"<w:comboBox/>"}"#).unwrap();
        let out = apply_content_control_value(
            &dd,
            &Js::parse(r#"{"kind":"dropdown","value":"B"}"#).unwrap(),
        )
        .unwrap();
        assert_eq!(
            out.stringify().unwrap(),
            r#"{"properties":{"sdtType":"comboBox","listItems":[{"value":"a","displayText":"A"},{"value":"b","displayText":"B"}],"rawPropertiesXml":"<w:comboBox w:lastValue=\"b\"/>","showingPlaceholder":false,"listLastValue":"b","controlState":{"value":"B","selectedValue":"b","selectedIndex":1,"placeholder":false}},"content":[{"type":"paragraph","content":[{"type":"run","content":[{"type":"text","text":"B"}]}]}]}"#
        );
        let date = Js::parse(r#"{"sdtType":"date","rawPropertiesXml":"<w:date><w:dateFormat w:val=\"MMMM d, yyyy\"/></w:date>"}"#).unwrap();
        let out = apply_content_control_value(
            &date,
            &Js::parse(r#"{"kind":"date","date":"2024-12-01T10:00"}"#).unwrap(),
        )
        .unwrap();
        assert_eq!(
            out.stringify().unwrap(),
            r#"{"properties":{"sdtType":"date","rawPropertiesXml":"<w:date w:fullDate=\"2024-12-01T00:00:00\"><w:dateFormat w:val=\"MMMM d, yyyy\"/></w:date>","showingPlaceholder":false,"dateFormat":"MMMM d, yyyy","dateState":{"fullDate":"2024-12-01T00:00:00","format":"MMMM d, yyyy"},"controlState":{"value":"2024-12-01","placeholder":false}},"content":[{"type":"paragraph","content":[{"type":"run","content":[{"type":"text","text":"December 1, 2024"}]}]}]}"#
        );
        assert!(
            apply_content_control_value(
                &date,
                &Js::parse(r#"{"kind":"date","date":"2024-13-01"}"#).unwrap()
            )
            .is_err()
        );
    }
}
