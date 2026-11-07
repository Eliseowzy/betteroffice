//! Port of `packages/docx/src/styles/styleResolver.ts`.
#![allow(unused_variables, unused_imports, dead_code, clippy::all)]

use crate::js::{Js, Result, SaveError};
use crate::paragraph_formatting_merge::merge_paragraph_formatting;
use crate::session::{Inheritance, Session};
use crate::text_formatting_merge::{merge_style_run_formatting, merge_text_formatting};
use crate::{arr, obj, throw};
use indexmap::IndexMap;
use std::collections::{HashMap, HashSet};
use std::rc::Rc;

thread_local! {
    /// `BUILTIN_NORMAL_STYLE` (one object, compared by identity).
    static BUILTIN_NORMAL_STYLE: Js = obj! {
        "styleId" => "Normal",
        "type" => "paragraph",
        "name" => "Normal",
        "default" => true,
        "pPr" => obj! { "spaceAfter" => 160, "lineSpacing" => 259, "lineSpacingRule" => "auto" },
    };
}

fn builtin_normal_style() -> Js {
    BUILTIN_NORMAL_STYLE.with(Js::clone)
}

/// `createStyleResolver` (src/styles/styleResolver.ts:328).
pub(crate) fn create_style_resolver(
    style_definitions: &Js,
) -> Result<crate::style_resolver::StyleResolver> {
    StyleResolver::new(style_definitions)
}

/// `class StyleResolver`.
pub(crate) struct StyleResolver {
    styles_by_id: IndexMap<String, Js>,
    doc_defaults: Js,
    default_paragraph_style: Js,
    default_table_style: Js,
    default_character_style: Js,
}

/// `a.localeCompare(b)` — approximated: case-insensitive first, lowercase before uppercase on ties.
fn locale_compare(a: &str, b: &str) -> std::cmp::Ordering {
    let primary = a.to_lowercase().cmp(&b.to_lowercase());
    if primary != std::cmp::Ordering::Equal {
        return primary;
    }
    for (x, y) in a.chars().zip(b.chars()) {
        if x != y {
            return if x.is_lowercase() {
                std::cmp::Ordering::Less
            } else {
                std::cmp::Ordering::Greater
            };
        }
    }
    a.len().cmp(&b.len())
}

fn sort_styles(styles: &mut Vec<Js>) {
    styles.sort_by(|a, b| {
        let pa = crate::js::to_number(&a.get("uiPriority").nullish_or(Js::Num(99.0)));
        let pb = crate::js::to_number(&b.get("uiPriority").nullish_or(Js::Num(99.0)));
        if pa != pb {
            return pa.partial_cmp(&pb).unwrap_or(std::cmp::Ordering::Equal);
        }
        let na = crate::js::to_js_string(&a.get("name").nullish_or(a.get("styleId")));
        let nb = crate::js::to_js_string(&b.get("name").nullish_or(b.get("styleId")));
        locale_compare(&na, &nb)
    });
}

impl StyleResolver {
    /// `StyleResolver.constructor` (src/styles/styleResolver.ts:68).
    pub(crate) fn new(style_definitions: &Js) -> Result<Self> {
        let mut styles_by_id = IndexMap::new();
        let doc_defaults = style_definitions.get("docDefaults");
        let styles = style_definitions.get("styles");
        if styles.truthy() {
            for style in styles.items() {
                let id = style.get("styleId");
                if id.truthy() {
                    styles_by_id.insert(crate::js::to_js_string(&id), style);
                }
            }
        }
        let mut this = StyleResolver {
            styles_by_id,
            doc_defaults,
            default_paragraph_style: Js::Undefined,
            default_table_style: Js::Undefined,
            default_character_style: Js::Undefined,
        };
        this.default_paragraph_style = this.find_default_style("paragraph")?;
        this.default_table_style = this.find_default_style("table")?;
        this.default_character_style = this.find_default_style("character")?;
        Ok(this)
    }

    /// `StyleResolver.getStyle` (src/styles/styleResolver.ts:90).
    pub(crate) fn get_style(&self, style_id: &str) -> Result<Js> {
        Ok(self.styles_by_id.get(style_id).cloned().unwrap_or_default())
    }

    /// `StyleResolver.getParagraphStyles` (src/styles/styleResolver.ts:97).
    pub(crate) fn get_paragraph_styles(&self) -> Result<Js> {
        self.styles_of_type("paragraph")
    }

    fn styles_of_type(&self, kind: &str) -> Result<Js> {
        let mut styles: Vec<Js> = self
            .styles_by_id
            .values()
            .filter(|s| {
                s.get("type").as_str() == Some(kind)
                    && !s.get("hidden").truthy()
                    && !s.get("semiHidden").truthy()
            })
            .cloned()
            .collect();
        sort_styles(&mut styles);
        Ok(Js::new_arr(styles))
    }

    /// `StyleResolver.hasParagraphStyle` (src/styles/styleResolver.ts:119).
    pub(crate) fn has_paragraph_style(&self, style_id: &str) -> Result<bool> {
        Ok(self.get_style(style_id)?.get("type").as_str() == Some("paragraph"))
    }

    /// `StyleResolver.resolveParagraphStyle` (src/styles/styleResolver.ts:130).
    pub(crate) fn resolve_paragraph_style(
        &self,
        style_id: Option<&str>,
        table_paragraph_formatting: &Js,
        table_run_formatting: &Js,
    ) -> Result<Js> {
        let result = Js::new_obj();
        let ppr = self.doc_defaults.get("pPr");
        if ppr.truthy() {
            result.set("paragraphFormatting", ppr.spread());
        }
        let rpr = self.doc_defaults.get("rPr");
        if rpr.truthy() {
            result.set("runFormatting", rpr.spread());
        }
        result.set(
            "runFormatting",
            merge_style_run_formatting(&result.get("runFormatting"), table_run_formatting)?,
        );

        let style = match style_id {
            Some(id) if !id.is_empty() => self.styles_by_id.get(id).cloned().unwrap_or_default(),
            _ => Js::Undefined,
        }
        .nullish_or(self.default_paragraph_style.clone());
        let builtin = builtin_normal_style();
        if style.strict_eq(&builtin) {
            self.merge_style_into_result(&result, &style)?;
        }
        if table_paragraph_formatting.truthy() {
            result.set(
                "paragraphFormatting",
                merge_paragraph_formatting(
                    &result.get("paragraphFormatting"),
                    table_paragraph_formatting,
                )?,
            );
        }
        if style.truthy() && !style.strict_eq(&builtin) {
            self.merge_style_into_result(&result, &style)?;
        }
        Ok(result)
    }

    /// `StyleResolver.getNextStyleId` (src/styles/styleResolver.ts:168).
    pub(crate) fn get_next_style_id(&self, style_id: Option<&str>) -> Result<Option<String>> {
        let Some(style_id) = style_id.filter(|s| !s.is_empty()) else {
            return Ok(None);
        };
        let next = self.get_style(style_id)?.get("next");
        Ok(if next.truthy() && next.as_str() != Some(style_id) {
            Some(crate::js::to_js_string(&next))
        } else {
            None
        })
    }

    /// `StyleResolver.getTableStyles` (src/styles/styleResolver.ts:177).
    pub(crate) fn get_table_styles(&self) -> Result<Js> {
        self.styles_of_type("table")
    }

    /// `StyleResolver.resolveRunStyle` (src/styles/styleResolver.ts:198).
    pub(crate) fn resolve_run_style(&self, style_id: Option<&str>) -> Result<Js> {
        let mut result = Js::new_obj();
        let rpr = self.doc_defaults.get("rPr");
        if rpr.truthy() {
            result = rpr.spread();
        }
        let crpr = self.default_character_style.get("rPr");
        if crpr.truthy() {
            result = self
                .merge_text_formatting(&result, &crpr)?
                .nullish_or(result.clone());
        }
        let nonempty = |r: Js| if r.key_count() > 0 { r } else { Js::Undefined };
        let Some(style_id) = style_id.filter(|s| !s.is_empty()) else {
            return Ok(nonempty(result));
        };
        let style_rpr = self.get_style(style_id)?.get("rPr");
        if !style_rpr.truthy() {
            return Ok(nonempty(result));
        }
        let merged = self.merge_text_formatting(&result, &style_rpr)?;
        Ok(if merged.truthy() && merged.key_count() > 0 {
            merged
        } else {
            Js::Undefined
        })
    }

    /// `StyleResolver.getRunStyleOwnProperties` (src/styles/styleResolver.ts:233).
    pub(crate) fn get_run_style_own_properties(&self, style_id: Option<&str>) -> Result<Js> {
        let Some(style_id) = style_id.filter(|s| !s.is_empty()) else {
            return Ok(Js::Undefined);
        };
        let rpr = self.get_style(style_id)?.get("rPr");
        if !rpr.truthy() {
            return Ok(Js::Undefined);
        }
        Ok(if rpr.keys().len() > 0 {
            rpr.spread()
        } else {
            Js::Undefined
        })
    }

    /// `StyleResolver.getDocDefaults` (src/styles/styleResolver.ts:245).
    pub(crate) fn get_doc_defaults(&self) -> Result<Js> {
        Ok(self.doc_defaults.clone())
    }

    /// `StyleResolver.getDefaultParagraphStyle` (src/styles/styleResolver.ts:252).
    pub(crate) fn get_default_paragraph_style(&self) -> Result<Js> {
        Ok(self.default_paragraph_style.clone())
    }

    /// `StyleResolver.getDefaultTableStyle` (src/styles/styleResolver.ts:264).
    pub(crate) fn get_default_table_style(&self) -> Result<Js> {
        Ok(self.default_table_style.clone())
    }

    /// `StyleResolver.getDefaultCharacterStyle` (src/styles/styleResolver.ts:276).
    pub(crate) fn get_default_character_style(&self) -> Result<Js> {
        Ok(self.default_character_style.clone())
    }

    /// `StyleResolver.hasStyle` (src/styles/styleResolver.ts:283).
    pub(crate) fn has_style(&self, style_id: &str) -> Result<bool> {
        Ok(self.styles_by_id.contains_key(style_id))
    }

    /// `StyleResolver.findDefaultStyle` (src/styles/styleResolver.ts:291).
    pub(crate) fn find_default_style(&self, r#type: &str) -> Result<Js> {
        for style in self.styles_by_id.values() {
            if style.get("type").as_str() == Some(r#type) && style.get("default").truthy() {
                return Ok(style.clone());
            }
        }
        if r#type == "paragraph" {
            return Ok(self
                .get_style("Normal")?
                .nullish_or(if self.doc_defaults.truthy() {
                    Js::Undefined
                } else {
                    builtin_normal_style()
                }));
        }
        Ok(Js::Undefined)
    }

    /// `StyleResolver.mergeStyleIntoResult` (src/styles/styleResolver.ts:308).
    pub(crate) fn merge_style_into_result(&self, result: &Js, style: &Js) -> Result<()> {
        let ppr = style.get("pPr");
        if ppr.truthy() {
            result.set(
                "paragraphFormatting",
                merge_paragraph_formatting(&result.get("paragraphFormatting"), &ppr)?,
            );
        }
        let rpr = style.get("rPr");
        if rpr.truthy() {
            result.set(
                "runFormatting",
                merge_style_run_formatting(&result.get("runFormatting"), &rpr)?,
            );
        }
        Ok(())
    }

    /// `StyleResolver.mergeTextFormatting` (src/styles/styleResolver.ts:317).
    pub(crate) fn merge_text_formatting(&self, target: &Js, source: &Js) -> Result<Js> {
        merge_text_formatting(target, source)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn resolves_small_styles_like_ts() {
        let defs = Js::parse(r#"{"docDefaults":{"rPr":{"fontSize":22,"bold":false},"pPr":{"spaceAfter":100}},"styles":[
{"styleId":"Normal","type":"paragraph","default":true,"pPr":{"lineSpacing":240},"rPr":{"color":{"rgb":"FF0000"}}},
{"styleId":"H1","type":"paragraph","name":"heading 1","uiPriority":9,"next":"Normal","pPr":{"spaceBefore":240,"numPr":{"ilvl":0}},"rPr":{"bold":true,"fontFamily":{"ascii":"Arial","asciiTheme":"majorHAnsi"}}},
{"styleId":"DPF","type":"character","default":true,"rPr":{"italic":true}},
{"styleId":"Strong","type":"character","rPr":{"bold":true,"color":{"auto":true}}}]}"#).unwrap();
        let r = StyleResolver::new(&defs).unwrap();
        let tpf = Js::parse(r#"{"spaceAfter":0,"numPr":{"numId":3}}"#).unwrap();
        let trf = Js::parse(r#"{"bold":true}"#).unwrap();
        assert_eq!(
            r.resolve_paragraph_style(Some("H1"), &tpf, &trf)
                .unwrap()
                .stringify()
                .unwrap(),
            r#"{"paragraphFormatting":{"spaceAfter":0,"numPr":{"numId":3,"ilvl":0},"spaceBefore":240},"runFormatting":{"fontSize":22,"bold":false,"fontFamily":{"ascii":"Arial","asciiTheme":"majorHAnsi"}}}"#
        );
        assert_eq!(
            r.resolve_paragraph_style(None, &Js::Undefined, &Js::Undefined)
                .unwrap()
                .stringify()
                .unwrap(),
            r#"{"paragraphFormatting":{"spaceAfter":100,"lineSpacing":240},"runFormatting":{"fontSize":22,"bold":false,"color":{"rgb":"FF0000"}}}"#
        );
        assert_eq!(
            r.resolve_run_style(Some("Strong"))
                .unwrap()
                .stringify()
                .unwrap(),
            r#"{"fontSize":22,"bold":true,"italic":true}"#
        );
        assert_eq!(
            r.get_next_style_id(Some("H1")).unwrap().as_deref(),
            Some("Normal")
        );
        let bare = StyleResolver::new(&Js::Undefined).unwrap();
        assert_eq!(
            bare.resolve_paragraph_style(Some("X"), &Js::Undefined, &Js::Undefined)
                .unwrap()
                .stringify()
                .unwrap(),
            r#"{"paragraphFormatting":{"spaceAfter":160,"lineSpacing":259,"lineSpacingRule":"auto"}}"#
        );
    }
}
