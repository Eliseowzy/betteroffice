//! Port of `packages/docx/src/yrs/saveFormatting.ts`.
#![allow(unused_variables, unused_imports, dead_code, clippy::all)]

use crate::js::{Js, Result, SaveError};
use crate::session::{Inheritance, Session};
use crate::{arr, obj, throw};
use indexmap::IndexMap;
use std::collections::{HashMap, HashSet};
use std::rc::Rc;

/// `v || undefined`.
fn or_u(v: Js) -> Js {
    if v.truthy() { v } else { Js::Undefined }
}

/// `v ?? undefined`.
fn nn(v: Js) -> Js {
    if v.is_nullish() { Js::Undefined } else { v }
}

/// `isStyleSourcedNumPr` (src/yrs/saveFormatting.ts:45).
pub(crate) fn is_style_sourced_num_pr(attrs: &Js) -> Result<bool> {
    Ok(!attrs.get("numPrFromStyle").is_nullish()
        && !attrs.get("numPr").is_nullish()
        && attrs.get("numPr").stringify() == attrs.get("numPrFromStyle").stringify())
}

/// `sameValue` (src/yrs/saveFormatting.ts:53).
pub(crate) fn same_value(left: &Js, right: &Js) -> Result<bool> {
    Ok(left.strict_eq(right) || left.stringify() == right.stringify())
}

/// `inheritsValue` (src/yrs/saveFormatting.ts:58).
pub(crate) fn inherits_value(attrs: &Js, orig: &Js, inherited: &Js, key: &Js) -> Result<bool> {
    let key = crate::js::to_js_string(key);
    Ok(orig.get(&key).is_undefined()
        && !inherited.get(&key).is_undefined()
        && same_value(&attrs.get(&key), &inherited.get(&key))?)
}

/// `withoutInherited` (src/yrs/saveFormatting.ts:75).
pub(crate) fn without_inherited(formatting: &Js, inherited: &Js) -> Result<Js> {
    let result = formatting.spread();
    for key in result.keys() {
        if matches!(
            key.as_str(),
            "styleId"
                | "numPr"
                | "lineSpacing"
                | "lineSpacingRule"
                | "indentFirstLine"
                | "hangingIndent"
        ) {
            continue;
        }
        if !inherited.get(&key).is_undefined()
            && same_value(&result.get(&key), &inherited.get(&key))?
        {
            result.delete(&key);
        }
    }
    let auto = || Js::from("auto");
    if !inherited.get("lineSpacing").is_undefined()
        && result
            .get("lineSpacing")
            .strict_eq(&inherited.get("lineSpacing"))
        && result
            .get("lineSpacingRule")
            .nullish_or(auto())
            .strict_eq(&inherited.get("lineSpacingRule").nullish_or(auto()))
    {
        result.delete("lineSpacing");
        result.delete("lineSpacingRule");
    }
    Ok(if result.entries().iter().any(|(_, v)| !v.is_undefined()) {
        result
    } else {
        Js::Undefined
    })
}

/// `paragraphAttrsToFormatting` (src/yrs/saveFormatting.ts:110).
pub(crate) fn paragraph_attrs_to_formatting(attrs: &Js, inherited: &Js) -> Result<Js> {
    let a = |k: &str| attrs.get(k);
    let orig = attrs.get("_originalFormatting");
    if orig.truthy() {
        let o = |k: &str| orig.get(k);
        let result = orig.spread();
        let inherits = |k: &str| inherits_value(attrs, &orig, inherited, &Js::from(k));
        if !inherits("alignment")? && !a("alignment").strict_eq(&or_u(o("alignment"))) {
            result.set("alignment", or_u(a("alignment")));
        }
        if is_style_sourced_num_pr(attrs)? {
            result.delete("numPr");
            result.delete("numPrFromStyle");
        } else if !a("numPr").strict_eq(&o("numPr"))
            && a("numPr").stringify() != o("numPr").stringify()
        {
            result.set("numPr", or_u(a("numPr")));
            result.delete("numPrFromStyle");
        }
        if !a("styleId").strict_eq(&or_u(o("styleId"))) {
            result.set("styleId", or_u(a("styleId")));
        }
        if !inherits("pageBreakBefore")?
            && !a("pageBreakBefore").strict_eq(&or_u(o("pageBreakBefore")))
        {
            result.set("pageBreakBefore", or_u(a("pageBreakBefore")));
        }
        for k in [
            "widowControl",
            "autoSpaceDE",
            "autoSpaceDN",
            "adjustRightInd",
        ] {
            if !inherits(k)? && !a(k).strict_eq(&nn(o(k))) {
                result.set(k, nn(a(k)));
            }
        }
        if !a("cnfStyle").is_nullish() && !same_value(&a("cnfStyle"), &o("cnfStyle"))? {
            result.set("cnfStyle", a("cnfStyle"));
        }
        if !inherits("bidi")? && !a("bidi").strict_eq(&or_u(o("bidi"))) {
            result.set("bidi", or_u(a("bidi")));
        }
        return Ok(result);
    }

    let has_formatting = a("alignment").truthy()
        || !a("spaceBefore").is_nullish()
        || !a("spaceAfter").is_nullish()
        || !a("spaceBeforeLines").is_nullish()
        || !a("spaceAfterLines").is_nullish()
        || !a("beforeAutospacing").is_nullish()
        || !a("afterAutospacing").is_nullish()
        || a("lineSpacing").truthy()
        || a("indentLeft").truthy()
        || a("indentRight").truthy()
        || a("indentFirstLine").truthy()
        || a("numPr").truthy()
        || a("styleId").truthy()
        || a("borders").truthy()
        || a("shading").truthy()
        || a("tabs").truthy()
        || !a("outlineLevel").is_nullish()
        || a("contextualSpacing").truthy()
        || a("pageBreakBefore").truthy()
        || !a("widowControl").is_nullish()
        || !a("autoSpaceDE").is_nullish()
        || !a("autoSpaceDN").is_nullish()
        || !a("adjustRightInd").is_nullish()
        || a("cnfStyle").truthy()
        || a("bidi").truthy();
    if !has_formatting {
        return Ok(Js::Undefined);
    }
    let formatting = obj! {
        "alignment" => or_u(a("alignment")),
        "spaceBefore" => nn(a("spaceBefore")),
        "spaceAfter" => nn(a("spaceAfter")),
        "spaceBeforeLines" => nn(a("spaceBeforeLines")),
        "spaceAfterLines" => nn(a("spaceAfterLines")),
        "beforeAutospacing" => nn(a("beforeAutospacing")),
        "afterAutospacing" => nn(a("afterAutospacing")),
        "lineSpacing" => or_u(a("lineSpacing")),
        "lineSpacingRule" => or_u(a("lineSpacingRule")),
        "indentLeft" => or_u(a("indentLeft")),
        "indentRight" => or_u(a("indentRight")),
        "indentFirstLine" => or_u(a("indentFirstLine")),
        "hangingIndent" => or_u(a("hangingIndent")),
        "numPr" => if is_style_sourced_num_pr(attrs)? { Js::Undefined } else { or_u(a("numPr")) },
        "styleId" => or_u(a("styleId")),
        "borders" => or_u(a("borders")),
        "shading" => or_u(a("shading")),
        "tabs" => or_u(a("tabs")),
        "outlineLevel" => nn(a("outlineLevel")),
        "contextualSpacing" => or_u(a("contextualSpacing")),
        "pageBreakBefore" => or_u(a("pageBreakBefore")),
        "widowControl" => nn(a("widowControl")),
        "autoSpaceDE" => nn(a("autoSpaceDE")),
        "autoSpaceDN" => nn(a("autoSpaceDN")),
        "adjustRightInd" => nn(a("adjustRightInd")),
        "cnfStyle" => or_u(a("cnfStyle")),
        "bidi" => or_u(a("bidi")),
    };
    if inherited.truthy() {
        without_inherited(&formatting, inherited)
    } else {
        Ok(formatting)
    }
}

fn width_of(attrs: &Js) -> Js {
    obj! {
        "value" => attrs.get("width").nullish_or(Js::Num(0.0)),
        "type" => attrs.get("widthType").or(Js::from("dxa")),
    }
}

/// `tableAttrsToFormatting` (src/yrs/saveFormatting.ts:241).
pub(crate) fn table_attrs_to_formatting(attrs: &Js) -> Result<Js> {
    let a = |k: &str| attrs.get(k);
    let margins = if same_value(&a("cellMargins"), &a("_originalResolvedMargins"))? {
        attrs.path(&["_originalFormatting", "cellMargins"])
    } else if a("cellMargins").truthy() {
        measurement_margins(&a("cellMargins"))?
    } else {
        Js::Undefined
    };
    let orig = a("_originalFormatting");
    if orig.truthy() {
        let o = |k: &str| orig.get(k);
        let result = orig.spread();
        if !a("styleId").strict_eq(&or_u(o("styleId"))) {
            result.set("styleId", or_u(a("styleId")));
        }
        if !a("justification").strict_eq(&or_u(o("justification"))) {
            result.set("justification", or_u(a("justification")));
        }
        if !a("floating").strict_eq(&or_u(o("floating"))) {
            result.set("floating", or_u(a("floating")));
        }
        if !a("tableLayout").strict_eq(&or_u(o("layout"))) {
            result.set("layout", or_u(a("tableLayout")));
        }
        if !a("look").strict_eq(&or_u(o("look"))) {
            result.set("look", or_u(a("look")));
        }
        if !a("bidi").strict_eq(&or_u(o("bidi"))) {
            result.set("bidi", or_u(a("bidi")));
        }
        if !a("width").strict_eq(&orig.path(&["width", "value"]))
            || !a("widthType").strict_eq(&orig.path(&["width", "type"]))
        {
            result.set(
                "width",
                if !a("width").is_nullish() || a("widthType").truthy() {
                    width_of(attrs)
                } else {
                    Js::Undefined
                },
            );
        }
        if !margins.strict_eq(&o("cellMargins")) {
            result.set("cellMargins", margins);
        }
        return Ok(result);
    }
    let has_formatting = a("styleId").truthy()
        || !a("width").is_nullish()
        || a("widthType").truthy()
        || a("justification").truthy()
        || a("tableLayout").truthy()
        || a("floating").truthy()
        || margins.truthy()
        || a("look").truthy()
        || a("bidi").truthy();
    if !has_formatting {
        return Ok(Js::Undefined);
    }
    Ok(obj! {
        "styleId" => or_u(a("styleId")),
        "width" => if !a("width").is_nullish() || a("widthType").truthy() { width_of(attrs) } else { Js::Undefined },
        "justification" => or_u(a("justification")),
        "layout" => or_u(a("tableLayout")),
        "floating" => or_u(a("floating")),
        "cellMargins" => margins,
        "look" => or_u(a("look")),
        "bidi" => or_u(a("bidi")),
    })
}

fn dxa(v: Js) -> Js {
    obj! { "value" => v, "type" => "dxa" }
}

/// `measurementMargins` (src/yrs/saveFormatting.ts:302).
pub(crate) fn measurement_margins(margins: &Js) -> Result<Js> {
    let m = |k: &str| {
        let v = margins.get(k);
        if v.is_nullish() {
            Js::Undefined
        } else {
            dxa(v)
        }
    };
    Ok(
        obj! { "top" => m("top"), "bottom" => m("bottom"), "left" => m("left"), "right" => m("right") },
    )
}

/// `tableRowAttrsToFormatting` (src/yrs/saveFormatting.ts:327).
pub(crate) fn table_row_attrs_to_formatting(attrs: &Js) -> Result<Js> {
    let a = |k: &str| attrs.get(k);
    let height = || {
        if a("height").truthy() {
            dxa(a("height"))
        } else {
            Js::Undefined
        }
    };
    let orig = a("_originalFormatting");
    if orig.truthy() {
        let result = orig.spread();
        if !a("height").strict_eq(&or_u(orig.path(&["height", "value"]))) {
            result.set("height", height());
        }
        if !a("heightRule").strict_eq(&or_u(orig.get("heightRule"))) {
            result.set("heightRule", or_u(a("heightRule")));
        }
        if !a("isHeader").strict_eq(&or_u(orig.get("header"))) {
            result.set("header", or_u(a("isHeader")));
        }
        return Ok(result);
    }
    if !a("height").truthy() && !a("isHeader").truthy() {
        return Ok(Js::Undefined);
    }
    Ok(obj! {
        "height" => height(),
        "heightRule" => or_u(a("heightRule")),
        "header" => or_u(a("isHeader")),
    })
}

/// `tableCellAttrsToFormatting` (src/yrs/saveFormatting.ts:379).
pub(crate) fn table_cell_attrs_to_formatting(attrs: &Js) -> Result<Js> {
    let a = |k: &str| attrs.get(k);
    let borders = if same_value(&a("borders"), &a("_originalResolvedBorders"))? {
        attrs.path(&["_originalFormatting", "borders"])
    } else {
        or_u(a("borders"))
    };
    let margins = if same_value(&a("margins"), &a("_originalResolvedMargins"))? {
        attrs.path(&["_originalFormatting", "margins"])
    } else if a("margins").truthy() {
        cell_margins(&a("margins"))?
    } else {
        Js::Undefined
    };
    let colspan_gt1 = crate::js::to_number(&a("colspan")) > 1.0;
    let cell_width =
        || obj! { "value" => a("width"), "type" => a("widthType").or(Js::from("dxa")) };
    let shading = || {
        if a("backgroundColor").truthy() {
            obj! { "fill" => obj! { "rgb" => a("backgroundColor") } }
        } else {
            Js::Undefined
        }
    };
    let orig = a("_originalFormatting");
    if orig.truthy() {
        let o = |k: &str| orig.get(k);
        let result = orig.spread();
        if colspan_gt1 {
            result.set("gridSpan", a("colspan"));
        }
        if !a("width").is_nullish() {
            result.set("width", cell_width());
        }
        if !a("verticalAlign").strict_eq(&or_u(o("verticalAlign"))) {
            result.set("verticalAlign", or_u(a("verticalAlign")));
        }
        if !a("backgroundColor")
            .nullish_or(Js::Null)
            .strict_eq(&a("_originalResolvedFill").nullish_or(Js::Null))
        {
            result.set("shading", shading());
        }
        if !borders.strict_eq(&o("borders")) {
            result.set("borders", borders);
        }
        if !margins.strict_eq(&o("margins")) {
            result.set("margins", margins);
        }
        if !a("textDirection").strict_eq(&or_u(o("textDirection"))) {
            result.set("textDirection", or_u(a("textDirection")));
        }
        return Ok(result);
    }
    let has_formatting = colspan_gt1
        || crate::js::to_number(&a("rowspan")) > 1.0
        || !a("width").is_nullish()
        || a("verticalAlign").truthy()
        || a("backgroundColor").truthy()
        || borders.truthy()
        || margins.truthy()
        || a("textDirection").truthy();
    if !has_formatting {
        return Ok(Js::Undefined);
    }
    Ok(obj! {
        "gridSpan" => if colspan_gt1 { a("colspan") } else { Js::Undefined },
        "width" => if !a("width").is_nullish() { cell_width() } else { Js::Undefined },
        "verticalAlign" => or_u(a("verticalAlign")),
        "textDirection" => or_u(a("textDirection")),
        "shading" => shading(),
        "borders" => borders,
        "margins" => margins,
    })
}

/// `cellMargins` (src/yrs/saveFormatting.ts:438).
pub(crate) fn cell_margins(margins: &Js) -> Result<Js> {
    let result = Js::new_obj();
    for k in ["top", "bottom", "left", "right"] {
        if !margins.get(k).is_nullish() {
            result.set(k, dxa(margins.get(k)));
        }
    }
    Ok(result)
}
