//! Port of `packages/docx/src/utils/units.ts`.
#![allow(unused_variables, unused_imports, dead_code, clippy::all)]

use crate::js::{Js, Result, SaveError};
use crate::session::{Inheritance, Session};
use crate::{arr, obj, throw};
use indexmap::IndexMap;
use std::collections::{HashMap, HashSet};
use std::rc::Rc;

const EMUS_PER_INCH: f64 = 914400.0;
const POINTS_PER_INCH: f64 = 72.0;
pub(crate) const PIXELS_PER_INCH: f64 = 96.0;
pub(crate) const TWIPS_PER_INCH: f64 = 1440.0;
const HALF_POINTS_PER_INCH: f64 = 144.0;
const EIGHTHS_PER_INCH: f64 = 576.0;
pub(crate) const AUTO_PARAGRAPH_SPACING_PX: f64 = 14.0;

/// `Math.round`.
fn js_round(x: f64) -> f64 {
    if !x.is_finite() {
        return x;
    }
    let r = x.floor();
    let v = if x - r >= 0.5 { r + 1.0 } else { r };
    if v == 0.0 && x < 0.0 { -0.0 } else { v }
}

/// `emuToPixels` (src/utils/units.ts:20).
pub(crate) fn emu_to_pixels(emu: Option<f64>) -> Result<f64> {
    let Some(emu) = emu else { return Ok(0.0) };
    if emu.is_nan() {
        return Ok(0.0);
    }
    Ok(js_round((emu * PIXELS_PER_INCH) / EMUS_PER_INCH))
}

/// `pixelsToEmu` (src/utils/units.ts:25).
pub(crate) fn pixels_to_emu(px: f64) -> Result<f64> {
    Ok(js_round((px / PIXELS_PER_INCH) * EMUS_PER_INCH))
}

/// `pointsToPixels` (src/utils/units.ts:29).
pub(crate) fn points_to_pixels(points: f64) -> Result<f64> {
    Ok((points / POINTS_PER_INCH) * PIXELS_PER_INCH)
}

/// `twipsToPixels` (src/utils/units.ts:63).
pub(crate) fn twips_to_pixels(twips: f64) -> Result<f64> {
    Ok((twips / TWIPS_PER_INCH) * PIXELS_PER_INCH)
}

/// `pixelsToTwips` (src/utils/units.ts:70).
pub(crate) fn pixels_to_twips(px: f64) -> Result<f64> {
    Ok((px / PIXELS_PER_INCH) * TWIPS_PER_INCH)
}

/// `emuToTwips` (src/utils/units.ts:81).
pub(crate) fn emu_to_twips(emu: f64) -> Result<f64> {
    Ok(js_round((emu / EMUS_PER_INCH) * TWIPS_PER_INCH))
}

/// `twipsToEmu` (src/utils/units.ts:88).
pub(crate) fn twips_to_emu(twips: f64) -> Result<f64> {
    Ok(js_round((twips / TWIPS_PER_INCH) * EMUS_PER_INCH))
}

/// `halfPointsToPixels` (src/utils/units.ts:101).
pub(crate) fn half_points_to_pixels(half_points: f64) -> Result<f64> {
    Ok((half_points / HALF_POINTS_PER_INCH) * PIXELS_PER_INCH)
}

/// `halfPointsToPoints` (src/utils/units.ts:108).
pub(crate) fn half_points_to_points(half_points: f64) -> Result<f64> {
    Ok(half_points / 2.0)
}

/// `pointsToHalfPoints` (src/utils/units.ts:115).
pub(crate) fn points_to_half_points(points: f64) -> Result<f64> {
    Ok(points * 2.0)
}

/// `eighthsToPixels` (src/utils/units.ts:128).
pub(crate) fn eighths_to_pixels(eighths: f64) -> Result<f64> {
    Ok((eighths / EIGHTHS_PER_INCH) * PIXELS_PER_INCH)
}

/// `roundPixels` (src/utils/units.ts:139). Callers pass 2 for the TS default.
pub(crate) fn round_pixels(px: f64, decimal_places: f64) -> Result<f64> {
    let factor = 10f64.powf(decimal_places);
    Ok(js_round(px * factor) / factor)
}

/// `clamp` (src/utils/units.ts:147).
pub(crate) fn clamp(value: f64, min: f64, max: f64) -> Result<f64> {
    Ok(js_min(js_max(value, min), max))
}

fn js_max(a: f64, b: f64) -> f64 {
    if a.is_nan() || b.is_nan() {
        f64::NAN
    } else {
        a.max(b)
    }
}
fn js_min(a: f64, b: f64) -> f64 {
    if a.is_nan() || b.is_nan() {
        f64::NAN
    } else {
        a.min(b)
    }
}

/// `formatPx` (src/utils/units.ts:158).
pub(crate) fn format_px(px: f64) -> Result<String> {
    Ok(format!(
        "{}px",
        crate::js::num_to_string(round_pixels(px, 2.0)?)
    ))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn conversions_match_ts() {
        assert_eq!(emu_to_pixels(Some(4762.5)).unwrap(), 1.0);
        assert_eq!(emu_to_pixels(Some(-4762.5)).unwrap(), -0.0);
        assert_eq!(emu_to_pixels(None).unwrap(), 0.0);
        assert_eq!(emu_to_twips(-317.5).unwrap(), 0.0);
        assert_eq!(twips_to_emu(1.5).unwrap(), 953.0);
        assert_eq!(pixels_to_emu(10.0).unwrap(), 95250.0);
        assert_eq!(format_px(1.005).unwrap(), "1px");
        assert_eq!(format_px(2.675).unwrap(), "2.68px");
        assert_eq!(format_px(-1.125).unwrap(), "-1.12px");
        assert_eq!(twips_to_pixels(1440.0).unwrap(), 96.0);
    }
}
