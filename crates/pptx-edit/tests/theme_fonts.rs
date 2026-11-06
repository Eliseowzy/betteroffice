use std::collections::BTreeMap;

use pptx_edit::{
    DeckSession, EditCtx, RunFontThemes, TextStylePatch, ThemeFontFace, ThemeFontSlot, ThemeFonts,
};

const DECK: &[u8] = include_bytes!("../../pptx-render/tests/fixtures/theme-fonts.pptx");

fn parts(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    ooxml_opc::unzip_parts(bytes).unwrap().into_iter().collect()
}

fn slide_xml(bytes: &[u8]) -> String {
    String::from_utf8(parts(bytes)["ppt/slides/slide1.xml"].clone()).unwrap()
}

fn with_slide_xml(edit: impl Fn(String) -> String) -> Vec<u8> {
    let mut parts = ooxml_opc::unzip_parts(DECK).unwrap();
    let (_, bytes) = parts
        .iter_mut()
        .find(|(path, _)| path == "ppt/slides/slide1.xml")
        .unwrap();
    *bytes = edit(String::from_utf8(bytes.clone()).unwrap()).into_bytes();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

/// `(slide id, shape id, story id)` of the text box with `source_id`.
fn target(session: &DeckSession, source_id: u32) -> (String, String, String) {
    let snapshot = session.snapshot().unwrap();
    let slide = &snapshot.slides[0];
    let shape = slide
        .shapes
        .iter()
        .find(|shape| shape.source_id == source_id)
        .unwrap();
    (
        slide.id.clone(),
        shape.id.clone(),
        shape.text_stories[0].id.clone(),
    )
}

fn first_run(session: &DeckSession, source_id: u32) -> (String, RunFontThemes) {
    let (slide, shape, story) = target(session, source_id);
    let (story, themes) = session
        .story_formatting_with_font_themes(&slide, &shape, &story)
        .unwrap();
    (
        story.paragraphs[0].runs[0]
            .style
            .font_family
            .clone()
            .unwrap(),
        themes[0][0],
    )
}

#[test]
fn theme_references_report_their_slot_beside_the_resolved_family() {
    let session = DeckSession::open(DECK, 7101).unwrap();
    let (family, themes) = first_run(&session, 2);
    assert_eq!(family, "Liberation Serif");
    assert_eq!(themes.latin, Some(ThemeFontSlot::Major));
    let (family, themes) = first_run(&session, 4);
    assert_eq!(family, "Liberation Sans");
    assert_eq!(themes.latin, Some(ThemeFontSlot::Minor));
    let (family, themes) = first_run(&session, 5);
    assert_eq!(family, "Arial");
    assert_eq!(themes.latin, None);
}

#[test]
fn an_inherited_theme_reference_reports_its_slot() {
    let deck = with_slide_xml(|xml| {
        let direct = r#"<a:latin typeface="+mj-lt"/></a:rPr><a:t>Major theme reference"#;
        assert_eq!(xml.matches(direct).count(), 1);
        xml.replacen(direct, "</a:rPr><a:t>Major theme reference", 1)
            .replacen(
                "<a:lstStyle/>",
                r#"<a:lstStyle><a:lvl1pPr><a:defRPr><a:latin typeface="+mj-lt"/><a:ea typeface="+mj-ea"/></a:defRPr></a:lvl1pPr></a:lstStyle>"#,
                1,
            )
    });
    let session = DeckSession::open(&deck, 7102).unwrap();
    let (family, themes) = first_run(&session, 2);
    assert_eq!(family, "Liberation Serif");
    assert_eq!(
        themes,
        RunFontThemes {
            latin: Some(ThemeFontSlot::Major),
            east_asian: Some(ThemeFontSlot::Major),
        }
    );
}

#[test]
fn the_theme_fonts_of_a_slide() {
    let session = DeckSession::open(DECK, 7103).unwrap();
    let (slide, _, _) = target(&session, 2);
    assert_eq!(
        session.theme_fonts(&slide).unwrap(),
        ThemeFonts {
            major: ThemeFontFace {
                latin: "Liberation Serif".into(),
                east_asian: None,
            },
            minor: ThemeFontFace {
                latin: "Liberation Sans".into(),
                east_asian: None,
            },
        }
    );
    assert!(session.theme_fonts("missing").is_err());
}

#[test]
fn a_theme_font_written_to_a_run_saves_as_its_reference() {
    let session = DeckSession::open(DECK, 7104).unwrap();
    let (_, _, story) = target(&session, 5);
    session
        .format_text(
            &EditCtx::local("test"),
            &story,
            0,
            8,
            &TextStylePatch {
                font_family: Some(ThemeFontSlot::Minor.reference("lt")),
                ..TextStylePatch::default()
            },
        )
        .unwrap();
    let (family, themes) = first_run(&session, 5);
    assert_eq!(family, "Liberation Sans");
    assert_eq!(themes.latin, Some(ThemeFontSlot::Minor));

    let saved = session.save().unwrap();
    let xml = slide_xml(&saved);
    assert!(
        xml.contains(r#"<a:latin typeface="+mn-lt"/></a:rPr><a:t>Explicit</a:t>"#),
        "{xml}"
    );
    assert!(
        xml.contains(r#"<a:latin typeface="Arial"/></a:rPr><a:t> font stays unchanged.</a:t>"#)
    );
    let untouched = slide_xml(DECK);
    let shape_two = |xml: &str| xml.split("<p:sp>").nth(1).unwrap().to_owned();
    assert_eq!(shape_two(&xml), shape_two(&untouched));

    let reopened = DeckSession::open(&saved, 7105).unwrap();
    let (family, themes) = first_run(&reopened, 5);
    assert_eq!(family, "Liberation Sans");
    assert_eq!(themes.latin, Some(ThemeFontSlot::Minor));
}

#[test]
fn theme_font_references_parse_by_script() {
    assert_eq!(
        ThemeFontSlot::from_reference("+mj-lt", "lt"),
        Some(ThemeFontSlot::Major)
    );
    assert_eq!(
        ThemeFontSlot::from_reference("+MN-EA", "ea"),
        Some(ThemeFontSlot::Minor)
    );
    assert_eq!(ThemeFontSlot::from_reference("+mj-ea", "lt"), None);
    assert_eq!(ThemeFontSlot::from_reference("Calibri", "lt"), None);
    assert_eq!(ThemeFontSlot::Major.reference("ea"), "+mj-ea");
}
