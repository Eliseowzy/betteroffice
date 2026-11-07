//! Opens once, then saves `--repeat` times, timing each save and its phases: for profiling.
//! `save_profile <source> <state> <metadata> <repeat>`.

use std::sync::Arc;
use std::time::Instant;

use docx_edit::session::{EditSessionCore, Host};
use docx_native_save::{OpenedSource, SaveOptions, save_docx};

fn main() {
    let args: Vec<String> = std::env::args().collect();
    let source = std::fs::read(&args[1]).unwrap();
    let state = std::fs::read(&args[2]).unwrap();
    let metadata = std::fs::read(&args[3]).unwrap();
    let repeat: usize = args.get(4).map_or(5, |n| n.parse().unwrap());
    let core = EditSessionCore::new(
        9,
        Host::new(Arc::new(|| 0), Arc::new(|| 1), Arc::new(|| 0.0)),
    );
    let opened = core.open_docx(&source, false, None, None).unwrap();
    core.bootstrap_peer(&state, &metadata, Some(source.clone()))
        .unwrap();
    let t = Instant::now();
    let opened = OpenedSource::after_open(source, &opened).unwrap();
    eprintln!("after_open {:.1} ms", t.elapsed().as_secs_f64() * 1e3);
    let options = SaveOptions {
        now: Some("2026-11-07T00:00:00.000Z".into()),
        record_saved_paragraph_ids: false,
        ..SaveOptions::default()
    };
    if std::env::var_os("ENGINE_ONLY").is_some() {
        for _ in 0..repeat {
            core.paragraph_save_plan().unwrap();
            core.paragraph_identities().unwrap();
            core.materialize_docx().unwrap();
            for story in core.story_ids() {
                core.story_segments(&story).unwrap();
            }
            core.list_sections_json().unwrap();
            core.list_comments().unwrap();
        }
        return;
    }
    for _ in 0..repeat {
        let t = Instant::now();
        let saved = save_docx(&core, &opened, &options).unwrap();
        eprintln!(
            "save {:.1} ms ({} bytes)",
            t.elapsed().as_secs_f64() * 1e3,
            saved.bytes.len()
        );
    }
    for (name, f) in [
        (
            "paragraph_save_plan",
            Box::new(|| {
                core.paragraph_save_plan().unwrap();
            }) as Box<dyn Fn()>,
        ),
        (
            "paragraph_identities",
            Box::new(|| {
                core.paragraph_identities().unwrap();
            }),
        ),
        (
            "materialize_docx",
            Box::new(|| {
                core.materialize_docx().unwrap();
            }),
        ),
        (
            "story_segments(body)",
            Box::new(|| {
                core.story_segments("body").unwrap();
            }),
        ),
        (
            "list_sections",
            Box::new(|| {
                core.list_sections_json().unwrap();
            }),
        ),
    ] {
        let t = Instant::now();
        f();
        eprintln!("  engine {name} {:.1} ms", t.elapsed().as_secs_f64() * 1e3);
    }
}
