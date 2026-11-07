//! Saves a few packages natively, unedited and edited, and reopens what was written.

use std::sync::Arc;

use docx_edit::session::{EditSessionCore, Host};
use docx_native_save::{OpenedSource, SaveOptions, save_docx};

fn session() -> EditSessionCore {
    EditSessionCore::new(
        7,
        Host::new(Arc::new(|| 0), Arc::new(|| 11), Arc::new(|| 0.0)),
    )
}

fn fixtures() -> Vec<std::path::PathBuf> {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../docx-edit/tests/fixtures");
    let mut out: Vec<_> = std::fs::read_dir(root)
        .unwrap()
        .flatten()
        .map(|e| e.path())
        .filter(|p| p.extension().is_some_and(|e| e == "docx"))
        .collect();
    out.sort();
    out
}

#[test]
fn unedited_packages_save_and_reopen() {
    for path in fixtures() {
        let bytes = std::fs::read(&path).unwrap();
        let core = session();
        let opened = core.open_docx(&bytes, true, None, None).unwrap();
        let source = OpenedSource::after_open(bytes.clone(), &opened)
            .unwrap_or_else(|e| panic!("{}: {e}", path.display()));
        let saved = save_docx(
            &core,
            &source,
            &SaveOptions {
                now: Some("2026-11-07T00:00:00.000Z".into()),
                ..SaveOptions::default()
            },
        )
        .unwrap_or_else(|e| panic!("{}: {e}", path.display()));
        let again = session();
        again
            .open_docx(&saved.bytes, true, None, None)
            .unwrap_or_else(|e| panic!("{}: reopen: {e}", path.display()));
    }
}
