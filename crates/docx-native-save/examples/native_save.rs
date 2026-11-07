//! `native_save --source a.docx --state s.bin --metadata m.bin --out o.docx [--wire w.json]
//! [--now ISO] [--no-modified-date] [--modified-by NAME] [--repeat N] [--json]`
//!
//! Mirrors the FinxWork desktop serializer: `openDocx(source, false)`, `bootstrapPeer(state,
//! metadata, source, host)`, `saveYrsDocx(session)`. With `--repeat N` the whole sequence runs N
//! times; `--json` prints the per-run timings (ms) on stdout.

use std::process::ExitCode;
use std::sync::Arc;
use std::sync::atomic::{AtomicU64, Ordering};
use std::time::{Instant, SystemTime, UNIX_EPOCH};

use docx_edit::session::{EditSessionCore, Host};
use docx_native_save::{OpenedSource, SaveOptions, save_docx};

fn now_ms() -> u64 {
    SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|d| d.as_millis() as u64)
        .unwrap_or(0)
}

fn new_session() -> EditSessionCore {
    static ENTROPY: AtomicU64 = AtomicU64::new(0);
    let seed = now_ms() ^ 0x9e37_79b9_7f4a_7c15;
    let host = Host::new(
        Arc::new(now_ms),
        Arc::new(move || {
            seed.wrapping_add(ENTROPY.fetch_add(0x9e37_79b9_7f4a_7c15, Ordering::Relaxed))
        }),
        Arc::new(|| now_ms() as f64),
    );
    EditSessionCore::new(1, host)
}

struct Args {
    source: String,
    state: String,
    metadata: String,
    out: String,
    wire: Option<String>,
    now: Option<String>,
    modified_by: Option<String>,
    update_modified_date: bool,
    repeat: usize,
    json: bool,
}

fn parse_args() -> Result<Args, String> {
    let mut it = std::env::args().skip(1);
    let (mut source, mut state, mut metadata, mut out) = (None, None, None, None);
    let mut a = Args {
        source: String::new(),
        state: String::new(),
        metadata: String::new(),
        out: String::new(),
        wire: None,
        now: None,
        modified_by: None,
        update_modified_date: true,
        repeat: 1,
        json: false,
    };
    while let Some(flag) = it.next() {
        let mut value = |name: &str| it.next().ok_or(format!("{name} needs a value"));
        match flag.as_str() {
            "--source" => source = Some(value("--source")?),
            "--state" => state = Some(value("--state")?),
            "--metadata" => metadata = Some(value("--metadata")?),
            "--out" => out = Some(value("--out")?),
            "--wire" => a.wire = Some(value("--wire")?),
            "--now" => a.now = Some(value("--now")?),
            "--modified-by" => a.modified_by = Some(value("--modified-by")?),
            "--repeat" => {
                a.repeat = value("--repeat")?
                    .parse::<usize>()
                    .map_err(|e| e.to_string())?
                    .max(1)
            }
            "--no-modified-date" => a.update_modified_date = false,
            "--json" => a.json = true,
            other => return Err(format!("unknown argument {other}")),
        }
    }
    a.source = source.ok_or("--source is required")?;
    a.state = state.ok_or("--state is required")?;
    a.metadata = metadata.ok_or("--metadata is required")?;
    a.out = out.ok_or("--out is required")?;
    Ok(a)
}

fn run() -> Result<(), String> {
    let args = parse_args()?;
    let read = |p: &str| std::fs::read(p).map_err(|e| format!("{p}: {e}"));
    let (source, state, metadata) = (
        read(&args.source)?,
        read(&args.state)?,
        read(&args.metadata)?,
    );
    let mut open_ms = Vec::new();
    let mut save_ms = Vec::new();
    let mut last = None;
    for _ in 0..args.repeat {
        let t0 = Instant::now();
        let core = new_session();
        let open_json = core
            .open_docx(&source, false, None, None)
            .map_err(|e| format!("open_docx: {e}"))?;
        core.bootstrap_peer(&state, &metadata, Some(source.clone()))
            .map_err(|e| format!("bootstrap_peer: {e}"))?;
        let opened = OpenedSource::after_open(source.clone(), &open_json)
            .map_err(|e| format!("after_open: {e}"))?;
        let t1 = Instant::now();
        let saved = save_docx(
            &core,
            &opened,
            &SaveOptions {
                update_modified_date: args.update_modified_date,
                modified_by: args.modified_by.clone(),
                now: args.now.clone(),
                capture_wire: args.wire.is_some(),
                ..SaveOptions::default()
            },
        )
        .map_err(|e| format!("save_docx: {e}"))?;
        let t2 = Instant::now();
        open_ms.push((t1 - t0).as_secs_f64() * 1e3);
        save_ms.push((t2 - t1).as_secs_f64() * 1e3);
        last = Some(saved);
    }
    let saved = last.expect("repeat >= 1");
    std::fs::write(&args.out, &saved.bytes).map_err(|e| format!("{}: {e}", args.out))?;
    if let (Some(path), Some(wire)) = (&args.wire, &saved.wire) {
        std::fs::write(path, wire).map_err(|e| format!("{path}: {e}"))?;
    }
    if args.json {
        let report = serde_json::json!({
            "repeat": args.repeat,
            "openBootstrapMs": open_ms,
            "saveMs": save_ms,
            "outBytes": saved.bytes.len(),
            "warnings": saved.warnings,
        });
        println!("{report}");
    }
    Ok(())
}

fn main() -> ExitCode {
    match run() {
        Ok(()) => ExitCode::SUCCESS,
        Err(e) => {
            eprintln!("native_save: {e}");
            ExitCode::FAILURE
        }
    }
}
