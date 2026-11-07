/**
 * Differential probe for the native (Rust) DOCX save. Test and tooling only, a no-op unless
 * `DOCX_NATIVE_SAVE_PROBE=<report.jsonl>` is set. `saveYrsDocx` calls {@link beginNativeSaveProbe}
 * before the JS save and the returned function after it; the probe writes the save's inputs to a
 * temp dir, runs `native_save`, compares the two packages part by part and appends one JSON line
 * to the report. It never changes what the JS save returns or throws.
 *
 * Env: DOCX_NATIVE_SAVE_PROBE (report path), DOCX_NATIVE_SAVE_CLI (CLI path, default
 * target/release/examples/native_save under the repo root).
 * @internal
 */

import { unzipContainer } from '../docx/wasm';
import type { YrsSession } from './index';

type WireListener = (wire: string) => void;
type ProbeGlobal = { __docxSaveWireHook?: WireListener };

/** Called by the save facade with the exact JSON handed to the package writer. */
export function reportSaveWire(wire: string): void {
  const hook = (globalThis as ProbeGlobal).__docxSaveWireHook;
  if (hook) {
    try { hook(wire); } catch { /* the probe never affects a save */ }
  }
}

function env(name: string): string | undefined {
  return typeof process === 'undefined' ? undefined : process.env?.[name];
}

const nodeImport = (spec: string): Promise<any> => import(/* @vite-ignore */ spec);

export type ProbeVerdict = 'identical' | 'same-parts-after-normalizing' | 'diff' | 'native-error';

export interface NativeComparison {
  verdict: ProbeVerdict;
  details: Record<string, unknown>;
  nativeMs?: { openBootstrap: number; save: number };
}

export interface NativeCompareInput {
  jsBytes: Uint8Array;
  jsWire?: string;
  source: Uint8Array;
  state: Uint8Array;
  metadata: Uint8Array;
  cliPath?: string;
}

const MODIFIED = /(<dcterms:modified\b[^>]*>)[^<]*(<\/dcterms:modified>)/;

function normalize(part: string, bytes: Uint8Array): Uint8Array {
  if (part !== 'docProps/core.xml') return bytes;
  const text = new TextDecoder().decode(bytes);
  return new TextEncoder().encode(text.replace(MODIFIED, '$1$2'));
}

function equalBytes(a: Uint8Array, b: Uint8Array): boolean {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i += 1) if (a[i] !== b[i]) return false;
  return true;
}

function firstDifference(a: Uint8Array, b: Uint8Array): number {
  const n = Math.min(a.length, b.length);
  for (let i = 0; i < n; i += 1) if (a[i] !== b[i]) return i;
  return n;
}

function context(bytes: Uint8Array, offset: number): string {
  const start = Math.max(0, offset - 100);
  return new TextDecoder().decode(bytes.subarray(start, Math.min(bytes.length, offset + 100)));
}

/** The first JSON path at which two values differ, or null. */
export function firstJsonDifference(a: unknown, b: unknown, path = '$'): string | null {
  if (Object.is(a, b)) return null;
  if (Array.isArray(a) && Array.isArray(b)) {
    for (let i = 0; i < Math.max(a.length, b.length); i += 1) {
      if (i >= a.length || i >= b.length) return `${path}[${i}]`;
      const d = firstJsonDifference(a[i], b[i], `${path}[${i}]`);
      if (d) return d;
    }
    return null;
  }
  if (a && b && typeof a === 'object' && typeof b === 'object' && !Array.isArray(a) && !Array.isArray(b)) {
    const ao = a as Record<string, unknown>;
    const bo = b as Record<string, unknown>;
    const keys = [...Object.keys(ao), ...Object.keys(bo).filter((k) => !(k in ao))];
    for (const key of keys) {
      if (!(key in ao) || !(key in bo)) return `${path}.${key}`;
      const d = firstJsonDifference(ao[key], bo[key], `${path}.${key}`);
      if (d) return d;
    }
    return null;
  }
  return path;
}

/** Runs the native CLI on the inputs and compares its package and wire with the JS save's. */
export async function compareWithNative(input: NativeCompareInput): Promise<NativeComparison> {
  const fs = await nodeImport('node:fs');
  const os = await nodeImport('node:os');
  const path = await nodeImport('node:path');
  const cp = await nodeImport('node:child_process');
  const root = path.resolve(import.meta.dir ?? '.', '../../../..');
  const cli = input.cliPath ?? env('DOCX_NATIVE_SAVE_CLI') ??
    path.join(root, 'target/release/examples/native_save');
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'native-save-'));
  try {
    const file = (name: string) => path.join(dir, name);
    fs.writeFileSync(file('a.docx'), input.source);
    fs.writeFileSync(file('s.bin'), input.state);
    fs.writeFileSync(file('m.bin'), input.metadata);
    const args = ['--source', file('a.docx'), '--state', file('s.bin'), '--metadata', file('m.bin'),
      '--out', file('o.docx'), '--wire', file('w.json'), '--json'];
    let jsRequest: any;
    if (input.jsWire !== undefined) {
      try { jsRequest = JSON.parse(input.jsWire); } catch { /* compared as text below */ }
    }
    if (jsRequest?.determinism?.now) args.push('--now', jsRequest.determinism.now);
    if (jsRequest?.options?.updateModifiedDate === false) args.push('--no-modified-date');
    if (typeof jsRequest?.options?.modifiedBy === 'string') args.push('--modified-by', jsRequest.options.modifiedBy);
    const run = cp.spawnSync(cli, args, { encoding: 'utf8', maxBuffer: 1 << 28 });
    if (run.error || run.status !== 0) {
      return {
        verdict: 'native-error',
        details: { status: run.status, signal: run.signal, error: run.error?.message,
          stderr: String(run.stderr ?? '').slice(0, 4000) },
      };
    }
    const nativeBytes = new Uint8Array(fs.readFileSync(file('o.docx')));
    let nativeMs: NativeComparison['nativeMs'];
    try {
      const timing = JSON.parse(String(run.stdout).trim().split('\n').pop() ?? '{}');
      nativeMs = { openBootstrap: timing.openBootstrapMs?.[0], save: timing.saveMs?.[0] };
    } catch { /* timing is optional */ }

    const js = unzipContainer(input.jsBytes);
    const rs = unzipContainer(nativeBytes);
    const jsNames = Object.keys(js);
    const rsNames = Object.keys(rs);
    const details: Record<string, unknown> = {
      zipBytesIdentical: equalBytes(input.jsBytes, nativeBytes),
      jsParts: jsNames.length, nativeParts: rsNames.length,
    };
    const onlyJs = jsNames.filter((n) => !(n in rs));
    const onlyNative = rsNames.filter((n) => !(n in js));
    const exact: string[] = [];
    const normalized: string[] = [];
    const differing: unknown[] = [];
    for (const name of jsNames) {
      if (!(name in rs)) continue;
      if (equalBytes(js[name]!, rs[name]!)) continue;
      if (equalBytes(normalize(name, js[name]!), normalize(name, rs[name]!))) { normalized.push(name); continue; }
      const offset = firstDifference(js[name]!, rs[name]!);
      differing.push({ part: name, jsLength: js[name]!.length, nativeLength: rs[name]!.length, offset,
        js: context(js[name]!, offset), native: context(rs[name]!, offset) });
      exact.push(name);
    }
    if (onlyJs.length) details.onlyJs = onlyJs;
    if (onlyNative.length) details.onlyNative = onlyNative;
    if (normalized.length) details.normalizedParts = normalized;
    if (differing.length) details.differingParts = differing;

    let nativeWire: string | undefined;
    try { nativeWire = fs.readFileSync(file('w.json'), 'utf8'); } catch { /* none */ }
    if (input.jsWire !== undefined && nativeWire !== undefined) {
      details.wireIdentical = input.jsWire === nativeWire;
      if (!details.wireIdentical) {
        let path: string | null = '(unparsable)';
        try { path = firstJsonDifference(JSON.parse(input.jsWire), JSON.parse(nativeWire)); } catch { /* text only */ }
        details.wireFirstDifference = path ?? '(same value, different text)';
      }
    } else {
      details.wireIdentical = null;
    }
    let verdict: ProbeVerdict = 'identical';
    if (onlyJs.length || onlyNative.length || differing.length) verdict = 'diff';
    else if (normalized.length) verdict = 'same-parts-after-normalizing';
    return { verdict, details, nativeMs };
  } catch (error) {
    return { verdict: 'native-error', details: { error: String((error as Error)?.stack ?? error) } };
  } finally {
    try { fs.rmSync(dir, { recursive: true, force: true }); } catch { /* best effort */ }
  }
}

/**
 * The test that is saving, best effort: bun's `expect.getState` is not exposed, so this is the
 * first `*.test.ts` frame of the synchronous call stack (file and line of the save call).
 */
function callSite(): { file?: string; name?: string } {
  const frames = String(new Error().stack ?? '').split('\n').slice(1);
  for (const frame of frames) {
    const m = /\(?((?:\/|file:\/\/)[^():]+\.test\.ts):(\d+)(?::\d+)?\)?/.exec(frame);
    if (m) return { file: m[1]!.replace('file://', ''), name: `line ${m[2]}` };
  }
  return {};
}

interface SaveLike { bytes: Uint8Array }

/**
 * Captures the save's inputs. Returns the function to await with the JS result, or undefined when
 * the probe is off or the save uses internal comment options.
 */
export function beginNativeSaveProbe(
  session: YrsSession,
  options: { commentIdentities?: unknown; commentBaseline?: unknown; commentDeletion?: unknown }
): ((saved: SaveLike) => Promise<void>) | undefined {
  const report = env('DOCX_NATIVE_SAVE_PROBE');
  if (!report) return undefined;
  if (options.commentIdentities || options.commentBaseline || options.commentDeletion) return undefined;
  try {
    const original = session.materializeDocx()?.originalBuffer;
    if (!original) return undefined;
    const source = new Uint8Array(original.slice(0));
    const state = Uint8Array.from(session.encodeState());
    const metadata = Uint8Array.from(session.encodePeerMetadata());
    const test = callSite();
    let wire: string | undefined;
    (globalThis as ProbeGlobal).__docxSaveWireHook = (w) => { wire = w; };
    return async (saved) => {
      delete (globalThis as ProbeGlobal).__docxSaveWireHook;
      try {
        const result = await compareWithNative({
          jsBytes: saved.bytes, jsWire: wire, source, state, metadata,
        });
        const fs = await nodeImport('node:fs');
        const line = { ...test, verdict: result.verdict, details: result.details, nativeMs: result.nativeMs };
        fs.appendFileSync(report, `${JSON.stringify(line)}\n`);
      } catch { /* the probe never affects a save */ }
    };
  } catch {
    delete (globalThis as ProbeGlobal).__docxSaveWireHook;
    return undefined;
  }
}
