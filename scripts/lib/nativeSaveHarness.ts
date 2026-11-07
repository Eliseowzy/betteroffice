/** Shared plumbing of the native-save differential and benchmark scripts. */

import { readFileSync, readdirSync, statSync } from 'node:fs';
import { join, relative, resolve } from 'node:path';
import { spawnSync } from 'node:child_process';
import { preloadEditWasm } from '../../packages/docx/src/wasm/edit';
import { preloadOpcWasm } from '../../packages/docx/src/wasm/opc';
import { createYrsSession, saveYrsDocx, type YrsSession } from '../../packages/docx/src/yrs';
import { compareWithNative, type NativeComparison } from '../../packages/docx/src/yrs/nativeSaveProbe';

export const ROOT = resolve(import.meta.dir, '../..');
export const CLI = process.env.DOCX_NATIVE_SAVE_CLI ?? join(ROOT, 'target/release/examples/native_save');
export { compareWithNative, createYrsSession, saveYrsDocx };
export type { NativeComparison, YrsSession };

export function buildCli(): void {
  const run = spawnSync('cargo', ['build', '--release', '-p', 'betteroffice-docx-native-save', '--example', 'native_save'],
    { cwd: ROOT, stdio: 'inherit' });
  if (run.status !== 0) throw new Error('cargo build of the native_save example failed');
}

export async function initWasm(): Promise<void> {
  const gen = join(ROOT, 'packages/docx/src/wasm/generated');
  await preloadEditWasm(new Uint8Array(readFileSync(join(gen, 'edit/docx_edit_bg.wasm'))));
  await preloadOpcWasm(new Uint8Array(readFileSync(join(gen, 'opc/ooxml_opc_bg.wasm'))));
}

/** Real `.docx` fixtures of the repo, largest first, at most `cap`: one per directory, plus the loose ones. */
export function findCorpus(cap = 24): string[] {
  const found: Array<{ path: string; size: number }> = [];
  const skip = new Set(['node_modules', 'target', '.git', 'dist', 'generated', '.claude']);
  const walk = (dir: string): void => {
    for (const name of readdirSync(dir)) {
      if (skip.has(name)) continue;
      const path = join(dir, name);
      const st = statSync(path, { throwIfNoEntry: false });
      if (!st) continue;
      if (st.isDirectory()) walk(path);
      else if (name.endsWith('.docx') && st.size > 0) found.push({ path, size: st.size });
    }
  };
  walk(ROOT);
  const byDir = new Map<string, typeof found>();
  for (const f of found) {
    const dir = f.path.slice(0, f.path.lastIndexOf('/'));
    byDir.set(dir, [...(byDir.get(dir) ?? []), f]);
  }
  const picked: typeof found = [];
  for (const files of byDir.values()) {
    files.sort((a, b) => b.size - a.size);
    picked.push(...(files.length > 2 ? files.slice(0, 1) : files));
  }
  return picked.sort((a, b) => b.size - a.size).slice(0, cap).map((f) => f.path);
}

export const rel = (path: string): string => relative(ROOT, path);

export interface Opened { session: YrsSession; source: Uint8Array }

/** `createYrsSession()` + `openDocx(source, true)`: how the TS tests (and editing) seed a session. */
export async function openForEditing(source: Uint8Array): Promise<Opened> {
  const session = await createYrsSession();
  session.openDocx(source, true);
  return { session, source };
}

/** What the desktop serializer holds after the edits: state, peer metadata, source. */
export function captureInputs(session: YrsSession, source: Uint8Array) {
  return {
    source,
    state: Uint8Array.from(session.encodeState()),
    metadata: Uint8Array.from(session.encodePeerMetadata()),
  };
}

type WireHost = { __docxSaveWireHook?: (wire: string) => void };

/** JS save (capturing its writer request) and the native save of the same inputs, compared. */
export async function saveAndCompare(session: YrsSession, source: Uint8Array): Promise<NativeComparison> {
  const inputs = captureInputs(session, source);
  let wire: string | undefined;
  (globalThis as WireHost).__docxSaveWireHook = (w) => { wire = w; };
  let saved;
  try {
    saved = await saveYrsDocx(session);
  } finally {
    delete (globalThis as WireHost).__docxSaveWireHook;
  }
  return compareWithNative({ ...inputs, jsBytes: saved.bytes, jsWire: wire, cliPath: CLI });
}

export function median(values: number[]): number {
  const s = [...values].sort((a, b) => a - b);
  const m = s.length >> 1;
  return s.length % 2 ? s[m]! : (s[m - 1]! + s[m]!) / 2;
}
