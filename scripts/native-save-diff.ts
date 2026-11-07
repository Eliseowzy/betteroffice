#!/usr/bin/env bun
/**
 * Differential probe runner: runs the TS save tests with `DOCX_NATIVE_SAVE_PROBE` set so every
 * `saveYrsDocx` is replayed through the native CLI and compared, then summarizes the report.
 *
 *   bun scripts/native-save-diff.ts [--report /tmp/native-save-probe.jsonl] [--no-build]
 *       [--build-wasm] [--summary-only] [--filter substring] [-- extra bun test args]
 *
 * --build-wasm runs `bun run build:docx-wasm` first (needed once; tests do not build it).
 */

import { existsSync, readFileSync, readdirSync, rmSync } from 'node:fs';
import { join } from 'node:path';
import { spawnSync } from 'node:child_process';
import { parseArgs } from 'node:util';
import { ROOT, CLI, buildCli, rel } from './lib/nativeSaveHarness';

/** Test files under packages/docx/src that are named like save tests or call `saveYrsDocx(`. */
export function saveTestFiles(filter?: string): string[] {
  const found: string[] = [];
  const walk = (dir: string): void => {
    for (const entry of readdirSync(dir, { withFileTypes: true })) {
      const path = join(dir, entry.name);
      if (entry.isDirectory()) { if (entry.name !== 'generated' && entry.name !== 'node_modules') walk(path); continue; }
      if (!entry.name.endsWith('.test.ts')) continue;
      if (/Save|residentSave/.test(entry.name) || readFileSync(path, 'utf8').includes('saveYrsDocx(')) found.push(path);
    }
  };
  walk(join(ROOT, 'packages/docx/src'));
  return found.filter((f) => !filter || f.includes(filter)).sort();
}

interface Row {
  file?: string; name?: string; verdict: string; details: Record<string, any>;
  nativeMs?: { openBootstrap: number; save: number };
}

export function summarize(report: string): string {
  const rows: Row[] = existsSync(report)
    ? readFileSync(report, 'utf8').split('\n').filter(Boolean).map((l) => JSON.parse(l))
    : [];
  const count = (key: (r: Row) => string) => {
    const m = new Map<string, number>();
    for (const r of rows) m.set(key(r), (m.get(key(r)) ?? 0) + 1);
    return [...m].sort((a, b) => b[1] - a[1]);
  };
  const out: string[] = [`${rows.length} probed saves`, '', 'Per verdict:'];
  for (const [v, n] of count((r) => r.verdict)) out.push(`  ${v}: ${n}`);
  out.push('', 'Per test file (verdict: count):');
  const files = new Map<string, Map<string, number>>();
  for (const r of rows) {
    const f = r.file ? rel(r.file) : '(unknown)';
    const m = files.get(f) ?? new Map();
    m.set(r.verdict, (m.get(r.verdict) ?? 0) + 1);
    files.set(f, m);
  }
  for (const [f, m] of [...files].sort()) {
    out.push(`  ${f}: ${[...m].map(([v, n]) => `${v} ${n}`).join(', ')}`);
  }
  const errors = count((r) => r.verdict === 'native-error'
    ? String(r.details.stderr ?? r.details.error ?? '').split('\n').find((l) => l.trim()) ?? '(no stderr)' : '').filter(([k]) => k);
  if (errors.length) {
    out.push('', 'native-error causes (first stderr line: count):');
    for (const [k, n] of errors.slice(0, 10)) out.push(`  ${n} x ${k.slice(0, 200)}`);
  }
  const diffs = rows.filter((r) => r.verdict === 'diff');
  if (diffs.length) {
    out.push('', 'Diffs:');
    for (const r of diffs) {
      const d = r.details;
      out.push(`  ${r.file ? rel(r.file) : '?'} :: ${r.name ?? '?'}`);
      if (d.onlyJs) out.push(`    parts only in JS: ${d.onlyJs.join(', ')}`);
      if (d.onlyNative) out.push(`    parts only in native: ${d.onlyNative.join(', ')}`);
      for (const p of d.differingParts ?? []) {
        out.push(`    ${p.part} @${p.offset} (js ${p.jsLength}B, native ${p.nativeLength}B)`);
        out.push(`      js:     ${JSON.stringify(p.js)}`);
        out.push(`      native: ${JSON.stringify(p.native)}`);
      }
      if (d.wireIdentical === false) out.push(`    wire differs first at ${d.wireFirstDifference}`);
    }
  }
  const wireDiff = rows.filter((r) => r.details?.wireIdentical === false).length;
  const timed = rows.filter((r) => r.nativeMs);
  out.push('', `wire identical: ${rows.filter((r) => r.details?.wireIdentical === true).length}, differs: ${wireDiff}, unknown: ${rows.filter((r) => r.details?.wireIdentical == null).length}`);
  if (timed.length) out.push(`native runs timed: ${timed.length}`);
  return out.join('\n');
}

if (import.meta.main) {
  const sep = process.argv.indexOf('--');
  const extra = sep >= 0 ? process.argv.slice(sep + 1) : [];
  const { values } = parseArgs({
    args: sep >= 0 ? process.argv.slice(2, sep) : process.argv.slice(2),
    options: {
      report: { type: 'string', default: '/tmp/native-save-probe.jsonl' },
      'no-build': { type: 'boolean' }, 'build-wasm': { type: 'boolean' },
      'summary-only': { type: 'boolean' }, filter: { type: 'string' },
    },
  });
  const report = values.report!;
  if (!values['summary-only']) {
    if (values['build-wasm']) {
      if (spawnSync('bun', ['run', 'build:docx-wasm'], { cwd: ROOT, stdio: 'inherit' }).status !== 0) process.exit(1);
    }
    if (!values['no-build']) buildCli();
    rmSync(report, { force: true });
    const files = saveTestFiles(values.filter);
    console.log(`running ${files.length} test files with DOCX_NATIVE_SAVE_PROBE=${report}`);
    const run = spawnSync('bun', ['test', '--isolate', ...extra, ...files], {
      cwd: ROOT, stdio: 'inherit',
      env: { ...process.env, DOCX_NATIVE_SAVE_PROBE: report, DOCX_NATIVE_SAVE_CLI: CLI },
    });
    console.log(`bun test exit status ${run.status} (non-zero means a JS test failed with the probe on)`);
  }
  console.log(`\n${summarize(report)}`);
}
