#!/usr/bin/env bun
/**
 * Native vs JS save benchmark.
 *
 *   bun scripts/native-save-bench.ts --out /tmp/bench.md [--runs 5] [--sizes 50,500,5000] [--real 3] [--no-node] [--build]
 *
 * Each document gets one small edit (insert text, comments kept), then the desktop serializer path
 * is timed (median of --runs after a warm-up): createYrsSession, openDocx(source, false),
 * bootstrapPeer(state, metadata, source, host) = "open"; saveYrsDocx = "save". Columns: this
 * checkout's JS under bun, the installed @betteroffice/docx under node (when it loads), and the
 * native_save CLI with --repeat (its first run is the warm-up). Peak RSS of the CLI comes from
 * `/usr/bin/time -l`.
 */

import { mkdirSync, mkdtempSync, readFileSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { spawnSync } from 'node:child_process';
import { parseArgs } from 'node:util';
import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../packages/docx/src/docx/rezip/parts';
import {
  CLI, ROOT, buildCli, captureInputs, createYrsSession, findCorpus, initWasm, median, openForEditing, rel,
  saveYrsDocx,
} from './lib/nativeSaveHarness';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const OFFICE = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
const RELNS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** A synthetic document: formatted paragraphs, a table every 25th, a comment every 1/4 of the body. */
export function syntheticDocx(paragraphs: number): Uint8Array {
  const body: string[] = [];
  const commentAt = new Map([0.1, 0.4, 0.7].map((f, i) => [Math.floor(paragraphs * f), i]));
  for (let i = 0; i < paragraphs; i += 1) {
    const c = commentAt.get(i);
    const text = `Paragraph ${i}: the quick brown fox jumps over the lazy dog, again and again. `;
    const runs = `<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">${text.slice(0, 20)}</w:t></w:r>` +
      `<w:r><w:rPr><w:i/><w:sz w:val="24"/></w:rPr><w:t xml:space="preserve">${text.slice(20)}</w:t></w:r>`;
    body.push(c === undefined ? `<w:p>${runs}</w:p>`
      : `<w:p><w:commentRangeStart w:id="${c}"/>${runs}<w:commentRangeEnd w:id="${c}"/><w:r><w:commentReference w:id="${c}"/></w:r></w:p>`);
    if (i % 25 === 24) {
      const cell = (t: string) => `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>${t}</w:t></w:r></w:p></w:tc>`;
      body.push(`<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/></w:tblPr><w:tblGrid>${'<w:gridCol w:w="3000"/>'.repeat(3)}</w:tblGrid>` +
        [0, 1].map((r) => `<w:tr>${[0, 1, 2].map((k) => cell(`r${r}c${k} @${i}`)).join('')}</w:tr>`).join('') + '</w:tbl>');
    }
  }
  const comments = [0, 1, 2].map((i) =>
    `<w:comment w:id="${i}" w:author="Bench" w:date="2026-10-01T10:00:00Z" w:initials="B"><w:p><w:r><w:t>Comment ${i}</w:t></w:r></w:p></w:comment>`).join('');
  const parts: PartsMap = new Map([
    ['[Content_Types].xml', toBytes(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${OFFICE}.document.main+xml"/><Override PartName="/word/comments.xml" ContentType="${OFFICE}.comments+xml"/></Types>`)],
    ['_rels/.rels', toBytes(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${RELNS}/officeDocument" Target="word/document.xml"/></Relationships>`)],
    ['word/_rels/document.xml.rels', toBytes(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${RELNS}/comments" Target="comments.xml"/></Relationships>`)],
    ['word/comments.xml', toBytes(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments xmlns:w="${W}">${comments}</w:comments>`)],
    ['word/document.xml', toBytes(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W}"><w:body>${body.join('')}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`)],
  ]);
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

interface Timing { open: number; save: number; total: number; outBytes: number }
interface Row { name: string; inBytes: number; js?: Timing; node?: Timing; native?: Timing & { rssMb?: number }; note?: string }

const med = (xs: number[]) => median(xs);
const fmt = (t?: Timing) => t ? `${t.open.toFixed(0)} / ${t.save.toFixed(0)} / ${t.total.toFixed(0)}` : 'n/a';

async function main() {
  const { values } = parseArgs({
    options: {
      out: { type: 'string' }, runs: { type: 'string', default: '5' }, sizes: { type: 'string', default: '50,500,5000' },
      real: { type: 'string', default: '3' }, 'no-node': { type: 'boolean' }, build: { type: 'boolean' },
    },
  });
  if (!values.out) throw new Error('--out <results.md> is required');
  const runs = Math.max(5, Number(values.runs));
  if (values.build) buildCli();
  await initWasm();
  const work = mkdtempSync(join(tmpdir(), 'native-save-bench-'));

  const docs: Array<{ name: string; source: Uint8Array }> = [];
  for (const n of values.sizes!.split(',').map(Number)) docs.push({ name: `synthetic ${n} paragraphs`, source: syntheticDocx(n) });
  for (const path of findCorpus(60).slice(0, Number(values.real))) docs.push({ name: rel(path), source: new Uint8Array(readFileSync(path)) });

  const rows: Row[] = [];
  const manifest: Array<{ name: string; source: string; state: string; metadata: string }> = [];
  for (const [i, doc] of docs.entries()) {
    const row: Row = { name: doc.name, inBytes: doc.source.length };
    rows.push(row);
    const dir = join(work, String(i));
    mkdirSync(dir);
    // The small edit, in a session seeded for editing.
    const { session: editing } = await openForEditing(doc.source);
    try {
      const p = editing.paragraphs('body').find((x) => x.text.length > 3);
      if (!p) { row.note = 'no body text to edit'; continue; }
      editing.insertText({ story: 'body', paraId: p.paraId, offset: 2 }, 'edit ');
    } catch (error) { row.note = `edit failed: ${(error as Error).message}`; continue; }
    const inputs = captureInputs(editing, doc.source);
    editing.destroy?.();
    const files = { source: join(dir, 'a.docx'), state: join(dir, 's.bin'), metadata: join(dir, 'm.bin') };
    writeFileSync(files.source, inputs.source); writeFileSync(files.state, inputs.state); writeFileSync(files.metadata, inputs.metadata);
    manifest.push({ name: doc.name, ...files });

    // (a) JS under bun, exactly the desktop serializer's calls.
    try {
      const open: number[] = []; const save: number[] = []; let outBytes = 0;
      for (let r = 0; r < runs + 1; r += 1) {
        const t0 = performance.now();
        const session = await createYrsSession();
        try {
          const host = session.openDocx(inputs.source, false);
          session.bootstrapPeer(inputs.state, inputs.metadata, inputs.source, host);
          const t1 = performance.now();
          const saved = await saveYrsDocx(session);
          const t2 = performance.now();
          outBytes = saved.bytes.length;
          if (r > 0) { open.push(t1 - t0); save.push(t2 - t1); }
        } finally { session.destroy(); }
      }
      row.js = { open: med(open), save: med(save), total: med(open) + med(save), outBytes };
    } catch (error) { row.note = `js failed: ${String((error as Error).message).slice(0, 120)}`; }

    // (b) native CLI.
    const run = spawnSync('/usr/bin/time', ['-l', CLI, '--source', files.source, '--state', files.state, '--metadata', files.metadata,
      '--out', join(dir, 'o.docx'), '--repeat', String(runs + 1), '--json'], { encoding: 'utf8', maxBuffer: 1 << 28 });
    if (run.status === 0) {
      const t = JSON.parse(run.stdout.trim().split('\n').pop()!);
      const open = med(t.openBootstrapMs.slice(1)); const save = med(t.saveMs.slice(1));
      const rss = /(\d+)\s+maximum resident set size/.exec(run.stderr);
      row.native = { open, save, total: open + save, outBytes: t.outBytes, rssMb: rss ? Number(rss[1]) / 1048576 : undefined };
    } else {
      row.note = `${row.note ?? ''} native: ${(run.stderr.split('\n').find((l) => /panicked|native_save:/.test(l)) ?? `exit ${run.status}`).slice(0, 140)}`.trim();
    }
  }

  // (c) node with the installed package.
  let nodeRss: number | undefined;
  if (!values['no-node'] && manifest.length) {
    const mpath = join(work, 'manifest.json');
    writeFileSync(mpath, JSON.stringify(manifest));
    const run = spawnSync('node', [join(ROOT, 'scripts/native-save-bench-node.mjs'), mpath, String(runs)],
      { encoding: 'utf8', maxBuffer: 1 << 28 });
    if (run.status === 0) {
      const res = JSON.parse(run.stdout.trim().split('\n').pop()!);
      nodeRss = res.__maxRssKb;
      for (const row of rows) {
        const r = res[row.name];
        if (r) row.node = { open: med(r.openBootstrapMs), save: med(r.saveMs), total: med(r.openBootstrapMs) + med(r.saveMs), outBytes: r.outBytes };
      }
    } else console.error('node variant failed:', run.stderr.slice(0, 600));
  }

  const lines = [
    '# Native save benchmark', '',
    `Median of ${runs} runs after one warm-up, milliseconds, cells are open+bootstrap / save / total. bun ${Bun.version}, node ${spawnSync('node', ['-v'], { encoding: 'utf8' }).stdout.trim()}.`,
    'JS (bun) is this checkout; JS (node) is the installed @betteroffice/docx (the desktop build); native is `native_save --repeat` in-process (no process start in the numbers).', '',
    '| document | source KB | JS bun | JS node | native | native/JS-bun save | out KB js / native | native peak RSS MB |',
    '|---|---|---|---|---|---|---|---|',
    ...rows.map((r) => `| ${r.name} | ${(r.inBytes / 1024).toFixed(1)} | ${fmt(r.js)} | ${fmt(r.node)} | ${fmt(r.native)} | ${r.js && r.native ? (r.native.save / r.js.save).toFixed(2) + 'x' : 'n/a'} | ${r.js ? (r.js.outBytes / 1024).toFixed(1) : '?'} / ${r.native ? (r.native.outBytes / 1024).toFixed(1) : '?'} | ${r.native?.rssMb?.toFixed(0) ?? 'n/a'} |`),
    '', `Peak RSS of the bun process (all JS runs, KB): ${process.resourceUsage().maxRSS}${nodeRss ? `; of the node process (KB): ${nodeRss}` : ''}.`,
    '', ...rows.filter((r) => r.note).map((r) => `- ${r.name}: ${r.note}`),
  ];
  writeFileSync(values.out, lines.join('\n') + '\n');
  console.log(lines.join('\n'));
}

if (import.meta.main) await main();
