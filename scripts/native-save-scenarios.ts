#!/usr/bin/env bun
/**
 * Corpus x scripted-edit differential of the native DOCX save against the TS save.
 *
 *   bun scripts/native-save-scenarios.ts --out /tmp/scenarios.md [--limit N] [--only scenarioName] [--build]
 *
 * Each cell opens a fixture with `openDocx(bytes, true)`, applies a scripted edit through the TS
 * session API, saves with JS (`saveYrsDocx`) and with `native_save` (inputs: source, encodeState,
 * encodePeerMetadata) and compares the packages part by part. An edit the document cannot take
 * (no such story, API refusal) is `skipped`, with the reason in the details.
 */

import { writeFileSync } from 'node:fs';
import { parseArgs } from 'node:util';
import {
  buildCli, findCorpus, initWasm, openForEditing, rel, saveAndCompare, type YrsSession,
} from './lib/nativeSaveHarness';
import { readFileSync } from 'node:fs';

const AUTHOR = { name: 'Tester', date: '2026-10-02T12:00:00Z' };
class Skip extends Error {}
const skip = (why: string): never => { throw new Skip(why); };

const loc = (story: string, paraId: string, offset: number) => ({ story, paraId, offset });
const range = (story: string, paraId: string, start: number, end: number) =>
  ({ story, start: { paraId, offset: start }, end: { paraId, offset: end } });

/** First body paragraph with some text. */
function bodyParagraph(s: YrsSession, min = 4) {
  const p = s.paragraphs('body').find((x) => x.text.length >= min);
  return p ?? skip('no body paragraph with text');
}

function tableTarget(s: YrsSession) {
  const p = s.paragraphs('body')[0] ?? skip('empty body');
  const receipt = s.insertTable(loc('body', p.paraId, 0), 2, 3);
  return receipt.table;
}

function storyWithPrefix(s: YrsSession, ...prefixes: string[]) {
  const id = s.storyIds().find((x) => prefixes.some((p) => x.startsWith(p)) && s.paragraphs(x).length > 0);
  return id ?? skip(`no story ${prefixes.join('/')}`);
}

function trackedInsert(s: YrsSession) {
  const p = bodyParagraph(s);
  s.insertText(loc('body', p.paraId, 1), 'tracked', AUTHOR);
}

function trackedDelete(s: YrsSession) {
  const p = bodyParagraph(s);
  s.deleteRange(range('body', p.paraId, 0, Math.min(3, p.text.length)), AUTHOR);
}

function resolveFirst(s: YrsSession, how: 'acceptChange' | 'rejectChange') {
  const revision = s.listRevisions()[0] ?? skip('no revision after the tracked edit');
  s[how]({ revisionId: revision.revisionId });
}

export const SCENARIOS: Record<string, (s: YrsSession) => void> = {
  insertText: (s) => { const p = bodyParagraph(s); s.insertText(loc('body', p.paraId, 1), 'Hello '); },
  deleteText: (s) => { const p = bodyParagraph(s); s.deleteRange(range('body', p.paraId, 0, 2)); },
  bold: (s) => { const p = bodyParagraph(s); s.toggleMark(range('body', p.paraId, 0, Math.min(5, p.text.length)), { type: 'bold' }); },
  italic: (s) => { const p = bodyParagraph(s); s.toggleMark(range('body', p.paraId, 0, Math.min(5, p.text.length)), { type: 'italic' }); },
  fontSize: (s) => { const p = bodyParagraph(s); s.toggleMark(range('body', p.paraId, 0, Math.min(5, p.text.length)), { type: 'fontSize', value: 18 }); },
  splitParagraph: (s) => { const p = bodyParagraph(s); s.splitParagraph(loc('body', p.paraId, 2)); },
  mergeParagraphs: (s) => {
    const ps = s.paragraphs('body');
    if (ps.length < 2) skip('single paragraph');
    s.mergeParagraphs('body', ps[0]!.paraId);
  },
  trackedInsert,
  trackedDelete,
  acceptRevision: (s) => { trackedInsert(s); resolveFirst(s, 'acceptChange'); },
  rejectRevision: (s) => { trackedDelete(s); resolveFirst(s, 'rejectChange'); },
  addComment: (s) => {
    const p = bodyParagraph(s);
    s.addComment([range('body', p.paraId, 0, Math.min(4, p.text.length))], AUTHOR.name, AUTHOR.date, 'A comment');
  },
  commentReply: (s) => {
    const p = bodyParagraph(s);
    const r = [range('body', p.paraId, 0, Math.min(4, p.text.length))];
    const parent = s.addComment(r, AUTHOR.name, AUTHOR.date, 'Parent');
    const reply = s.addComment(r, 'Replier', AUTHOR.date, 'Reply');
    s.updateComment(reply.commentId, { parentId: parent.commentId });
  },
  resolveComment: (s) => {
    const p = bodyParagraph(s);
    const c = s.addComment([range('body', p.paraId, 0, Math.min(4, p.text.length))], AUTHOR.name, AUTHOR.date, 'Done soon');
    s.updateComment(c.commentId, { done: true });
  },
  deleteComment: (s) => {
    const p = bodyParagraph(s);
    const c = s.addComment([range('body', p.paraId, 0, Math.min(4, p.text.length))], AUTHOR.name, AUTHOR.date, 'Gone');
    s.applyRawOps('body', [{ op: 'removeComment', id: c.commentId }]);
  },
  insertTable: (s) => { tableTarget(s); },
  insertTableRow: (s) => { const t = tableTarget(s); s.insertRow({ ...t, row: 0, column: 0 }, 'below'); },
  deleteTableRow: (s) => {
    const t = tableTarget(s);
    const cell = { ...t, row: 1, column: 0 };
    s.deleteRow({ anchor: cell, head: cell });
  },
  mergeCells: (s) => {
    const t = tableTarget(s);
    s.mergeCells({ anchor: { ...t, row: 0, column: 0 }, head: { ...t, row: 0, column: 1 } });
  },
  editHeaderFooter: (s) => {
    const story = storyWithPrefix(s, 'hf:');
    s.insertText(loc(story, s.paragraphs(story)[0]!.paraId, 0), 'Header ');
  },
  editFootnote: (s) => {
    const story = storyWithPrefix(s, 'fn', 'footnote', 'note');
    s.insertText(loc(story, s.paragraphs(story)[0]!.paraId, 0), 'Note ');
  },
  sectionMargins: (s) => { s.setSectionProperties(s.listSections()[0]?.closingParagraphId ?? skip('no section'), { margins: { left: 1000, right: 1200 } }); },
  sectionOrientation: (s) => {
    const first = s.listSections()[0] ?? skip('no section');
    const w = Number(first.resolved.pageWidth ?? 12240);
    const h = Number(first.resolved.pageHeight ?? 15840);
    s.setSectionProperties(first.closingParagraphId, { orientation: 'landscape', pageWidth: Math.max(w, h), pageHeight: Math.min(w, h) });
  },
  toggleList: (s) => {
    const p = bodyParagraph(s);
    s.setParagraphAttrs(range('body', p.paraId, 0, 0), { numPr: { numId: 1, ilvl: 0 } });
  },
  insertHyperlink: (s) => {
    const p = bodyParagraph(s);
    s.setHyperlink(range('body', p.paraId, 0, Math.min(4, p.text.length)), { href: 'https://example.com/' });
  },
  pageBreak: (s) => { const p = bodyParagraph(s); s.insertPageBreak(loc('body', p.paraId, 2)); },
};

async function main() {
  const { values } = parseArgs({
    options: { out: { type: 'string' }, limit: { type: 'string' }, only: { type: 'string' }, build: { type: 'boolean' } },
  });
  if (!values.out) throw new Error('--out <report.md> is required');
  if (values.build) buildCli();
  await initWasm();
  let docs = findCorpus(values.limit ? Number(values.limit) : 24);
  const names = Object.keys(SCENARIOS).filter((n) => !values.only || values.only.split(',').includes(n));
  console.log(`corpus (${docs.length}):\n${docs.map((d) => `  ${rel(d)}`).join('\n')}`);

  type Cell = { verdict: string; note?: string };
  const grid = new Map<string, Cell>();
  const notes: string[] = [];
  const totals: Record<string, number> = {};
  for (const doc of docs) {
    const source = new Uint8Array(readFileSync(doc));
    for (const name of names) {
      let cell: Cell;
      try {
        const { session } = await openForEditing(source);
        try {
          SCENARIOS[name]!(session);
        } catch (error) {
          if (error instanceof Skip) throw error;
          throw new Skip(`edit refused: ${String((error as Error)?.message ?? error).slice(0, 160)}`);
        }
        let result;
        try {
          result = await saveAndCompare(session, source);
        } finally {
          session.destroy?.();
        }
        cell = { verdict: result.verdict };
        if (result.verdict !== 'identical') {
          notes.push(`### ${rel(doc)} / ${name}: ${result.verdict}\n\n\`\`\`json\n${JSON.stringify(result.details, null, 1).slice(0, 3000)}\n\`\`\`\n`);
        }
      } catch (error) {
        cell = error instanceof Skip ? { verdict: 'skipped', note: error.message }
          : { verdict: 'js-error', note: String((error as Error)?.message ?? error).slice(0, 200) };
      }
      grid.set(`${doc}\0${name}`, cell);
      totals[cell.verdict] = (totals[cell.verdict] ?? 0) + 1;
    }
    console.log(`${rel(doc)} done`);
  }

  const short: Record<string, string> = {
    identical: 'ok', 'same-parts-after-normalizing': 'norm', diff: 'DIFF', 'native-error': 'ERR', skipped: 'skip', 'js-error': 'JSERR',
  };
  const lines = [
    '# Native save scenarios', '',
    `Totals: ${Object.entries(totals).map(([k, v]) => `${k} ${v}`).join(', ')}`, '',
    'ok = identical, norm = same after masking dcterms:modified, DIFF = packages differ, ERR = native CLI failed, skip = edit unsupported on the doc, JSERR = JS save failed.', '',
    `| document | ${names.join(' | ')} |`, `|---|${names.map(() => '---').join('|')}|`,
    ...docs.map((d) => `| ${rel(d)} | ${names.map((n) => short[grid.get(`${d}\0${n}`)!.verdict]).join(' | ')} |`),
    '', '## Skips and JS errors', '',
    ...docs.flatMap((d) => names.flatMap((n) => {
      const c = grid.get(`${d}\0${n}`)!;
      return c.note ? [`- ${rel(d)} / ${n}: ${c.verdict}: ${c.note}`] : [];
    })),
    '', '## Non-identical details', '', ...notes,
  ];
  writeFileSync(values.out, lines.join('\n'));
  console.log(`wrote ${values.out}: ${JSON.stringify(totals)}`);
}

if (import.meta.main) await main();
