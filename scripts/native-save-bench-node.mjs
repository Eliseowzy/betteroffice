#!/usr/bin/env node
// Node timing of the desktop serializer path with the installed @betteroffice/docx build.
//   node scripts/native-save-bench-node.mjs <manifest.json> [runs]
// manifest: [{ name, source, state, metadata }] (file paths). Prints JSON { name: {openBootstrapMs:[], saveMs:[]} }.
// Resolves the package from NATIVE_SAVE_NODE_PKG_DIR (default: the FinxWork word worktree's desktop app).
import { readFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { pathToFileURL } from 'node:url';

const appDir = process.env.NATIVE_SAVE_NODE_PKG_DIR
  ?? '/Users/eliseo/FinxWork/.claude/worktrees/word/apps/desktop';
const require = createRequire(`${appDir}/package.json`);
const { createYrsSession, saveYrsDocx } = await import(pathToFileURL(require.resolve('@betteroffice/docx/yrs')).href);
const manifest = JSON.parse(readFileSync(process.argv[2], 'utf8'));
const runs = Number(process.argv[3] ?? 6);
const result = {};
for (const doc of manifest) {
  const source = new Uint8Array(readFileSync(doc.source));
  const state = new Uint8Array(readFileSync(doc.state));
  const metadata = new Uint8Array(readFileSync(doc.metadata));
  const openBootstrapMs = [];
  const saveMs = [];
  let outBytes = 0;
  for (let i = 0; i < runs + 1; i += 1) { // run 0 is the warm-up
    const t0 = performance.now();
    const session = await createYrsSession();
    try {
      const host = session.openDocx(source, false);
      session.bootstrapPeer(state, metadata, source, host);
      const t1 = performance.now();
      const saved = await saveYrsDocx(session);
      const t2 = performance.now();
      outBytes = saved.bytes.length;
      if (i > 0) { openBootstrapMs.push(t1 - t0); saveMs.push(t2 - t1); }
    } finally {
      session.destroy();
    }
  }
  result[doc.name] = { openBootstrapMs, saveMs, outBytes };
}
result.__maxRssKb = process.resourceUsage().maxRSS;
console.log(JSON.stringify(result));
