// Builds the pptx and docx packages from this checkout and packs them as npm tarballs for a host app.
//
//   bun scripts/finx-pack.ts [--out <dir>] [--skip-build] [--allow-dirty]
//
// Output: <dir>/betteroffice-<pkg>-<version>.tgz for every package in PACKAGES, plus
// manifest.json and SHA256SUMS. The version is `<pptx version>-finx.g<commit>`; the
// packed manifests pin each other to exactly that version.
import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { mkdirSync, readFileSync, readdirSync, statSync, writeFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

// Dependency order: fonts and the i18n packages have no workspace dependencies in this set.
// @betteroffice/fonts is repacked: this checkout's docx imports `@betteroffice/fonts/progressive`,
// which the published 0.4.2 lacks. Its font packs stay on npm. The host overrides it too.
const PACKAGES = ['fonts', 'pptx-i18n', 'pptx', 'pptx-react', 'docx-i18n', 'docx', 'docx-react'] as const;
const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const args = process.argv.slice(2);
const flag = (name: string) => args.includes(name);
const outIndex = args.indexOf('--out');
if (outIndex !== -1 && (!args[outIndex + 1] || args[outIndex + 1].startsWith('--'))) throw new Error('--out requires a directory');
const out = resolve(outIndex !== -1 ? args[outIndex + 1] : join(root, 'dist-finx'));

function run(command: string, commandArgs: string[], cwd = root, capture = false): string {
  const result = spawnSync(command, commandArgs, { cwd, encoding: 'utf8', stdio: capture ? ['ignore', 'pipe', 'inherit'] : 'inherit' });
  if (result.status !== 0) throw new Error(`${command} ${commandArgs.join(' ')} failed (${result.status})`);
  return result.stdout?.trim() ?? '';
}

const sha = run('git', ['rev-parse', '--short=8', 'HEAD'], root, true);
const dirty = run('git', ['status', '--porcelain'], root, true) !== '';
if (dirty && !flag('--allow-dirty')) throw new Error('the checkout has uncommitted changes; commit them or pass --allow-dirty');

const manifestPath = (name: string) => join(root, 'packages', name, 'package.json');
const base = (JSON.parse(readFileSync(manifestPath('pptx'), 'utf8')) as { version: string }).version.replace(/-.*$/, '');
// The `g` keeps an all-digit commit from being an invalid numeric semver identifier.
const version = `${base}-finx.g${sha}${dirty ? '.dirty' : ''}`;
const names = new Map(PACKAGES.map((name) => [(JSON.parse(readFileSync(manifestPath(name), 'utf8')) as { name: string }).name, name]));

if (!flag('--skip-build')) {
  run('bun', ['run', 'build:pptx-wasm']);
  run('bun', ['run', 'build:fonts']); // pptx-react's type build imports it
  run('bun', ['run', 'build:pptx']);
  run('bun', ['run', 'build:docx-wasm']);
  run('bun', ['run', 'build:docx']);
}

for (const name of PACKAGES) {
  const dist = join(root, 'packages', name, 'dist');
  if (!statSync(dist, { throwIfNoEntry: false })?.isDirectory()) throw new Error(`packages/${name}/dist is missing; run without --skip-build`);
}
for (const wasm of ['pptx_wasm_bg.wasm', 'pptx_wasm_lean_bg.wasm', 'metafile_svg_wasm_bg.wasm']) {
  if (!statSync(join(root, 'packages/pptx/dist/generated', wasm), { throwIfNoEntry: false })?.isFile()) throw new Error(`packages/pptx/dist/generated/${wasm} is missing`);
}

if (statSync(out, { throwIfNoEntry: false }) && readdirSync(out).length > 0) throw new Error(`output directory is not empty: ${out}`);
mkdirSync(out, { recursive: true });

const DEPENDENCY_FIELDS = ['dependencies', 'devDependencies', 'peerDependencies', 'optionalDependencies'] as const;
const packed: { name: string; version: string; file: string; bytes: number; sha256: string }[] = [];
for (const dir of PACKAGES) {
  const path = manifestPath(dir);
  const original = readFileSync(path, 'utf8');
  const manifest = JSON.parse(original) as Record<string, any>;
  manifest.version = version;
  for (const field of DEPENDENCY_FIELDS) {
    for (const dependency of Object.keys(manifest[field] ?? {})) {
      if (names.has(dependency) && String(manifest[field][dependency]).startsWith('workspace:')) manifest[field][dependency] = version;
    }
  }
  try {
    writeFileSync(path, `${JSON.stringify(manifest, null, 2)}\n`);
    // `bun pm pack` resolves the remaining workspace: ranges (optional font peers) to real versions.
    run('bun', ['pm', 'pack', '--destination', out], join(root, 'packages', dir));
  } finally {
    writeFileSync(path, original);
  }
}

for (const file of readdirSync(out).filter((entry) => entry.endsWith('.tgz')).sort()) {
  const listing = run('tar', ['-xOzf', join(out, file), 'package/package.json'], root, true);
  if (listing.includes('workspace:')) throw new Error(`${file} still contains a workspace: range`);
  const { name, version: packedVersion } = JSON.parse(listing) as { name: string; version: string };
  if (packedVersion !== version) throw new Error(`${file} has version ${packedVersion}, expected ${version}`);
  const bytes = readFileSync(join(out, file));
  packed.push({ name, version, file, bytes: bytes.length, sha256: createHash('sha256').update(bytes).digest('hex') });
}
if (packed.length !== PACKAGES.length) throw new Error(`expected ${PACKAGES.length} tarballs, found ${packed.length}`);

writeFileSync(join(out, 'manifest.json'), `${JSON.stringify({ version, commit: run('git', ['rev-parse', 'HEAD'], root, true), dirty, packages: packed }, null, 2)}\n`);
writeFileSync(join(out, 'SHA256SUMS'), packed.map((entry) => `${entry.sha256}  ${entry.file}`).join('\n') + '\n');
for (const entry of packed) console.log(`${entry.file}  ${(entry.bytes / 1024 / 1024).toFixed(2)} MiB`);
console.log(`version ${version} -> ${out}`);
