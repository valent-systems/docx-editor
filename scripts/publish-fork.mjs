#!/usr/bin/env node
/**
 * Publish the packages to a forked npm scope (e.g. `@sqren`) WITHOUT committing
 * the rename into source — so the tree stays `@eigenpal`-scoped and mergeable
 * with upstream.
 *
 * The previous ad-hoc publish only rescoped each package's top-level `name`. It
 * left the inter-package `dependencies` and EVERY bundled `import` in `dist/`
 * pointing at `@eigenpal/docx-editor-*`, so `@sqren/docx-editor-react` declared a
 * dependency on `@eigenpal/docx-editor-core@<fork-version>` (which doesn't
 * exist) and its code imported `@eigenpal/docx-editor-core` at runtime. The
 * result installed-fails and would runtime-fail. This script rescopes ALL three:
 *
 *   1. package.json `name`
 *   2. package.json `dependencies`/`peerDependencies`/`optionalDependencies`
 *      keys that reference a sibling fork package (and pins them to the fork
 *      version)
 *   3. every `@eigenpal/docx-editor-*` (and `@eigenpal/nuxt-docx-editor`) import
 *      specifier inside the built `dist/`
 *
 * It edits each package.json in place only for the duration of `npm publish`,
 * then restores it (`dist/` is gitignored, so its rewrites don't dirty the tree).
 *
 * Usage:
 *   node scripts/publish-fork.mjs --version 1.7.0-fork.0 --tag fork            # dry-run (default)
 *   node scripts/publish-fork.mjs --version 1.7.0-fork.0 --tag fork --publish  # actually publish
 *   node scripts/publish-fork.mjs --version 1.7.0-fork.0 --scope @sqren --publish --skip-build
 *
 * Run `bun run build:packages` first (or omit --skip-build to let it build).
 */

import { execSync } from 'node:child_process';
import { readFileSync, writeFileSync, readdirSync, statSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');

// ── args ─────────────────────────────────────────────────────────────────────
function arg(name, fallback) {
  const i = process.argv.indexOf(`--${name}`);
  if (i === -1) return fallback;
  const next = process.argv[i + 1];
  return next && !next.startsWith('--') ? next : true;
}
const TO_SCOPE = String(arg('scope', '@sqren')).replace(/\/$/, '');
const VERSION = arg('version', null);
const NPM_TAG = String(arg('tag', 'fork'));
const PUBLISH = arg('publish', false) === true;
const SKIP_BUILD = arg('skip-build', false) === true;

if (!VERSION) {
  console.error('Missing --version (e.g. --version 1.7.0-fork.0)');
  process.exit(1);
}

// Publishable packages, in dependency order (deps publish before dependents).
const PACKAGES = [
  'packages/i18n',
  'packages/core',
  'packages/agents',
  'packages/react',
  'packages/vue',
  'packages/nuxt',
];

const FROM_SCOPE = '@eigenpal';
// Rescope a package name from the source scope to the fork scope. Matches both
// `@eigenpal/docx-editor-*` and `@eigenpal/nuxt-docx-editor`.
// Global for whole-file/string replacement…
const SIBLING_RE = new RegExp(`${FROM_SCOPE}/((?:nuxt-)?docx-editor[\\w-]*)`, 'g');
const rescope = (s) => s.replace(SIBLING_RE, `${TO_SCOPE}/$1`);
// …and a separate stateless, anchored test for dependency keys (a global regex
// is stateful under repeated .test()).
const isSiblingDep = (key) => new RegExp(`^${FROM_SCOPE}/(?:nuxt-)?docx-editor`).test(key);

const DIST_EXT = /\.(m?js|c?js|d\.m?ts|d\.c?ts|map|json)$/;

function* walk(dir) {
  for (const entry of readdirSync(dir)) {
    const p = join(dir, entry);
    if (statSync(p).isDirectory()) yield* walk(p);
    else yield p;
  }
}

function rewriteDistImports(pkgDir) {
  const dist = join(pkgDir, 'dist');
  let count = 0;
  try {
    statSync(dist);
  } catch {
    console.warn(`  ! no dist/ in ${pkgDir} — did you build?`);
    return 0;
  }
  for (const file of walk(dist)) {
    if (!DIST_EXT.test(file)) continue;
    const before = readFileSync(file, 'utf8');
    const after = rescope(before);
    if (after !== before) {
      writeFileSync(file, after);
      count++;
    }
  }
  return count;
}

const log = (...a) => console.log(...a);
const run = (cmd, cwd) => execSync(cmd, { cwd, stdio: 'inherit' });

if (!SKIP_BUILD) {
  log('› Building packages…');
  run('bun run build:packages', ROOT);
}

log(`\n› Fork publish → scope ${TO_SCOPE}, version ${VERSION}, tag ${NPM_TAG}, ${PUBLISH ? 'PUBLISH' : 'DRY-RUN'}\n`);

const restore = [];
let failed = false;
try {
  for (const rel of PACKAGES) {
    const pkgDir = join(ROOT, rel);
    const pkgJsonPath = join(pkgDir, 'package.json');
    const original = readFileSync(pkgJsonPath, 'utf8');
    const pkg = JSON.parse(original);
    restore.push([pkgJsonPath, original]);

    const newName = rescope(pkg.name);
    pkg.name = newName;
    pkg.version = VERSION;

    // Rescope + pin sibling fork deps; leave third-party deps untouched.
    for (const field of ['dependencies', 'peerDependencies', 'optionalDependencies']) {
      const deps = pkg[field];
      if (!deps) continue;
      const next = {};
      for (const [key, range] of Object.entries(deps)) {
        if (isSiblingDep(key)) {
          next[rescope(key)] = VERSION; // sibling fork packages publish at the same version
        } else {
          next[key] = range;
        }
      }
      pkg[field] = next;
    }

    writeFileSync(pkgJsonPath, JSON.stringify(pkg, null, 2) + '\n');
    const rewritten = rewriteDistImports(pkgDir);
    log(`  ${pkg.name}@${VERSION}  (rescoped ${rewritten} dist file${rewritten === 1 ? '' : 's'})`);

    const publishCmd = PUBLISH
      ? `npm publish --access public --tag ${NPM_TAG}`
      : `npm publish --access public --tag ${NPM_TAG} --dry-run`;
    run(publishCmd, pkgDir);
  }
} catch (err) {
  failed = true;
  console.error('\n✗ Publish aborted:', err.message);
} finally {
  // Restore source package.json files (dist/ is gitignored — leave it).
  for (const [path, content] of restore) writeFileSync(path, content);
  log('\n› Restored source package.json files.');
}

process.exit(failed ? 1 : 0);
