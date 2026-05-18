#!/usr/bin/env node
// Run API Extractor over every published subpath in packages/core/package.json's
// exports map. Generates one `etc/<slug>.api.md` snapshot per entry. Pass
// `--local` to write/update snapshots; default is CI mode (compare and fail
// on drift).
//
// Why programmatic instead of one config file per subpath: core ships ~60
// public entry points. Per-file configs would clutter the repo root. The
// runner derives each invocation from package.json — adding a new subpath
// to the exports map automatically extends the API surface check.
//
// Perf: one CompilerState built from the first entry is reused for all 61
// invocations via `additionalEntryPoints`, so we only parse tsconfig + walk
// the dist tree once instead of 61 times.

import { CompilerState, Extractor, ExtractorConfig } from '@microsoft/api-extractor';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const packageRoot = path.resolve(__dirname, '..');
const packageJsonPath = path.join(packageRoot, 'package.json');
const etcDir = path.join(packageRoot, 'etc');
const tempDir = path.join(packageRoot, 'temp');

const isLocal = process.argv.includes('--local');

const pkg = JSON.parse(fs.readFileSync(packageJsonPath, 'utf8'));

function slugForKey(key) {
  if (key === '.') return 'index';
  return key.replace(/^\.\//, '').replace(/\//g, '-');
}

function sourcePathForEntry(value) {
  // Resolve to the src path that produced the dist .d.ts, so drift output
  // can tell a contributor where to look. tsup writes `./dist/<x>.js` from
  // `./src/<x>.ts` (or `./src/<x>/index.ts` for directory entries).
  const importPath = typeof value.import === 'string' ? value.import : null;
  if (!importPath) return null;
  const distRel = importPath.replace(/^\.\/dist\//, '').replace(/\.m?js$/, '');
  const direct = path.join('src', `${distRel}.ts`);
  if (fs.existsSync(path.join(packageRoot, direct))) return direct;
  const indexed = path.join('src', distRel, 'index.ts');
  if (fs.existsSync(path.join(packageRoot, indexed))) return indexed;
  return direct; // best-effort fallback so the hint is still useful
}

function entriesFromExports(exportsMap) {
  const entries = [];
  for (const [key, value] of Object.entries(exportsMap)) {
    if (key === './package.json') continue;
    if (typeof value !== 'object' || value === null) continue;
    if (typeof value.types !== 'string') continue;
    entries.push({
      key,
      dts: value.types,
      slug: slugForKey(key),
      src: sourcePathForEntry(value),
    });
  }
  return entries;
}

const targets = entriesFromExports(pkg.exports);

fs.mkdirSync(etcDir, { recursive: true });
fs.mkdirSync(tempDir, { recursive: true });

const tsdocMessageReporting = {
  'tsdoc-undefined-tag': { logLevel: 'none' },
};
const extractorMessageReporting = {
  'ae-forgotten-export': { logLevel: 'none' },
  'ae-missing-release-tag': { logLevel: 'warning' },
};

function buildConfig({ dts, slug }) {
  const dtsPath = path.resolve(packageRoot, dts);
  const configObject = {
    mainEntryPointFilePath: dtsPath,
    apiReport: {
      enabled: true,
      reportFolder: etcDir,
      reportFileName: `${slug}.api.md`,
      reportTempFolder: tempDir,
    },
    docModel: { enabled: false },
    dtsRollup: { enabled: false },
    tsdocMetadata: { enabled: false },
    compiler: { tsconfigFilePath: path.join(packageRoot, 'tsconfig.json') },
    messages: {
      extractorMessageReporting,
      tsdocMessageReporting,
    },
    projectFolder: packageRoot,
  };
  return ExtractorConfig.prepare({
    configObject,
    configObjectFullPath: packageJsonPath,
    packageJsonFullPath: packageJsonPath,
  });
}

const present = [];
const skipped = [];
for (const target of targets) {
  const dtsPath = path.resolve(packageRoot, target.dts);
  if (fs.existsSync(dtsPath)) {
    present.push(target);
  } else {
    skipped.push(target);
  }
}

if (present.length === 0) {
  console.error('No built .d.ts files found. Run `bun run build` first.');
  process.exit(1);
}

// Build ONE CompilerState that knows about every entry point, then share it
// across all 61 Extractor.invoke() calls. ~5-10x faster than letting each
// invocation parse the same tsconfig fresh.
const firstConfig = buildConfig(present[0]);
const additionalEntryPoints = present.slice(1).map(
  (t) => buildConfig(t).mainEntryPointFilePath
);
const compilerState = CompilerState.create(firstConfig, {
  additionalEntryPoints,
});

let totalErrors = 0;
let totalWarnings = 0;
const driftedEntries = [];

for (const target of present) {
  const extractorConfig = buildConfig(target);
  const result = Extractor.invoke(extractorConfig, {
    localBuild: isLocal,
    showVerboseMessages: false,
    compilerState,
    messageCallback: (message) => {
      message.handled = true;
    },
  });

  totalErrors += result.errorCount;
  totalWarnings += result.warningCount;

  if (!isLocal && result.apiReportChanged) {
    driftedEntries.push(target);
  }
}

console.log(`API Extractor: ${present.length} entries scanned`);
console.log(`  errors: ${totalErrors}`);
console.log(`  warnings: ${totalWarnings}`);
if (skipped.length > 0) console.log(`  skipped: ${skipped.length}`);

// Fail CI when an entry listed in package.json `exports` has no built dist —
// catches a half-built tree before it silently masks API surface drift.
if (!isLocal && skipped.length > 0) {
  console.error(`\nMissing dist files for ${skipped.length} entr${skipped.length === 1 ? 'y' : 'ies'}:`);
  for (const t of skipped) console.error(`  - ${t.key} → ${t.dts}`);
  console.error(`\nFix: bun run --filter '@eigenpal/docx-editor-core' build`);
  process.exit(1);
}

if (driftedEntries.length > 0) {
  console.error(
    `\nAPI surface drift in ${driftedEntries.length} entr${driftedEntries.length === 1 ? 'y' : 'ies'}:`
  );
  for (const t of driftedEntries) {
    const where = t.src ? ` (${t.src})` : '';
    console.error(`  - ${t.slug}${where}`);
  }
  console.error(`\nFix: bun run api:extract`);
  console.error(`Then commit the updated etc/*.api.md files.`);
  process.exit(1);
}

if (totalErrors > 0) {
  process.exit(1);
}
