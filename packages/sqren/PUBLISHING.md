# Publishing `@sqren/docx-editor`

`@sqren/docx-editor` is a **single bundled distribution** of this docx-editor
fork (the read-only DOCX renderer). `packages/sqren` is almost empty — its
`src/index.ts` just re-exports `@eigenpal/docx-editor-react`, and tsup **inlines**
`@eigenpal/docx-editor-{core,react,i18n,agents}` plus their third-party deps
(docxtemplater, jszip, dompurify, clsx, sonner, @radix-ui/react-select, …) into
one self-contained package. Only `react`, `react-dom`, and `prosemirror-*` stay
as **peer** dependencies.

Consumers therefore install **one** dependency. This is deliberately a bundle
and not four `@sqren/*` packages — see [Gotchas](#gotchas).

## Prerequisites

- Logged in to npm as the owner of the `@sqren` scope — check with `npm whoami`
  (should print `sqren`).
- [`bun`](https://bun.sh) installed (this repo's toolchain), and `bun install`
  has been run at the repo root.
- npm 2FA / OTP ready if your account requires it for publish.

## Release steps

### 1. Make your changes

Edit the underlying source under `packages/{core,react,i18n,agents}/src`. You
rarely touch `packages/sqren` itself (it's just the re-export entry + build
config).

### 2. Rebuild the underlying `@eigenpal/*` packages from source ⚠️

The bundle re-exports the **built** `@eigenpal/docx-editor-react` `dist/`, and
the committed `dist/` can lag behind `src/`. **Always rebuild before
publishing** or you will ship stale code:

```sh
# from the repo root
( cd packages/i18n   && bunx tsup )
( cd packages/core   && NODE_OPTIONS=--max-old-space-size=8192 bunx tsup )
( cd packages/agents && bunx tsup )
( cd packages/react  && bunx tsup \
    && bunx tailwindcss -c ./tailwind.config.js \
         -i ./src/styles/editor.css -o ./dist/styles.css --minify )
```

The root `bun run build:packages` also works, but it additionally builds
vue/nuxt and runs api-extractor / doc-injection steps that can fail on
unrelated packages. The four `tsup` invocations above are the reliable subset
the bundle actually needs. `packages/react` also needs its Tailwind-built
`dist/styles.css` (the bundle copies it — see step 4).

### 3. Bump the version

```sh
cd packages/sqren
npm version patch        # or minor / major
```

`@sqren/docx-editor` is versioned **independently** of the `@eigenpal/*` line
(it started at `1.0.0`); use plain semver. npm refuses to republish an existing
version.

### 4. Build + sanity-check the bundle

```sh
bun run build            # tsup + copies ../react/dist/styles.css → dist/styles.css
                         # (prepublishOnly runs this automatically on publish too)

npm pack --dry-run       # tarball should contain README.md + dist/ only (files: ["dist"])

# confirm it is self-contained — this should print "self-contained ✓":
grep -rqE "(from|require\()['\"]@(eigenpal|sqren)/" dist/*.js dist/*.cjs \
  && echo "NOT self-contained — scoped runtime import leaked!" \
  || echo "self-contained ✓"
```

Only `react` / `react-dom` / `prosemirror-*` should remain external (the peer
deps). Everything else must be inlined.

### 5. Publish

```sh
npm publish              # public — publishConfig.access is "public"
                         # prepublishOnly rebuilds first
```

Verify it landed:

```sh
npm view @sqren/docx-editor version
```

### 6. Update consumers

For example the CLM sandbox (`React-OOXML-Office-sandbox`):

```sh
# bump "@sqren/docx-editor" in its package.json to the new version, then:
npm install
```

For local development against an **unpublished** build, the sandbox provides
`npm run docxeditor:local` (installs `file:../docx-editor/packages/sqren`) and
`npm run docxeditor:npm` (switches back to the registry release).

## Gotchas

- **Stale dist** — always do step 2. The committed `@eigenpal/*` `dist/` may
  predate your `src/` changes (e.g. the custom `acceptChange` / `rejectChange`),
  and the bundle inlines the dist, so skipping the rebuild silently ships old
  code.
- **Do NOT publish the four packages renamed to `@sqren/*`.** Renaming only the
  package *names* (and not their internal cross-package `dependencies`) produces
  **un-installable** packages whose deps resolve to a nonexistent
  `@eigenpal/docx-editor-core@<fork-version>`. The bundle sidesteps this by
  inlining everything. If you see `@sqren/docx-editor-{react,core,i18n,agents}`
  prerelease versions on the registry, those are the broken artifacts — run
  `npm deprecate` on them.
- **CSS ships pre-expanded.** `dist/styles.css` is the Tailwind-built
  `packages/react/dist/styles.css`, copied in by the build. Consumers
  `import "@sqren/docx-editor/styles.css"`; they must **not** rely on their own
  Tailwind to scan `node_modules` (it won't generate the editor's classes).
- **Peer deps are a shared singleton.** `react` / `react-dom` / `prosemirror-*`
  must resolve to a single instance shared with any companion engine (e.g. the
  headless `@beyondwork/docx-react-component`). Consumers enforce this with Vite
  `dedupe` + single-copy resolution.
- **Consumer `file:` lockfile pin.** After a consumer runs `docxeditor:local`
  (a `file:` install), its `package-lock.json` keeps the `file:` resolution; a
  plain `npm install` won't switch back to the registry even with a `^x.y.z`
  range, because the linked local build satisfies it. Force it:
  `rm -rf node_modules/@sqren`, strip the `@sqren/docx-editor` entries from
  `package-lock.json`, then `npm install`.
```
