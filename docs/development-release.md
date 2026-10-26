# Development and release notes

## Build from source

1. Install dependencies:

   ```bash
   npm install
   ```

2. Build the plugin:

   ```bash
   npm run build
   ```

3. Copy the generated files from `dist` into your Obsidian vault plugin
   folder:

   ```text
   dist/main.js
   dist/docx-engine.cjs
   dist/manifest.json
   dist/styles.css
   ```

## Installation and upgrades

### Install from a release zip

1. Download `obsidian-word-reader-X.Y.Z.zip` from the GitHub release page.
2. Create the plugin folder in your vault:

   ```text
   YourVault/.obsidian/plugins/word-reader/
   ```

3. Extract the zip into that folder. All four files must be directly inside
   the plugin folder:

   ```text
   YourVault/.obsidian/plugins/word-reader/main.js
   YourVault/.obsidian/plugins/word-reader/docx-engine.cjs
   YourVault/.obsidian/plugins/word-reader/manifest.json
   YourVault/.obsidian/plugins/word-reader/styles.css
   ```

   `docx-engine.cjs` is the verified DOCX rendering engine. Keep it next to
   `main.js`; do not rename, move, or delete it.

4. Restart Obsidian or reload community plugins.
5. In Obsidian, open Settings, enable Community plugins, then enable Office
   Reader.
6. Open a `.docx` file to verify the engine loads. If `docx-engine.cjs` is
   missing or damaged, the plugin shows a load error instead of rendering.

### Install individual release assets

Download all four assets from the same GitHub release (`main.js`,
`docx-engine.cjs`, `manifest.json`, `styles.css`) and place them directly in
the plugin folder shown above. Do not mix files from different releases.

> If you install or update through the Obsidian community plugin browser, it
> may download only `main.js`, `manifest.json`, and `styles.css`. Download
> `docx-engine.cjs` from the same release and place it beside `main.js`, or
> use the release zip instead.

### Upgrade from an older release

1. Disable Office Reader in Obsidian.
2. Replace all four files in the plugin folder with the files from the new
   release.
3. Keep `data.json` if it exists.
4. Re-enable the plugin.

Do not replace only `main.js`; `main.js` and `docx-engine.cjs` must come from
the same release.

## Local release

Create and validate a local release package:

```bash
npm run release
```

The release command runs TypeScript checks, builds the plugin, validates
version consistency, creates the installable zip, and extracts the current
changelog section for release notes.

Expected output:

```text
release/obsidian-word-reader-X.Y.Z.zip
release/CHANGELOG-X.Y.Z.md
```

The zip root contains only the files Obsidian needs:

```text
main.js
docx-engine.cjs
manifest.json
styles.css
```

Release artifacts are ignored by Git and should not be committed.

## GitHub auto release

GitHub Actions creates a release automatically when a version tag without a
`v` prefix is pushed:

```bash
git tag 3.6.0
git push origin 2.5.0
```

The workflow validates that the tag matches `package.json`,
`manifest.json`, and `package-lock.json`, checks the 500 KiB bundle budget,
builds the plugin, creates
`release/obsidian-word-reader-3.6.0.zip`, extracts the matching
`CHANGELOG.md` section, and uploads `main.js`, `docx-engine.cjs`,
`manifest.json`, `styles.css`, and the zip to the GitHub Release.

## Development commands

Install dependencies:

```bash
npm install
```

Run a production build:

```bash
npm run build
```

Run TypeScript checks:

```bash
npm run typecheck
```

Run unit and integration tests:

```bash
npm test
```

Run the production security scan after building:

```bash
npm run build
npm run security:scan
```

Run every automated quality gate:

```bash
npm run check
```

Validate release metadata and zip contents after packaging:

```bash
npm run release:zip
npm run release:check
```

The initial community-directory submission and release review steps are in
[COMMUNITY_PLUGIN_CHECKLIST.md](COMMUNITY_PLUGIN_CHECKLIST.md).

CI and release workflows use Node.js 20.19.0.
