# xelper PPTX prerelease channel

This channel creates one immutable release set from official
`@betteroffice/pptx-react@0.5.1` source commit
`4e982c3aea0065578604c3fe7727938533888f66`:

- `@davidranaway/betteroffice-pptx@0.5.1-xelper.1`
- `@davidranaway/betteroffice-pptx-react@0.5.1-xelper.5`

The ordered patches retain reviewed PR #2698 slide-strip navigation and add the
reviewed behavior from PRs #2731, #2732, and #2733: resize handles while editing
text with ordered/cancellable commits, cursor-anchored pinch zoom, and
authoritative non-mutating live resize previews for text-bearing shapes.
`channel.json` records every reviewed commit and PR merge base. Changesets,
release metadata, and unrelated cross-product paths are excluded.

One separately recorded compatibility patch adapts those reviewed changes to
official 0.5.1 without claiming the whole later commit. It takes only the
Promise-aware `handlers.command` result bridge and return annotation from
upstream commit `fe7a97514f14c1d6bfa0be1a69b1364403e173b0`, and makes the
preview regression accept the one-or-more valid paint commits produced by the
older painter while still requiring command coverage before preview removal.
`channel.json` records the exact source path, source lines, hunk checksum,
patch checksum, path set, and post-patch hashes. No later raster architecture
or font work is included.

The official 0.5.1 source predates the later shared `pptx-engine`/native facade.
The #2733 renderer contract and its parity/non-mutation assertion are therefore
backported at the authoritative 0.5.1 `pptx-wasm` renderer boundary. The React
package keeps the import identity `@betteroffice/pptx` through this exact npm
alias:

```json
"@betteroffice/pptx": "npm:@davidranaway/betteroffice-pptx@0.5.1-xelper.1"
```

Official `@betteroffice/fonts@0.5.1` and
`@betteroffice/pptx-i18n@0.5.1` remain exact dependencies. No core code is
bundled into the React package.

A separate build-environment patch pins `systeminformation@5.33.16` (MIT) to
remediate `GHSA-3f7j-7jwc-fxfj`. It changes only the generated build workspace
root manifest and lockfile, before the frozen install; it never changes either
staged package manifest.

## Build and inspect

From the repository root:

```bash
pwd
test -f package.json
bun install --frozen-lockfile
BUN_CONFIG_PRELOAD='' bun test scripts/package-xelper-pptx-react.test.ts
bun scripts/package-xelper-pptx-react.mjs
```

The packaging command validates the immutable source, reviewed heads, commits
and merge bases, ordered patch SHA-256 values, exact path sets, and post-patch
hashes. It applies everything to a clean official archive, builds the Wasm
prerequisite, runs the smallest focused core/React regressions for all four PRs,
runs the core and React typechecks/builds, prints the complete unfiltered audit,
blocks high/critical and unclassified findings, inspects both npm pack file
sets, and writes two ignored tarballs plus release-set provenance under
`artifacts/`. The invoking checkout must remain unchanged.

## Core package bootstrap and Trusted Publishing

The new public core scope must exist before npm can bind its Trusted Publisher.
Do not bootstrap it with either reviewed release-set version.

1. From the exact official 0.5.1 source and fork repository, create a minimal,
   immutable bootstrap package version such as
   `@davidranaway/betteroffice-pptx@0.5.1-xelper.bootstrap.0`. Keep its
   `repository.url` set to `https://github.com/davidranaway/betteroffice.git`,
   retain Apache-2.0 licensing, publish with `--access public --tag xelper`, and
   do not move `latest`.
2. In npm package settings, configure the Trusted Publisher for
   `davidranaway/betteroffice`, workflow
   `publish-xelper-pptx-react.yml`, environment unset unless the workflow later
   declares the exact same environment.
3. Complete the first OIDC publication within two days of that configuration.
   The bootstrap publish does not prove or bind OIDC.
4. Before creating the release tag, verify that both reviewed targets are still
   absent and that no `NPM_TOKEN` secret is configured for the workflow.

The release workflow is triggered only by
`@davidranaway/betteroffice-pptx-react@0.5.1-xelper.5`. It validates and builds
both artifacts, checks both registry targets before any upload, publishes core
first and React second with provenance and dist-tag `xelper`, refuses npm
tokens, and never moves `latest`. Never reuse or move a pushed release tag.
