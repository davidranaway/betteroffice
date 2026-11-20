# xelper pptx-react prerelease

This channel builds `@davidranaway/betteroffice-pptx-react@0.5.1-xelper.1`
from the official `@betteroffice/pptx-react@0.5.1` tag plus the single ordered
pptx-react source backport recorded in `channel.json`. That backport carries
the reviewed PR #2698 behavior in full: plain Up/Down thumbnail navigation and
activation, repeated and boundary-key handling, modifier isolation, preserved
Alt+Up/Down reordering, and pointer focus retention. A separate ordered
build-environment patch pins `systeminformation@5.33.16` in only the generated
root manifest and lockfile to remediate `GHSA-3f7j-7jwc-fxfj`; it is applied
before the frozen install and never changes the staged package manifest or
published dependencies. The channel never copies package source from the
current branch.

## Build and inspect

From the repository root:

```bash
pwd
test -f package.json
bun install --frozen-lockfile
BUN_CONFIG_PRELOAD='' bun test scripts/package-xelper-pptx-react.test.ts
bun scripts/package-xelper-pptx-react.mjs
```

The packaging command validates the source tag, reviewed commit, both ordered
patches, their exact path sets, and pre/post-patch checksums; creates a clean
temporary archive checkout; applies the build-only security patch before a
frozen install; runs the backported focus test, pptx-react typecheck and build,
the complete unfiltered dependency audit and unsuppressed high-severity gate,
and npm pack inspection; then writes the ignored tarball and its provenance
record under `artifacts/`. It refuses unexpected inputs and verifies that the
invoking checkout was not mutated.

## First publish, then Trusted Publishing

The npm package name does not exist yet. The workflow intentionally refuses to
create it. After reviewing the generated tarball, the owner performs the first
publish interactively:

```bash
npm login
npm publish xelper-prereleases/pptx-react/artifacts/davidranaway-betteroffice-pptx-react-0.5.1-xelper.1.tgz \
  --access public --tag xelper
```

Then configure npm Trusted Publishing for the fork repository and
`.github/workflows/publish-xelper-pptx-react.yml`. Subsequent reviewed versions
are published with provenance and dist-tag `xelper` by pushing the exact tag
`@davidranaway/betteroffice-pptx-react@<version>`. Never reuse or move a
published tag.
