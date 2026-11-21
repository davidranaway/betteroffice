#!/usr/bin/env bun

import { createHash } from 'node:crypto';
import {
  cpSync,
  existsSync,
  lstatSync,
  mkdirSync,
  readFileSync,
  readdirSync,
  rmSync,
  writeFileSync,
} from 'node:fs';
import { dirname, join, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { spawnSync } from 'node:child_process';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const channelDirectory = join(root, 'xelper-prereleases', 'pptx-react');
const channelPath = join(channelDirectory, 'channel.json');
const workDirectory = join(channelDirectory, '.work');
const sourceDirectory = join(workDirectory, 'source');
const stageDirectory = join(workDirectory, 'packages');
const artifactsDirectory = join(channelDirectory, 'artifacts');
const packageManagerVersion = '1.4.3';
process.env.PATH = `${dirname(process.execPath)}:${process.env.PATH ?? ''}`;

function fail(message) {
  throw new Error(message);
}

function readJson(path) {
  return JSON.parse(readFileSync(path, 'utf8'));
}

function sha256(path) {
  return createHash('sha256').update(readFileSync(path)).digest('hex');
}

function sha256Text(value) {
  return createHash('sha256').update(value).digest('hex');
}

function run(command, args, options = {}) {
  const result = spawnSync(command, args, {
    cwd: options.cwd ?? root,
    encoding: 'utf8',
    env: options.env ?? process.env,
    stdio: options.capture ? 'pipe' : 'inherit',
  });
  if (result.error) throw result.error;
  if (result.status !== 0 && !options.allowFailure) {
    const detail = options.capture ? `\n${result.stdout}${result.stderr}` : '';
    fail(`${command} ${args.join(' ')} failed with exit code ${result.status}.${detail}`);
  }
  return result;
}

function git(args, options = {}) {
  return run('git', args, { ...options, cwd: root });
}

function gitOutput(args) {
  return git(args, { capture: true }).stdout.trim();
}

function listFiles(directory, prefix = '') {
  const files = [];
  for (const entry of readdirSync(directory, { withFileTypes: true })) {
    const path = join(directory, entry.name);
    const name = prefix ? `${prefix}/${entry.name}` : entry.name;
    if (entry.isSymbolicLink()) fail(`Unexpected symlink in staged package: ${name}`);
    if (entry.isDirectory()) files.push(...listFiles(path, name));
    else if (entry.isFile()) files.push(name);
    else fail(`Unexpected staged filesystem entry: ${name}`);
  }
  return files.sort();
}

export function loadChannel() {
  const channel = readJson(channelPath);
  const core = channel.packages?.core;
  const react = channel.packages?.react;
  const compatibility = channel.compatibilityAdaptations?.[0];
  if (channel.schemaVersion !== 2 ||
      channel.releaseTag !== '@davidranaway/betteroffice-pptx-react@0.5.1-xelper.5' ||
      channel.repository !== 'https://github.com/davidranaway/betteroffice.git' ||
      channel.distTag !== 'xelper' ||
      channel.source?.tag !== '@betteroffice/pptx-react@0.5.1' ||
      channel.source?.commit !== '4e982c3aea0065578604c3fe7727938533888f66' ||
      core?.name !== '@davidranaway/betteroffice-pptx' ||
      core?.version !== '0.5.1-xelper.1' ||
      react?.name !== '@davidranaway/betteroffice-pptx-react' ||
      react?.version !== '0.5.1-xelper.5' ||
      channel.backports?.map(({ upstreamPullRequest }) => upstreamPullRequest).join(',') !==
        '2698,2731,2732,2733' ||
      channel.compatibilityAdaptations?.length !== 1 ||
      compatibility.sourceHunk?.commit !== 'fe7a97514f14c1d6bfa0be1a69b1364403e173b0' ||
      compatibility.sourceHunk?.path !== 'packages/pptx/src/session/commandHost.ts' ||
      JSON.stringify(compatibility.sourceHunk?.lines) !== JSON.stringify([1572, 1595, 1596]) ||
      compatibility.sourceHunk?.sha256 !==
        '0108c75a9764d9dcf2e48ffec768037c954682a400b42a848934d6b1200b116a' ||
      compatibility.testAdaptation?.path !==
        'packages/pptx-react/src/worker/peerlessS6.test.tsx') {
    fail('channel.json does not describe the reviewed dual-package 0.5.1 release set.');
  }
  if (JSON.stringify(core.dependencies) !== JSON.stringify({ '@betteroffice/fonts': '0.5.1' }) ||
      JSON.stringify(react.dependencies) !== JSON.stringify({
        '@betteroffice/fonts': '0.5.1',
        '@betteroffice/pptx': 'npm:@davidranaway/betteroffice-pptx@0.5.1-xelper.1',
        '@betteroffice/pptx-i18n': '0.5.1',
      })) {
    fail('channel.json does not preserve the exact 0.5.1 dependency contract and core alias.');
  }
  const securityPatch = channel.buildEnvironment?.patches?.[0];
  if (channel.buildEnvironment?.patches?.length !== 1 ||
      securityPatch.advisory !== 'GHSA-3f7j-7jwc-fxfj' ||
      securityPatch.package !== 'systeminformation' ||
      securityPatch.version !== '5.33.16' ||
      securityPatch.license !== 'MIT') {
    fail('channel.json does not retain the reviewed systeminformation security patch.');
  }
  return channel;
}

function provenanceFor(channel, packageKey) {
  const packageDefinition = channel.packages[packageKey];
  return {
    schemaVersion: 2,
    releaseTag: channel.releaseTag,
    package: {
      name: packageDefinition.name,
      version: packageDefinition.version,
      repository: channel.repository,
      distTag: channel.distTag,
    },
    source: channel.source,
    releaseSet: Object.fromEntries(Object.entries(channel.packages)
      .map(([key, value]) => [key, { name: value.name, version: value.version }])),
    backports: channel.backports.map(({
      upstreamPullRequest, headRef, headCommit, mergeBase, behavior, reviewedCommits,
      adaptation, excludedReviewedPaths, patch, sha256: digest,
    }) => ({
      upstreamPullRequest,
      headRef,
      headCommit,
      mergeBase,
      behavior,
      reviewedCommits,
      ...(adaptation ? { adaptation } : {}),
      excludedReviewedPaths,
      patch,
      sha256: digest,
    })),
    compatibilityAdaptations: channel.compatibilityAdaptations.map(({
      behavior, sourceHunk, testAdaptation, patch, sha256: digest,
    }) => ({
      behavior,
      sourceHunk,
      testAdaptation,
      patch,
      sha256: digest,
    })),
    buildEnvironment: {
      securityPatches: channel.buildEnvironment.patches.map(
        ({ advisory, package: dependency, version, license, patch, sha256: digest }) => ({
          advisory,
          package: dependency,
          version,
          license,
          patch,
          sha256: digest,
        }),
      ),
    },
  };
}

export function stagePackage({ channel, packageKey, sourcePackage, stage }) {
  const definition = channel.packages[packageKey];
  sourcePackage ??= join(sourceDirectory, definition.sourceDirectory);
  stage ??= join(stageDirectory, packageKey);
  const sourceManifestPath = join(sourcePackage, 'package.json');
  const sourceManifestBytes = readFileSync(sourceManifestPath, 'utf8');
  const sourceManifest = JSON.parse(sourceManifestBytes);
  if (sourceManifest.name !== definition.sourceName || sourceManifest.version !== '0.5.1') {
    fail(`The source package is not official ${definition.sourceName}@0.5.1.`);
  }
  const expectedSourceDependencies = packageKey === 'core'
    ? { '@betteroffice/fonts': 'workspace:^' }
    : {
        '@betteroffice/fonts': 'workspace:^',
        '@betteroffice/pptx': 'workspace:*',
        '@betteroffice/pptx-i18n': 'workspace:*',
      };
  if (JSON.stringify(sourceManifest.dependencies) !== JSON.stringify(expectedSourceDependencies)) {
    fail('The official source manifest has unexpected runtime dependencies.');
  }
  if (sourceManifest.license !== 'Apache-2.0' ||
      sourceManifest.repository?.url !== 'https://github.com/openooxml/betteroffice.git' ||
      sourceManifest.repository?.directory !== definition.sourceDirectory) {
    fail('The official package license or repository provenance is unexpected.');
  }
  if (!existsSync(join(sourcePackage, 'dist', 'index.js')) ||
      !existsSync(join(sourcePackage, 'dist', 'index.d.ts'))) {
    fail('pptx-react build output is incomplete.');
  }

  rmSync(stage, { recursive: true, force: true });
  mkdirSync(stage, { recursive: true });
  cpSync(join(sourcePackage, 'dist'), join(stage, 'dist'), { recursive: true, dereference: false });
  cpSync(join(sourcePackage, 'LICENSE'), join(stage, 'LICENSE'), { dereference: false });

  const stagedManifest = structuredClone(sourceManifest);
  stagedManifest.name = definition.name;
  stagedManifest.version = definition.version;
  stagedManifest.repository = {
    ...stagedManifest.repository,
    url: channel.repository,
  };
  stagedManifest.dependencies = structuredClone(definition.dependencies);
  stagedManifest.files = ['dist', 'LICENSE', 'XELPER-PROVENANCE.json'];
  writeFileSync(join(stage, 'package.json'), `${JSON.stringify(stagedManifest, null, 2)}\n`);
  writeFileSync(
    join(stage, 'XELPER-PROVENANCE.json'),
    `${JSON.stringify(provenanceFor(channel, packageKey), null, 2)}\n`,
  );

  if (readFileSync(sourceManifestPath, 'utf8') !== sourceManifestBytes) {
    fail('Staging mutated the source package manifest.');
  }
  const stagedFiles = listFiles(stage);
  const required = ['LICENSE', 'XELPER-PROVENANCE.json', 'dist/index.d.ts', 'dist/index.js', 'package.json'];
  for (const path of required) {
    if (!stagedFiles.includes(path)) fail(`Staged package is missing ${path}.`);
  }
  for (const path of stagedFiles) {
    if (required.includes(path)) continue;
    if (!path.startsWith('dist/')) {
      fail(`Unexpected staged package file: ${path}`);
    }
  }
  return { manifest: stagedManifest, files: stagedFiles, provenance: provenanceFor(channel, packageKey) };
}

function validateInputs(channel) {
  if (process.versions.bun !== packageManagerVersion) {
    fail(`Use locked Bun ${packageManagerVersion}; found ${process.versions.bun ?? 'no Bun runtime'}.`);
  }
  const sourceCommit = gitOutput(['rev-list', '-n', '1', `${channel.source.tag}^{commit}`]);
  if (sourceCommit !== channel.source.commit) {
    fail(`${channel.source.tag} resolves to ${sourceCommit}, not reviewed source ${channel.source.commit}.`);
  }
  for (const backport of channel.backports) {
    const headCommit = gitOutput(['rev-parse', `${backport.headRef}^{commit}`]);
    if (headCommit !== backport.headCommit) {
      fail(`${backport.headRef} resolves to ${headCommit}, not reviewed head ${backport.headCommit}.`);
    }
    git(['cat-file', '-e', `${backport.mergeBase}^{commit}`]);
    if (git(['merge-base', '--is-ancestor', backport.mergeBase, backport.headCommit],
      { allowFailure: true }).status !== 0) {
      fail(`Recorded merge base is not an ancestor of PR #${backport.upstreamPullRequest}.`);
    }
    for (const commit of backport.reviewedCommits) git(['cat-file', '-e', `${commit}^{commit}`]);
  }
  for (const adaptation of channel.compatibilityAdaptations) {
    const { commit, path, lines, sha256: expected } = adaptation.sourceHunk;
    git(['cat-file', '-e', `${commit}^{commit}`]);
    const source = gitOutput(['show', `${commit}:${path}`]).split('\n');
    const hunk = `${lines.map((line) => source[line - 1]?.trim() ?? '').join('\n')}\n`;
    if (sha256Text(hunk) !== expected) {
      fail(`Compatibility source hunk checksum mismatch: ${commit}:${path}`);
    }
  }

  for (const patch of [
    ...channel.buildEnvironment.patches,
    ...channel.backports,
    ...channel.compatibilityAdaptations,
  ]) {
    const patchPath = join(channelDirectory, patch.patch);
    if (sha256(patchPath) !== patch.sha256) fail(`Patch checksum mismatch: ${patch.patch}`);
    const patchPaths = gitOutput(['apply', '--unidiff-zero', '--numstat', patchPath])
      .split('\n')
      .filter(Boolean)
      .map((line) => line.split('\t').at(-1))
      .sort();
    if (JSON.stringify(patchPaths) !== JSON.stringify([...patch.paths].sort())) {
      fail(`The recorded patch touches an unreviewed or unexpected path: ${patch.patch}`);
    }
    if ('upstreamPullRequest' in patch && patchPaths.some((path) => path.startsWith('.changeset/') ||
      !['bindings/', 'crates/pptx-', 'packages/pptx/', 'packages/pptx-react/', 'shared/office-session/']
        .some((prefix) => path.startsWith(prefix)))) {
      fail(`Patch escapes the approved PPTX source, test, shared, or core path set: ${patch.patch}`);
    }
  }
}

function prepareSource(channel) {
  if (existsSync(workDirectory) && lstatSync(workDirectory).isSymbolicLink()) {
    fail('Refusing a symlinked prerelease work directory.');
  }
  rmSync(workDirectory, { recursive: true, force: true });
  mkdirSync(sourceDirectory, { recursive: true });
  const archive = join(workDirectory, 'source.tar');
  git(['archive', '--format=tar', `--output=${archive}`, channel.source.tag]);
  run('tar', ['-xf', archive, '-C', sourceDirectory]);
  rmSync(archive);

  const sourcePrefix = relative(root, sourceDirectory);
  for (const patch of channel.buildEnvironment.patches) {
    for (const [path, expected] of Object.entries(patch.prePatchSha256)) {
      const actual = sha256(join(sourceDirectory, path));
      if (actual !== expected) fail(`Pre-patch checksum mismatch for ${path}.`);
    }
    const patchPath = join(channelDirectory, patch.patch);
    git(['apply', '--unidiff-zero', '--check', `--directory=${sourcePrefix}`, patchPath]);
    git(['apply', '--unidiff-zero', `--directory=${sourcePrefix}`, patchPath]);
    for (const [path, expected] of Object.entries(patch.postPatchSha256)) {
      const actual = sha256(join(sourceDirectory, path));
      if (actual !== expected) fail(`Build-environment post-patch checksum mismatch for ${path}.`);
    }
  }

  for (const backport of channel.backports) {
    const patchPath = join(channelDirectory, backport.patch);
    git(['apply', '--unidiff-zero', '--check', `--directory=${sourcePrefix}`, patchPath]);
    git(['apply', '--unidiff-zero', `--directory=${sourcePrefix}`, patchPath]);
    for (const [path, expected] of Object.entries(backport.postPatchSha256)) {
      const actual = sha256(join(sourceDirectory, path));
      if (actual !== expected) fail(`Post-patch checksum mismatch for ${path}.`);
    }
  }
  for (const adaptation of channel.compatibilityAdaptations) {
    const patchPath = join(channelDirectory, adaptation.patch);
    git(['apply', '--unidiff-zero', '--check', `--directory=${sourcePrefix}`, patchPath]);
    git(['apply', '--unidiff-zero', `--directory=${sourcePrefix}`, patchPath]);
    for (const [path, expected] of Object.entries(adaptation.postPatchSha256)) {
      const actual = sha256(join(sourceDirectory, path));
      if (actual !== expected) fail(`Compatibility post-patch checksum mismatch for ${path}.`);
    }
  }
}

function inspectPack(channel, packageKey) {
  const definition = channel.packages[packageKey];
  const stage = join(stageDirectory, packageKey);
  const dryRun = run('npm', ['pack', '--dry-run', '--json'], {
    cwd: stage,
    capture: true,
  });
  const report = JSON.parse(dryRun.stdout);
  if (!Array.isArray(report) || report.length !== 1) fail('npm pack returned an unexpected report.');
  const packed = report[0];
  if (packed.name !== definition.name || packed.version !== definition.version) {
    fail('npm pack resolved an unexpected package name or version.');
  }
  const files = packed.files.map(({ path }) => path).sort();
  const stagedFiles = listFiles(stage);
  if (JSON.stringify(files) !== JSON.stringify(stagedFiles)) {
    fail(`npm pack file set differs from staged files:\n${files.join('\n')}`);
  }
  return packed;
}

function buildPackage(channel) {
  const bun = process.execPath;
  console.log('[1/10] Installing the official tag lockfile');
  run(bun, ['install', '--frozen-lockfile'], { cwd: sourceDirectory });
  const securedDependency = readJson(join(sourceDirectory, 'node_modules', 'systeminformation', 'package.json'));
  if (securedDependency.version !== '5.33.16' || securedDependency.license !== 'MIT') {
    fail('The generated build workspace did not resolve systeminformation@5.33.16 (MIT).');
  }
  console.log('[2/10] Building the tagged PPTX Wasm prerequisite');
  run(bun, ['run', 'build:pptx-wasm'], { cwd: sourceDirectory });

  console.log('[3/10] Running focused core and React regressions');
  const test = (paths, pattern) => run(bun, [
    'test', ...paths, ...(pattern ? ['--test-name-pattern', pattern] : []),
  ], {
    cwd: sourceDirectory,
    env: { ...process.env, BUN_CONFIG_PRELOAD: '' },
  });
  test(['packages/pptx-react/src/worker/PeerlessPptxSurface.test.tsx'],
    'strip_arrow_keys_focus_activate_and_preserve_reorder');
  test(['packages/pptx-react/src/worker/peerlessInsertionOutline.test.tsx'],
    'resize released|pending released resize|new shape gesture|grouped child text');
  test(['shared/office-session/commands/session.test.ts'],
    'cancels a command waiting for its predecessor');
  test(['packages/pptx-react/src/usePptxPinchZoom.test.tsx']);
  test(['packages/pptx-react/src/worker/PeerlessShapeResizePreview.test.tsx']);
  test(['packages/pptx/src/session/commandS6.test.ts'],
    'relative resize|shape rectangle previews');
  test(['packages/pptx-react/src/worker/peerlessS6.test.tsx'],
    'text editing keeps resize handles|text_resize_preview|released text-edit resize');
  run('cargo', [
    'test', '-p', 'pptx-wasm',
    'shape_rect_preview_reflows_text_without_changing_session_state',
  ], { cwd: sourceDirectory });

  console.log('[4/10] Typechecking core and React');
  run(bun, ['run', '--filter', '@betteroffice/pptx', 'typecheck'], { cwd: sourceDirectory });
  run(bun, ['run', '--filter', '@betteroffice/pptx-react', 'typecheck'], { cwd: sourceDirectory });

  console.log('[5/10] Building core and React');
  run(bun, ['run', '--filter', '@betteroffice/pptx', 'build'], { cwd: sourceDirectory });
  run(bun, ['run', '--filter', '@betteroffice/pptx-react', 'build'], { cwd: sourceDirectory });

  console.log('[6/10] Auditing the complete resolved dependency set');
  const audit = run(bun, ['audit'], { cwd: sourceDirectory, allowFailure: true, capture: true });
  process.stdout.write(audit.stdout);
  process.stderr.write(audit.stderr);
  if (/severity\\s*[:=]\\s*[\"']?(unknown|unclassified)/i.test(`${audit.stdout}\n${audit.stderr}`)) {
    fail('The complete dependency audit contains an unclassified finding.');
  }
  run(bun, ['audit', '--audit-level=high'], { cwd: sourceDirectory });

  console.log('[7/10] Staging both immutable packages');
  rmSync(stageDirectory, { recursive: true, force: true });
  for (const packageKey of ['core', 'react']) {
    stagePackage({ channel, packageKey });
    inspectPack(channel, packageKey);
  }

  console.log('[8/10] Creating both npm tarballs');
  rmSync(artifactsDirectory, { recursive: true, force: true });
  mkdirSync(artifactsDirectory, { recursive: true });
  const tarballs = {};
  for (const packageKey of ['core', 'react']) {
    const pack = run('npm', ['pack', '--json', '--pack-destination', artifactsDirectory], {
      cwd: join(stageDirectory, packageKey),
      capture: true,
    });
    const report = JSON.parse(pack.stdout);
    if (!Array.isArray(report) || report.length !== 1) {
      fail(`npm pack did not create exactly one ${packageKey} tarball.`);
    }
    const tarball = join(artifactsDirectory, report[0].filename);
    tarballs[packageKey] = {
      file: relative(channelDirectory, tarball),
      sha256: sha256(tarball),
      size: report[0].size,
      unpackedSize: report[0].unpackedSize,
      integrity: report[0].integrity,
    };
  }

  console.log('[9/10] Recording release-set provenance');
  writeFileSync(
    join(artifactsDirectory, 'provenance.json'),
    `${JSON.stringify({
      schemaVersion: 2,
      releaseTag: channel.releaseTag,
      source: channel.source,
      tarballs,
    }, null, 2)}\n`,
  );
  console.log('[10/10] Created the reviewed dual-package release set');
  return Object.values(tarballs).map(({ file }) => join(channelDirectory, file));
}

function validateTag(channel, tag) {
  const expected = channel.releaseTag;
  if (tag !== expected) fail(`Release tag must be exactly ${expected}; received ${tag}.`);
  console.log(`Validated release tag ${tag}.`);
}

function main() {
  const channel = loadChannel();
  if (process.argv[2] === '--validate-tag' && process.argv.length === 4) {
    validateTag(channel, process.argv[3]);
    return;
  }
  if (process.argv.length !== 2) fail('Unexpected arguments. Run without arguments or use --validate-tag <tag>.');

  const statusBefore = gitOutput(['status', '--porcelain=v1', '--untracked-files=all']);
  console.log('[0/10] Validating pinned source and backport inputs');
  validateInputs(channel);
  rmSync(artifactsDirectory, { recursive: true, force: true });
  prepareSource(channel);
  const tarballs = buildPackage(channel);
  const statusAfter = gitOutput(['status', '--porcelain=v1', '--untracked-files=all']);
  if (statusAfter !== statusBefore) {
    fail('Packaging mutated the source checkout; inspect git status before retrying.');
  }
  console.log(`\nReady:\n${tarballs.map((tarball) => relative(root, tarball)).join('\n')}`);
}

if (resolve(process.argv[1] ?? '') === fileURLToPath(import.meta.url)) {
  try {
    main();
  } catch (error) {
    console.error(error instanceof Error ? error.message : error);
    process.exitCode = 1;
  }
}
