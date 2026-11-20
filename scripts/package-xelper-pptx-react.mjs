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
const stageDirectory = join(workDirectory, 'package');
const artifactsDirectory = join(channelDirectory, 'artifacts');
const sourcePackageDirectory = join(sourceDirectory, 'packages', 'pptx-react');
const packageManagerVersion = '1.4.3';

function fail(message) {
  throw new Error(message);
}

function readJson(path) {
  return JSON.parse(readFileSync(path, 'utf8'));
}

function sha256(path) {
  return createHash('sha256').update(readFileSync(path)).digest('hex');
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
  const expected = {
    schemaVersion: 1,
    name: '@davidranaway/betteroffice-pptx-react',
    version: '0.5.1-xelper.4',
    repository: 'https://github.com/davidranaway/betteroffice.git',
    distTag: 'xelper',
    tag: '@betteroffice/pptx-react@0.5.1',
    commit: '4e982c3aea0065578604c3fe7727938533888f66',
    reviewedCommits: [
      '541770e783cd972527b0f7e4592965dabcb9229e',
      'aeb5b2c6591efa80906d84d55ea8e14aaf344a4f',
    ],
    buildAdvisory: 'GHSA-3f7j-7jwc-fxfj',
    buildPackage: 'systeminformation',
    buildVersion: '5.33.16',
    buildLicense: 'MIT',
  };
  if (channel.schemaVersion !== expected.schemaVersion ||
      channel.package?.name !== expected.name ||
      channel.package?.version !== expected.version ||
      channel.package?.repository !== expected.repository ||
      channel.package?.distTag !== expected.distTag ||
      channel.source?.tag !== expected.tag ||
      channel.source?.commit !== expected.commit ||
      channel.backports?.length !== 1 ||
      channel.backports[0]?.behavior !== 'Full slide-strip keyboard navigation and focus retention' ||
      JSON.stringify(channel.backports[0]?.reviewedCommits) !== JSON.stringify(expected.reviewedCommits) ||
      channel.buildEnvironment?.patches?.length !== 1 ||
      channel.buildEnvironment.patches[0]?.advisory !== expected.buildAdvisory ||
      channel.buildEnvironment.patches[0]?.package !== expected.buildPackage ||
      channel.buildEnvironment.patches[0]?.version !== expected.buildVersion ||
      channel.buildEnvironment.patches[0]?.license !== expected.buildLicense) {
    fail('channel.json does not describe the reviewed, pptx-react-only 0.5.1 channel.');
  }
  if (JSON.stringify(channel.dependencies) !== JSON.stringify({
    '@betteroffice/pptx': '0.5.1',
    '@betteroffice/pptx-i18n': '0.5.1',
  })) {
    fail('channel.json must pin only the two official pptx-react 0.5.1 runtime dependencies.');
  }
  return channel;
}

function provenanceFor(channel) {
  return {
    schemaVersion: 1,
    package: channel.package,
    source: {
      repository: 'https://github.com/openooxml/betteroffice.git',
      tag: channel.source.tag,
      commit: channel.source.commit,
    },
    backports: channel.backports.map(({ upstreamPullRequest, behavior, reviewedCommits, patch, sha256: digest }) => ({
      upstreamPullRequest,
      behavior,
      reviewedCommits,
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

export function stagePackage({ channel, sourcePackage = sourcePackageDirectory, stage = stageDirectory }) {
  const sourceManifestPath = join(sourcePackage, 'package.json');
  const sourceManifestBytes = readFileSync(sourceManifestPath, 'utf8');
  const sourceManifest = JSON.parse(sourceManifestBytes);
  if (sourceManifest.name !== '@betteroffice/pptx-react' || sourceManifest.version !== '0.5.1') {
    fail('The source package is not official @betteroffice/pptx-react@0.5.1.');
  }
  if (JSON.stringify(sourceManifest.dependencies) !== JSON.stringify({
    '@betteroffice/pptx': 'workspace:*',
    '@betteroffice/pptx-i18n': 'workspace:*',
  })) {
    fail('The official source manifest has unexpected runtime dependencies.');
  }
  if (sourceManifest.license !== 'Apache-2.0' ||
      sourceManifest.repository?.url !== 'https://github.com/openooxml/betteroffice.git' ||
      sourceManifest.repository?.directory !== 'packages/pptx-react') {
    fail('The official package license or repository provenance is unexpected.');
  }
  if (!existsSync(join(sourcePackage, 'dist', 'index.js')) ||
      !existsSync(join(sourcePackage, 'dist', 'index.d.ts'))) {
    fail('pptx-react build output is incomplete.');
  }

  rmSync(stage, { recursive: true, force: true });
  mkdirSync(stage, { recursive: true });
  cpSync(join(sourcePackage, 'dist'), join(stage, 'dist'), { recursive: true, dereference: false });

  const stagedManifest = structuredClone(sourceManifest);
  stagedManifest.name = channel.package.name;
  stagedManifest.version = channel.package.version;
  stagedManifest.repository = {
    ...stagedManifest.repository,
    url: channel.package.repository,
  };
  stagedManifest.dependencies = structuredClone(channel.dependencies);
  stagedManifest.files = ['dist', 'XELPER-PROVENANCE.json'];
  writeFileSync(join(stage, 'package.json'), `${JSON.stringify(stagedManifest, null, 2)}\n`);
  writeFileSync(
    join(stage, 'XELPER-PROVENANCE.json'),
    `${JSON.stringify(provenanceFor(channel), null, 2)}\n`,
  );

  if (readFileSync(sourceManifestPath, 'utf8') !== sourceManifestBytes) {
    fail('Staging mutated the source package manifest.');
  }
  const stagedFiles = listFiles(stage);
  const required = ['XELPER-PROVENANCE.json', 'dist/index.d.ts', 'dist/index.js', 'package.json'];
  for (const path of required) {
    if (!stagedFiles.includes(path)) fail(`Staged package is missing ${path}.`);
  }
  for (const path of stagedFiles) {
    if (required.includes(path)) continue;
    if (!path.startsWith('dist/') || (!path.endsWith('.js') && !path.endsWith('.d.ts'))) {
      fail(`Unexpected staged package file: ${path}`);
    }
  }
  return { manifest: stagedManifest, files: stagedFiles, provenance: provenanceFor(channel) };
}

function validateInputs(channel) {
  if (process.versions.bun !== packageManagerVersion) {
    fail(`Use locked Bun ${packageManagerVersion}; found ${process.versions.bun ?? 'no Bun runtime'}.`);
  }
  const sourceCommit = gitOutput(['rev-list', '-n', '1', `${channel.source.tag}^{commit}`]);
  if (sourceCommit !== channel.source.commit) {
    fail(`${channel.source.tag} resolves to ${sourceCommit}, not reviewed source ${channel.source.commit}.`);
  }
  for (const commit of channel.backports[0].reviewedCommits) {
    git(['cat-file', '-e', `${commit}^{commit}`]);
  }

  for (const patch of [...channel.buildEnvironment.patches, ...channel.backports]) {
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
  }
  const backport = channel.backports[0];
  const reviewedPaths = [...new Set(backport.reviewedCommits.flatMap((commit) => gitOutput([
    'diff-tree', '--no-commit-id', '--name-only', '-r', commit,
  ]).split('\n').filter((path) => path.startsWith('packages/pptx-react/'))))].sort();
  if (JSON.stringify(reviewedPaths) !== JSON.stringify([...backport.paths].sort())) {
    fail('The reviewed upstream commit path set has changed or is unexpected.');
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

  const backport = channel.backports[0];
  const patchPath = join(channelDirectory, backport.patch);
  git(['apply', '--unidiff-zero', '--check', `--directory=${sourcePrefix}`, patchPath]);
  git(['apply', '--unidiff-zero', `--directory=${sourcePrefix}`, patchPath]);
  for (const [path, expected] of Object.entries(backport.postPatchSha256)) {
    const actual = sha256(join(sourceDirectory, path));
    if (actual !== expected) fail(`Post-patch checksum mismatch for ${path}.`);
  }
}

function inspectPack(channel) {
  const dryRun = run('npm', ['pack', '--dry-run', '--json'], {
    cwd: stageDirectory,
    capture: true,
  });
  const report = JSON.parse(dryRun.stdout);
  if (!Array.isArray(report) || report.length !== 1) fail('npm pack returned an unexpected report.');
  const packed = report[0];
  if (packed.name !== channel.package.name || packed.version !== channel.package.version) {
    fail('npm pack resolved an unexpected package name or version.');
  }
  const files = packed.files.map(({ path }) => path).sort();
  const stagedFiles = listFiles(stageDirectory);
  if (JSON.stringify(files) !== JSON.stringify(stagedFiles)) {
    fail(`npm pack file set differs from staged files:\n${files.join('\n')}`);
  }
  return packed;
}

function buildPackage(channel) {
  const bun = process.execPath;
  console.log('[1/8] Installing the official tag lockfile');
  run(bun, ['install', '--frozen-lockfile'], { cwd: sourceDirectory });
  const securedDependency = readJson(join(sourceDirectory, 'node_modules', 'systeminformation', 'package.json'));
  if (securedDependency.version !== '5.33.16' || securedDependency.license !== 'MIT') {
    fail('The generated build workspace did not resolve systeminformation@5.33.16 (MIT).');
  }
  console.log('[2/8] Building the tagged PPTX Wasm prerequisite');
  run(bun, ['run', 'build:pptx-wasm'], { cwd: sourceDirectory });
  console.log('[3/8] Running the 0.5.1 focus regression');
  run(bun, [
    'test',
    'packages/pptx-react/src/worker/PeerlessPptxSurface.test.tsx',
    '--test-name-pattern',
    'strip_arrow_keys_focus_activate_and_preserve_reorder',
  ], {
    cwd: sourceDirectory,
    env: { ...process.env, BUN_CONFIG_PRELOAD: '' },
  });
  console.log('[4/8] Typechecking pptx-react');
  run(bun, ['run', '--filter', '@betteroffice/pptx-react', 'typecheck'], { cwd: sourceDirectory });
  console.log('[5/8] Building pptx-react');
  run(bun, ['run', '--filter', '@betteroffice/pptx-react', 'build'], { cwd: sourceDirectory });

  console.log('[6/8] Auditing the complete resolved dependency set');
  run(bun, ['audit'], { cwd: sourceDirectory, allowFailure: true });
  run(bun, ['audit', '--audit-level=high'], { cwd: sourceDirectory });

  console.log('[7/8] Staging and inspecting npm pack output');
  stagePackage({ channel });
  inspectPack(channel);
  rmSync(artifactsDirectory, { recursive: true, force: true });
  mkdirSync(artifactsDirectory, { recursive: true });
  const pack = run('npm', ['pack', '--json', '--pack-destination', artifactsDirectory], {
    cwd: stageDirectory,
    capture: true,
  });
  const report = JSON.parse(pack.stdout);
  if (!Array.isArray(report) || report.length !== 1) fail('npm pack did not create exactly one tarball.');
  const tarball = join(artifactsDirectory, report[0].filename);
  const artifactProvenance = {
    ...provenanceFor(channel),
    tarball: {
      file: relative(channelDirectory, tarball),
      sha256: sha256(tarball),
      size: report[0].size,
      unpackedSize: report[0].unpackedSize,
      integrity: report[0].integrity,
    },
  };
  writeFileSync(
    join(artifactsDirectory, 'provenance.json'),
    `${JSON.stringify(artifactProvenance, null, 2)}\n`,
  );
  console.log('[8/8] Created the reviewed prerelease tarball');
  return tarball;
}

function validateTag(channel, tag) {
  const expected = `${channel.package.name}@${channel.package.version}`;
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
  console.log('[0/8] Validating pinned source and backport inputs');
  validateInputs(channel);
  rmSync(artifactsDirectory, { recursive: true, force: true });
  prepareSource(channel);
  const tarball = buildPackage(channel);
  const statusAfter = gitOutput(['status', '--porcelain=v1', '--untracked-files=all']);
  if (statusAfter !== statusBefore) {
    fail('Packaging mutated the source checkout; inspect git status before retrying.');
  }
  console.log(`\nReady: ${relative(root, tarball)}`);
}

if (resolve(process.argv[1] ?? '') === fileURLToPath(import.meta.url)) {
  try {
    main();
  } catch (error) {
    console.error(error instanceof Error ? error.message : error);
    process.exitCode = 1;
  }
}
