import { afterEach, expect, test } from 'bun:test';
import { mkdtempSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { loadChannel, stagePackage } from './package-xelper-pptx-react.mjs';

const temporaryDirectories: string[] = [];

afterEach(() => {
  for (const directory of temporaryDirectories.splice(0)) {
    rmSync(directory, { recursive: true, force: true });
  }
});

test('stages the reviewed personal-scope package without mutating official source', () => {
  const directory = mkdtempSync(join(tmpdir(), 'betteroffice-pptx-react-channel-'));
  temporaryDirectories.push(directory);
  const source = join(directory, 'source');
  const stage = join(directory, 'stage');
  mkdirSync(join(source, 'dist'), { recursive: true });
  const sourceManifest = {
    name: '@betteroffice/pptx-react',
    version: '0.5.1',
    type: 'module',
    main: './dist/index.js',
    types: './dist/index.d.ts',
    exports: {
      '.': { types: './dist/index.d.ts', import: './dist/index.js' },
    },
    files: ['dist'],
    dependencies: {
      '@betteroffice/pptx': 'workspace:*',
      '@betteroffice/pptx-i18n': 'workspace:*',
    },
    peerDependencies: {
      react: '^18.0.0 || ^19.0.0',
      'react-dom': '^18.0.0 || ^19.0.0',
    },
    license: 'Apache-2.0',
    repository: {
      type: 'git',
      url: 'https://github.com/openooxml/betteroffice.git',
      directory: 'packages/pptx-react',
    },
  };
  const manifestBytes = `${JSON.stringify(sourceManifest, null, 2)}\n`;
  writeFileSync(join(source, 'package.json'), manifestBytes);
  writeFileSync(join(source, 'dist', 'index.js'), 'export const editor = true;\n');
  writeFileSync(join(source, 'dist', 'index.d.ts'), 'export declare const editor: true;\n');
  writeFileSync(join(source, 'dist', 'chunk-ABC.js'), 'export {};\n');

  const result = stagePackage({ channel: loadChannel(), sourcePackage: source, stage });

  expect(result.manifest).toMatchObject({
    name: '@davidranaway/betteroffice-pptx-react',
    version: '0.5.1-xelper.4',
    dependencies: {
      '@betteroffice/pptx': '0.5.1',
      '@betteroffice/pptx-i18n': '0.5.1',
    },
    license: 'Apache-2.0',
    repository: {
      url: 'https://github.com/davidranaway/betteroffice.git',
      directory: 'packages/pptx-react',
    },
  });
  expect(result.manifest.peerDependencies).toEqual(sourceManifest.peerDependencies);
  expect(result.manifest.dependencies).toEqual({
    '@betteroffice/pptx': '0.5.1',
    '@betteroffice/pptx-i18n': '0.5.1',
  });
  expect(result.files).toEqual([
    'XELPER-PROVENANCE.json',
    'dist/chunk-ABC.js',
    'dist/index.d.ts',
    'dist/index.js',
    'package.json',
  ]);
  expect(result.provenance).toMatchObject({
    source: {
      repository: 'https://github.com/openooxml/betteroffice.git',
      tag: '@betteroffice/pptx-react@0.5.1',
      commit: '4e982c3aea0065578604c3fe7727938533888f66',
    },
    backports: [{
      upstreamPullRequest: 2698,
      behavior: 'Full slide-strip keyboard navigation and focus retention',
      reviewedCommits: [
        '541770e783cd972527b0f7e4592965dabcb9229e',
        'aeb5b2c6591efa80906d84d55ea8e14aaf344a4f',
      ],
    }],
    buildEnvironment: {
      securityPatches: [{
        advisory: 'GHSA-3f7j-7jwc-fxfj',
        package: 'systeminformation',
        version: '5.33.16',
        license: 'MIT',
        patch: 'patches/build/0001-pin-systeminformation-5.33.16.patch',
      }],
    },
  });
  expect(readFileSync(join(source, 'package.json'), 'utf8')).toBe(manifestBytes);
  expect(readFileSync(join(source, 'dist', 'index.js'), 'utf8')).toBe('export const editor = true;\n');
});
