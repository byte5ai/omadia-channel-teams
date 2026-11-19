#!/usr/bin/env node
/**
 * typecheck-against-sdk.mjs — typecheck this plugin against ANOTHER checkout
 * of the kernel's SDK packages than the one `tsconfig.json` points at.
 *
 *     node scripts/typecheck-against-sdk.mjs <dir with middleware/packages>
 *
 * Why: the plugin ships independently of the kernel and is built by whoever
 * rolls it out against the kernel packages they have. 0.36.6 used a key
 * (`supports.aiLabel`) that the SDK of the kernel's main branch did not have
 * yet, and its build failed there. Before a release, run this against the SDK
 * of the kernel's `main` (the OLDEST SDK the release is built against) in
 * addition to the normal typecheck against the newest one:
 *
 *     git -C ../odoo-bot archive origin/main middleware/packages/harness-channel-sdk \
 *       middleware/packages/plugin-api | tar -x -C /tmp/sdk-main
 *     node scripts/typecheck-against-sdk.mjs /tmp/sdk-main
 *
 * The SDK is read from SOURCE (`src/index.ts`), so no build of the kernel
 * packages is needed. Only errors in this plugin's own files count.
 */

import { spawnSync } from 'node:child_process';
import { existsSync, rmSync, writeFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const pkgRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const sdkRoot = process.argv[2] ? resolve(process.argv[2]) : undefined;
if (!sdkRoot) {
  console.error('usage: node scripts/typecheck-against-sdk.mjs <dir with middleware/packages>');
  process.exit(2);
}
const packages = join(sdkRoot, 'middleware', 'packages');
const sdkIndex = join(packages, 'harness-channel-sdk', 'src', 'index.ts');
const apiIndex = join(packages, 'plugin-api', 'src', 'index.ts');
for (const file of [sdkIndex, apiIndex]) {
  if (!existsSync(file)) {
    console.error(`not found: ${file}`);
    process.exit(2);
  }
}

const configPath = join(pkgRoot, 'tsconfig.sdk-check.tmp.json');
writeFileSync(
  configPath,
  JSON.stringify(
    {
      extends: './tsconfig.json',
      compilerOptions: {
        noEmit: true,
        composite: false,
        declaration: false,
        declarationMap: false,
        rootDir: '/',
        paths: {
          '@omadia/plugin-api': [apiIndex],
          '@omadia/channel-sdk': [sdkIndex],
          '@omadia/plugin-ui-helpers': ['../odoo-bot/middleware/packages/harness-ui-helpers/dist/index.d.ts'],
          '@omadia/orchestrator': ['../odoo-bot/middleware/packages/harness-orchestrator/dist/index.d.ts'],
          '@omadia/diagrams': ['../odoo-bot/middleware/packages/harness-diagrams/dist/index.d.ts'],
          '@omadia/integration-microsoft365': ['../omadia-m365-connector/dist/index.d.ts'],
        },
      },
      include: ['src/**/*.ts'],
    },
    null,
    2,
  ),
);

let output = '';
try {
  const res = spawnSync('npx', ['tsc', '-p', configPath], { cwd: pkgRoot, encoding: 'utf8' });
  output = `${res.stdout ?? ''}${res.stderr ?? ''}`;
} finally {
  rmSync(configPath, { force: true });
}

const errors = output.split('\n').filter((line) => /error TS\d+/.test(line));
// The SDK sources are compiled along; their own errors (missing dependencies
// of the kernel packages in this repo) are not this plugin's.
const own = errors.filter((line) => line.startsWith('src/'));
const foreign = errors.length - own.length;
const version = (dir) => {
  try {
    return JSON.parse(spawnSync('cat', [join(packages, dir, 'package.json')], { encoding: 'utf8' }).stdout).version;
  } catch {
    return '?';
  }
};
console.log(`channel-sdk ${version('harness-channel-sdk')}, plugin-api ${version('plugin-api')} at ${sdkRoot}`);
if (own.length > 0) {
  console.error(own.join('\n'));
  console.error(`✗ ${String(own.length)} error(s) in this plugin against that SDK`);
  process.exit(1);
}
console.log(`✓ this plugin typechecks against that SDK (${String(foreign)} error(s) inside the SDK sources ignored)`);
