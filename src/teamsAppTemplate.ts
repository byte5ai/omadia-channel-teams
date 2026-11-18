// What the shipped agent app package template says about itself (#2043):
// its revision and the resource-specific permissions it requests. The
// readiness check compares an install against it, so an install made from an
// earlier package is named (`package_outdated`) instead of looking fine.

import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

export interface TeamsAppTemplateInfo {
  /** `appPackage/TEMPLATE_REVISION.txt`; the middle part of an agent app's
   *  version `1.<template revision>.<identity revision>`. */
  revision: number;
  /** RSC permissions the template requests, in template order. */
  permissions: readonly string[];
}

/** The file holding the template revision. It carries an extension because
 *  the kernel's plugin zip ingest refuses extension-less files (0.36.0 shipped
 *  it as `TEMPLATE_REVISION` and could not be installed). */
export const TEMPLATE_REVISION_FILE = 'TEMPLATE_REVISION.txt';

/** `appPackage/` next to `dist/` in the installed plugin. */
export function defaultAppPackageDir(): string {
  return join(dirname(fileURLToPath(import.meta.url)), '..', 'appPackage');
}

/**
 * Read revision and permissions from an `appPackage/` directory. Returns
 * `undefined` when either file is missing or unusable: the caller then makes
 * no statement about outdated packages rather than a wrong one.
 */
export function readTeamsAppTemplateInfo(dir: string = defaultAppPackageDir()): TeamsAppTemplateInfo | undefined {
  try {
    const revisionText = readFileSync(join(dir, TEMPLATE_REVISION_FILE), 'utf8').trim();
    if (!/^\d{1,6}$/.test(revisionText)) return undefined;
    const template = readFileSync(join(dir, 'manifest.json.template'), 'utf8');
    // The template is not JSON until rendered (it has unquoted placeholders),
    // so the permission block is read by pattern.
    const block = /"resourceSpecific"\s*:\s*\[([\s\S]*?)\]/.exec(template)?.[1];
    if (block === undefined) return undefined;
    const permissions = [...block.matchAll(/"name"\s*:\s*"([^"]+)"/g)].flatMap((match) => (match[1] ? [match[1]] : []));
    return { revision: Number(revisionText), permissions };
  } catch {
    return undefined;
  }
}

/** The template revision an agent app version was built from: the middle
 *  part of `1.<template revision>.<identity revision>`. Versions from before
 *  the scheme (`1.0.<n>`) read as revision 0. `undefined` for anything else. */
export function templateRevisionOfAppVersion(version: string | undefined): number | undefined {
  const match = /^1\.(\d{1,6})\.\d+$/.exec(version?.trim() ?? '');
  return match?.[1] !== undefined ? Number(match[1]) : undefined;
}
