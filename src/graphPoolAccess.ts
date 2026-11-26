import type { Pool } from 'pg';

/**
 * `graphPool`, or undefined when the kernel has none or refuses it because
 * the operator has not granted `permissions.sql` (omadia#838, ADR-0050 C7:
 * the kernel throws `SqlPermissionError`). A refusal is logged with the
 * remedy and never thrown, so the channel activates without a database
 * instead of not at all.
 */
export function resolveGraphPool(
  services: { get<T>(name: string): T | undefined },
  warn: (message: string) => void,
): Pool | undefined {
  try {
    return services.get<Pool>('graphPool');
  } catch (err) {
    // Only the kernel's refusal for a missing grant is survivable here; any
    // other error (an undeclared service, a broken registry) is a real fault.
    if (!(err instanceof Error) || err.name !== 'SqlPermissionError') throw err;
    warn(
      `teams: no database access (${err.message}). Grant the plugin's database permission on its page in the admin UI (needs a kernel that runs channel migrations, omadia#838); until then question cards, meeting state and conversation references are kept in memory only.`,
    );
    return undefined;
  }
}
