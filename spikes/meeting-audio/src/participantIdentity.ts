// Spike S3' (byte5ai/omadia#584, omadia ADR-0287): map a participant of a
// Teams meeting, as the ACS Calling SDK names it, to the session roster.
//
// SPIKE CODE. Not wired into the plugin, not exported from the package
// entry point. It answers one question with tests: can a voice in the call be
// tied to the same roster entry (and therefore the same principal) as that
// person's chat messages, without guessing?
//
// The rules the prototype pins down:
//   - Only the Entra object id binds. A Teams user in the call is
//     `8:orgid:<oid>` (public cloud); the roster carries the same oid as
//     `userRef.id` (kind `teams-aad`) or as `externalId`. A match returns the
//     roster's own `userRef`, so the kernel derives the same principal for a
//     spoken contribution as for a chat message (`channelUserPrincipal`).
//   - A display name never binds. Two people can share a name and an
//     anonymous guest can type any name; the name is carried as a label only.
//   - The agent's own tenant user is recognised and never attributed as a
//     speaker of the meeting.
//   - Phone numbers are not carried: a dial-in caller is "phone", the number
//     stays in the call.
//   - Sovereign clouds (`8:dod:`, `8:gcch:`) never match a public-cloud
//     roster, even with the same oid.
// Speech attributed this way is still unconfirmed (omadia ADR-0026): the
// mapping says who spoke, not who authored a decision.

/** The identifier shapes of `@azure/communication-common` that a remote
 *  participant of a Teams meeting can carry (structural mirror, no SDK
 *  dependency). */
export type AcsParticipantIdentifier =
  | {
      readonly kind: 'microsoftTeamsUser';
      readonly microsoftTeamsUserId: string;
      readonly isAnonymous?: boolean;
      readonly cloud?: 'public' | 'dod' | 'gcch';
      readonly rawId?: string;
    }
  | { readonly kind: 'communicationUser'; readonly communicationUserId: string }
  | { readonly kind: 'phoneNumber'; readonly phoneNumber: string; readonly rawId?: string }
  | {
      readonly kind: 'microsoftTeamsApp';
      readonly teamsAppId: string;
      readonly cloud?: 'public' | 'dod' | 'gcch';
      readonly rawId?: string;
    }
  | { readonly kind: 'unknown'; readonly id: string };

/** One roster member, the fields of channel-sdk `ConversationParticipant`
 *  this mapping reads. */
export interface RosterMember {
  readonly userRef: {
    readonly kind: string;
    readonly id: string;
    readonly displayName?: string;
    readonly email?: string;
  };
  readonly externalId?: string | null;
  readonly isBot?: boolean;
}

export type ParticipantResolution =
  /** A tenant user who is on the roster; `member` is the roster's own entry. */
  | { readonly kind: 'roster'; readonly aadObjectId: string; readonly member: RosterMember }
  /** The agent's own tenant user. Never a speaker of the meeting. */
  | { readonly kind: 'self'; readonly aadObjectId: string }
  /** A tenant user the roster does not list (joined the call, not the chat,
   *  or the roster is partial). Attributable by oid, not by principal yet. */
  | { readonly kind: 'tenant_user_unlisted'; readonly aadObjectId: string; readonly label?: string }
  /** A Teams user of a sovereign cloud; never matched against this roster. */
  | { readonly kind: 'other_cloud'; readonly cloud: string; readonly label?: string }
  /** Anonymous guest (Teams visitor) or an ACS identity. */
  | { readonly kind: 'anonymous'; readonly label?: string }
  /** Dial-in caller. The number is deliberately not carried. */
  | { readonly kind: 'phone'; readonly label?: string }
  /** Another app or bot in the call. */
  | { readonly kind: 'app'; readonly appId: string; readonly label?: string }
  | { readonly kind: 'unknown'; readonly label?: string };

export interface ResolveParticipantInput {
  readonly identifier: AcsParticipantIdentifier;
  /** `remoteParticipant.displayName`; a label, never used to bind. */
  readonly displayName?: string;
}

export interface ResolveParticipantContext {
  readonly roster: readonly RosterMember[];
  /** Entra object id of the agent's own tenant user. */
  readonly selfAadObjectId?: string;
}

const ENTRA_OBJECT_ID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

/** Raw-id prefixes of `@azure/communication-common` (`identifierModels`). */
const RAW_ID_PREFIXES: readonly { readonly prefix: string; readonly parse: (rest: string) => AcsParticipantIdentifier }[] = [
  { prefix: '8:orgid:', parse: (rest) => ({ kind: 'microsoftTeamsUser', microsoftTeamsUserId: rest, cloud: 'public' }) },
  { prefix: '8:dod:', parse: (rest) => ({ kind: 'microsoftTeamsUser', microsoftTeamsUserId: rest, cloud: 'dod' }) },
  { prefix: '8:gcch:', parse: (rest) => ({ kind: 'microsoftTeamsUser', microsoftTeamsUserId: rest, cloud: 'gcch' }) },
  {
    prefix: '8:teamsvisitor:',
    parse: (rest) => ({ kind: 'microsoftTeamsUser', microsoftTeamsUserId: rest, isAnonymous: true, cloud: 'public' }),
  },
  { prefix: '28:orgid:', parse: (rest) => ({ kind: 'microsoftTeamsApp', teamsAppId: rest, cloud: 'public' }) },
  { prefix: '28:dod:', parse: (rest) => ({ kind: 'microsoftTeamsApp', teamsAppId: rest, cloud: 'dod' }) },
  { prefix: '28:gcch:', parse: (rest) => ({ kind: 'microsoftTeamsApp', teamsAppId: rest, cloud: 'gcch' }) },
  { prefix: '8:acs:', parse: (rest) => ({ kind: 'communicationUser', communicationUserId: `8:acs:${rest}` }) },
  { prefix: '8:spool:', parse: (rest) => ({ kind: 'communicationUser', communicationUserId: `8:spool:${rest}` }) },
  { prefix: '4:', parse: (rest) => ({ kind: 'phoneNumber', phoneNumber: rest.startsWith('+') ? rest : `+${rest}` }) },
];

/** Parse a raw id (`8:orgid:<oid>`, `4:+49…`, …) as the SDK's
 *  `createIdentifierFromRawId` does, for the prefixes a meeting can carry. */
export function identifierFromRawId(rawId: string): AcsParticipantIdentifier {
  const trimmed = rawId.trim();
  for (const { prefix, parse } of RAW_ID_PREFIXES) {
    if (trimmed.startsWith(prefix) && trimmed.length > prefix.length) {
      return parse(trimmed.slice(prefix.length));
    }
  }
  return { kind: 'unknown', id: trimmed };
}

function normaliseOid(raw: string | null | undefined): string | undefined {
  const value = raw?.trim().toLowerCase();
  return value && ENTRA_OBJECT_ID.test(value) ? value : undefined;
}

function rosterOid(member: RosterMember): string | undefined {
  const fromRef = member.userRef.kind === 'teams-aad' ? normaliseOid(member.userRef.id) : undefined;
  return fromRef ?? normaliseOid(member.externalId);
}

function labelOf(displayName: string | undefined): { label?: string } {
  const label = displayName?.trim();
  return label ? { label } : {};
}

function resolveTeamsUser(
  identifier: Extract<AcsParticipantIdentifier, { kind: 'microsoftTeamsUser' }>,
  displayName: string | undefined,
  ctx: ResolveParticipantContext,
): ParticipantResolution {
  if (identifier.isAnonymous) return { kind: 'anonymous', ...labelOf(displayName) };
  const cloud = identifier.cloud ?? 'public';
  if (cloud !== 'public') return { kind: 'other_cloud', cloud, ...labelOf(displayName) };
  const oid = normaliseOid(identifier.microsoftTeamsUserId);
  if (!oid) return { kind: 'unknown', ...labelOf(displayName) };
  if (oid === normaliseOid(ctx.selfAadObjectId)) return { kind: 'self', aadObjectId: oid };
  const member = ctx.roster.find((candidate) => !candidate.isBot && rosterOid(candidate) === oid);
  return member
    ? { kind: 'roster', aadObjectId: oid, member }
    : { kind: 'tenant_user_unlisted', aadObjectId: oid, ...labelOf(displayName) };
}

/** Resolve one remote participant against the roster. Pure; never throws. */
export function resolveParticipant(
  input: ResolveParticipantInput,
  ctx: ResolveParticipantContext,
): ParticipantResolution {
  const { identifier, displayName } = input;
  switch (identifier.kind) {
    case 'microsoftTeamsUser':
      return resolveTeamsUser(identifier, displayName, ctx);
    case 'communicationUser':
      return { kind: 'anonymous', ...labelOf(displayName) };
    case 'phoneNumber':
      return { kind: 'phone', ...labelOf(displayName) };
    case 'microsoftTeamsApp':
      return { kind: 'app', appId: identifier.teamsAppId, ...labelOf(displayName) };
    case 'unknown': {
      const parsed = identifierFromRawId(identifier.id);
      return parsed.kind === 'unknown'
        ? { kind: 'unknown', ...labelOf(displayName) }
        : resolveParticipant({ identifier: parsed, displayName }, ctx);
    }
  }
}

/** Whether a resolution may appear as a speaker of the meeting at all. */
export function isMeetingSpeaker(resolution: ParticipantResolution): boolean {
  return resolution.kind !== 'self' && resolution.kind !== 'app';
}
