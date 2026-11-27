// Spike S3' (byte5ai/omadia#584): participant identifier → roster, and T1
// attribution from dominant-speaker changes. The spike modules are not part
// of the package entry point, so they are imported from source (esbuild
// bundles them into the test).
import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

import {
  identifierFromRawId,
  isMeetingSpeaker,
  resolveParticipant,
  type RosterMember,
} from '../spikes/meeting-audio/src/participantIdentity.js';
import { attributeSegment } from '../spikes/meeting-audio/src/speakerAttribution.js';

const ANNA_OID = '6f1c2b8e-1d2a-4c3b-9e8f-0a1b2c3d4e5f';
const BERT_OID = '0b7e3c5d-2e4f-4a6b-8c9d-1e2f3a4b5c6d';
const SELF_OID = 'a1b2c3d4-e5f6-4a7b-8c9d-0e1f2a3b4c5d';
const BOT_APP_ID = '11111111-2222-4333-8444-555555555555';

const ANNA: RosterMember = {
  userRef: { kind: 'teams-aad', id: ANNA_OID, displayName: 'Anna Schmidt', email: 'anna@example.com' },
  externalId: '29:anna-per-bot',
};
const BERT_BY_EXTERNAL_ID: RosterMember = {
  userRef: { kind: 'teams-aad', id: '29:bert-per-bot', displayName: 'Bert Meyer' },
  externalId: BERT_OID.toUpperCase(),
};
const AGENT_BOT: RosterMember = { userRef: { kind: 'teams-aad', id: SELF_OID, displayName: 'Agent' }, isBot: true };
const ROSTER = [ANNA, BERT_BY_EXTERNAL_ID, AGENT_BOT];
const ctx = { roster: ROSTER, selfAadObjectId: SELF_OID };

describe("S3' participant → roster", () => {
  it('binds a tenant user to the roster entry by Entra object id and keeps its userRef', () => {
    const r = resolveParticipant(
      { identifier: { kind: 'microsoftTeamsUser', microsoftTeamsUserId: ANNA_OID.toUpperCase(), cloud: 'public' }, displayName: 'Anna' },
      ctx,
    );
    assert.equal(r.kind, 'roster');
    assert.ok(r.kind === 'roster');
    assert.equal(r.member, ANNA, 'the roster entry itself, so the principal equals the chat principal');
    assert.equal(r.aadObjectId, ANNA_OID);
  });

  it('matches a roster entry whose oid is only in externalId', () => {
    const r = resolveParticipant({ identifier: identifierFromRawId(`8:orgid:${BERT_OID}`) }, ctx);
    assert.equal(r.kind, 'roster');
    assert.ok(r.kind === 'roster');
    assert.equal(r.member, BERT_BY_EXTERNAL_ID);
  });

  it('never binds by display name', () => {
    const r = resolveParticipant(
      { identifier: { kind: 'microsoftTeamsUser', microsoftTeamsUserId: '22222222-3333-4444-8555-666666666666' }, displayName: 'Anna Schmidt' },
      ctx,
    );
    assert.deepEqual(r, { kind: 'tenant_user_unlisted', aadObjectId: '22222222-3333-4444-8555-666666666666', label: 'Anna Schmidt' });
  });

  it('treats an anonymous visitor as anonymous even with a roster name', () => {
    assert.deepEqual(resolveParticipant({ identifier: identifierFromRawId(`8:teamsvisitor:${ANNA_OID}`), displayName: 'Anna Schmidt' }, ctx), {
      kind: 'anonymous',
      label: 'Anna Schmidt',
    });
    assert.equal(resolveParticipant({ identifier: identifierFromRawId('8:acs:res_guest-1') }, ctx).kind, 'anonymous');
  });

  it('recognises the agent itself and other apps, which are never speakers', () => {
    const self = resolveParticipant({ identifier: identifierFromRawId(`8:orgid:${SELF_OID}`) }, ctx);
    assert.deepEqual(self, { kind: 'self', aadObjectId: SELF_OID });
    const app = resolveParticipant({ identifier: identifierFromRawId(`28:orgid:${BOT_APP_ID}`), displayName: 'Recorder' }, ctx);
    assert.deepEqual(app, { kind: 'app', appId: BOT_APP_ID, label: 'Recorder' });
    assert.equal(isMeetingSpeaker(self), false);
    assert.equal(isMeetingSpeaker(app), false);
  });

  it('does not carry a dial-in number', () => {
    const r = resolveParticipant({ identifier: identifierFromRawId('4:+4930123456'), displayName: '+49 30 123***' }, ctx);
    assert.deepEqual(r, { kind: 'phone', label: '+49 30 123***' });
    assert.equal(JSON.stringify(r).includes('4930123456'), false);
  });

  it('never matches a sovereign-cloud user against the public roster', () => {
    const r = resolveParticipant({ identifier: identifierFromRawId(`8:gcch:${ANNA_OID}`) }, ctx);
    assert.deepEqual(r, { kind: 'other_cloud', cloud: 'gcch' });
  });

  it('skips bot roster entries and malformed ids', () => {
    const asBot = resolveParticipant({ identifier: { kind: 'microsoftTeamsUser', microsoftTeamsUserId: SELF_OID } }, { roster: ROSTER });
    assert.equal(asBot.kind, 'tenant_user_unlisted', 'a bot entry with the same oid is not a human speaker');
    assert.equal(resolveParticipant({ identifier: { kind: 'microsoftTeamsUser', microsoftTeamsUserId: 'not-an-oid' } }, ctx).kind, 'unknown');
    assert.equal(resolveParticipant({ identifier: { kind: 'unknown', id: 'x:y' } }, ctx).kind, 'unknown');
  });
});

describe('T1 attribution from dominant-speaker changes', () => {
  const ANNA_RAW = `8:orgid:${ANNA_OID}`;
  const BERT_RAW = `8:orgid:${BERT_OID}`;

  it('gives a segment to the speaker who held the floor for most of it', () => {
    const changes = [
      { atMs: 0, speakerRawId: ANNA_RAW },
      { atMs: 9_000, speakerRawId: BERT_RAW },
    ];
    assert.deepEqual(attributeSegment({ startMs: 1_000, endMs: 10_000 }, changes), { speakerRawId: ANNA_RAW, share: 8 / 9 });
  });

  it('leaves a crosstalk segment unattributed', () => {
    const changes = [
      { atMs: 0, speakerRawId: ANNA_RAW },
      { atMs: 5_000, speakerRawId: BERT_RAW },
    ];
    const r = attributeSegment({ startMs: 0, endMs: 10_000 }, changes);
    assert.equal(r.speakerRawId, undefined);
    assert.equal(r.share, 0.5);
  });

  it('accepts unordered changes and silence gaps', () => {
    const changes = [
      { atMs: 4_000, speakerRawId: BERT_RAW },
      { atMs: 0 },
      { atMs: 1_000, speakerRawId: BERT_RAW },
    ];
    assert.equal(attributeSegment({ startMs: 0, endMs: 5_000 }, changes).speakerRawId, BERT_RAW);
    assert.deepEqual(attributeSegment({ startMs: 0, endMs: 0 }, changes), { share: 0 });
    assert.deepEqual(attributeSegment({ startMs: 0, endMs: 900 }, changes), { share: 0 });
  });
});
