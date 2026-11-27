// Spike A-S1 (byte5ai/omadia#584): token path for the agent's tenant user,
// against a simulated Entra token endpoint and ACS identity service.
import { after, before, describe, it } from 'node:test';
import { strict as assert } from 'node:assert';
import { createHash, createHmac } from 'node:crypto';

import {
  ACS_TEAMS_SCOPES,
  assertMeetingAudioSpikeEnabled,
  exchangeForAcsToken,
  MeetingAudioSpikeError,
  refreshEntraToken,
  signAcsRequest,
  type FetchLike,
} from '../spikes/meeting-audio/src/teamsUserToken.js';

interface Call {
  url: string;
  headers: Record<string, string>;
  body: string;
}

function fakeFetch(status: number, body: unknown, calls: Call[] = []): FetchLike {
  return async (url, init) => {
    calls.push({ url, headers: init.headers, body: init.body });
    return { status, text: async () => (typeof body === 'string' ? body : JSON.stringify(body)) };
  };
}

const ACCESS_KEY = Buffer.from('not-a-real-acs-key-for-tests').toString('base64');
const FIXED_DATE = new Date('2026-11-27T12:00:00Z');

const previousFlag = process.env.TEAMS_MEETING_AUDIO_SPIKE;
before(() => {
  process.env.TEAMS_MEETING_AUDIO_SPIKE = '1';
});
after(() => {
  if (previousFlag === undefined) delete process.env.TEAMS_MEETING_AUDIO_SPIKE;
  else process.env.TEAMS_MEETING_AUDIO_SPIKE = previousFlag;
});

describe('A-S1 feature flag', () => {
  it('refuses to run without TEAMS_MEETING_AUDIO_SPIKE=1', () => {
    assert.throws(() => assertMeetingAudioSpikeEnabled({}), MeetingAudioSpikeError);
    assert.throws(() => assertMeetingAudioSpikeEnabled({ TEAMS_MEETING_AUDIO_SPIKE: 'true' }), MeetingAudioSpikeError);
    assert.doesNotThrow(() => assertMeetingAudioSpikeEnabled({ TEAMS_MEETING_AUDIO_SPIKE: '1' }));
  });
});

describe('A-S1 calls to Microsoft are gated', () => {
  it('refuses both calls without the flag and never reaches fetch', async () => {
    const saved = process.env.TEAMS_MEETING_AUDIO_SPIKE;
    delete process.env.TEAMS_MEETING_AUDIO_SPIKE;
    try {
      const calls: Call[] = [];
      await assert.rejects(
        refreshEntraToken({ tenantId: 't', clientId: 'c', refreshToken: 'r' }, fakeFetch(200, {}, calls)),
        MeetingAudioSpikeError,
      );
      await assert.rejects(
        exchangeForAcsToken(
          { endpoint: 'https://acs-test.communication.azure.com', accessKey: ACCESS_KEY, entraAccessToken: 'x', clientId: 'c', userObjectId: 'u' },
          fakeFetch(200, {}, calls),
        ),
        MeetingAudioSpikeError,
      );
      assert.equal(calls.length, 0);
    } finally {
      if (saved !== undefined) process.env.TEAMS_MEETING_AUDIO_SPIKE = saved;
    }
  });
});

describe('A-S1 Entra refresh-token grant', () => {
  it('asks for the two ACS scopes plus offline_access and returns the rolled refresh token', async () => {
    const calls: Call[] = [];
    const tokens = await refreshEntraToken(
      { tenantId: 'tenant-1', clientId: 'client-1', refreshToken: 'rt-old' },
      fakeFetch(200, { access_token: 'at', refresh_token: 'rt-new', expires_in: 3600 }, calls),
    );
    assert.deepEqual(tokens, { accessToken: 'at', refreshToken: 'rt-new', expiresInSec: 3600 });
    assert.equal(calls[0]?.url, 'https://login.microsoftonline.com/tenant-1/oauth2/v2.0/token');
    const form = new URLSearchParams(calls[0]?.body);
    assert.equal(form.get('grant_type'), 'refresh_token');
    assert.equal(form.get('scope'), [...ACS_TEAMS_SCOPES, 'offline_access'].join(' '));
  });

  it('reports the Entra error code without echoing the token', async () => {
    await assert.rejects(
      refreshEntraToken(
        { tenantId: 't', clientId: 'c', refreshToken: 'rt-secret' },
        fakeFetch(400, { error: 'invalid_grant', error_description: 'AADSTS50173 … rt-secret' }),
      ),
      (err: unknown) =>
        err instanceof MeetingAudioSpikeError &&
        err.code === 'invalid_grant' &&
        err.status === 400 &&
        !err.message.includes('rt-secret'),
    );
  });
});

describe('A-S1 ACS exchange', () => {
  it('signs the request and sends token, appId and userId', async () => {
    const calls: Call[] = [];
    const result = await exchangeForAcsToken(
      {
        endpoint: 'https://acs-test.communication.azure.com',
        accessKey: ACCESS_KEY,
        entraAccessToken: 'entra-at',
        clientId: 'client-1',
        userObjectId: 'oid-1',
        now: () => FIXED_DATE,
      },
      fakeFetch(200, { token: 'acs-token', expiresOn: '2026-11-27T13:00:00Z' }, calls),
    );
    assert.deepEqual(result, { token: 'acs-token', expiresOn: '2026-11-27T13:00:00Z' });
    const call = calls[0];
    assert.ok(call);
    assert.equal(call.url, 'https://acs-test.communication.azure.com/teamsUser/:exchangeAccessToken?api-version=2023-10-01');
    assert.deepEqual(JSON.parse(call.body), { token: 'entra-at', appId: 'client-1', userId: 'oid-1' });

    const hash = createHash('sha256').update(call.body).digest('base64');
    const toSign = `POST\n/teamsUser/:exchangeAccessToken?api-version=2023-10-01\n${FIXED_DATE.toUTCString()};acs-test.communication.azure.com;${hash}`;
    const expected = createHmac('sha256', Buffer.from(ACCESS_KEY, 'base64')).update(toSign).digest('base64');
    assert.equal(call.headers['x-ms-content-sha256'], hash);
    assert.equal(call.headers.authorization, `HMAC-SHA256 SignedHeaders=x-ms-date;host;x-ms-content-sha256&Signature=${expected}`);
  });

  it('surfaces the ACS error code (e.g. a token whose oid is not the user)', async () => {
    await assert.rejects(
      exchangeForAcsToken(
        { endpoint: 'https://acs-test.communication.azure.com', accessKey: ACCESS_KEY, entraAccessToken: 'x', clientId: 'c', userObjectId: 'u' },
        fakeFetch(401, { error: { code: 'InvalidAccessToken', message: 'oid mismatch' } }),
      ),
      (err: unknown) => err instanceof MeetingAudioSpikeError && err.code === 'InvalidAccessToken' && err.status === 401,
    );
  });

  it('refuses a non-https endpoint and a response without a token', async () => {
    await assert.rejects(
      exchangeForAcsToken(
        { endpoint: 'http://acs-test.communication.azure.com', accessKey: ACCESS_KEY, entraAccessToken: 'x', clientId: 'c', userObjectId: 'u' },
        fakeFetch(200, {}),
      ),
      /https/,
    );
    await assert.rejects(
      exchangeForAcsToken(
        { endpoint: 'https://acs-test.communication.azure.com', accessKey: ACCESS_KEY, entraAccessToken: 'x', clientId: 'c', userObjectId: 'u' },
        fakeFetch(200, { expiresOn: 'x' }),
      ),
      /no token/,
    );
  });

  it('signs the query string into the path', () => {
    const headers = signAcsRequest({
      method: 'post',
      url: new URL('https://h.example/a?b=1'),
      body: '',
      accessKey: ACCESS_KEY,
      date: FIXED_DATE,
    });
    const expected = createHmac('sha256', Buffer.from(ACCESS_KEY, 'base64'))
      .update(`POST\n/a?b=1\n${FIXED_DATE.toUTCString()};h.example;${createHash('sha256').update('').digest('base64')}`)
      .digest('base64');
    assert.ok(headers.authorization?.endsWith(expected));
  });
});
