import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

import {
  createTeamsDefaultChannelKeyProvider,
  resolveTeamsDefaultChannelKey,
} from '../src/teamsDefaultChannelKey.js';
import type { TeamsBotIdentity } from '../src/teamsBotIdentity.js';

function bot(appId: string, botSlug: string): TeamsBotIdentity {
  return {
    appId,
    tenantId: 'tenant',
    appType: 'SingleTenant',
    appPasswordSecretRef: `${botSlug}_password`,
    botSlug,
    displayName: botSlug,
  };
}

const AGENT_BOT = bot('AAAA-agent', 'mira');
const DEPLOYMENT_BOT = bot('BBBB-Deployment', 'default');

describe('Teams default bot key (kernel Conductor routing)', () => {
  it('answers the Microsoft 365 app registration when the plugin serves it, lowercase', () => {
    assert.equal(
      resolveTeamsDefaultChannelKey([' bbbb-deployment '], [AGENT_BOT, DEPLOYMENT_BOT]),
      '28:bbbb-deployment',
    );
  });

  it('takes the config value before the env value, like the kernel did', () => {
    assert.equal(
      resolveTeamsDefaultChannelKey(['bbbb-deployment', 'aaaa-agent'], [AGENT_BOT, DEPLOYMENT_BOT]),
      '28:bbbb-deployment',
    );
    assert.equal(
      resolveTeamsDefaultChannelKey([undefined, 'BBBB-deployment'], [AGENT_BOT, DEPLOYMENT_BOT]),
      '28:bbbb-deployment',
    );
  });

  it('never names an operator-listed agent bot as the default', () => {
    assert.equal(resolveTeamsDefaultChannelKey([undefined, '  '], [AGENT_BOT, DEPLOYMENT_BOT]), undefined);
    assert.equal(resolveTeamsDefaultChannelKey(['cccc-unserved'], [AGENT_BOT]), undefined);
    assert.equal(resolveTeamsDefaultChannelKey(['cccc-unserved', 'aaaa-agent'], [AGENT_BOT]), undefined);
  });

  it('provider reads the config per call and serves channel type teams', () => {
    let scalar: string | undefined = 'bbbb-deployment';
    const provider = createTeamsDefaultChannelKeyProvider(() => [scalar], [AGENT_BOT, DEPLOYMENT_BOT]);
    assert.equal(provider.channelType, 'teams');
    assert.equal(provider.defaultChannelKey(), '28:bbbb-deployment');
    scalar = undefined;
    assert.equal(provider.defaultChannelKey(), undefined);
  });
});
