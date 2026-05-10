import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';
import { aiLabelEntity, buildAnswerCard } from '@omadia/channel-teams';
import type { RunTracePayload } from '@omadia/orchestrator';

const sampleTrace: RunTracePayload = {
  scope: 'teams-test',
  userId: 'u1',
  startedAt: '2026-04-19T10:00:00.000Z',
  finishedAt: '2026-04-19T10:00:08.400Z',
  durationMs: 8400,
  status: 'success',
  iterations: 2,
  orchestratorToolCalls: [
    {
      callId: 'orch-1',
      toolName: 'query_knowledge_graph',
      durationMs: 30,
      isError: false,
      agentContext: 'orchestrator',
    },
  ],
  agentInvocations: [
    {
      index: 0,
      agentName: 'query_odoo_hr',
      durationMs: 4200,
      subIterations: 3,
      status: 'success',
      toolCalls: [
        {
          callId: 's-1',
          toolName: 'odoo_execute',
          durationMs: 420,
          isError: false,
          agentContext: 'query_odoo_hr',
        },
        {
          callId: 's-2',
          toolName: 'odoo_execute',
          durationMs: 180,
          isError: false,
          agentContext: 'query_odoo_hr',
        },
      ],
    },
    {
      index: 1,
      agentName: 'query_confluence_playbook',
      durationMs: 2100,
      subIterations: 1,
      status: 'error',
      toolCalls: [
        {
          callId: 's-3',
          toolName: 'confluence_search',
          durationMs: 1800,
          isError: true,
          agentContext: 'query_confluence_playbook',
        },
      ],
    },
  ],
};

describe('aiLabelEntity', () => {
  it('returns the AIGeneratedContent schema.org entity Teams expects', () => {
    const e = aiLabelEntity() as Record<string, unknown>;
    assert.equal(e['type'], 'https://schema.org/Message');
    assert.deepEqual(e['additionalType'], ['AIGeneratedContent']);
  });
});

describe('buildAnswerCard', () => {
  it('returns an Adaptive Card attachment with the answer text', () => {
    const att = buildAnswerCard({ answer: 'Das ist die Antwort.' });
    assert.equal(att.contentType, 'application/vnd.microsoft.card.adaptive');
    const json = JSON.stringify(att.content);
    assert.match(json, /Das ist die Antwort\./);
    assert.match(json, /AI-generated/);
  });

  it('includes a ToggleVisibility action when a run trace is present', () => {
    const att = buildAnswerCard({
      answer: 'OK',
      runTrace: sampleTrace,
    });
    const content = att.content as { actions?: unknown[]; body: unknown[] };
    assert.ok(Array.isArray(content.actions) && content.actions.length > 0);
    const action = content.actions[0] as { type: string; targetElements: string[] };
    assert.equal(action.type, 'Action.ToggleVisibility');
    assert.deepEqual(action.targetElements, ['trace-container']);
  });

  it('embeds every agent name and sub-tool in the trace container at full level', () => {
    const att = buildAnswerCard({ answer: 'OK', runTrace: sampleTrace });
    const json = JSON.stringify(att.content);
    assert.match(json, /query_odoo_hr/);
    assert.match(json, /query_confluence_playbook/);
    assert.match(json, /odoo_execute/);
    assert.match(json, /confluence_search/);
  });

  it('degrades gracefully on oversized traces — card stays under the Teams cap', () => {
    // Synthesise a trace with thousands of sub-calls.
    const huge: RunTracePayload = {
      ...sampleTrace,
      agentInvocations: [
        {
          ...sampleTrace.agentInvocations[0]!,
          toolCalls: Array.from({ length: 2000 }, (_, i) => ({
            callId: `s-${String(i)}`,
            toolName: 'odoo_execute',
            durationMs: 100,
            isError: false,
            agentContext: 'query_odoo_hr',
          })),
        },
      ],
    };
    const att = buildAnswerCard({ answer: 'OK', runTrace: huge });
    const serialized = JSON.stringify(att.content);
    assert.ok(
      Buffer.byteLength(serialized, 'utf8') < 28_000,
      `card size ${String(Buffer.byteLength(serialized, 'utf8'))} exceeds Teams cap`,
    );
  });

  it('omits the trace section entirely when no trace is passed', () => {
    const att = buildAnswerCard({ answer: 'nur antwort' });
    const content = att.content as { actions?: unknown[] };
    assert.equal(content.actions, undefined);
  });

  // -------------------------------------------------------------------------
  // OB-81 · Palaia capture-disclosure section
  // -------------------------------------------------------------------------

  it('renders a ToggleVisibility action for capture-disclosure parallel to the trace', () => {
    const att = buildAnswerCard({
      answer: 'OK',
      captureDisclosure: {
        persisted: true,
        reasons: ['scorer-skipped:level=minimal', 'privacy-strip:1'],
        entryType: 'memory',
        visibility: 'team',
        significance: null,
        embedded: true,
        privacyBlocksStripped: 1,
        hintTagsProcessed: 0,
      },
    });
    const content = att.content as { actions?: unknown[]; body: unknown[] };
    assert.ok(Array.isArray(content.actions));
    const disclosureAction = (content.actions ?? []).find(
      (a) =>
        (a as { targetElements?: string[] }).targetElements?.[0] ===
        'capture-disclosure-container',
    );
    assert.ok(
      disclosureAction !== undefined,
      'expected a ToggleVisibility action targeting capture-disclosure-container',
    );
    const json = JSON.stringify(att.content);
    assert.match(json, /Memory-Auswirkung/);
    assert.match(json, /persisted/);
    // The summary should also surface the entry type for at-a-glance scan.
    assert.match(json, /memory/);
  });

  it('renders the disclosure body with significance + reason markers', () => {
    const att = buildAnswerCard({
      answer: 'OK',
      captureDisclosure: {
        persisted: false,
        reasons: ['dropped:significance<threshold'],
        entryType: null,
        visibility: null,
        significance: 0.08,
        embedded: false,
        privacyBlocksStripped: 0,
        hintTagsProcessed: 0,
      },
    });
    const json = JSON.stringify(att.content);
    assert.match(json, /verworfen/);
    assert.match(json, /0\.08/);
    assert.match(json, /dropped:significance<threshold/);
  });

  it('omits the disclosure section entirely when captureDisclosure is undefined', () => {
    const att = buildAnswerCard({ answer: 'plain' });
    const json = JSON.stringify(att.content);
    assert.doesNotMatch(json, /capture-disclosure-container/);
    assert.doesNotMatch(json, /Memory-Auswirkung/);
  });

  it('drops disclosure (but keeps answer) when card grows past the Teams cap', () => {
    const longAnswer = 'x'.repeat(25_000);
    const att = buildAnswerCard({
      answer: longAnswer,
      captureDisclosure: {
        persisted: true,
        reasons: ['ok'],
        entryType: 'memory',
        visibility: 'team',
        significance: 0.9,
        embedded: true,
        privacyBlocksStripped: 0,
        hintTagsProcessed: 0,
      },
    });
    const serialized = JSON.stringify(att.content);
    assert.ok(
      Buffer.byteLength(serialized, 'utf8') < 28_000,
      `card size ${String(Buffer.byteLength(serialized, 'utf8'))} exceeds Teams cap`,
    );
    assert.doesNotMatch(serialized, /capture-disclosure-container/);
  });
});
