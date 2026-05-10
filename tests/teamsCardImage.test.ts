import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';
import { buildAnswerCard } from '@omadia/channel-teams';

interface AdaptiveCard {
  $schema: string;
  type: 'AdaptiveCard';
  version: string;
  body: Array<Record<string, unknown>>;
  actions?: Array<Record<string, unknown>>;
}

function cardOf(attachment: { content: unknown }): AdaptiveCard {
  return attachment.content as AdaptiveCard;
}

describe('buildAnswerCard attachments', () => {
  it('does not emit any Image element when no attachments are passed', () => {
    const att = buildAnswerCard({ answer: 'Hallo' });
    const card = cardOf(att);
    const images = card.body.filter((b) => b['type'] === 'Image');
    assert.equal(images.length, 0);
  });

  it('emits an Image element for each diagram attachment', () => {
    const att = buildAnswerCard({
      answer: 'Hier das Diagramm:',
      attachments: [
        {
          kind: 'image',
          url: 'http://localhost:3979/diagrams/byte5/abc.png?exp=1&sig=aa',
          altText: 'flow diagram',
          diagramKind: 'mermaid',
          cacheHit: false,
        },
      ],
    });
    const card = cardOf(att);
    const images = card.body.filter((b) => b['type'] === 'Image') as Array<
      Record<string, unknown>
    >;
    assert.equal(images.length, 1);
    assert.equal(images[0]?.['altText'], 'flow diagram');
    assert.match(String(images[0]?.['url']), /diagrams\/byte5\/abc\.png/);
    const action = images[0]?.['selectAction'] as Record<string, unknown>;
    assert.equal(action?.['type'], 'Action.OpenUrl');
  });

  it('caps the number of images at 3', () => {
    const many = Array.from({ length: 10 }, (_, i) => ({
      kind: 'image' as const,
      url: `http://localhost:3979/diagrams/byte5/${String(i)}.png?exp=1&sig=aa`,
      altText: `img-${String(i)}`,
      diagramKind: 'mermaid',
      cacheHit: false,
    }));
    const att = buildAnswerCard({ answer: 'viele', attachments: many });
    const card = cardOf(att);
    const images = card.body.filter((b) => b['type'] === 'Image');
    assert.equal(images.length, 3);
  });
});
