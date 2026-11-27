import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

import { toStoredTurnAttachments, topicDecisionUploads } from '@omadia/channel-teams';

// omadia#1962 (ADR-0178): the kernel ingests storage keys only from the
// structured `storedAttachments` field, never from the `[attachments-info]`
// text. The bot states exactly what it persisted from the activity.

describe('toStoredTurnAttachments', () => {
  it('states every persisted file with key, name, type and size', () => {
    const out = toStoredTurnAttachments([
      {
        storageKey: 'teams-attachments/t/c/1-a.pdf',
        fileName: 'a.pdf',
        contentType: 'application/pdf',
        sizeBytes: 2048,
        sha256: 'x',
        source: 'teams.file',
        signedUrl: 'https://bot.example/attachments/a?sig=1',
      },
      {
        storageKey: 'teams-attachments/t/c/2-b.png',
        fileName: 'b.png',
        contentType: 'image/png',
        sizeBytes: 10,
        sha256: 'y',
        source: 'teams.file',
      },
    ]);
    assert.deepEqual(out, [
      {
        fileName: 'a.pdf',
        contentType: 'application/pdf',
        storageKey: 'teams-attachments/t/c/1-a.pdf',
        sizeBytes: 2048,
        signedUrl: 'https://bot.example/attachments/a?sig=1',
      },
      {
        fileName: 'b.png',
        contentType: 'image/png',
        storageKey: 'teams-attachments/t/c/2-b.png',
        sizeBytes: 10,
      },
    ]);
  });

  it('states nothing when nothing was persisted', () => {
    assert.deepEqual(toStoredTurnAttachments([]), []);
  });
});

// omadia#2304: a message that triggered the topic card keeps its uploads on the
// server-side pending entry; the click's turn gets them, never from the card.
describe('topicDecisionUploads', () => {
  const upload = {
    fileName: 'notiz.wav',
    contentType: 'audio/wav',
    storageKey: 'teams-attachments/t/c/1-notiz.wav',
    sizeBytes: 1024,
  };
  const pending = { userMessage: 'Transkribiere das', askedAt: 1, storedAttachments: [upload] };

  it('hands the pending uploads to the turn the click starts', () => {
    assert.deepEqual(topicDecisionUploads(pending, 'Transkribiere das'), [upload]);
  });

  it('hands nothing when the card names another message', () => {
    assert.equal(topicDecisionUploads(pending, 'etwas anderes'), undefined);
  });

  it('hands nothing without a pending entry or without uploads', () => {
    assert.equal(topicDecisionUploads(undefined, 'Transkribiere das'), undefined);
    assert.equal(
      topicDecisionUploads({ userMessage: 'x', askedAt: 1 }, 'x'),
      undefined,
    );
  });
});
