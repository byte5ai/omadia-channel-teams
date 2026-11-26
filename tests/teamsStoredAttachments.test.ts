import { describe, it } from 'node:test';
import { strict as assert } from 'node:assert';

import { toStoredTurnAttachments } from '@omadia/channel-teams';

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
