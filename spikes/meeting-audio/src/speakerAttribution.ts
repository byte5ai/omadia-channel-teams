// Spike S3' / attribution T1 (byte5ai/omadia#584, omadia ADR-0287): attribute
// a transcript segment of the MIXED meeting stream to a participant, from the
// SDK's dominant-speaker changes.
//
// SPIKE CODE, not wired. The transcription model does not diarise, so the
// speaker comes from the call: the ACS `DominantSpeakers` feature reports an
// ordered list whenever it changes; the head of that list is taken as the
// speaker from that moment on. A segment goes to the participant who held
// the floor for most of it, and to nobody when no one held it for at least
// `minShare` of the segment (crosstalk, interjections): an unattributed
// segment is better than a wrong name (ADR-0026).

/** One dominant-speaker change: from `atMs` on, `speakerRawId` holds the
 *  floor (`undefined` = nobody). Offsets share the audio stream's clock. */
export interface FloorChange {
  readonly atMs: number;
  readonly speakerRawId?: string;
}

export interface SegmentWindow {
  readonly startMs: number;
  readonly endMs: number;
}

export interface SegmentAttribution {
  /** Raw id of the participant, absent when no one held the floor long enough. */
  readonly speakerRawId?: string;
  /** Share of the segment the chosen (or, when absent, the leading) speaker held, 0..1. */
  readonly share: number;
}

/** Default: a speaker must hold the floor for 60 % of a segment. */
export const DEFAULT_MIN_FLOOR_SHARE = 0.6;

function sortedChanges(changes: readonly FloorChange[]): readonly FloorChange[] {
  return [...changes].sort((a, b) => a.atMs - b.atMs);
}

function floorTimePerSpeaker(changes: readonly FloorChange[], window: SegmentWindow): ReadonlyMap<string, number> {
  const ordered = sortedChanges(changes);
  const totals = new Map<string, number>();
  ordered.forEach((change, index) => {
    const next = ordered[index + 1];
    const from = Math.max(change.atMs, window.startMs);
    const to = Math.min(next?.atMs ?? Number.POSITIVE_INFINITY, window.endMs);
    if (!change.speakerRawId || to <= from) return;
    totals.set(change.speakerRawId, (totals.get(change.speakerRawId) ?? 0) + (to - from));
  });
  return totals;
}

/** Attribute one segment. Pure; never throws. */
export function attributeSegment(
  window: SegmentWindow,
  changes: readonly FloorChange[],
  minShare: number = DEFAULT_MIN_FLOOR_SHARE,
): SegmentAttribution {
  const length = window.endMs - window.startMs;
  if (!(length > 0)) return { share: 0 };
  let leader: { readonly rawId: string; readonly ms: number } | undefined;
  for (const [rawId, ms] of floorTimePerSpeaker(changes, window)) {
    if (!leader || ms > leader.ms) leader = { rawId, ms };
  }
  if (!leader) return { share: 0 };
  const share = leader.ms / length;
  return share >= minShare ? { speakerRawId: leader.rawId, share } : { share };
}
