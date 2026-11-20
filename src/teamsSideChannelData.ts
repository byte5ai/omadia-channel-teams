/**
 * The side channel for data on a Teams answer (spec 2042 §5.13, criteria
 * 25/35, kernel #2085): rows an agent shows people without its model reading
 * them arrive as `SemanticAnswer.data`, with `textWithoutData` as the same text
 * minus the markdown rendering the kernel put into `text` for text-only
 * connectors.
 *
 * Both fields land with the kernel's S4 stream and are not on every channel
 * SDK this plugin is compiled against, so the shapes are declared here
 * structurally and read feature-detect style: the code compiles before and
 * after S4 lands, and an answer without them renders exactly as it did.
 */

/** Rows a table shows; the rest is a "weitere Zeilen" line. Same number as the kernel's `SIDE_CHANNEL_TEXT_MAX_ROWS`. */
export const SIDE_CHANNEL_MAX_ROWS = 50;
/** A cell is cut here, as the kernel's text rendering does. */
export const SIDE_CHANNEL_MAX_CELL_CHARS = 200;

type Cell = string | number | boolean | null;

/** Structural mirror of the kernel's `TableData`. */
export interface SideChannelTable {
  readonly kind: 'table';
  readonly columns: readonly string[];
  readonly rows: readonly (readonly Cell[])[];
  readonly caption?: string;
}

/** What a channel that renders the data natively takes from an answer. */
export interface SideChannelData {
  readonly tables: readonly SideChannelTable[];
  /** `text` without the data rendering: shown next to the native tables. */
  readonly textWithoutData: string;
}

function isCell(value: unknown): value is Cell {
  return value === null || ['string', 'number', 'boolean'].includes(typeof value);
}

function asTable(payload: unknown): SideChannelTable | undefined {
  if (payload === null || typeof payload !== 'object') return undefined;
  const candidate = payload as Record<string, unknown>;
  if (candidate['kind'] !== 'table') return undefined;
  const { columns, rows, caption } = candidate;
  if (!Array.isArray(columns) || !columns.every((c) => typeof c === 'string')) return undefined;
  if (!Array.isArray(rows) || !rows.every((r) => Array.isArray(r) && r.every(isCell))) return undefined;
  return {
    kind: 'table',
    columns: columns as string[],
    rows: rows as Cell[][],
    ...(typeof caption === 'string' && caption.length > 0 ? { caption } : {}),
  };
}

/**
 * The side-channel data of an answer, or `undefined` when the answer renders
 * from `text` alone: no `data`, no `textWithoutData` (without it the tables are
 * already inside `text` and drawing them again would show them twice), or
 * nothing readable in `data`.
 */
export function readSideChannelData(answer: unknown): SideChannelData | undefined {
  if (answer === null || typeof answer !== 'object' || !('data' in answer)) return undefined;
  const { data, textWithoutData } = answer as { data?: unknown; textWithoutData?: unknown };
  if (!Array.isArray(data) || data.length === 0 || typeof textWithoutData !== 'string') return undefined;
  const tables = data.map(asTable).filter((table): table is SideChannelTable => table !== undefined);
  // A payload this plugin cannot draw is still in `text`: then `text` stays the source.
  if (tables.length !== data.length) return undefined;
  return { tables, textWithoutData };
}

/** The cell as shown: empty for null, one line, cut at {@link SIDE_CHANNEL_MAX_CELL_CHARS}. */
export function sideChannelCellText(value: Cell): string {
  const text = value === null ? '' : String(value).replace(/\r?\n/g, ' ');
  return text.length > SIDE_CHANNEL_MAX_CELL_CHARS ? `${text.slice(0, SIDE_CHANNEL_MAX_CELL_CHARS - 1)}…` : text;
}

/** Characters the tables put on a card: what the card-or-plain-text decision counts. */
export function sideChannelLength(tables: readonly SideChannelTable[]): number {
  return tables.reduce(
    (sum, table) =>
      sum +
      table.columns.join('').length +
      table.rows.slice(0, SIDE_CHANNEL_MAX_ROWS).reduce((rows, row) => rows + row.map(sideChannelCellText).join('').length, 0),
    0,
  );
}
