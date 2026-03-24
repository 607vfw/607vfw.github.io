/**
 * Google Apps Script: Dedupe Registrations sheet.
 *
 * - Assumes a header row with these columns:
 *   timestamp, operation_id, operation_name, discord, callsign, role, aircraft,
 *   experience, notes, notify, updated_at
 * - Keeps ONLY the most recently updated row per key:
 *   operation_id + discord + callsign
 *
 * How to use:
 * 1) Open the Google Sheet → Extensions → Apps Script
 * 2) Paste this file into the project
 * 3) Run `dedupeRegistrations()` (authorize once)
 *
 * Notes:
 * - This is intentionally manual/controlled to avoid accidental data loss.
 */

function dedupeRegistrations() {
  const SHEET_NAME = 'Registrations';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) throw new Error(`Sheet not found: ${SHEET_NAME}`);

  const range = sh.getDataRange();
  const values = range.getValues();
  if (values.length < 2) return; // header only

  const header = values[0].map(String);

  const idx = (name) => {
    const i = header.indexOf(name);
    if (i < 0) throw new Error(`Missing column: ${name}`);
    return i;
  };

  const iOp = idx('operation_id');
  const iDiscord = idx('discord');
  const iCallsign = idx('callsign');
  const iUpdated = idx('updated_at');

  const norm = (v) => String(v ?? '').trim();
  const keyOf = (row) => `${norm(row[iOp])}||${norm(row[iDiscord])}||${norm(row[iCallsign])}`;

  // Map key -> { rowIndex, updatedAt }
  const keep = new Map();

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const key = keyOf(row);
    if (!key || key === '||||') continue;

    const updatedRaw = row[iUpdated];
    const updated = updatedRaw instanceof Date ? updatedRaw.getTime() : Date.parse(String(updatedRaw));
    const updatedTs = Number.isFinite(updated) ? updated : 0;

    const existing = keep.get(key);
    if (!existing || updatedTs >= existing.updatedTs) {
      keep.set(key, { rowIndex: r + 1, updatedTs }); // 1-based sheet row
    }
  }

  // Build a set of rows to delete (all data rows except chosen keepers)
  const keepRows = new Set(Array.from(keep.values()).map(x => x.rowIndex));
  const toDelete = [];
  for (let r = 2; r <= values.length; r++) {
    if (!keepRows.has(r)) toDelete.push(r);
  }

  // Delete from bottom to top so indices stay valid
  toDelete.sort((a, b) => b - a);
  for (const r of toDelete) {
    sh.deleteRow(r);
  }

  Logger.log(`Deduped. Kept ${keepRows.size} unique registrations. Deleted ${toDelete.length} rows.`);
}
