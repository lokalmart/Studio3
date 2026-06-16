import XLSX from 'xlsx';
import fs from 'node:fs';

export const EXCEL_CELL_LIMIT = 32767;
export const SAFE_CELL_LIMIT = 32700;

export function readWorkbook(filePath) {
  return XLSX.readFile(filePath, { cellDates: true, dense: false });
}

export function sheetToRows(workbook, sheetName) {
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) return [];
  return XLSX.utils.sheet_to_json(sheet, { defval: '', raw: false });
}

export function rowsToSheet(rows) {
  return XLSX.utils.json_to_sheet(rows || []);
}

export function writeWorkbook(sheets, filePath) {
  const wb = XLSX.utils.book_new();
  for (const item of sheets) {
    const name = sanitizeSheetName(item.name);
    const rows = item.rows?.length ? item.rows : [{ empty: '' }];
    XLSX.utils.book_append_sheet(wb, rowsToSheet(rows), name);
  }
  XLSX.writeFile(wb, filePath, { compression: true });
  return filePath;
}

export function sanitizeSheetName(name) {
  const safe = String(name || 'Sheet')
    .replace(/[\\/?*\[\]:]/g, '_')
    .slice(0, 31);
  return safe || 'Sheet';
}

export function isHelperSheet(sheetName) {
  const s = String(sheetName || '').toLowerCase();
  return [
    'readme', 'readme_import', 'dashboard', 'validation_report', 'prompt', 'prompt_handbook',
    'relationship_map', 'ai_memory_index', 'task_database', 'chatter_project', 'chatter_tasks',
    'task_hierarchy', 'context', 'notes', 'log', 'logs'
  ].some(x => s === x || s.startsWith(`${x}_`));
}

export function inferModelFromRows(sheetName, rows) {
  const rowModel = rows.find(r => r._model || r.model)?.['_model'] || rows.find(r => r._model || r.model)?.model;
  if (rowModel) return String(rowModel).trim();
  if (String(sheetName).includes('.')) return String(sheetName).trim();
  return '';
}

export function safeCellValue(value) {
  if (value === null || value === undefined) return '';
  let text;
  if (value instanceof Date) return value.toISOString();
  if (typeof value === 'object') text = JSON.stringify(value);
  else text = String(value);
  if (text.length > SAFE_CELL_LIMIT) {
    return text.slice(0, SAFE_CELL_LIMIT - 40) + ' …[TRUNCATED_BY_STUDIO2]';
  }
  return text;
}

export function safeExportRow(row) {
  const out = {};
  const truncated = [];
  for (const [key, value] of Object.entries(row || {})) {
    const before = value === null || value === undefined ? '' : (typeof value === 'object' ? JSON.stringify(value) : String(value));
    const after = safeCellValue(value);
    out[key] = after;
    if (before.length > SAFE_CELL_LIMIT) truncated.push(key);
  }
  if (truncated.length) out._studio2_truncated_fields = truncated.join(',');
  return out;
}

export function saveUploadedBuffer(buffer, filepath) {
  fs.writeFileSync(filepath, buffer);
  return filepath;
}
