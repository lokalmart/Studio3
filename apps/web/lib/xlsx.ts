import * as XLSX from 'xlsx';
import type { SheetData, WorkbookState } from './types';

const CONTEXT_SHEETS = new Set([
  'readme', 'readme_import', 'dashboard', 'validation_report', 'prompt', 'prompt_handbook',
  'relationship_map', 'ai_memory_index', 'task_database', 'chatter_project', 'chatter_tasks',
  'task_hierarchy', 'context', 'notes', 'log', 'logs'
]);

export function inferModel(sheetName: string, rows: Record<string, any>[]) {
  const row = rows.find(r => r._model || r.model);
  const model = row?._model || row?.model;
  if (model) return String(model).trim();
  if (sheetName.includes('.')) return sheetName.trim();
  return '';
}

export function classifySheet(sheetName: string, model: string, columns: string[]): SheetData['kind'] {
  const lower = sheetName.toLowerCase();
  if (!model && (CONTEXT_SHEETS.has(lower) || [...CONTEXT_SHEETS].some(s => lower.startsWith(`${s}_`)))) return 'context';
  if (['product.template', 'product.product'].includes(model)) return 'product';
  if (model === 'res.partner') return 'contact';
  if (model.startsWith('project.')) return 'project';
  if (model === 'knowledge.article') return 'knowledge';
  if (model.startsWith('sale.')) return 'sales';
  if (columns.some(c => ['list_price', 'default_code', 'barcode'].includes(c)) && !model) return 'product';
  return model ? 'dynamic' : 'context';
}

export async function fileToWorkbookState(file: File | Blob, fileName = 'workbook.xlsx'): Promise<WorkbookState> {
  const buffer = await file.arrayBuffer();
  const wb = XLSX.read(buffer, { type: 'array', cellDates: true });
  const sheets: SheetData[] = wb.SheetNames.map(name => {
    const rows = XLSX.utils.sheet_to_json<Record<string, any>>(wb.Sheets[name], { defval: '', raw: false });
    const columns = rows.length ? Array.from(new Set(rows.flatMap(r => Object.keys(r)))) : [];
    const model = inferModel(name, rows);
    return { name, model, rows, columns, kind: classifySheet(name, model, columns) };
  });
  return { fileName, sheets };
}

export function workbookStateToBlob(state: WorkbookState): Blob {
  const wb = XLSX.utils.book_new();
  state.sheets.forEach(sheet => {
    const rows = sheet.rows.length ? sheet.rows : [{ empty: '' }];
    const ws = XLSX.utils.json_to_sheet(rows);
    XLSX.utils.book_append_sheet(wb, ws, sanitizeSheetName(sheet.name));
  });
  const out = XLSX.write(wb, { type: 'array', bookType: 'xlsx', compression: true });
  return new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}

export function sanitizeSheetName(name: string) {
  return (name || 'Sheet').replace(/[\\/?*\[\]:]/g, '_').slice(0, 31) || 'Sheet';
}

export function updateCell(state: WorkbookState, sheetIndex: number, rowIndex: number, key: string, value: any): WorkbookState {
  const copy = structuredClone(state);
  copy.sheets[sheetIndex].rows[rowIndex][key] = value;
  if (!copy.sheets[sheetIndex].columns.includes(key)) copy.sheets[sheetIndex].columns.push(key);
  return copy;
}

export function addColumn(state: WorkbookState, sheetIndex: number, key: string): WorkbookState {
  const copy = structuredClone(state);
  if (!key || copy.sheets[sheetIndex].columns.includes(key)) return copy;
  copy.sheets[sheetIndex].columns.push(key);
  copy.sheets[sheetIndex].rows.forEach(r => { r[key] = r[key] ?? ''; });
  return copy;
}

export function localValidate(sheet: SheetData) {
  const errors: string[] = [];
  const warnings: string[] = [];
  if (sheet.kind !== 'context' && !sheet.model) errors.push('Sheet data tidak punya _model dan nama sheet bukan model Odoo.');
  sheet.rows.forEach((row, idx) => {
    const line = idx + 2;
    if (sheet.kind !== 'context' && !(row._model || sheet.model)) errors.push(`Row ${line}: _model kosong.`);
    Object.entries(row).forEach(([key, value]) => {
      const text = value == null ? '' : String(value);
      if (text.length > 32700) warnings.push(`Row ${line}: ${key} melebihi batas Excel dan akan dipotong saat export.`);
    });
  });
  return { errors, warnings };
}
