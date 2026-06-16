'use client';

import React, { useEffect, useMemo, useState } from 'react';
import * as XLSX from 'xlsx';

type Conn = {
  url: string;
  db: string;
  username: string;
  password: string;
};

type Row = Record<string, any>;
type SheetState = { name: string; model: string; rows: Row[]; columns: string[]; kind: EditorKind; helper: boolean };
type EditorKind = 'product' | 'contact' | 'project' | 'knowledge' | 'sales' | 'dynamic' | 'helper';
type OdooField = { string?: string; type?: string; required?: boolean; readonly?: boolean; relation?: string; selection?: Array<[string, string]> };

type LogItem = { time: string; level: 'info' | 'ok' | 'warn' | 'error'; message: string; detail?: any };

const STORAGE_KEY = 'studio2_v9_settings';
const EXCEL_LIMIT = 32767;
const SAFE_LIMIT = 32000;
const HELPER_SHEETS = new Set(['readme', 'readme_import', 'dashboard', 'validation_report', 'prompt', 'ai_memory_index', 'task_database', 'relationship_map', 'chatter_project', 'chatter_tasks', 'task_hierarchy']);
const META_COLUMNS = new Set(['_model', '__action', '_external_id', 'external_id', 'id', 'x_studio2_odoo_id', '__rownum__', '_studio2_truncated_fields', '_studio2_note', '_studio2_error']);

const defaultConn: Conn = { url: '', db: '', username: '', password: '' };

function now() {
  return new Date().toLocaleTimeString('id-ID', { hour12: false });
}

function detectModel(sheetName: string, rows: Row[]) {
  const fromRow = rows.find(r => r._model)?._model;
  if (fromRow) return String(fromRow).trim();
  const normalized = sheetName.toLowerCase().trim();
  const known: Record<string, string> = {
    contacts: 'res.partner',
    contact: 'res.partner',
    partner: 'res.partner',
    partners: 'res.partner',
    product: 'product.template',
    products: 'product.template',
    project: 'project.project',
    projects: 'project.project',
    task: 'project.task',
    tasks: 'project.task',
    knowledge: 'knowledge.article',
    articles: 'knowledge.article',
    sales: 'sale.order',
    orders: 'sale.order'
  };
  return known[normalized] || sheetName;
}

function detectKind(model: string, sheetName: string): EditorKind {
  const m = String(model || '').toLowerCase();
  const s = String(sheetName || '').toLowerCase();
  if (HELPER_SHEETS.has(s) || m.includes('helper') || m.includes('readme')) return 'helper';
  if (m === 'res.partner') return 'contact';
  if (m === 'product.template' || m === 'product.product' || m === 'product.category' || m === 'product.public.category') return 'product';
  if (m.startsWith('project.')) return 'project';
  if (m === 'knowledge.article') return 'knowledge';
  if (m.startsWith('sale.')) return 'sales';
  return 'dynamic';
}

function makeColumns(rows: Row[]) {
  const set = new Set<string>();
  rows.forEach(row => Object.keys(row).forEach(k => set.add(k)));
  return Array.from(set);
}

function sanitizeWorkbookValue(value: any) {
  if (value === null || value === undefined) return '';
  if (Array.isArray(value) || (typeof value === 'object' && value !== null)) return JSON.stringify(value).slice(0, SAFE_LIMIT);
  const text = String(value);
  return text.length > EXCEL_LIMIT ? text.slice(0, SAFE_LIMIT) : value;
}

function normalizeRowsForSheet(rows: Row[]) {
  return rows.map((row, idx) => {
    const out: Row = { ...row };
    if (!out.__rownum__) out.__rownum__ = idx + 2;
    return out;
  });
}

function rowsToSheetState(name: string, rawRows: Row[]): SheetState {
  const rows = normalizeRowsForSheet(rawRows);
  const model = detectModel(name, rows);
  const kind = detectKind(model, name);
  const helper = kind === 'helper' || HELPER_SHEETS.has(name.toLowerCase());
  const columns = makeColumns(rows);
  if (!columns.includes('_model') && !helper) columns.unshift('_model');
  if (!columns.includes('__action') && !helper) columns.unshift('__action');
  if (!columns.includes('_external_id') && !helper) columns.unshift('_external_id');
  return { name, model, rows, columns, kind, helper };
}

function chunk<T>(arr: T[], size: number) {
  const out: T[][] = [];
  for (let i = 0; i < arr.length; i += size) out.push(arr.slice(i, i + size));
  return out;
}

function parseCsvFields(text: string) {
  return text.split(',').map(x => x.trim()).filter(Boolean);
}

function importantFields(kind: EditorKind, model: string) {
  if (kind === 'contact') return ['name', 'display_name', 'phone', 'mobile', 'email', 'street', 'city', 'supplier_rank', 'customer_rank', 'is_company'];
  if (kind === 'product') return ['name', 'default_code', 'barcode', 'list_price', 'standard_price', 'categ_id', 'public_categ_ids', 'image_1920', 'description_sale', 'sale_ok', 'purchase_ok'];
  if (kind === 'project') return ['name', 'project_id', 'parent_id', 'stage_id', 'user_id', 'user_ids', 'partner_id', 'date_deadline', 'description', 'priority', 'sequence'];
  if (kind === 'knowledge') return ['name', 'body', 'body_html', 'parent_id', 'category'];
  if (kind === 'sales') return ['name', 'partner_id', 'date_order', 'state', 'amount_total', 'invoice_status'];
  return ['display_name', 'name', 'create_date', 'write_date'];
}

export default function HomePage() {
  const [mode, setMode] = useState<'import' | 'export'>('import');
  const [settingsOpen, setSettingsOpen] = useState(false);
  const [conn, setConn] = useState<Conn>(defaultConn);
  const [logs, setLogs] = useState<LogItem[]>([]);
  const [busy, setBusy] = useState(false);

  const [sheets, setSheets] = useState<SheetState[]>([]);
  const [activeSheetIndex, setActiveSheetIndex] = useState(0);
  const activeSheet = sheets[activeSheetIndex];
  const [schema, setSchema] = useState<Record<string, OdooField> | null>(null);
  const [schemaModel, setSchemaModel] = useState('');
  const [batchSize, setBatchSize] = useState(20);
  const [selectedRows, setSelectedRows] = useState<Record<number, boolean>>({});

  const [exportModel, setExportModel] = useState('res.partner');
  const [exportFields, setExportFields] = useState('name,display_name,email,phone,mobile,street,city,customer_rank,supplier_rank');
  const [exportDomain, setExportDomain] = useState('[]');
  const [scanRecords, setScanRecords] = useState<Row[]>([]);
  const [scanCount, setScanCount] = useState(0);
  const [scanOffset, setScanOffset] = useState(0);
  const [selectedIds, setSelectedIds] = useState<Record<number, boolean>>({});

  useEffect(() => {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (raw) setConn({ ...defaultConn, ...JSON.parse(raw) });
    } catch {}
  }, []);

  useEffect(() => {
    try { localStorage.setItem(STORAGE_KEY, JSON.stringify(conn)); } catch {}
  }, [conn]);

  useEffect(() => {
    setSchema(null);
    setSchemaModel('');
    setSelectedRows({});
  }, [activeSheetIndex]);

  function addLog(level: LogItem['level'], message: string, detail?: any) {
    setLogs(prev => [{ time: now(), level, message, detail }, ...prev].slice(0, 160));
  }

  async function callOdoo(payload: Row) {
    const res = await fetch('/api/odoo', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ ...payload, connection: conn })
    });
    const data = await res.json();
    if (!data.ok) throw new Error(data.error || 'API error');
    return data;
  }

  async function testConnection() {
    setBusy(true);
    try {
      const data = await callOdoo({ action: 'test' });
      addLog('ok', `Koneksi Odoo berhasil. UID ${data.uid}, kontak ${data.partner_count}.`);
    } catch (e: any) {
      addLog('error', `Koneksi gagal: ${e.message}`);
    } finally {
      setBusy(false);
    }
  }

  async function loadSchema(model?: string) {
    const target = model || activeSheet?.model || exportModel;
    if (!target) return;
    setBusy(true);
    try {
      const data = await callOdoo({ action: 'schema', model: target });
      setSchema(data.fields);
      setSchemaModel(target);
      addLog('ok', `Schema dimuat: ${target} (${Object.keys(data.fields || {}).length} field).`);
    } catch (e: any) {
      addLog('error', `Gagal memuat schema ${target}: ${e.message}`);
    } finally {
      setBusy(false);
    }
  }

  async function handleFile(file: File) {
    setBusy(true);
    try {
      const buffer = await file.arrayBuffer();
      const wb = XLSX.read(buffer, { type: 'array', cellDates: false });
      const parsed: SheetState[] = wb.SheetNames.map(name => {
        const rows = XLSX.utils.sheet_to_json<Row>(wb.Sheets[name], { defval: '' });
        return rowsToSheetState(name, rows);
      });
      setSheets(parsed);
      setActiveSheetIndex(0);
      addLog('ok', `XLSX dibaca: ${parsed.length} sheet dari ${file.name}.`);
    } catch (e: any) {
      addLog('error', `Gagal membaca XLSX: ${e.message}`);
    } finally {
      setBusy(false);
    }
  }

  function updateCell(rowIndex: number, col: string, value: any) {
    setSheets(prev => prev.map((sheet, si) => {
      if (si !== activeSheetIndex) return sheet;
      const rows = sheet.rows.map((row, ri) => ri === rowIndex ? { ...row, [col]: value } : row);
      const columns = sheet.columns.includes(col) ? sheet.columns : [...sheet.columns, col];
      return { ...sheet, rows, columns };
    }));
  }

  function addColumn() {
    const name = prompt('Nama kolom baru, misalnya x_review_status atau vendor_external_id');
    if (!name || !activeSheet) return;
    setSheets(prev => prev.map((sheet, si) => si === activeSheetIndex ? { ...sheet, columns: sheet.columns.includes(name) ? sheet.columns : [...sheet.columns, name] } : sheet));
  }

  function addRow() {
    if (!activeSheet) return;
    setSheets(prev => prev.map((sheet, si) => {
      if (si !== activeSheetIndex) return sheet;
      const row: Row = { _model: sheet.model, __action: 'upsert', _external_id: '', __rownum__: sheet.rows.length + 2 };
      return { ...sheet, rows: [...sheet.rows, row] };
    }));
  }

  function validateActiveSheet() {
    if (!activeSheet) return [];
    const issues: string[] = [];
    if (activeSheet.helper) return issues;
    if (!activeSheet.model) issues.push('Model belum jelas. Isi _model atau ubah nama sheet sesuai model Odoo.');
    activeSheet.rows.forEach((row, i) => {
      if (!row._model && !activeSheet.model) issues.push(`Row ${i + 2}: _model kosong.`);
      if (!row.__action) issues.push(`Row ${i + 2}: __action kosong; default tetap upsert.`);
      if (!row._external_id && !row.x_studio2_odoo_id && !row.id) issues.push(`Row ${i + 2}: _external_id kosong. Aman untuk create, kurang aman untuk update.`);
    });
    if (schema && schemaModel === activeSheet.model) {
      const unknown = activeSheet.columns.filter(c => !META_COLUMNS.has(c) && !c.endsWith('_external_id') && !c.endsWith('_external_ids') && !schema[c]);
      if (unknown.length) issues.push(`Kolom tidak dikenal di ${activeSheet.model}: ${unknown.slice(0, 30).join(', ')}${unknown.length > 30 ? '...' : ''}`);
      const required = Object.entries(schema).filter(([, v]) => v.required && !v.readonly).map(([k]) => k);
      const missing = required.filter(f => !activeSheet.columns.includes(f) && !activeSheet.columns.includes(`${f}_external_id`));
      if (missing.length) issues.push(`Field wajib belum ada di sheet: ${missing.slice(0, 20).join(', ')}`);
    }
    return issues;
  }

  const issues = useMemo(() => validateActiveSheet(), [activeSheet, schema, schemaModel]);

  async function importActiveSheet() {
    if (!activeSheet) return;
    if (activeSheet.helper) {
      addLog('warn', 'Sheet helper/context tidak diimport. Pilih sheet model Odoo.');
      return;
    }
    const chosen = activeSheet.rows.filter((_, idx) => Object.keys(selectedRows).length ? selectedRows[idx] : true);
    if (!chosen.length) {
      addLog('warn', 'Tidak ada row dipilih untuk import.');
      return;
    }
    const parts = chunk(chosen, Math.max(1, Math.min(50, batchSize)));
    setBusy(true);
    addLog('info', `Mulai import ${chosen.length} row dari ${activeSheet.name} dalam ${parts.length} batch.`);
    let totalCreated = 0, totalUpdated = 0, totalFailed = 0;
    for (let i = 0; i < parts.length; i++) {
      try {
        const data = await callOdoo({ action: 'import_batch', model: activeSheet.model, rows: parts[i] });
        totalCreated += Number(data.created || 0);
        totalUpdated += Number(data.updated || 0);
        totalFailed += Number(data.failed || 0);
        addLog(data.failed ? 'warn' : 'ok', `Batch ${i + 1}/${parts.length}: created ${data.created}, updated ${data.updated}, failed ${data.failed}.`, data.results);
        if (data.failed) break;
      } catch (e: any) {
        totalFailed += parts[i].length;
        addLog('error', `Batch ${i + 1} gagal: ${e.message}`);
        break;
      }
    }
    addLog(totalFailed ? 'warn' : 'ok', `Import selesai: created ${totalCreated}, updated ${totalUpdated}, failed ${totalFailed}.`);
    setBusy(false);
  }

  function downloadWorkbook(fileName = 'studio2_export.xlsx') {
    if (!sheets.length) return;
    const wb = XLSX.utils.book_new();
    sheets.forEach(sheet => {
      const cleanRows = sheet.rows.map(row => {
        const out: Row = {};
        sheet.columns.forEach(col => out[col] = sanitizeWorkbookValue(row[col]));
        return out;
      });
      const ws = XLSX.utils.json_to_sheet(cleanRows.length ? cleanRows : [{}], { header: sheet.columns.length ? sheet.columns : undefined });
      XLSX.utils.book_append_sheet(wb, ws, sheet.name.slice(0, 31) || 'Sheet1');
    });
    XLSX.writeFile(wb, fileName);
    addLog('ok', `XLSX diunduh: ${fileName}`);
  }

  async function scanModel(reset = true) {
    setBusy(true);
    try {
      let domain: any[] = [];
      try { domain = JSON.parse(exportDomain || '[]'); } catch { throw new Error('Domain harus JSON array, contoh: [] atau [["is_company","=",true]]'); }
      const offset = reset ? 0 : scanOffset;
      const kind = detectKind(exportModel, exportModel);
      const fields = Array.from(new Set([...importantFields(kind, exportModel), ...parseCsvFields(exportFields)])).filter(Boolean);
      const data = await callOdoo({ action: 'record_scan', model: exportModel, fields, domain, offset, limit: 80 });
      setScanCount(data.count || 0);
      setScanOffset(offset + (data.records?.length || 0));
      setScanRecords(prev => reset ? data.records : [...prev, ...data.records]);
      if (reset) setSelectedIds({});
      addLog('ok', `Scan ${exportModel}: ${data.records.length} record dimuat dari ${data.count}.`);
    } catch (e: any) {
      addLog('error', `Scan gagal: ${e.message}`);
    } finally {
      setBusy(false);
    }
  }

  async function exportSelectedRecords() {
    const ids = Object.entries(selectedIds).filter(([, v]) => v).map(([id]) => Number(id));
    if (!ids.length) {
      addLog('warn', 'Belum ada record dipilih untuk export.');
      return;
    }
    setBusy(true);
    try {
      const fields = parseCsvFields(exportFields);
      const data = await callOdoo({ action: 'export_records', model: exportModel, ids, fields });
      const sheet = rowsToSheetState(data.sheet || exportModel, data.rows || []);
      setSheets([sheet]);
      setActiveSheetIndex(0);
      setMode('import');
      addLog('ok', `Export ${ids.length} record dari ${exportModel} masuk editor XLSX.`);
    } catch (e: any) {
      addLog('error', `Export gagal: ${e.message}`);
    } finally {
      setBusy(false);
    }
  }

  async function exportProject() {
    const id = prompt('Masukkan ID project Odoo, contoh: 73');
    if (!id) return;
    setBusy(true);
    try {
      const data = await callOdoo({ action: 'export_project', project_id: Number(id) });
      const nextSheets = Object.entries(data.sheets || {}).map(([name, rows]) => rowsToSheetState(name, rows as Row[]));
      setSheets(nextSheets);
      setActiveSheetIndex(0);
      setMode('import');
      addLog('ok', `Project ${id} diexport ke ${nextSheets.length} sheet dan masuk editor.`);
    } catch (e: any) {
      addLog('error', `Export project gagal: ${e.message}`);
    } finally {
      setBusy(false);
    }
  }

  const visibleColumns = useMemo(() => {
    if (!activeSheet) return [];
    const priority = ['_model', '__action', '_external_id', 'x_studio2_odoo_id', ...importantFields(activeSheet.kind, activeSheet.model)];
    const set = new Set(activeSheet.columns);
    return [...priority.filter(c => set.has(c)), ...activeSheet.columns.filter(c => !priority.includes(c) && c !== '__rownum__')];
  }, [activeSheet]);

  const connectionOk = conn.url && conn.db && conn.username && conn.password;

  return (
    <main className="min-h-screen p-4 md:p-6">
      <section className="mx-auto max-w-7xl space-y-4">
        <header className="rounded-[28px] border border-slate-200 bg-white/85 p-4 shadow-soft backdrop-blur md:p-5">
          <div className="flex flex-col gap-4 md:flex-row md:items-center md:justify-between">
            <div>
              <div className="flex flex-wrap items-center gap-2">
                <span className="badge bg-slate-900 text-white">Studio2 v9</span>
                <span className="badge bg-emerald-50 text-emerald-700">Vercel-only</span>
                <span className="badge bg-blue-50 text-blue-700">Browser-side XLSX</span>
              </div>
              <h1 className="mt-3 text-2xl font-black tracking-tight md:text-3xl">Lokalmart Odoo XLSX Studio</h1>
              <p className="mt-1 max-w-3xl text-sm text-slate-600">Import/export dibuat bertahap agar cocok untuk Vercel gratis: preview dan editor di browser, API hanya mengirim batch kecil ke Odoo.</p>
            </div>
            <div className="flex flex-wrap gap-2">
              <button className={`btn ${mode === 'import' ? 'btn-primary' : 'btn-soft'}`} onClick={() => setMode('import')}>Import / Editor</button>
              <button className={`btn ${mode === 'export' ? 'btn-primary' : 'btn-soft'}`} onClick={() => setMode('export')}>Export</button>
              <button className="btn btn-soft" onClick={() => setSettingsOpen(!settingsOpen)}>⚙ Koneksi</button>
            </div>
          </div>
          {settingsOpen && (
            <div className="mt-5 grid gap-3 rounded-3xl border border-slate-200 bg-slate-50 p-4 md:grid-cols-2 lg:grid-cols-5">
              <input className="small-input" placeholder="Odoo URL: https://xxx.odoo.com" value={conn.url} onChange={e => setConn({ ...conn, url: e.target.value })} />
              <input className="small-input" placeholder="Database" value={conn.db} onChange={e => setConn({ ...conn, db: e.target.value })} />
              <input className="small-input" placeholder="Email/Username" value={conn.username} onChange={e => setConn({ ...conn, username: e.target.value })} />
              <input className="small-input" placeholder="Password/API key" type="password" value={conn.password} onChange={e => setConn({ ...conn, password: e.target.value })} />
              <button className="btn btn-green" disabled={busy || !connectionOk} onClick={testConnection}>{busy ? 'Memproses...' : 'Tes Koneksi'}</button>
            </div>
          )}
        </header>

        <div className="grid gap-4 lg:grid-cols-[320px_1fr]">
          <aside className="space-y-4">
            {mode === 'import' ? (
              <ImportPanel busy={busy} sheets={sheets} activeSheetIndex={activeSheetIndex} setActiveSheetIndex={setActiveSheetIndex} onFile={handleFile} addRow={addRow} addColumn={addColumn} download={() => downloadWorkbook('studio2_edited.xlsx')} importActiveSheet={importActiveSheet} batchSize={batchSize} setBatchSize={setBatchSize} loadSchema={() => loadSchema()} />
            ) : (
              <ExportPanel busy={busy} model={exportModel} setModel={setExportModel} fields={exportFields} setFields={setExportFields} domain={exportDomain} setDomain={setExportDomain} scan={() => scanModel(true)} loadMore={() => scanModel(false)} exportSelected={exportSelectedRecords} exportProject={exportProject} scanCount={scanCount} scanRecords={scanRecords} scanOffset={scanOffset} selectedIds={selectedIds} setSelectedIds={setSelectedIds} loadSchema={() => loadSchema(exportModel)} />
            )}
            <LogPanel logs={logs} />
          </aside>

          <section className="min-w-0 rounded-[28px] border border-slate-200 bg-white/90 p-4 shadow-soft backdrop-blur">
            {mode === 'export' ? (
              <RecordPicker records={scanRecords} selectedIds={selectedIds} setSelectedIds={setSelectedIds} scanCount={scanCount} />
            ) : activeSheet ? (
              <Editor
                sheet={activeSheet}
                columns={visibleColumns}
                schema={schemaModel === activeSheet.model ? schema : null}
                issues={issues}
                selectedRows={selectedRows}
                setSelectedRows={setSelectedRows}
                updateCell={updateCell}
              />
            ) : (
              <EmptyState />
            )}
          </section>
        </div>
      </section>
    </main>
  );
}

function ImportPanel(props: {
  busy: boolean; sheets: SheetState[]; activeSheetIndex: number; setActiveSheetIndex: (n: number) => void; onFile: (f: File) => void;
  addRow: () => void; addColumn: () => void; download: () => void; importActiveSheet: () => void; batchSize: number; setBatchSize: (n: number) => void; loadSchema: () => void;
}) {
  return (
    <div className="rounded-[28px] border border-slate-200 bg-white/90 p-4 shadow-soft">
      <h2 className="text-lg font-black">Import / Editor</h2>
      <p className="mt-1 text-sm text-slate-600">Upload XLSX, edit, validasi, lalu import per batch kecil.</p>
      <label className="mt-4 block cursor-pointer rounded-3xl border-2 border-dashed border-slate-300 bg-slate-50 p-5 text-center hover:bg-slate-100">
        <input className="hidden" type="file" accept=".xlsx,.xls" onChange={e => e.target.files?.[0] && props.onFile(e.target.files[0])} />
        <div className="text-3xl">📄</div>
        <div className="mt-2 font-bold">Pilih XLSX</div>
        <div className="text-xs text-slate-500">Parsing dilakukan di browser</div>
      </label>
      {props.sheets.length > 0 && (
        <div className="mt-4 space-y-2">
          <div className="text-xs font-black uppercase text-slate-500">Sheet</div>
          <div className="max-h-56 space-y-2 overflow-auto pr-1 studio-scroll">
            {props.sheets.map((sheet, i) => (
              <button key={sheet.name} onClick={() => props.setActiveSheetIndex(i)} className={`w-full rounded-2xl border p-3 text-left ${i === props.activeSheetIndex ? 'border-slate-900 bg-slate-900 text-white' : 'border-slate-200 bg-white hover:bg-slate-50'}`}>
                <div className="flex items-center justify-between gap-2">
                  <div className="truncate font-bold">{sheet.name}</div>
                  <span className={`badge ${sheet.helper ? 'bg-amber-100 text-amber-800' : 'bg-emerald-100 text-emerald-800'}`}>{sheet.rows.length}</span>
                </div>
                <div className={`mt-1 text-xs ${i === props.activeSheetIndex ? 'text-slate-200' : 'text-slate-500'}`}>{sheet.helper ? 'context/helper' : sheet.model}</div>
              </button>
            ))}
          </div>
          <div className="grid grid-cols-2 gap-2 pt-2">
            <button className="btn btn-soft" onClick={props.addRow}>+ Row</button>
            <button className="btn btn-soft" onClick={props.addColumn}>+ Kolom</button>
            <button className="btn btn-soft" onClick={props.loadSchema}>Schema</button>
            <button className="btn btn-soft" onClick={props.download}>Download</button>
          </div>
          <div className="pt-2">
            <label className="text-xs font-bold text-slate-500">Batch size</label>
            <input className="small-input mt-1 w-full" type="number" min={1} max={50} value={props.batchSize} onChange={e => props.setBatchSize(Number(e.target.value))} />
          </div>
          <button className="btn btn-green mt-2 w-full" disabled={props.busy} onClick={props.importActiveSheet}>Import Sheet Aktif</button>
        </div>
      )}
    </div>
  );
}

function ExportPanel(props: {
  busy: boolean; model: string; setModel: (s: string) => void; fields: string; setFields: (s: string) => void; domain: string; setDomain: (s: string) => void;
  scan: () => void; loadMore: () => void; exportSelected: () => void; exportProject: () => void; scanCount: number; scanOffset: number; scanRecords: Row[];
  selectedIds: Record<number, boolean>; setSelectedIds: (x: Record<number, boolean>) => void; loadSchema: () => void;
}) {
  const selectedCount = Object.values(props.selectedIds).filter(Boolean).length;
  return (
    <div className="rounded-[28px] border border-slate-200 bg-white/90 p-4 shadow-soft">
      <h2 className="text-lg font-black">Export</h2>
      <p className="mt-1 text-sm text-slate-600">Scan record dulu, pilih record dan field, lalu hasilnya masuk editor XLSX.</p>
      <div className="mt-4 space-y-3">
        <div>
          <label className="text-xs font-bold text-slate-500">Model Odoo</label>
          <input className="small-input mt-1 w-full" value={props.model} onChange={e => props.setModel(e.target.value)} placeholder="res.partner" />
        </div>
        <div>
          <label className="text-xs font-bold text-slate-500">Fields export, pisahkan koma</label>
          <textarea className="small-input mt-1 h-24 w-full" value={props.fields} onChange={e => props.setFields(e.target.value)} />
        </div>
        <div>
          <label className="text-xs font-bold text-slate-500">Domain JSON</label>
          <input className="small-input mt-1 w-full" value={props.domain} onChange={e => props.setDomain(e.target.value)} placeholder="[]" />
        </div>
        <div className="grid grid-cols-2 gap-2">
          <button className="btn btn-primary" disabled={props.busy} onClick={props.scan}>Scan</button>
          <button className="btn btn-soft" disabled={props.busy} onClick={props.loadSchema}>Schema</button>
          <button className="btn btn-soft" disabled={props.busy || props.scanRecords.length >= props.scanCount} onClick={props.loadMore}>Load More</button>
          <button className="btn btn-soft" disabled={props.busy} onClick={props.exportProject}>Export Project</button>
        </div>
        <div className="rounded-2xl bg-slate-50 p-3 text-sm text-slate-600">Terseleksi <b>{selectedCount}</b> record dari {props.scanRecords.length}/{props.scanCount} yang sudah discan.</div>
        <button className="btn btn-green w-full" disabled={props.busy || selectedCount === 0} onClick={props.exportSelected}>Export Record Terpilih</button>
      </div>
    </div>
  );
}

function Editor(props: {
  sheet: SheetState; columns: string[]; schema: Record<string, OdooField> | null; issues: string[]; selectedRows: Record<number, boolean>; setSelectedRows: (x: Record<number, boolean>) => void; updateCell: (row: number, col: string, value: any) => void;
}) {
  const { sheet } = props;
  return (
    <div className="space-y-4">
      <div className="flex flex-col gap-3 md:flex-row md:items-start md:justify-between">
        <div>
          <div className="flex flex-wrap gap-2">
            <span className="badge bg-slate-900 text-white">{labelKind(sheet.kind)}</span>
            <span className="badge bg-slate-100 text-slate-700">{sheet.model}</span>
            <span className="badge bg-blue-50 text-blue-700">{sheet.rows.length} rows</span>
            <span className="badge bg-purple-50 text-purple-700">{props.columns.length} columns</span>
          </div>
          <h2 className="mt-2 text-xl font-black">{sheet.name}</h2>
          <p className="text-sm text-slate-600">Editor mengikuti model Odoo. Sheet helper/context tidak akan diimport.</p>
        </div>
      </div>

      {sheet.helper && (
        <div className="rounded-3xl border border-amber-200 bg-amber-50 p-4 text-sm text-amber-900">Sheet ini dibaca sebagai context/helper. Isinya bisa diedit dan didownload, tapi tidak dikirim ke Odoo.</div>
      )}

      {props.issues.length > 0 && (
        <div className="rounded-3xl border border-rose-200 bg-rose-50 p-4 text-sm text-rose-900">
          <div className="font-black">Validasi menemukan catatan:</div>
          <ul className="mt-2 list-disc space-y-1 pl-5">
            {props.issues.slice(0, 10).map((x, i) => <li key={i}>{x}</li>)}
          </ul>
          {props.issues.length > 10 && <div className="mt-2">+ {props.issues.length - 10} catatan lain.</div>}
        </div>
      )}

      <div className="overflow-auto rounded-3xl border border-slate-200 studio-scroll" style={{ maxHeight: '68vh' }}>
        <table className="min-w-full border-collapse text-sm">
          <thead className="sticky top-0 z-10 bg-slate-100">
            <tr>
              <th className="w-12 border-b border-r border-slate-200 p-2 text-left">
                <input type="checkbox" onChange={e => {
                  const next: Record<number, boolean> = {};
                  if (e.target.checked) sheet.rows.forEach((_, i) => next[i] = true);
                  props.setSelectedRows(next);
                }} />
              </th>
              <th className="border-b border-r border-slate-200 p-2 text-left text-xs font-black uppercase text-slate-500">#</th>
              {props.columns.map(col => (
                <th key={col} className={`whitespace-nowrap border-b border-r border-slate-200 p-2 text-left text-xs font-black uppercase ${props.schema && !META_COLUMNS.has(col) && !props.schema[col] && !col.endsWith('_external_id') && !col.endsWith('_external_ids') ? 'bg-rose-50 text-rose-700' : 'text-slate-500'}`}>
                  <div>{col}</div>
                  {props.schema?.[col] && <div className="mt-1 normal-case text-slate-400">{props.schema[col].type}{props.schema[col].relation ? ` → ${props.schema[col].relation}` : ''}</div>}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {sheet.rows.slice(0, 500).map((row, ri) => (
              <tr key={ri} className={props.selectedRows[ri] ? 'bg-blue-50/60' : 'hover:bg-slate-50'}>
                <td className="border-r border-t border-slate-200 p-2"><input type="checkbox" checked={Boolean(props.selectedRows[ri])} onChange={e => props.setSelectedRows({ ...props.selectedRows, [ri]: e.target.checked })} /></td>
                <td className="border-r border-t border-slate-200 p-2 text-xs text-slate-500">{ri + 1}</td>
                {props.columns.map(col => (
                  <td key={col} className="min-w-[180px] border-r border-t border-slate-200 align-top">
                    <CellInput value={row[col] ?? ''} field={props.schema?.[col]} onChange={value => props.updateCell(ri, col, value)} />
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
        {sheet.rows.length > 500 && <div className="p-4 text-center text-sm text-slate-500">Editor menampilkan 500 row pertama agar browser tetap ringan. Import tetap memakai row yang ada di state.</div>}
      </div>
    </div>
  );
}

function CellInput({ value, field, onChange }: { value: any; field?: OdooField; onChange: (v: any) => void }) {
  if (field?.type === 'boolean') {
    return <div className="flex h-full items-center p-2"><input type="checkbox" checked={['true', '1', 'yes', true].includes(value)} onChange={e => onChange(e.target.checked)} /></div>;
  }
  if (field?.type === 'selection' && field.selection) {
    return (
      <select className="cell-input" value={value ?? ''} onChange={e => onChange(e.target.value)}>
        <option value="">—</option>
        {field.selection.map(([k, label]) => <option key={k} value={k}>{label}</option>)}
      </select>
    );
  }
  if (field?.type === 'text' || field?.type === 'html' || String(value || '').length > 120) {
    return <textarea className="cell-input min-h-[52px]" value={value ?? ''} onChange={e => onChange(e.target.value)} />;
  }
  if (field?.type === 'integer' || field?.type === 'float' || field?.type === 'monetary') {
    return <input className="cell-input" type="number" value={value ?? ''} onChange={e => onChange(e.target.value)} />;
  }
  if (field?.type === 'date') {
    return <input className="cell-input" type="date" value={String(value || '').slice(0, 10)} onChange={e => onChange(e.target.value)} />;
  }
  return <input className="cell-input" value={value ?? ''} onChange={e => onChange(e.target.value)} />;
}

function RecordPicker({ records, selectedIds, setSelectedIds, scanCount }: { records: Row[]; selectedIds: Record<number, boolean>; setSelectedIds: (x: Record<number, boolean>) => void; scanCount: number }) {
  const cols = makeColumns(records).filter(c => c !== 'id').slice(0, 10);
  return (
    <div className="space-y-4">
      <div>
        <div className="flex flex-wrap gap-2">
          <span className="badge bg-slate-900 text-white">Record Picker</span>
          <span className="badge bg-blue-50 text-blue-700">{records.length}/{scanCount}</span>
        </div>
        <h2 className="mt-2 text-xl font-black">Pilih record yang akan diexport</h2>
        <p className="text-sm text-slate-600">Export hanya record yang dicentang supaya XLSX tidak terlalu besar dan tidak menabrak limit Vercel.</p>
      </div>
      {!records.length ? (
        <div className="rounded-3xl border border-dashed border-slate-300 bg-slate-50 p-10 text-center text-slate-500">Belum ada record. Klik Scan di panel kiri.</div>
      ) : (
        <div className="overflow-auto rounded-3xl border border-slate-200 studio-scroll" style={{ maxHeight: '70vh' }}>
          <table className="min-w-full text-sm">
            <thead className="sticky top-0 bg-slate-100">
              <tr>
                <th className="border-b border-r border-slate-200 p-2"><input type="checkbox" onChange={e => {
                  const next: Record<number, boolean> = {};
                  if (e.target.checked) records.forEach(r => next[Number(r.id)] = true);
                  setSelectedIds(next);
                }} /></th>
                <th className="border-b border-r border-slate-200 p-2 text-left">ID</th>
                {cols.map(c => <th key={c} className="border-b border-r border-slate-200 p-2 text-left">{c}</th>)}
              </tr>
            </thead>
            <tbody>
              {records.map(row => (
                <tr key={row.id} className={selectedIds[Number(row.id)] ? 'bg-blue-50' : 'hover:bg-slate-50'}>
                  <td className="border-r border-t border-slate-200 p-2"><input type="checkbox" checked={Boolean(selectedIds[Number(row.id)])} onChange={e => setSelectedIds({ ...selectedIds, [Number(row.id)]: e.target.checked })} /></td>
                  <td className="border-r border-t border-slate-200 p-2 font-bold">{row.id}</td>
                  {cols.map(c => <td key={c} className="max-w-sm truncate border-r border-t border-slate-200 p-2">{String(row[c] ?? '')}</td>)}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

function LogPanel({ logs }: { logs: LogItem[] }) {
  return (
    <div className="rounded-[28px] border border-slate-200 bg-white/90 p-4 shadow-soft">
      <h2 className="text-lg font-black">Log</h2>
      <div className="mt-3 max-h-80 space-y-2 overflow-auto pr-1 studio-scroll">
        {logs.length === 0 ? <div className="text-sm text-slate-500">Belum ada aktivitas.</div> : logs.map((log, i) => (
          <details key={i} className={`rounded-2xl border p-3 text-sm ${log.level === 'error' ? 'border-rose-200 bg-rose-50 text-rose-900' : log.level === 'warn' ? 'border-amber-200 bg-amber-50 text-amber-900' : log.level === 'ok' ? 'border-emerald-200 bg-emerald-50 text-emerald-900' : 'border-slate-200 bg-slate-50 text-slate-700'}`}>
            <summary className="cursor-pointer font-semibold">[{log.time}] {log.message}</summary>
            {log.detail && <pre className="mt-2 max-h-56 overflow-auto whitespace-pre-wrap text-xs">{JSON.stringify(log.detail, null, 2)}</pre>}
          </details>
        ))}
      </div>
    </div>
  );
}

function EmptyState() {
  return (
    <div className="flex min-h-[60vh] items-center justify-center rounded-3xl border-2 border-dashed border-slate-300 bg-slate-50 text-center">
      <div className="max-w-lg p-8">
        <div className="text-5xl">🧰</div>
        <h2 className="mt-4 text-2xl font-black">Mulai dari XLSX atau export record Odoo</h2>
        <p className="mt-2 text-slate-600">Upload file untuk import, atau buka menu Export untuk scan record dan membuat XLSX baru dari database Odoo.</p>
      </div>
    </div>
  );
}

function labelKind(kind: EditorKind) {
  const map: Record<EditorKind, string> = {
    product: 'Product Editor',
    contact: 'Contact Editor',
    project: 'Project Editor',
    knowledge: 'Knowledge Editor',
    sales: 'Sales Editor',
    dynamic: 'Dynamic Odoo Editor',
    helper: 'Context Sheet'
  };
  return map[kind];
}
