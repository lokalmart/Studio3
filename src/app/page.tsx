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

function now() { return new Date().toLocaleTimeString('id-ID', { hour12: false }); }
function detectModel(sheetName: string, rows: Row[]) {
  const fromRow = rows.find(r => r._model)?._model;
  if (fromRow) return String(fromRow).trim();
  const normalized = sheetName.toLowerCase().trim();
  const known: Record<string, string> = {
    contacts: 'res.partner', contact: 'res.partner', partner: 'res.partner', partners: 'res.partner',
    product: 'product.template', products: 'product.template', project: 'project.project', projects: 'project.project',
    task: 'project.task', tasks: 'project.task', knowledge: 'knowledge.article', articles: 'knowledge.article',
    sales: 'sale.order', orders: 'sale.order'
  };
  return known[normalized] || sheetName;
}
function detectKind(model: string, sheetName: string): EditorKind {
  const m = String(model || '').toLowerCase(); const s = String(sheetName || '').toLowerCase();
  if (HELPER_SHEETS.has(s) || m.includes('helper') || m.includes('readme')) return 'helper';
  if (m === 'res.partner') return 'contact';
  if (m === 'product.template' || m === 'product.product' || m === 'product.category' || m === 'product.public.category') return 'product';
  if (m.startsWith('project.')) return 'project';
  if (m === 'knowledge.article') return 'knowledge';
  if (m.startsWith('sale.')) return 'sales';
  return 'dynamic';
}
function makeColumns(rows: Row[]) { const set = new Set<string>(); rows.forEach(row => Object.keys(row).forEach(k => set.add(k))); return Array.from(set); }
function sanitizeWorkbookValue(value: any) {
  if (value === null || value === undefined) return '';
  if (Array.isArray(value) || (typeof value === 'object' && value !== null)) return JSON.stringify(value).slice(0, SAFE_LIMIT);
  const text = String(value);
  return text.length > EXCEL_LIMIT ? text.slice(0, SAFE_LIMIT) : value;
}
function normalizeRowsForSheet(rows: Row[]) { return rows.map((row, idx) => ({ ...row, __rownum__: row.__rownum__ || idx + 2 })); }
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
function chunk<T>(arr: T[], size: number) { const out: T[][] = []; for (let i = 0; i < arr.length; i += size) out.push(arr.slice(i, i + size)); return out; }
function parseCsvFields(text: string) { return text.split(',').map(x => x.trim()).filter(Boolean); }
function importantFields(kind: EditorKind, model: string) {
  if (kind === 'contact') return ['name', 'display_name', 'phone', 'mobile', 'email', 'street', 'city', 'supplier_rank', 'customer_rank', 'is_company'];
  if (kind === 'product') return ['name', 'default_code', 'barcode', 'list_price', 'standard_price', 'categ_id', 'public_categ_ids', 'image_1920', 'description_sale', 'sale_ok', 'purchase_ok'];
  if (kind === 'project') return ['name', 'project_id', 'parent_id', 'stage_id', 'user_id', 'user_ids', 'partner_id', 'date_deadline', 'description', 'priority', 'sequence'];
  if (kind === 'knowledge') return ['name', 'body', 'body_html', 'parent_id', 'category'];
  if (kind === 'sales') return ['name', 'partner_id', 'date_order', 'state', 'amount_total', 'invoice_status'];
  return ['display_name', 'name', 'create_date', 'write_date'];
}
function kindIcon(kind: EditorKind) {
  return ({ product: '◈', contact: '◎', project: '▦', knowledge: '✦', sales: '₿', dynamic: '◇', helper: '✧' } as Record<EditorKind, string>)[kind];
}
function labelKind(kind: EditorKind) {
  return ({ product: 'Product', contact: 'Contact', project: 'Project', knowledge: 'Knowledge', sales: 'Sales', dynamic: 'Dynamic', helper: 'Context' } as Record<EditorKind, string>)[kind];
}

type ModelPreset = { key: string; label: string; model: string; kind: EditorKind; description: string; fields: string; tone: string };
const MODEL_PRESETS: ModelPreset[] = [
  { key: 'contacts', label: 'Contacts', model: 'res.partner', kind: 'contact', description: 'Customer, supplier, member, vendor, dan mitra Lokalmart.', fields: 'name,display_name,email,phone,mobile,street,street2,city,state_id,country_id,customer_rank,supplier_rank,is_company,comment', tone: 'violet' },
  { key: 'products', label: 'Products', model: 'product.template', kind: 'product', description: 'Produk, harga, barcode, kategori, foto URL, dan data katalog.', fields: 'name,default_code,barcode,list_price,standard_price,categ_id,public_categ_ids,sale_ok,purchase_ok,website_published,description_sale,image_1920', tone: 'cyan' },
  { key: 'projects', label: 'Projects', model: 'project.project', kind: 'project', description: 'Project utama seperti Ground Zero, Soraya Kitchen, dan Pilot.', fields: 'name,display_name,partner_id,user_id,active,date_start,date,description,privacy_visibility,stage_id', tone: 'amber' },
  { key: 'tasks', label: 'Tasks', model: 'project.task', kind: 'project', description: 'Task, subtask, parent, stage, deadline, PIC, dan hierarchy.', fields: 'name,display_name,project_id,parent_id,stage_id,user_ids,partner_id,date_deadline,priority,sequence,description', tone: 'amber' },
  { key: 'knowledge', label: 'Knowledge', model: 'knowledge.article', kind: 'knowledge', description: 'Artikel knowledge, SOP, konteks AI, dan rujukan project.', fields: 'name,display_name,parent_id,body,body_html,create_date,write_date', tone: 'fuchsia' },
  { key: 'sales', label: 'Sales', model: 'sale.order', kind: 'sales', description: 'Sales order, customer, status invoice, total, dan tanggal order.', fields: 'name,partner_id,date_order,state,invoice_status,amount_untaxed,amount_tax,amount_total,validity_date,note', tone: 'emerald' },
  { key: 'categories', label: 'Categories', model: 'product.category', kind: 'product', description: 'Kategori teknis internal produk Lokalmart.', fields: 'name,display_name,parent_id,complete_name,property_cost_method,property_valuation', tone: 'cyan' },
  { key: 'website_categories', label: 'Web Categories', model: 'product.public.category', kind: 'product', description: 'Kategori katalog website/eCommerce untuk pelanggan.', fields: 'name,display_name,parent_id,sequence,website_id', tone: 'cyan' }
];
function findPresetByModel(model: string) { return MODEL_PRESETS.find(p => p.model === model); }
function toneClass(tone: string) { return `tone-${tone}`; }

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
  const [exportFields, setExportFields] = useState(MODEL_PRESETS[0].fields);
  const [exportDomain, setExportDomain] = useState('[]');
  const [scanRecords, setScanRecords] = useState<Row[]>([]);
  const [scanCount, setScanCount] = useState(0);
  const [scanOffset, setScanOffset] = useState(0);
  const [selectedIds, setSelectedIds] = useState<Record<number, boolean>>({});

  useEffect(() => { try { const raw = localStorage.getItem(STORAGE_KEY); if (raw) setConn({ ...defaultConn, ...JSON.parse(raw) }); } catch {} }, []);
  useEffect(() => { try { localStorage.setItem(STORAGE_KEY, JSON.stringify(conn)); } catch {} }, [conn]);
  useEffect(() => { setSchema(null); setSchemaModel(''); setSelectedRows({}); }, [activeSheetIndex]);

  function addLog(level: LogItem['level'], message: string, detail?: any) { setLogs(prev => [{ time: now(), level, message, detail }, ...prev].slice(0, 160)); }
  async function callOdoo(payload: Row) {
    const res = await fetch('/api/odoo', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ ...payload, connection: conn }) });
    const data = await res.json(); if (!data.ok) throw new Error(data.error || 'API error'); return data;
  }
  async function testConnection() { setBusy(true); try { const data = await callOdoo({ action: 'test' }); addLog('ok', `Koneksi Odoo berhasil. UID ${data.uid}, kontak ${data.partner_count}.`); } catch (e: any) { addLog('error', `Koneksi gagal: ${e.message}`); } finally { setBusy(false); } }
  async function loadSchema(model?: string) {
    const target = model || activeSheet?.model || exportModel; if (!target) return; setBusy(true);
    try { const data = await callOdoo({ action: 'schema', model: target }); setSchema(data.fields); setSchemaModel(target); addLog('ok', `Schema dimuat: ${target} (${Object.keys(data.fields || {}).length} field).`); }
    catch (e: any) { addLog('error', `Gagal memuat schema ${target}: ${e.message}`); } finally { setBusy(false); }
  }
  async function handleFile(file: File) {
    setBusy(true); try {
      const buffer = await file.arrayBuffer(); const wb = XLSX.read(buffer, { type: 'array', cellDates: false });
      const parsed: SheetState[] = wb.SheetNames.map(name => rowsToSheetState(name, XLSX.utils.sheet_to_json<Row>(wb.Sheets[name], { defval: '' })));
      setSheets(parsed); setActiveSheetIndex(0); setMode('import'); addLog('ok', `XLSX dibaca: ${parsed.length} sheet dari ${file.name}.`);
    } catch (e: any) { addLog('error', `Gagal membaca XLSX: ${e.message}`); } finally { setBusy(false); }
  }
  function updateCell(rowIndex: number, col: string, value: any) { setSheets(prev => prev.map((sheet, si) => si !== activeSheetIndex ? sheet : { ...sheet, rows: sheet.rows.map((row, ri) => ri === rowIndex ? { ...row, [col]: value } : row), columns: sheet.columns.includes(col) ? sheet.columns : [...sheet.columns, col] })); }
  function addColumn() { const name = prompt('Nama kolom baru, misalnya x_review_status atau vendor_external_id'); if (!name || !activeSheet) return; setSheets(prev => prev.map((sheet, si) => si === activeSheetIndex ? { ...sheet, columns: sheet.columns.includes(name) ? sheet.columns : [...sheet.columns, name] } : sheet)); }
  function addRow() { if (!activeSheet) return; setSheets(prev => prev.map((sheet, si) => si === activeSheetIndex ? { ...sheet, rows: [...sheet.rows, { _model: sheet.model, __action: 'upsert', _external_id: '', __rownum__: sheet.rows.length + 2 }] } : sheet)); }
  function validateActiveSheet() {
    if (!activeSheet) return [];
    const issues: string[] = []; if (activeSheet.helper) return issues;
    if (!activeSheet.model) issues.push('Model belum jelas. Isi _model atau ubah nama sheet sesuai model Odoo.');
    activeSheet.rows.forEach((row, i) => { if (!row._model && !activeSheet.model) issues.push(`Row ${i + 2}: _model kosong.`); if (!row.__action) issues.push(`Row ${i + 2}: __action kosong; default tetap upsert.`); if (!row._external_id && !row.x_studio2_odoo_id && !row.id) issues.push(`Row ${i + 2}: _external_id kosong. Aman untuk create, kurang aman untuk update.`); });
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
    if (!activeSheet) return; if (activeSheet.helper) return addLog('warn', 'Sheet helper/context tidak diimport. Pilih sheet model Odoo.');
    const chosen = activeSheet.rows.filter((_, idx) => Object.keys(selectedRows).length ? selectedRows[idx] : true); if (!chosen.length) return addLog('warn', 'Tidak ada row dipilih untuk import.');
    const parts = chunk(chosen, Math.max(1, Math.min(50, batchSize))); setBusy(true); addLog('info', `Mulai import ${chosen.length} row dari ${activeSheet.name} dalam ${parts.length} batch.`);
    let totalCreated = 0, totalUpdated = 0, totalFailed = 0;
    for (let i = 0; i < parts.length; i++) {
      try { const data = await callOdoo({ action: 'import_batch', model: activeSheet.model, rows: parts[i] }); totalCreated += Number(data.created || 0); totalUpdated += Number(data.updated || 0); totalFailed += Number(data.failed || 0); addLog(data.failed ? 'warn' : 'ok', `Batch ${i + 1}/${parts.length}: created ${data.created}, updated ${data.updated}, failed ${data.failed}.`, data.results); if (data.failed) break; }
      catch (e: any) { totalFailed += parts[i].length; addLog('error', `Batch ${i + 1} gagal: ${e.message}`); break; }
    }
    addLog(totalFailed ? 'warn' : 'ok', `Import selesai: created ${totalCreated}, updated ${totalUpdated}, failed ${totalFailed}.`); setBusy(false);
  }
  function downloadWorkbook(fileName = 'studio2_export.xlsx') {
    if (!sheets.length) return; const wb = XLSX.utils.book_new();
    sheets.forEach(sheet => { const cleanRows = sheet.rows.map(row => { const out: Row = {}; sheet.columns.forEach(col => out[col] = sanitizeWorkbookValue(row[col])); return out; }); const ws = XLSX.utils.json_to_sheet(cleanRows.length ? cleanRows : [{}], { header: sheet.columns.length ? sheet.columns : undefined }); XLSX.utils.book_append_sheet(wb, ws, sheet.name.slice(0, 31) || 'Sheet1'); });
    XLSX.writeFile(wb, fileName); addLog('ok', `XLSX diunduh: ${fileName}`);
  }
  async function scanModel(reset = true) {
    setBusy(true); try {
      let domain: any[] = []; try { domain = JSON.parse(exportDomain || '[]'); } catch { throw new Error('Domain harus JSON array, contoh: [] atau [["is_company","=",true]]'); }
      const offset = reset ? 0 : scanOffset; const kind = detectKind(exportModel, exportModel); const fields = Array.from(new Set([...importantFields(kind, exportModel), ...parseCsvFields(exportFields)])).filter(Boolean);
      const data = await callOdoo({ action: 'record_scan', model: exportModel, fields, domain, offset, limit: 80 });
      setScanCount(data.count || 0); setScanOffset(offset + (data.records?.length || 0)); setScanRecords(prev => reset ? data.records : [...prev, ...data.records]); if (reset) setSelectedIds({});
      addLog('ok', `Scan ${exportModel}: ${data.records.length} record dimuat dari ${data.count}.`);
    } catch (e: any) { addLog('error', `Scan gagal: ${e.message}`); } finally { setBusy(false); }
  }
  async function exportSelectedRecords() {
    const ids = Object.entries(selectedIds).filter(([, v]) => v).map(([id]) => Number(id)); if (!ids.length) return addLog('warn', 'Belum ada record dipilih untuk export.');
    setBusy(true); try { const data = await callOdoo({ action: 'export_records', model: exportModel, ids, fields: parseCsvFields(exportFields) }); const sheet = rowsToSheetState(data.sheet || exportModel, data.rows || []); setSheets([sheet]); setActiveSheetIndex(0); setMode('import'); addLog('ok', `Export ${ids.length} record dari ${exportModel} masuk editor XLSX.`); }
    catch (e: any) { addLog('error', `Export gagal: ${e.message}`); } finally { setBusy(false); }
  }
  async function exportProject() {
    const id = prompt('Masukkan ID project Odoo, contoh: 73'); if (!id) return; setBusy(true);
    try { const data = await callOdoo({ action: 'export_project', project_id: Number(id) }); const nextSheets = Object.entries(data.sheets || {}).map(([name, rows]) => rowsToSheetState(name, rows as Row[])); setSheets(nextSheets); setActiveSheetIndex(0); setMode('import'); addLog('ok', `Project ${id} diexport ke ${nextSheets.length} sheet dan masuk editor.`); }
    catch (e: any) { addLog('error', `Export project gagal: ${e.message}`); } finally { setBusy(false); }
  }
  const visibleColumns = useMemo(() => {
    if (!activeSheet) return [];
    const priority = ['_model', '__action', '_external_id', 'x_studio2_odoo_id', ...importantFields(activeSheet.kind, activeSheet.model)]; const set = new Set(activeSheet.columns);
    return [...priority.filter(c => set.has(c)), ...activeSheet.columns.filter(c => !priority.includes(c) && c !== '__rownum__')];
  }, [activeSheet]);
  const connectionOk = Boolean(conn.url && conn.db && conn.username && conn.password);
  const selectedImportCount = Object.values(selectedRows).filter(Boolean).length;
  const selectedExportCount = Object.values(selectedIds).filter(Boolean).length;

  return (
    <main className="app-bg min-h-screen pb-24 text-white md:pb-6">
      <div className="mx-auto flex min-h-screen w-full max-w-[1320px] flex-col gap-4 px-3 py-3 sm:px-5 lg:px-6">
        <AppHeader mode={mode} setMode={setMode} settingsOpen={settingsOpen} setSettingsOpen={setSettingsOpen} connectionOk={connectionOk} busy={busy} />
        {settingsOpen && <ConnectionCard conn={conn} setConn={setConn} testConnection={testConnection} busy={busy} connectionOk={connectionOk} />}
        {mode === 'import' ? (
          <ImportExperience busy={busy} sheets={sheets} activeSheet={activeSheet} activeSheetIndex={activeSheetIndex} setActiveSheetIndex={setActiveSheetIndex} onFile={handleFile} addRow={addRow} addColumn={addColumn} download={() => downloadWorkbook('studio2_edited.xlsx')} importActiveSheet={importActiveSheet} batchSize={batchSize} setBatchSize={setBatchSize} loadSchema={() => loadSchema()} selectedCount={selectedImportCount} issues={issues} columns={visibleColumns} schema={schemaModel === activeSheet?.model ? schema : null} selectedRows={selectedRows} setSelectedRows={setSelectedRows} updateCell={updateCell} />
        ) : (
          <ExportExperience busy={busy} model={exportModel} setModel={setExportModel} fields={exportFields} setFields={setExportFields} domain={exportDomain} setDomain={setExportDomain} scan={() => scanModel(true)} loadMore={() => scanModel(false)} exportSelected={exportSelectedRecords} exportProject={exportProject} scanCount={scanCount} scanOffset={scanOffset} records={scanRecords} selectedIds={selectedIds} setSelectedIds={setSelectedIds} loadSchema={() => loadSchema(exportModel)} selectedCount={selectedExportCount} />
        )}
        <LogDrawer logs={logs} />
      </div>
      <MobileNav mode={mode} setMode={setMode} settingsOpen={settingsOpen} setSettingsOpen={setSettingsOpen} />
    </main>
  );
}

function AppHeader({ mode, setMode, settingsOpen, setSettingsOpen, connectionOk, busy }: { mode: 'import' | 'export'; setMode: (m: 'import' | 'export') => void; settingsOpen: boolean; setSettingsOpen: (v: boolean) => void; connectionOk: boolean; busy: boolean }) {
  return (
    <header className="hero-card">
      <div className="flex items-start justify-between gap-3">
        <div className="flex min-w-0 items-center gap-3">
          <div className="brand-mark">LM</div>
          <div className="min-w-0">
            <div className="flex flex-wrap gap-1.5">
              <span className="mini-badge">Studio2 v9.3</span>
              <span className="mini-badge soft">Mobile Command</span>
              <span className={connectionOk ? 'mini-badge ok' : 'mini-badge warn'}>{connectionOk ? 'Odoo ready' : 'Need setup'}</span>
            </div>
            <h1 className="mt-2 text-2xl font-black leading-none tracking-tight sm:text-4xl">Lokalmart Studio</h1>
            <p className="mt-2 max-w-2xl text-xs leading-5 text-slate-300 sm:text-sm">Import, export, validasi, dan edit data Odoo lewat workflow yang ringan—bukan spreadsheet mentah.</p>
          </div>
        </div>
        <div className="hidden items-center gap-2 rounded-3xl border border-white/10 bg-white/5 p-1.5 md:flex">
          <button className={`mode-btn ${mode === 'import' ? 'active' : ''}`} onClick={() => setMode('import')}>Import</button>
          <button className={`mode-btn ${mode === 'export' ? 'active' : ''}`} onClick={() => setMode('export')}>Export</button>
          <button className={`mode-btn ${settingsOpen ? 'active' : ''}`} onClick={() => setSettingsOpen(!settingsOpen)}>⚙</button>
        </div>
      </div>
      {busy && <div className="mt-4 h-1 overflow-hidden rounded-full bg-white/10"><div className="loading-line" /></div>}
    </header>
  );
}

function MobileNav({ mode, setMode, settingsOpen, setSettingsOpen }: { mode: 'import' | 'export'; setMode: (m: 'import' | 'export') => void; settingsOpen: boolean; setSettingsOpen: (v: boolean) => void }) {
  return <nav className="mobile-nav md:hidden"><button className={mode === 'import' ? 'on' : ''} onClick={() => setMode('import')}>⇧<span>Import</span></button><button className={mode === 'export' ? 'on' : ''} onClick={() => setMode('export')}>⇩<span>Export</span></button><button className={settingsOpen ? 'on' : ''} onClick={() => setSettingsOpen(!settingsOpen)}>⚙<span>Koneksi</span></button></nav>;
}

function ConnectionCard({ conn, setConn, testConnection, busy, connectionOk }: { conn: Conn; setConn: (c: Conn) => void; testConnection: () => void; busy: boolean; connectionOk: boolean }) {
  return <section className="glass-card p-4"><div className="mb-3 flex items-center justify-between"><div><div className="eyebrow">Connection vault</div><h2 className="text-xl font-black">Target Odoo</h2></div><span className={connectionOk ? 'status ok' : 'status warn'}>{connectionOk ? 'Siap' : 'Belum lengkap'}</span></div><div className="grid gap-3 sm:grid-cols-2 xl:grid-cols-5"><FloatingInput label="Odoo URL" placeholder="https://xxx.odoo.com" value={conn.url} onChange={v => setConn({ ...conn, url: v })} /><FloatingInput label="Database" placeholder="nama database" value={conn.db} onChange={v => setConn({ ...conn, db: v })} /><FloatingInput label="Username" placeholder="email/username" value={conn.username} onChange={v => setConn({ ...conn, username: v })} /><FloatingInput label="Password / API Key" placeholder="localStorage" type="password" value={conn.password} onChange={v => setConn({ ...conn, password: v })} /><button className="primary-btn min-h-[58px]" disabled={busy || !connectionOk} onClick={testConnection}>Tes Koneksi</button></div></section>;
}

function FloatingInput({ label, value, onChange, placeholder, type = 'text' }: { label: string; value: string; onChange: (v: string) => void; placeholder?: string; type?: string }) {
  return <label className="input-card"><span>{label}</span><input placeholder={placeholder} type={type} value={value} onChange={e => onChange(e.target.value)} /></label>;
}

function ImportExperience(props: { busy: boolean; sheets: SheetState[]; activeSheet?: SheetState; activeSheetIndex: number; setActiveSheetIndex: (n: number) => void; onFile: (f: File) => void; addRow: () => void; addColumn: () => void; download: () => void; importActiveSheet: () => void; batchSize: number; setBatchSize: (n: number) => void; loadSchema: () => void; selectedCount: number; issues: string[]; columns: string[]; schema: Record<string, OdooField> | null; selectedRows: Record<number, boolean>; setSelectedRows: (x: Record<number, boolean>) => void; updateCell: (row: number, col: string, value: any) => void }) {
  if (!props.sheets.length) return <StartImport onFile={props.onFile} />;
  return <section className="space-y-4"><MobileStepper labels={['Upload', 'Sheet', 'Validasi', 'Import']} active={props.activeSheet ? 2 : 1} /><SheetSwitcher sheets={props.sheets} activeIndex={props.activeSheetIndex} setActiveIndex={props.setActiveSheetIndex} /><ActionDock primary="Import" primaryAction={props.importActiveSheet} disabled={props.busy} items={[['Schema', props.loadSchema], ['+ Row', props.addRow], ['+ Kolom', props.addColumn], ['Download', props.download]]} extra={<label className="tiny-upload"><input className="hidden" type="file" accept=".xlsx,.xls" onChange={e => e.target.files?.[0] && props.onFile(e.target.files[0])} />Ganti XLSX</label>} /><Editor sheet={props.activeSheet!} columns={props.columns} schema={props.schema} issues={props.issues} selectedRows={props.selectedRows} setSelectedRows={props.setSelectedRows} updateCell={props.updateCell} /></section>;
}

function StartImport({ onFile }: { onFile: (f: File) => void }) {
  return <section className="start-grid"><div className="start-card"><div className="app-chip">Import Mission</div><h2 className="mt-4 text-3xl font-black sm:text-5xl">Buka XLSX, lalu rapikan sebelum masuk Odoo.</h2><p className="mt-4 max-w-xl text-sm leading-6 text-slate-300">Studio membaca file di browser, mengenali model, lalu menampilkan editor yang sesuai: contact, product, project, knowledge, atau dynamic Odoo.</p><label className="upload-hero mt-6"><input className="hidden" type="file" accept=".xlsx,.xls" onChange={e => e.target.files?.[0] && onFile(e.target.files[0])} /><span>＋</span><b>Pilih XLSX</b><small>Preview lokal · aman untuk Vercel</small></label></div><div className="feature-stack"><Feature n="01" title="Preview dulu" text="Tidak langsung import. Semua sheet dibaca sebagai draft kerja." /><Feature n="02" title="Editor per model" text="res.partner jadi contact, product jadi product, project jadi project." /><Feature n="03" title="Batch aman" text="Import dipecah kecil agar tidak menabrak timeout Vercel." /></div></section>;
}

function Feature({ n, title, text }: { n: string; title: string; text: string }) { return <div className="feature-card"><span>{n}</span><b>{title}</b><p>{text}</p></div>; }
function MobileStepper({ labels, active }: { labels: string[]; active: number }) { return <div className="stepper">{labels.map((label, i) => <div key={label} className={i <= active ? 'done' : ''}><span>{i + 1}</span><b>{label}</b></div>)}</div>; }
function SheetSwitcher({ sheets, activeIndex, setActiveIndex }: { sheets: SheetState[]; activeIndex: number; setActiveIndex: (n: number) => void }) { return <div className="sheet-strip">{sheets.map((s, i) => <button key={s.name} className={i === activeIndex ? 'on' : ''} onClick={() => setActiveIndex(i)}><span>{kindIcon(s.kind)}</span><b>{s.name}</b><small>{s.rows.length} row · {s.model}</small></button>)}</div>; }
function ActionDock({ primary, primaryAction, disabled, items, extra }: { primary: string; primaryAction: () => void; disabled?: boolean; items: Array<[string, () => void]>; extra?: React.ReactNode }) { return <div className="action-dock"><button className="primary-btn" disabled={disabled} onClick={primaryAction}>{primary}</button><div className="quick-actions">{items.map(([label, fn]) => <button key={label} onClick={fn}>{label}</button>)}{extra}</div></div>; }

function ExportExperience(props: { busy: boolean; model: string; setModel: (s: string) => void; fields: string; setFields: (s: string) => void; domain: string; setDomain: (s: string) => void; scan: () => void; loadMore: () => void; exportSelected: () => void; exportProject: () => void; scanCount: number; scanOffset: number; records: Row[]; selectedIds: Record<number, boolean>; setSelectedIds: (x: Record<number, boolean>) => void; loadSchema: () => void; selectedCount: number }) {
  return <section className="space-y-4"><MobileStepper labels={['Model', 'Scan', 'Pilih', 'Export']} active={props.records.length ? (props.selectedCount ? 2 : 1) : 0} /><div className="mission-layout"><MissionSelector model={props.model} setModel={props.setModel} setFields={props.setFields} fields={props.fields} setFieldsText={props.setFields} domain={props.domain} setDomain={props.setDomain} scan={props.scan} busy={props.busy} loadSchema={props.loadSchema} exportProject={props.exportProject} /><div className="min-w-0"><ActionDock primary="Export Terpilih" primaryAction={props.exportSelected} disabled={props.busy || props.selectedCount === 0} items={[['Scan', props.scan], ['Load more', props.loadMore], ['Schema', props.loadSchema], ['Project', props.exportProject]]} /><RecordPicker records={props.records} selectedIds={props.selectedIds} setSelectedIds={props.setSelectedIds} scanCount={props.scanCount} /></div></div></section>;
}

function MissionSelector(props: { model: string; setModel: (s: string) => void; setFields: (s: string) => void; fields: string; setFieldsText: (s: string) => void; domain: string; setDomain: (s: string) => void; scan: () => void; busy: boolean; loadSchema: () => void; exportProject: () => void }) {
  const activePreset = findPresetByModel(props.model);
  return <div className="mission-card"><div className="eyebrow">Export mission</div><h2 className="mt-1 text-2xl font-black">Pilih data</h2><p className="mt-2 text-sm leading-6 text-slate-300">Checklist model utama. Custom tetap ada, tapi tidak mengganggu layar utama.</p><div className="model-grid mt-4">{MODEL_PRESETS.map(p => <button key={p.key} className={`model-card ${toneClass(p.tone)} ${props.model === p.model ? 'on' : ''}`} onClick={() => { props.setModel(p.model); props.setFields(p.fields); }}><span className="check">{props.model === p.model ? '✓' : ''}</span><b>{kindIcon(p.kind)} {p.label}</b><small>{p.model}</small><em>{p.description}</em></button>)}</div><details className="advanced mt-4"><summary>Advanced model & fields</summary><div className="mt-3 space-y-3"><FloatingInput label="Custom model" value={props.model} onChange={props.setModel} placeholder="res.partner" /><label className="input-card"><span>Fields export</span><textarea value={props.fields} onChange={e => props.setFieldsText(e.target.value)} /></label><FloatingInput label="Domain JSON" value={props.domain} onChange={props.setDomain} placeholder="[]" /></div></details><div className="mt-4 grid grid-cols-2 gap-2"><button className="primary-btn" disabled={props.busy || !props.model} onClick={props.scan}>Scan</button><button className="secondary-btn" disabled={props.busy || !props.model} onClick={props.loadSchema}>Schema</button></div><button className="secondary-btn mt-2 w-full" disabled={props.busy} onClick={props.exportProject}>Export Project by ID</button><div className="mt-3 rounded-2xl border border-white/10 bg-white/5 p-3 text-xs text-slate-300"><b>{activePreset?.label || 'Custom'}</b><br />{parseCsvFields(props.fields).length} field siap diexport</div></div>;
}

function Editor(props: { sheet: SheetState; columns: string[]; schema: Record<string, OdooField> | null; issues: string[]; selectedRows: Record<number, boolean>; setSelectedRows: (x: Record<number, boolean>) => void; updateCell: (row: number, col: string, value: any) => void }) {
  const { sheet } = props;
  const selected = Object.values(props.selectedRows).filter(Boolean).length;
  return <div className="editor-shell"><div className="editor-head"><div><div className="flex flex-wrap gap-2"><span className="mini-badge">{kindIcon(sheet.kind)} {labelKind(sheet.kind)}</span><span className="mini-badge soft">{sheet.model}</span><span className="mini-badge ok">{sheet.rows.length} rows</span><span className="mini-badge soft">{props.columns.length} fields</span></div><h2 className="mt-3 text-2xl font-black">{sheet.name}</h2><p className="mt-1 text-sm text-slate-300">Mobile menampilkan row sebagai kartu. Desktop tetap punya grid untuk edit cepat.</p></div><div className="summary-grid"><Metric label="Issue" value={String(props.issues.length)} /><Metric label="Selected" value={String(selected)} /><Metric label="Schema" value={props.schema ? 'ON' : 'OFF'} /></div></div>{sheet.helper && <Alert tone="amber" text="Sheet ini context/helper. Bisa diedit dan didownload, tapi tidak dikirim ke Odoo." />}{props.issues.length > 0 && <div className="issue-card"><b>Validasi menemukan catatan</b><ul>{props.issues.slice(0, 8).map((x, i) => <li key={i}>{x}</li>)}</ul>{props.issues.length > 8 && <small>+ {props.issues.length - 8} catatan lain</small>}</div>}<div className="md:hidden"><CardRows sheet={sheet} columns={props.columns} schema={props.schema} selectedRows={props.selectedRows} setSelectedRows={props.setSelectedRows} updateCell={props.updateCell} /></div><div className="hidden md:block"><GridRows sheet={sheet} columns={props.columns} schema={props.schema} selectedRows={props.selectedRows} setSelectedRows={props.setSelectedRows} updateCell={props.updateCell} /></div></div>;
}
function Alert({ text, tone }: { text: string; tone: 'amber' }) { return <div className={`alert-${tone}`}>{text}</div>; }
function Metric({ label, value }: { label: string; value: string }) { return <div className="metric"><small>{label}</small><b>{value}</b></div>; }

function CardRows(props: { sheet: SheetState; columns: string[]; schema: Record<string, OdooField> | null; selectedRows: Record<number, boolean>; setSelectedRows: (x: Record<number, boolean>) => void; updateCell: (row: number, col: string, value: any) => void }) {
  const priority = props.columns.slice(0, 10);
  return <div className="row-cards">{props.sheet.rows.slice(0, 80).map((row, ri) => <details key={ri} className="row-card" open={ri === 0}><summary><input type="checkbox" checked={Boolean(props.selectedRows[ri])} onChange={e => props.setSelectedRows({ ...props.selectedRows, [ri]: e.target.checked })} onClick={e => e.stopPropagation()} /><span>Row {ri + 1}</span><b>{String(row.name || row.display_name || row._external_id || props.sheet.model || '').slice(0, 40) || 'Record'}</b></summary><div className="row-fields">{priority.map(col => <label key={col}><span>{col}<small>{props.schema?.[col]?.type || 'xlsx'}</small></span><CellInput value={row[col] ?? ''} field={props.schema?.[col]} onChange={v => props.updateCell(ri, col, v)} /></label>)}</div></details>)}{props.sheet.rows.length > 80 && <div className="alert-amber">Mobile menampilkan 80 row pertama agar ringan. Pakai desktop/grid untuk edit massal.</div>}</div>;
}

function GridRows(props: { sheet: SheetState; columns: string[]; schema: Record<string, OdooField> | null; selectedRows: Record<number, boolean>; setSelectedRows: (x: Record<number, boolean>) => void; updateCell: (row: number, col: string, value: any) => void }) {
  return <div className="table-shell"><table><thead><tr><th><input type="checkbox" onChange={e => { const next: Record<number, boolean> = {}; if (e.target.checked) props.sheet.rows.forEach((_, i) => next[i] = true); props.setSelectedRows(next); }} /></th><th>#</th>{props.columns.map(col => <th key={col} className={props.schema && !META_COLUMNS.has(col) && !props.schema[col] && !col.endsWith('_external_id') && !col.endsWith('_external_ids') ? 'danger' : ''}><b>{col}</b>{props.schema?.[col] && <small>{props.schema[col].type}{props.schema[col].relation ? ` → ${props.schema[col].relation}` : ''}</small>}</th>)}</tr></thead><tbody>{props.sheet.rows.slice(0, 500).map((row, ri) => <tr key={ri} className={props.selectedRows[ri] ? 'selected' : ''}><td><input type="checkbox" checked={Boolean(props.selectedRows[ri])} onChange={e => props.setSelectedRows({ ...props.selectedRows, [ri]: e.target.checked })} /></td><td>{ri + 1}</td>{props.columns.map(col => <td key={col}><CellInput value={row[col] ?? ''} field={props.schema?.[col]} onChange={v => props.updateCell(ri, col, v)} /></td>)}</tr>)}</tbody></table>{props.sheet.rows.length > 500 && <div className="p-4 text-center text-sm text-slate-400">Ditampilkan 500 row pertama agar browser tetap ringan.</div>}</div>;
}

function CellInput({ value, field, onChange }: { value: any; field?: OdooField; onChange: (v: any) => void }) {
  if (field?.type === 'boolean') return <input className="check-input" type="checkbox" checked={['true', '1', 'yes', true].includes(value)} onChange={e => onChange(e.target.checked)} />;
  if (field?.type === 'selection' && field.selection) return <select className="cell-input" value={value ?? ''} onChange={e => onChange(e.target.value)}><option value="">—</option>{field.selection.map(([k, label]) => <option key={k} value={k}>{label}</option>)}</select>;
  if (field?.type === 'text' || field?.type === 'html' || String(value || '').length > 120) return <textarea className="cell-input min-h-[70px]" value={value ?? ''} onChange={e => onChange(e.target.value)} />;
  if (field?.type === 'integer' || field?.type === 'float' || field?.type === 'monetary') return <input className="cell-input" type="number" value={value ?? ''} onChange={e => onChange(e.target.value)} />;
  if (field?.type === 'date') return <input className="cell-input" type="date" value={String(value || '').slice(0, 10)} onChange={e => onChange(e.target.value)} />;
  return <input className="cell-input" value={value ?? ''} onChange={e => onChange(e.target.value)} />;
}

function RecordPicker({ records, selectedIds, setSelectedIds, scanCount }: { records: Row[]; selectedIds: Record<number, boolean>; setSelectedIds: (x: Record<number, boolean>) => void; scanCount: number }) {
  const [query, setQuery] = useState('');
  const cols = makeColumns(records).filter(c => c !== 'id').slice(0, 6);
  const selectedCount = Object.values(selectedIds).filter(Boolean).length;
  const filtered = records.filter(row => !query.trim() || String(Object.values(row).join(' ')).toLowerCase().includes(query.toLowerCase()));
  const selectVisible = () => { const next = { ...selectedIds }; filtered.forEach(r => next[Number(r.id)] = true); setSelectedIds(next); };
  const clearVisible = () => { const next = { ...selectedIds }; filtered.forEach(r => delete next[Number(r.id)]); setSelectedIds(next); };
  if (!records.length) return <div className="empty-panel"><div>⌕</div><h2>Belum ada record</h2><p>Pilih model di atas, lalu tekan Scan. Record akan muncul sebagai checklist card yang mudah dipilih.</p></div>;
  return <div className="record-section"><div className="record-top"><label className="search-card"><span>Cari record</span><input value={query} onChange={e => setQuery(e.target.value)} placeholder="nama, email, ID, telepon..." /></label><div className="record-actions"><button onClick={selectVisible}>Select visible</button><button onClick={clearVisible}>Clear visible</button><button onClick={() => setSelectedIds({})}>Clear all</button></div></div><div className="record-stats"><Metric label="Scanned" value={`${records.length}/${scanCount}`} /><Metric label="Visible" value={String(filtered.length)} /><Metric label="Selected" value={String(selectedCount)} /></div><div className="record-grid">{filtered.slice(0, 240).map(row => { const id = Number(row.id); const selected = Boolean(selectedIds[id]); const title = String(row.display_name || row.name || row.email || `Record #${id}`); const subtitle = String(row.email || row.phone || row.mobile || row.city || 'Odoo record'); return <button key={id} type="button" onClick={() => setSelectedIds({ ...selectedIds, [id]: !selected })} className={`record-card ${selected ? 'on' : ''}`}><span className="select-dot">{selected ? '✓' : ''}</span><span className="min-w-0 flex-1 text-left"><b>{title}</b><small>{subtitle}</small>{cols.slice(0, 3).map(c => row[c] ? <em key={c}><strong>{c}</strong>{String(row[c]).slice(0, 64)}</em> : null)}</span><i>#{id}</i></button>; })}</div>{filtered.length > 240 && <div className="alert-amber">Ditampilkan 240 record pertama. Gunakan search atau Load More bertahap.</div>}</div>;
}

function LogDrawer({ logs }: { logs: LogItem[] }) {
  const [open, setOpen] = useState(false); const last = logs[0];
  return <section className="log-drawer"><button onClick={() => setOpen(!open)}><span>Activity Log</span><b>{logs.length}</b><small>{last ? `${last.level}: ${last.message}` : 'idle'}</small></button>{open && <div className="log-list">{logs.length === 0 ? <p>Belum ada aktivitas.</p> : logs.map((log, i) => <details key={i} className={`log-item ${log.level}`}><summary>[{log.time}] {log.message}</summary>{log.detail && <pre>{JSON.stringify(log.detail, null, 2)}</pre>}</details>)}</div>}</section>;
}
