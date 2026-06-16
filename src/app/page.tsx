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

function kindIcon(kind: EditorKind) {
  const map: Record<EditorKind, string> = {
    product: '◈',
    contact: '◎',
    project: '▦',
    knowledge: '✦',
    sales: '₿',
    dynamic: '◇',
    helper: '✧'
  };
  return map[kind];
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

function kindTone(kind: EditorKind) {
  const map: Record<EditorKind, string> = {
    product: 'from-cyan-400/25 to-emerald-400/10 border-cyan-300/30',
    contact: 'from-violet-400/25 to-sky-400/10 border-violet-300/30',
    project: 'from-amber-400/25 to-orange-400/10 border-amber-300/30',
    knowledge: 'from-fuchsia-400/25 to-indigo-400/10 border-fuchsia-300/30',
    sales: 'from-emerald-400/25 to-lime-400/10 border-emerald-300/30',
    dynamic: 'from-slate-400/20 to-cyan-400/10 border-white/15',
    helper: 'from-amber-400/20 to-slate-400/10 border-amber-200/30'
  };
  return map[kind];
}


type ModelPreset = {
  key: string;
  label: string;
  model: string;
  kind: EditorKind;
  description: string;
  fields: string;
  tone: string;
};

const MODEL_PRESETS: ModelPreset[] = [
  {
    key: 'contacts',
    label: 'Contacts',
    model: 'res.partner',
    kind: 'contact',
    description: 'Customer, supplier, member, vendor, dan mitra Lokalmart.',
    fields: 'name,display_name,email,phone,mobile,street,street2,city,state_id,country_id,customer_rank,supplier_rank,is_company,comment',
    tone: 'violet'
  },
  {
    key: 'products',
    label: 'Products',
    model: 'product.template',
    kind: 'product',
    description: 'Produk, harga, barcode, kategori, foto URL, dan data katalog.',
    fields: 'name,default_code,barcode,list_price,standard_price,categ_id,public_categ_ids,sale_ok,purchase_ok,website_published,description_sale,image_1920',
    tone: 'cyan'
  },
  {
    key: 'projects',
    label: 'Projects',
    model: 'project.project',
    kind: 'project',
    description: 'Project utama seperti Ground Zero, Soraya Kitchen, dan Pilot.',
    fields: 'name,display_name,partner_id,user_id,active,date_start,date,description,privacy_visibility,stage_id',
    tone: 'amber'
  },
  {
    key: 'tasks',
    label: 'Tasks',
    model: 'project.task',
    kind: 'project',
    description: 'Task, subtask, parent, stage, deadline, PIC, dan hierarchy.',
    fields: 'name,display_name,project_id,parent_id,stage_id,user_ids,partner_id,date_deadline,priority,sequence,description',
    tone: 'amber'
  },
  {
    key: 'knowledge',
    label: 'Knowledge',
    model: 'knowledge.article',
    kind: 'knowledge',
    description: 'Artikel knowledge, SOP, konteks AI, dan rujukan project.',
    fields: 'name,display_name,parent_id,body,body_html,create_date,write_date',
    tone: 'fuchsia'
  },
  {
    key: 'sales',
    label: 'Sales',
    model: 'sale.order',
    kind: 'sales',
    description: 'Sales order, customer, status invoice, total, dan tanggal order.',
    fields: 'name,partner_id,date_order,state,invoice_status,amount_untaxed,amount_tax,amount_total,validity_date,note',
    tone: 'emerald'
  },
  {
    key: 'categories',
    label: 'Categories',
    model: 'product.category',
    kind: 'product',
    description: 'Kategori teknis internal produk Lokalmart.',
    fields: 'name,display_name,parent_id,complete_name,property_cost_method,property_valuation',
    tone: 'cyan'
  },
  {
    key: 'website_categories',
    label: 'Web Categories',
    model: 'product.public.category',
    kind: 'product',
    description: 'Kategori katalog website/eCommerce untuk pelanggan.',
    fields: 'name,display_name,parent_id,sequence,website_id',
    tone: 'cyan'
  }
];

function findPresetByModel(model: string) {
  return MODEL_PRESETS.find(p => p.model === model);
}

function modelToneClasses(tone: string) {
  const map: Record<string, string> = {
    cyan: 'model-tone-cyan',
    violet: 'model-tone-violet',
    amber: 'model-tone-amber',
    fuchsia: 'model-tone-fuchsia',
    emerald: 'model-tone-emerald'
  };
  return map[tone] || 'model-tone-default';
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
      const required = Object.entries(schema as Record<string, OdooField>).filter(([, v]) => v.required && !v.readonly).map(([k]) => k);
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
  const compactRail = mode === 'import' ? sheets.length > 0 : scanRecords.length > 0;
  const selectedImportCount = Object.values(selectedRows).filter(Boolean).length;
  const selectedExportCount = Object.values(selectedIds).filter(Boolean).length;

  return (
    <main className="web3-bg min-h-screen overflow-hidden text-white">
      <div className="orb orb-one" />
      <div className="orb orb-two" />
      <div className="orb orb-three" />
      <section className="relative mx-auto flex min-h-screen max-w-[1600px] flex-col gap-4 p-3 md:p-5">
        <header className="glass-panel border-white/10 p-4 md:p-5">
          <div className="flex flex-col gap-4 xl:flex-row xl:items-center xl:justify-between">
            <div className="min-w-0">
              <div className="flex flex-wrap items-center gap-2">
                <span className="badge badge-dark">Studio2 v9.2</span>
                <span className="badge badge-cyan">Command UI</span>
                <span className="badge badge-green">Vercel-safe batch</span>
                <span className={connectionOk ? 'status-pill status-ok' : 'status-pill status-warn'}>
                  <span className="pulse-dot" />{connectionOk ? 'Odoo target ready' : 'Koneksi belum lengkap'}
                </span>
              </div>
              <h1 className="mt-3 text-3xl font-black tracking-tight md:text-5xl">
                Lokalmart <span className="text-gradient">Command Studio</span>
              </h1>
              <p className="mt-2 max-w-4xl text-sm leading-6 text-slate-300 md:text-base">
                Command center untuk operasi data Odoo: pilih mission, scan record, edit seperti native workspace, validasi schema, lalu jalankan import/export batch kecil yang aman untuk Vercel.
              </p>
            </div>
            <div className="flex flex-wrap gap-2 rounded-3xl border border-white/10 bg-white/5 p-2 backdrop-blur-xl">
              <button className={`nav-tab ${mode === 'import' ? 'nav-tab-active' : ''}`} onClick={() => setMode('import')}>Import & Editor</button>
              <button className={`nav-tab ${mode === 'export' ? 'nav-tab-active' : ''}`} onClick={() => setMode('export')}>Export</button>
              <button className={`nav-tab ${settingsOpen ? 'nav-tab-active' : ''}`} onClick={() => setSettingsOpen(!settingsOpen)}>⚙ Koneksi</button>
            </div>
          </div>
          {settingsOpen && (
            <div className="mt-5 grid gap-3 rounded-[28px] border border-cyan-300/15 bg-slate-950/60 p-4 shadow-2xl shadow-cyan-950/20 md:grid-cols-2 xl:grid-cols-5">
              <FloatingInput label="Odoo URL" placeholder="https://xxx.odoo.com" value={conn.url} onChange={value => setConn({ ...conn, url: value })} />
              <FloatingInput label="Database" placeholder="nama database" value={conn.db} onChange={value => setConn({ ...conn, db: value })} />
              <FloatingInput label="Username" placeholder="email/username" value={conn.username} onChange={value => setConn({ ...conn, username: value })} />
              <FloatingInput label="Password / API Key" placeholder="disimpan di browser" type="password" value={conn.password} onChange={value => setConn({ ...conn, password: value })} />
              <button className="btn btn-neo h-full min-h-[58px]" disabled={busy || !connectionOk} onClick={testConnection}>{busy ? 'Memproses...' : 'Tes Koneksi'}</button>
            </div>
          )}
        </header>

        <div className="grid flex-1 gap-4 lg:grid-cols-[88px_1fr] xl:grid-cols-[var(--rail)_minmax(0,1fr)]" style={{ ['--rail' as any]: compactRail ? '118px' : '390px' }}>
          <aside className="min-w-0 space-y-4">
            {mode === 'import' ? (
              <ImportPanel compact={compactRail} busy={busy} sheets={sheets} activeSheetIndex={activeSheetIndex} setActiveSheetIndex={setActiveSheetIndex} onFile={handleFile} addRow={addRow} addColumn={addColumn} download={() => downloadWorkbook('studio2_edited.xlsx')} importActiveSheet={importActiveSheet} batchSize={batchSize} setBatchSize={setBatchSize} loadSchema={() => loadSchema()} selectedCount={selectedImportCount} />
            ) : (
              <ExportPanel compact={compactRail} busy={busy} model={exportModel} setModel={setExportModel} fields={exportFields} setFields={setExportFields} domain={exportDomain} setDomain={setExportDomain} scan={() => scanModel(true)} loadMore={() => scanModel(false)} exportSelected={exportSelectedRecords} exportProject={exportProject} scanCount={scanCount} scanRecords={scanRecords} scanOffset={scanOffset} selectedIds={selectedIds} loadSchema={() => loadSchema(exportModel)} selectedCount={selectedExportCount} />
            )}
            <LogPanel compact={compactRail} logs={logs} />
          </aside>

          <section className="workspace-panel min-w-0">
            <WorkflowBar mode={mode} activeSheet={activeSheet} sheets={sheets} records={scanRecords.length} selectedImportCount={selectedImportCount} selectedExportCount={selectedExportCount} issues={issues.length} />
            <div className="mt-4 min-h-[62vh]">
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
            </div>
          </section>
        </div>
      </section>
    </main>
  );
}

function FloatingInput({ label, value, onChange, placeholder, type = 'text' }: { label: string; value: string; onChange: (v: string) => void; placeholder?: string; type?: string }) {
  return (
    <label className="field-shell">
      <span>{label}</span>
      <input placeholder={placeholder} type={type} value={value} onChange={e => onChange(e.target.value)} />
    </label>
  );
}

function WorkflowBar(props: { mode: 'import' | 'export'; activeSheet?: SheetState; sheets: SheetState[]; records: number; selectedImportCount: number; selectedExportCount: number; issues: number }) {
  const steps = props.mode === 'import'
    ? [
      ['Upload', props.sheets.length > 0],
      ['Preview', Boolean(props.activeSheet)],
      ['Validate', Boolean(props.activeSheet) && props.issues === 0],
      ['Import', props.selectedImportCount > 0 || Boolean(props.activeSheet)]
    ]
    : [
      ['Model', true],
      ['Scan', props.records > 0],
      ['Select', props.selectedExportCount > 0],
      ['Export', false]
    ];
  return (
    <div className="flex flex-col gap-3 rounded-[30px] border border-white/10 bg-slate-950/35 p-3 md:flex-row md:items-center md:justify-between">
      <div className="flex min-w-0 items-center gap-3">
        <div className="grid h-12 w-12 place-items-center rounded-2xl border border-cyan-300/20 bg-cyan-300/10 text-xl shadow-lg shadow-cyan-950/20">⌁</div>
        <div className="min-w-0">
          <div className="text-xs font-black uppercase tracking-[.22em] text-cyan-200/80">Active workflow</div>
          <div className="truncate text-lg font-black text-white">{props.mode === 'import' ? props.activeSheet?.name || 'Import pipeline' : 'Export record pipeline'}</div>
        </div>
      </div>
      <div className="flex flex-wrap gap-2">
        {steps.map(([label, done], i) => (
          <div key={String(label)} className={`step-chip ${done ? 'step-done' : ''}`}>
            <span>{i + 1}</span>{label}
          </div>
        ))}
      </div>
    </div>
  );
}

function ImportPanel(props: {
  compact: boolean; busy: boolean; sheets: SheetState[]; activeSheetIndex: number; setActiveSheetIndex: (n: number) => void; onFile: (f: File) => void;
  addRow: () => void; addColumn: () => void; download: () => void; importActiveSheet: () => void; batchSize: number; setBatchSize: (n: number) => void; loadSchema: () => void; selectedCount: number;
}) {
  if (props.compact) {
    const active = props.sheets[props.activeSheetIndex];
    return (
      <div className="rail-card">
        <div className="rail-title">Import</div>
        <label className="rail-action cursor-pointer" title="Pilih XLSX">
          <input className="hidden" type="file" accept=".xlsx,.xls" onChange={e => e.target.files?.[0] && props.onFile(e.target.files[0])} />
          <span>＋</span>
        </label>
        <button className="rail-action" onClick={props.loadSchema} title="Load schema">⌘</button>
        <button className="rail-action" onClick={props.addColumn} title="Tambah kolom">▥</button>
        <button className="rail-action" onClick={props.download} title="Download XLSX">⇩</button>
        <button className="rail-action rail-go" disabled={props.busy} onClick={props.importActiveSheet} title="Import sheet aktif">↗</button>
        <div className="mt-2 w-full space-y-2">
          {props.sheets.slice(0, 6).map((sheet, i) => (
            <button key={sheet.name} onClick={() => props.setActiveSheetIndex(i)} className={`mini-sheet ${i === props.activeSheetIndex ? 'mini-sheet-active' : ''}`} title={`${sheet.name} · ${sheet.model}`}>
              {kindIcon(sheet.kind)}
            </button>
          ))}
        </div>
        <div className="rail-caption">{active?.rows.length || 0} rows</div>
      </div>
    );
  }

  return (
    <div className="control-card">
      <div className="flex items-center justify-between gap-3">
        <div>
          <div className="eyebrow">Import pipeline</div>
          <h2 className="text-xl font-black">Upload → Edit → Import</h2>
        </div>
        <div className="icon-tile">⇧</div>
      </div>
      <p className="mt-2 text-sm leading-6 text-slate-300">XLSX dibaca di browser. Setelah sheet aktif, panel ini otomatis mengecil agar editor jadi fokus utama.</p>
      <label className="dropzone mt-5 block cursor-pointer">
        <input className="hidden" type="file" accept=".xlsx,.xls" onChange={e => e.target.files?.[0] && props.onFile(e.target.files[0])} />
        <div className="text-4xl">◇</div>
        <div className="mt-2 text-lg font-black">Drop / pilih XLSX</div>
        <div className="text-xs text-slate-400">Preview lokal, aman untuk Vercel gratis</div>
      </label>
      {props.sheets.length > 0 && (
        <div className="mt-5 space-y-3">
          <div className="flex items-center justify-between">
            <div className="eyebrow">Detected sheets</div>
            <span className="badge badge-cyan">{props.sheets.length}</span>
          </div>
          <div className="max-h-72 space-y-2 overflow-auto pr-1 studio-scroll">
            {props.sheets.map((sheet, i) => (
              <button key={sheet.name} onClick={() => props.setActiveSheetIndex(i)} className={`sheet-card bg-gradient-to-br ${kindTone(sheet.kind)} ${i === props.activeSheetIndex ? 'sheet-card-active' : ''}`}>
                <div className="flex items-start justify-between gap-3">
                  <div className="min-w-0">
                    <div className="truncate text-sm font-black text-white">{kindIcon(sheet.kind)} {sheet.name}</div>
                    <div className="mt-1 truncate text-xs text-slate-300">{sheet.helper ? 'context/helper' : sheet.model}</div>
                  </div>
                  <span className="badge badge-dark">{sheet.rows.length}</span>
                </div>
              </button>
            ))}
          </div>
          <div className="grid grid-cols-2 gap-2 pt-1">
            <button className="btn btn-ghost" onClick={props.addRow}>+ Row</button>
            <button className="btn btn-ghost" onClick={props.addColumn}>+ Kolom</button>
            <button className="btn btn-ghost" onClick={props.loadSchema}>Schema</button>
            <button className="btn btn-ghost" onClick={props.download}>Download</button>
          </div>
          <FloatingInput label="Batch size" type="number" value={String(props.batchSize)} onChange={v => props.setBatchSize(Number(v))} />
          <button className="btn btn-neo w-full" disabled={props.busy} onClick={props.importActiveSheet}>Import Sheet Aktif</button>
        </div>
      )}
    </div>
  );
}

function ExportPanel(props: {
  compact: boolean; busy: boolean; model: string; setModel: (s: string) => void; fields: string; setFields: (s: string) => void; domain: string; setDomain: (s: string) => void;
  scan: () => void; loadMore: () => void; exportSelected: () => void; exportProject: () => void; scanCount: number; scanOffset: number; scanRecords: Row[];
  selectedIds: Record<number, boolean>; loadSchema: () => void; selectedCount: number;
}) {
  const activePreset = findPresetByModel(props.model);
  const fieldCount = parseCsvFields(props.fields).length;

  if (props.compact) {
    return (
      <div className="rail-card command-rail">
        <div className="rail-title">Export</div>
        <div className="compact-model-badge" title={props.model}>{activePreset ? kindIcon(activePreset.kind) : '◇'}</div>
        <button className="rail-action" disabled={props.busy} onClick={props.scan} title="Scan ulang">⌕</button>
        <button className="rail-action" disabled={props.busy} onClick={props.loadSchema} title="Schema">⌘</button>
        <button className="rail-action" disabled={props.busy || props.scanRecords.length >= props.scanCount} onClick={props.loadMore} title="Load more">…</button>
        <button className="rail-action" disabled={props.busy} onClick={props.exportProject} title="Export project">▦</button>
        <button className="rail-action rail-go" disabled={props.busy || props.selectedCount === 0} onClick={props.exportSelected} title="Export selected">⇩</button>
        <div className="rail-caption">{props.selectedCount}/{props.scanRecords.length}</div>
      </div>
    );
  }

  return (
    <div className="control-card command-control">
      <div className="flex items-start justify-between gap-3">
        <div>
          <div className="eyebrow">Export command</div>
          <h2 className="text-2xl font-black leading-tight">Pilih target data</h2>
          <p className="mt-2 text-sm leading-6 text-slate-300">Tidak perlu mengetik model. Pilih dari checklist, scan record, centang data, lalu export ke editor.</p>
        </div>
        <div className="icon-tile">⌕</div>
      </div>

      <div className="mt-5 grid gap-2">
        {MODEL_PRESETS.map(preset => {
          const active = props.model === preset.model;
          return (
            <button
              key={preset.key}
              type="button"
              onClick={() => {
                props.setModel(preset.model);
                props.setFields(preset.fields);
              }}
              className={`model-check-card ${modelToneClasses(preset.tone)} ${active ? 'model-check-active' : ''}`}
            >
              <span className={`check-ring ${active ? 'check-ring-on' : ''}`}>{active ? '✓' : ''}</span>
              <span className="min-w-0 flex-1">
                <span className="flex items-center gap-2 text-sm font-black text-white">
                  <span>{kindIcon(preset.kind)}</span>
                  <span className="truncate">{preset.label}</span>
                </span>
                <span className="mt-1 block truncate text-xs text-slate-400">{preset.model}</span>
                <span className="mt-1 block text-left text-[11px] leading-4 text-slate-300/85">{preset.description}</span>
              </span>
            </button>
          );
        })}
      </div>

      <div className="mt-4 rounded-3xl border border-white/10 bg-white/[.045] p-3">
        <div className="flex items-center justify-between gap-3">
          <div className="min-w-0">
            <div className="text-[10px] font-black uppercase tracking-[.2em] text-cyan-200/80">Current mission</div>
            <div className="mt-1 truncate text-sm font-black text-white">{activePreset?.label || 'Custom model'} · {props.model || 'belum dipilih'}</div>
          </div>
          <span className="badge badge-dark">{fieldCount} fields</span>
        </div>
        <div className="mt-3 flex flex-wrap gap-1.5">
          {parseCsvFields(props.fields).slice(0, 12).map(field => <span className="field-pill" key={field}>{field}</span>)}
          {fieldCount > 12 && <span className="field-pill field-pill-more">+{fieldCount - 12}</span>}
        </div>
      </div>

      <details className="advanced-drawer mt-3">
        <summary>Advanced: custom model, field list, domain</summary>
        <div className="mt-3 space-y-3">
          <FloatingInput label="Custom model" value={props.model} onChange={props.setModel} placeholder="contoh: res.partner" />
          <label className="field-shell">
            <span>Fields export</span>
            <textarea value={props.fields} onChange={e => props.setFields(e.target.value)} />
          </label>
          <FloatingInput label="Domain JSON" value={props.domain} onChange={props.setDomain} placeholder="[]" />
        </div>
      </details>

      <div className="mt-4 grid grid-cols-2 gap-2">
        <button className="btn btn-neo" disabled={props.busy || !props.model} onClick={props.scan}>Scan Record</button>
        <button className="btn btn-ghost" disabled={props.busy || !props.model} onClick={props.loadSchema}>Schema</button>
        <button className="btn btn-ghost" disabled={props.busy || props.scanRecords.length >= props.scanCount} onClick={props.loadMore}>Load More</button>
        <button className="btn btn-ghost" disabled={props.busy} onClick={props.exportProject}>Project Export</button>
      </div>

      <div className="stat-grid mt-3">
        <Metric label="Scanned" value={`${props.scanRecords.length}/${props.scanCount}`} />
        <Metric label="Selected" value={String(props.selectedCount)} />
      </div>
      <button className="btn btn-neo mt-3 w-full" disabled={props.busy || props.selectedCount === 0} onClick={props.exportSelected}>Export Record Terpilih</button>
    </div>
  );
}

function Metric({ label, value }: { label: string; value: string }) {
  return (
    <div className="rounded-2xl border border-white/10 bg-white/5 p-3">
      <div className="text-[10px] font-black uppercase tracking-[.2em] text-slate-400">{label}</div>
      <div className="mt-1 text-lg font-black text-white">{value}</div>
    </div>
  );
}

function Editor(props: {
  sheet: SheetState; columns: string[]; schema: Record<string, OdooField> | null; issues: string[]; selectedRows: Record<number, boolean>; setSelectedRows: (x: Record<number, boolean>) => void; updateCell: (row: number, col: string, value: any) => void;
}) {
  const { sheet } = props;
  return (
    <div className="space-y-4">
      <div className={`rounded-[30px] border bg-gradient-to-br p-4 ${kindTone(sheet.kind)}`}>
        <div className="flex flex-col gap-4 md:flex-row md:items-start md:justify-between">
          <div className="min-w-0">
            <div className="flex flex-wrap gap-2">
              <span className="badge badge-dark">{kindIcon(sheet.kind)} {labelKind(sheet.kind)}</span>
              <span className="badge badge-muted">{sheet.model}</span>
              <span className="badge badge-cyan">{sheet.rows.length} rows</span>
              <span className="badge badge-green">{props.columns.length} fields</span>
            </div>
            <h2 className="mt-3 truncate text-2xl font-black text-white">{sheet.name}</h2>
            <p className="mt-1 text-sm text-slate-300">Editor membaca tipe field Odoo saat schema dimuat: boolean, selection, date, number, text/html, dan relation hint.</p>
          </div>
          <div className="grid grid-cols-3 gap-2 text-center">
            <Metric label="Issues" value={String(props.issues.length)} />
            <Metric label="Selected" value={String(Object.values(props.selectedRows).filter(Boolean).length)} />
            <Metric label="Schema" value={props.schema ? 'ON' : 'OFF'} />
          </div>
        </div>
      </div>

      {sheet.helper && (
        <div className="alert-card alert-amber">Sheet ini dibaca sebagai context/helper. Isinya bisa diedit dan didownload, tapi tidak dikirim ke Odoo.</div>
      )}

      {props.issues.length > 0 && (
        <div className="alert-card alert-rose">
          <div className="font-black">Validasi menemukan catatan:</div>
          <ul className="mt-2 list-disc space-y-1 pl-5">
            {props.issues.slice(0, 10).map((x, i) => <li key={i}>{x}</li>)}
          </ul>
          {props.issues.length > 10 && <div className="mt-2">+ {props.issues.length - 10} catatan lain.</div>}
        </div>
      )}

      <div className="table-shell studio-scroll">
        <table className="min-w-full border-collapse text-sm">
          <thead className="sticky top-0 z-10">
            <tr>
              <th className="table-head w-12">
                <input type="checkbox" onChange={e => {
                  const next: Record<number, boolean> = {};
                  if (e.target.checked) sheet.rows.forEach((_, i) => next[i] = true);
                  props.setSelectedRows(next);
                }} />
              </th>
              <th className="table-head text-left">#</th>
              {props.columns.map(col => (
                <th key={col} className={`table-head text-left ${props.schema && !META_COLUMNS.has(col) && !props.schema[col] && !col.endsWith('_external_id') && !col.endsWith('_external_ids') ? 'table-head-danger' : ''}`}>
                  <div className="whitespace-nowrap">{col}</div>
                  {props.schema?.[col] && <div className="mt-1 whitespace-nowrap text-[10px] normal-case text-cyan-200/70">{props.schema[col].type}{props.schema[col].relation ? ` → ${props.schema[col].relation}` : ''}</div>}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {sheet.rows.slice(0, 500).map((row, ri) => (
              <tr key={ri} className={props.selectedRows[ri] ? 'row-selected' : 'row-normal'}>
                <td className="table-cell text-center"><input type="checkbox" checked={Boolean(props.selectedRows[ri])} onChange={e => props.setSelectedRows({ ...props.selectedRows, [ri]: e.target.checked })} /></td>
                <td className="table-cell text-xs text-slate-400">{ri + 1}</td>
                {props.columns.map(col => (
                  <td key={col} className="table-cell min-w-[190px] align-top">
                    <CellInput value={row[col] ?? ''} field={props.schema?.[col]} onChange={value => props.updateCell(ri, col, value)} />
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
        {sheet.rows.length > 500 && <div className="p-4 text-center text-sm text-slate-400">Editor menampilkan 500 row pertama agar browser tetap ringan. Import tetap memakai row yang ada di state.</div>}
      </div>
    </div>
  );
}

function CellInput({ value, field, onChange }: { value: any; field?: OdooField; onChange: (v: any) => void }) {
  if (field?.type === 'boolean') {
    return <div className="flex h-full items-center p-3"><input type="checkbox" checked={['true', '1', 'yes', true].includes(value)} onChange={e => onChange(e.target.checked)} /></div>;
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
    return <textarea className="cell-input min-h-[62px]" value={value ?? ''} onChange={e => onChange(e.target.value)} />;
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
  const [query, setQuery] = useState('');
  const cols = makeColumns(records).filter(c => c !== 'id').slice(0, 8);
  const selectedCount = Object.values(selectedIds).filter(Boolean).length;
  const filtered = records.filter(row => {
    if (!query.trim()) return true;
    const q = query.toLowerCase();
    return String(row.display_name || row.name || row.email || row.phone || row.id || '').toLowerCase().includes(q)
      || cols.some(c => String(row[c] ?? '').toLowerCase().includes(q));
  });
  const selectVisible = () => {
    const next = { ...selectedIds };
    filtered.forEach(r => next[Number(r.id)] = true);
    setSelectedIds(next);
  };
  const clearVisible = () => {
    const next = { ...selectedIds };
    filtered.forEach(r => delete next[Number(r.id)]);
    setSelectedIds(next);
  };

  return (
    <div className="space-y-4">
      <div className="command-hero">
        <div className="flex flex-col gap-4 lg:flex-row lg:items-center lg:justify-between">
          <div className="min-w-0">
            <div className="flex flex-wrap gap-2">
              <span className="badge badge-dark">⌕ Record Command Deck</span>
              <span className="badge badge-cyan">{records.length}/{scanCount}</span>
              <span className="badge badge-green">{selectedCount} selected</span>
            </div>
            <h2 className="mt-3 text-3xl font-black tracking-tight text-white">Pilih record seperti memilih asset</h2>
            <p className="mt-1 max-w-2xl text-sm leading-6 text-slate-300">Tidak ada tabel mentah di tahap keputusan. Scan record tampil sebagai kartu checklist; detail spreadsheet baru muncul setelah data masuk editor.</p>
          </div>
          <div className="grid min-w-[230px] grid-cols-2 gap-2">
            <Metric label="Visible" value={String(filtered.length)} />
            <Metric label="Selected" value={String(selectedCount)} />
          </div>
        </div>
      </div>

      {!records.length ? (
        <div className="empty-state min-h-[56vh]">
          <div className="text-6xl">⌕</div>
          <h2 className="mt-4 text-2xl font-black">Belum ada record discan</h2>
          <p className="mt-2 max-w-xl text-slate-300">Pilih model dari checklist di panel kiri, klik Scan Record, lalu centang record yang ingin masuk XLSX.</p>
        </div>
      ) : (
        <div className="space-y-3">
          <div className="record-toolbar">
            <label className="search-shell">
              <span>Search</span>
              <input value={query} onChange={e => setQuery(e.target.value)} placeholder="Cari nama, email, telepon, ID, atau teks record..." />
            </label>
            <div className="flex flex-wrap gap-2">
              <button className="btn btn-ghost" onClick={selectVisible}>Select visible</button>
              <button className="btn btn-ghost" onClick={clearVisible}>Clear visible</button>
              <button className="btn btn-ghost" onClick={() => setSelectedIds({})}>Clear all</button>
            </div>
          </div>

          <div className="record-grid studio-scroll">
            {filtered.slice(0, 240).map(row => {
              const id = Number(row.id);
              const selected = Boolean(selectedIds[id]);
              const title = String(row.display_name || row.name || row.email || `Record #${id}`);
              const subtitle = String(row.email || row.phone || row.mobile || row.city || row.state_id || 'Odoo record');
              return (
                <button key={id} type="button" onClick={() => setSelectedIds({ ...selectedIds, [id]: !selected })} className={`record-card ${selected ? 'record-card-on' : ''}`}>
                  <span className={`check-ring ${selected ? 'check-ring-on' : ''}`}>{selected ? '✓' : ''}</span>
                  <span className="min-w-0 flex-1 text-left">
                    <span className="flex items-center justify-between gap-3">
                      <span className="truncate text-sm font-black text-white">{title}</span>
                      <span className="rounded-full border border-white/10 bg-white/5 px-2 py-1 text-[10px] font-black text-cyan-100">#{id}</span>
                    </span>
                    <span className="mt-1 block truncate text-xs text-slate-400">{subtitle}</span>
                    <span className="mt-3 grid gap-1.5">
                      {cols.slice(0, 4).map(c => row[c] ? (
                        <span key={c} className="record-meta">
                          <b>{c}</b><span>{String(row[c])}</span>
                        </span>
                      ) : null)}
                    </span>
                  </span>
                </button>
              );
            })}
          </div>
          {filtered.length > 240 && <div className="alert-card alert-amber">Ditampilkan 240 record pertama dari hasil filter agar browser tetap ringan. Persempit pencarian atau gunakan Load More bertahap.</div>}
        </div>
      )}
    </div>
  );
}

function LogPanel({ logs, compact }: { logs: LogItem[]; compact: boolean }) {
  if (compact) {
    const last = logs[0];
    return (
      <div className="rail-card min-h-0">
        <div className="rail-title">Log</div>
        <div className={`log-orb ${last?.level || 'info'}`}>{logs.length}</div>
        <div className="rail-caption">{last ? last.level.toUpperCase() : 'idle'}</div>
      </div>
    );
  }
  return (
    <div className="control-card">
      <div className="flex items-center justify-between">
        <h2 className="text-xl font-black">Activity Log</h2>
        <span className="badge badge-muted">{logs.length}</span>
      </div>
      <div className="mt-4 max-h-80 space-y-2 overflow-auto pr-1 studio-scroll">
        {logs.length === 0 ? <div className="rounded-2xl border border-white/10 bg-white/5 p-4 text-sm text-slate-400">Belum ada aktivitas.</div> : logs.map((log, i) => (
          <details key={i} className={`log-card ${log.level}`}>
            <summary className="cursor-pointer font-semibold">[{log.time}] {log.message}</summary>
            {log.detail && <pre className="mt-2 max-h-56 overflow-auto whitespace-pre-wrap text-xs text-slate-300">{JSON.stringify(log.detail, null, 2)}</pre>}
          </details>
        ))}
      </div>
    </div>
  );
}

function EmptyState() {
  return (
    <div className="empty-state min-h-[64vh]">
      <div className="mx-auto grid h-24 w-24 place-items-center rounded-[32px] border border-cyan-300/20 bg-cyan-300/10 text-5xl shadow-2xl shadow-cyan-950/40">⌁</div>
      <h2 className="mt-6 text-3xl font-black text-white">Pilih workflow di kiri</h2>
      <p className="mx-auto mt-3 max-w-xl text-slate-300">Upload XLSX untuk import/editor, atau masuk Export untuk scan record Odoo lalu membuat XLSX baru. Begitu data aktif, panel kiri otomatis menjadi compact agar workspace fokus.</p>
      <div className="mt-6 flex flex-wrap justify-center gap-2">
        <span className="badge badge-cyan">schema-aware</span>
        <span className="badge badge-green">mobile-first</span>
        <span className="badge badge-muted">batch import</span>
      </div>
    </div>
  );
}
