'use client';

import React, { ChangeEvent, useEffect, useMemo, useState } from 'react';
import * as XLSX from 'xlsx';

type Conn = { url: string; db: string; username: string; password: string };
type Row = Record<string, any>;
type EditorKind = 'product' | 'contact' | 'project' | 'knowledge' | 'sales' | 'dynamic' | 'helper';
type Mission = 'home' | 'import' | 'export' | 'review' | 'settings';
type ImportStep = 'upload' | 'review' | 'editor' | 'execute';
type ExportStep = 'choose' | 'records' | 'fields' | 'preview';
type OdooField = { string?: string; type?: string; required?: boolean; readonly?: boolean; relation?: string; selection?: Array<[string, string]> };
type SheetState = { name: string; model: string; rows: Row[]; columns: string[]; kind: EditorKind; helper: boolean };
type Issue = { level: 'error' | 'warn' | 'ok'; title: string; detail?: string };
type LogItem = { time: string; level: 'info' | 'ok' | 'warn' | 'error'; message: string; detail?: any };
type ModelPreset = { key: string; label: string; model: string; kind: EditorKind; description: string; fields: string; risk: 'safe' | 'careful' | 'advanced'; icon: string };
type ExportMode = 'single' | 'bundle';
type BundlePreset = { key: string; label: string; primaryModel: string; kind: EditorKind; description: string; fields: string; icon: string; includes?: Array<{ key: string; label: string; default?: boolean }> };

const STORAGE_KEY = 'studio2_v10_2_settings';
const SAFE_LIMIT = 32000;
const HELPER_SHEETS = new Set(['readme', 'readme_import', 'dashboard', 'validation_report', 'task_database', 'relationship_map', 'readme_export', 'README_EXPORT'.toLowerCase(), 'chatter_project', 'chatter_tasks', 'task_hierarchy', 'logs']);
const META_COLUMNS = new Set(['_model', '__action', '_external_id', 'external_id', 'id', 'x_studio2_odoo_id', '__rownum__', '_studio2_truncated_fields', '_studio2_note', '_studio2_error']);
const defaultConn: Conn = { url: '', db: '', username: '', password: '' };

const MODEL_PRESETS: ModelPreset[] = [
  { key: 'contacts', label: 'Contacts', model: 'res.partner', kind: 'contact', description: 'Pelanggan, supplier, agen, UMKM, vendor, dan member.', fields: 'name,display_name,email,phone,mobile,street,street2,city,state_id,country_id,customer_rank,supplier_rank,is_company,category_id,comment', risk: 'safe', icon: '◎' },
  { key: 'products', label: 'Products', model: 'product.template', kind: 'product', description: 'Produk, harga, barcode, kategori, vendor, foto URL, dan publish.', fields: 'name,default_code,barcode,list_price,standard_price,categ_id,public_categ_ids,sale_ok,purchase_ok,website_published,description_sale,image_1920', risk: 'careful', icon: '◈' },
  { key: 'projects', label: 'Projects', model: 'project.project', kind: 'project', description: 'Project utama: Ground Zero, Soraya Kitchen, Pilot, Ekspansi.', fields: 'name,display_name,partner_id,user_id,active,date_start,date,description,privacy_visibility,stage_id', risk: 'safe', icon: '▦' },
  { key: 'tasks', label: 'Tasks', model: 'project.task', kind: 'project', description: 'Task, subtask, parent, stage, deadline, PIC, hierarchy.', fields: 'name,display_name,project_id,parent_id,stage_id,user_ids,partner_id,date_deadline,priority,sequence,description', risk: 'careful', icon: '☷' },
  { key: 'knowledge', label: 'Knowledge', model: 'knowledge.article', kind: 'knowledge', description: 'SOP, dokumentasi project, rujukan operasional, dan artikel Knowledge.', fields: 'name,display_name,parent_id,body,body_html,create_date,write_date', risk: 'advanced', icon: '✦' },
  { key: 'sales', label: 'Sales', model: 'sale.order', kind: 'sales', description: 'Sales order, customer, state, invoice, total, dan tanggal.', fields: 'name,partner_id,date_order,state,invoice_status,amount_untaxed,amount_tax,amount_total,validity_date,note', risk: 'advanced', icon: '₿' },
  { key: 'categories', label: 'Categories', model: 'product.category', kind: 'product', description: 'Kategori teknis internal produk Lokalmart.', fields: 'name,display_name,parent_id,complete_name,property_cost_method,property_valuation', risk: 'safe', icon: '◇' },
  { key: 'web_categories', label: 'Web Categories', model: 'product.public.category', kind: 'product', description: 'Kategori katalog website / ecommerce.', fields: 'name,display_name,parent_id,sequence,website_id', risk: 'safe', icon: '✧' }
];


const BUNDLE_PRESETS: BundlePreset[] = [
  { key: 'project', label: 'Project Bundle', primaryModel: 'project.project', kind: 'project', icon: '▦', description: 'Project utama beserta task, subtask, milestone, update, stage, partner, dan responsible user.', fields: 'name,display_name,partner_id,user_id,date_start,date,description,privacy_visibility,stage_id,active' },
  { key: 'contact', label: 'Contact Bundle', primaryModel: 'res.partner', kind: 'contact', icon: '◎', description: 'Contact/customer/supplier beserta child address, tag, dan relasi bisnis opsional.', fields: 'name,display_name,email,phone,mobile,street,street2,city,state_id,country_id,customer_rank,supplier_rank,is_company,category_id,comment', includes: [{ key: 'sales', label: 'Ikutkan sales order terkait' }, { key: 'projects', label: 'Ikutkan task/project terkait' }] },
  { key: 'product', label: 'Product Bundle', primaryModel: 'product.template', kind: 'product', icon: '◈', description: 'Produk beserta varian, vendor/supplierinfo, kategori teknis, kategori website, dan satuan.', fields: 'name,display_name,default_code,barcode,list_price,standard_price,categ_id,public_categ_ids,uom_id,uom_po_id,sale_ok,purchase_ok,website_published,description_sale,active' },
  { key: 'sales', label: 'Sales Bundle', primaryModel: 'sale.order', kind: 'sales', icon: '₿', description: 'Sales order beserta order line, customer, product, dan user sales.', fields: 'name,display_name,partner_id,date_order,state,amount_total,invoice_status,user_id,validity_date,note' },
  { key: 'knowledge', label: 'Knowledge Bundle', primaryModel: 'knowledge.article', kind: 'knowledge', icon: '✦', description: 'Artikel Knowledge beserta child article dan parent reference jika ada.', fields: 'name,display_name,parent_id,body,body_html,create_date,write_date' }
];

function now() { return new Date().toLocaleTimeString('id-ID', { hour12: false }); }
function compact(text: any, max = 72) { const s = String(text ?? ''); return s.length > max ? `${s.slice(0, max)}…` : s; }
function parseCsvFields(text: string) { return text.split(',').map(x => x.trim()).filter(Boolean); }
function dedupe<T>(arr: T[]) { return Array.from(new Set(arr)); }
function sanitizeWorkbookValue(value: any) {
  if (value === null || value === undefined) return '';
  if (Array.isArray(value) || (typeof value === 'object' && value !== null)) return JSON.stringify(value).slice(0, SAFE_LIMIT);
  const text = String(value);
  return text.length > SAFE_LIMIT ? text.slice(0, SAFE_LIMIT) : value;
}
function normalizeRows(rows: Row[]) { return rows.map((row, index) => ({ ...row, __rownum__: row.__rownum__ || index + 2 })); }
function makeColumns(rows: Row[]) { const set = new Set<string>(); rows.forEach(row => Object.keys(row).forEach(key => set.add(key))); return Array.from(set); }
function detectModel(sheetName: string, rows: Row[]) {
  const rowModel = rows.find(row => row._model)?._model;
  if (rowModel) return String(rowModel).trim();
  const s = sheetName.toLowerCase().trim();
  const known: Record<string, string> = { contacts: 'res.partner', contact: 'res.partner', partner: 'res.partner', partners: 'res.partner', product: 'product.template', products: 'product.template', project: 'project.project', projects: 'project.project', task: 'project.task', tasks: 'project.task', knowledge: 'knowledge.article', articles: 'knowledge.article', sales: 'sale.order', orders: 'sale.order' };
  return known[s] || sheetName;
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
function rowsToSheetState(name: string, rawRows: Row[]): SheetState {
  const rows = normalizeRows(rawRows);
  const model = detectModel(name, rows);
  const kind = detectKind(model, name);
  const helper = kind === 'helper' || HELPER_SHEETS.has(name.toLowerCase());
  const columns = makeColumns(rows);
  if (!helper) {
    if (!columns.includes('_model')) columns.unshift('_model');
    if (!columns.includes('__action')) columns.unshift('__action');
    if (!columns.includes('_external_id')) columns.unshift('_external_id');
  }
  return { name, model, rows, columns, kind, helper };
}
function importantFields(kind: EditorKind, model: string) {
  if (kind === 'contact') return ['name', 'display_name', 'phone', 'mobile', 'email', 'street', 'city', 'supplier_rank', 'customer_rank', 'is_company', 'comment'];
  if (kind === 'product') return ['name', 'default_code', 'barcode', 'list_price', 'standard_price', 'categ_id', 'public_categ_ids', 'website_published', 'description_sale', 'image_1920'];
  if (kind === 'project') return ['name', 'project_id', 'parent_id', 'stage_id', 'user_id', 'user_ids', 'partner_id', 'date_deadline', 'description', 'priority', 'sequence'];
  if (kind === 'knowledge') return ['name', 'body', 'body_html', 'parent_id', 'category'];
  if (kind === 'sales') return ['name', 'partner_id', 'date_order', 'state', 'amount_total', 'invoice_status'];
  return ['display_name', 'name', 'create_date', 'write_date'];
}
function labelKind(kind: EditorKind) {
  return ({ product: 'Product', contact: 'Contact', project: 'Project', knowledge: 'Knowledge', sales: 'Sales', dynamic: 'Dynamic', helper: 'Context' } as Record<EditorKind, string>)[kind];
}
function iconKind(kind: EditorKind) {
  return ({ product: '◈', contact: '◎', project: '▦', knowledge: '✦', sales: '₿', dynamic: '◇', helper: '✧' } as Record<EditorKind, string>)[kind];
}
function chunk<T>(arr: T[], size: number) { const out: T[][] = []; for (let i = 0; i < arr.length; i += size) out.push(arr.slice(i, i + size)); return out; }

export default function HomePage() {
  const [mission, setMission] = useState<Mission>('home');
  const [importStep, setImportStep] = useState<ImportStep>('upload');
  const [exportStep, setExportStep] = useState<ExportStep>('choose');
  const [conn, setConn] = useState<Conn>(defaultConn);
  const [busy, setBusy] = useState(false);
  const [logs, setLogs] = useState<LogItem[]>([]);
  const [logOpen, setLogOpen] = useState(false);
  const [sheets, setSheets] = useState<SheetState[]>([]);
  const [activeSheetIndex, setActiveSheetIndex] = useState(0);
  const [schema, setSchema] = useState<Record<string, OdooField> | null>(null);
  const [schemaModel, setSchemaModel] = useState('');
  const [selectedRows, setSelectedRows] = useState<Record<number, boolean>>({});
  const [editorMode, setEditorMode] = useState<'cards' | 'grid'>('cards');
  const [batchSize, setBatchSize] = useState(20);
  const [selectedModelKeys, setSelectedModelKeys] = useState<Record<string, boolean>>({ contacts: true });
  const [exportMode, setExportMode] = useState<ExportMode>('bundle');
  const [activeBundleKey, setActiveBundleKey] = useState('project');
  const [bundleIncludes, setBundleIncludes] = useState<Record<string, boolean>>({});
  const [activeModel, setActiveModel] = useState('project.project');
  const [exportFields, setExportFields] = useState(BUNDLE_PRESETS[0].fields);
  const [customModel, setCustomModel] = useState('');
  const [customFields, setCustomFields] = useState('name,display_name,create_date,write_date');
  const [domain, setDomain] = useState('[]');
  const [scanRecords, setScanRecords] = useState<Row[]>([]);
  const [scanCount, setScanCount] = useState(0);
  const [scanOffset, setScanOffset] = useState(0);
  const [selectedIds, setSelectedIds] = useState<Record<number, boolean>>({});
  const activeSheet = sheets[activeSheetIndex];

  useEffect(() => {
    try { const raw = localStorage.getItem(STORAGE_KEY); if (raw) setConn({ ...defaultConn, ...JSON.parse(raw) }); } catch {}
  }, []);
  useEffect(() => { try { localStorage.setItem(STORAGE_KEY, JSON.stringify(conn)); } catch {} }, [conn]);
  useEffect(() => { setSchema(null); setSchemaModel(''); setSelectedRows({}); }, [activeSheetIndex]);

  function addLog(level: LogItem['level'], message: string, detail?: any) { setLogs(prev => [{ time: now(), level, message, detail }, ...prev].slice(0, 180)); }
  async function callOdoo(payload: Row) {
    const res = await fetch('/api/odoo', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ ...payload, connection: conn }) });
    const data = await res.json();
    if (!data.ok) throw new Error(data.error || 'Odoo API error');
    return data;
  }
  async function testConnection() {
    setBusy(true);
    try { const data = await callOdoo({ action: 'test' }); addLog('ok', `Target Odoo tersambung. UID ${data.uid}, contacts ${data.partner_count}.`); }
    catch (e: any) { addLog('error', `Koneksi gagal: ${e.message}`); }
    finally { setBusy(false); }
  }
  async function loadSchema(model = activeSheet?.model) {
    if (!model) return;
    setBusy(true);
    try { const data = await callOdoo({ action: 'schema', model }); setSchema(data.fields); setSchemaModel(model); addLog('ok', `Schema ${model} dimuat: ${Object.keys(data.fields || {}).length} field.`); }
    catch (e: any) { addLog('error', `Schema gagal: ${e.message}`); }
    finally { setBusy(false); }
  }
  async function handleFile(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (!file) return;
    setBusy(true);
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: 'array', cellDates: true });
      const nextSheets = wb.SheetNames.map(name => rowsToSheetState(name, XLSX.utils.sheet_to_json<Row>(wb.Sheets[name], { defval: '' })));
      setSheets(nextSheets);
      setActiveSheetIndex(0);
      setImportStep('review');
      setMission('import');
      addLog('ok', `${file.name} dibaca: ${nextSheets.length} sheet.`);
    } catch (e: any) { addLog('error', `Gagal membaca XLSX: ${e.message}`); }
    finally { setBusy(false); event.target.value = ''; }
  }
  function updateCell(rowIndex: number, column: string, value: any) {
    setSheets(prev => prev.map((sheet, sheetIndex) => {
      if (sheetIndex !== activeSheetIndex) return sheet;
      const rows = sheet.rows.map((row, index) => index === rowIndex ? { ...row, [column]: value } : row);
      const columns = sheet.columns.includes(column) ? sheet.columns : [...sheet.columns, column];
      return { ...sheet, rows, columns };
    }));
  }
  function addRow() {
    if (!activeSheet) return;
    setSheets(prev => prev.map((sheet, index) => index === activeSheetIndex ? { ...sheet, rows: [...sheet.rows, { _model: sheet.model, __action: 'upsert', __rownum__: sheet.rows.length + 2 }] } : sheet));
  }
  function addColumn() {
    if (!activeSheet) return;
    const name = prompt('Nama kolom Odoo / kolom XLSX baru:');
    if (!name) return;
    setSheets(prev => prev.map((sheet, index) => index === activeSheetIndex ? { ...sheet, columns: sheet.columns.includes(name) ? sheet.columns : [...sheet.columns, name], rows: sheet.rows.map(row => ({ ...row, [name]: row[name] ?? '' })) } : sheet));
  }
  function validateSheet(sheet = activeSheet): Issue[] {
    if (!sheet) return [];
    const issues: Issue[] = [];
    if (sheet.helper) return [{ level: 'ok', title: 'Context/helper sheet', detail: 'Sheet ini dibaca sebagai konteks, bukan kandidat import ke Odoo.' }];
    if (!sheet.model) issues.push({ level: 'error', title: 'Model belum jelas', detail: 'Isi kolom _model atau ubah nama sheet menjadi technical model Odoo.' });
    const missingAction = sheet.rows.filter(row => !row.__action).length;
    const missingExternal = sheet.rows.filter(row => !row._external_id && !row.x_studio2_odoo_id && !row.id).length;
    if (missingAction) issues.push({ level: 'warn', title: `${missingAction} row belum punya __action`, detail: 'Studio akan memakai upsert sebagai default, tapi lebih baik tetap eksplisit.' });
    if (missingExternal) issues.push({ level: 'warn', title: `${missingExternal} row belum punya external ID`, detail: 'Aman untuk create, kurang aman untuk update ulang.' });
    if (schema && schemaModel === sheet.model) {
      const unknown = sheet.columns.filter(c => !META_COLUMNS.has(c) && !c.endsWith('_external_id') && !c.endsWith('_external_ids') && !schema[c]);
      if (unknown.length) issues.push({ level: 'warn', title: `${unknown.length} kolom tidak dikenal`, detail: unknown.slice(0, 18).join(', ') });
      const required = Object.entries(schema).filter(([, v]) => v.required && !v.readonly).map(([k]) => k);
      const missingRequired = required.filter(f => !sheet.columns.includes(f) && !sheet.columns.includes(`${f}_external_id`));
      if (missingRequired.length) issues.push({ level: 'error', title: `${missingRequired.length} field wajib belum ada`, detail: missingRequired.slice(0, 18).join(', ') });
    } else {
      issues.push({ level: 'warn', title: 'Schema belum dimuat', detail: 'Klik Cek Schema supaya Studio bisa membedakan field valid, required, relation, dan readonly.' });
    }
    if (!issues.length) issues.push({ level: 'ok', title: 'Sheet terlihat aman', detail: 'Tidak ada error fatal dari pemeriksaan cepat.' });
    return issues;
  }
  const issues = useMemo(() => validateSheet(activeSheet), [activeSheet, schema, schemaModel]);
  const visibleColumns = useMemo(() => {
    if (!activeSheet) return [];
    const priority = ['_model', '__action', '_external_id', 'x_studio2_odoo_id', ...importantFields(activeSheet.kind, activeSheet.model)];
    const set = new Set(activeSheet.columns);
    return [...priority.filter(c => set.has(c)), ...activeSheet.columns.filter(c => !priority.includes(c) && c !== '__rownum__')];
  }, [activeSheet]);
  function downloadWorkbook(name = 'studio2_dataset.xlsx') {
    if (!sheets.length) return;
    const wb = XLSX.utils.book_new();
    sheets.forEach(sheet => {
      const rows = sheet.rows.map(row => { const out: Row = {}; sheet.columns.forEach(col => out[col] = sanitizeWorkbookValue(row[col])); return out; });
      const ws = XLSX.utils.json_to_sheet(rows.length ? rows : [{}], { header: sheet.columns.length ? sheet.columns : undefined });
      XLSX.utils.book_append_sheet(wb, ws, sheet.name.slice(0, 31) || 'Sheet1');
    });
    XLSX.writeFile(wb, name);
    addLog('ok', `Download XLSX: ${name}`);
  }
  async function importActiveSheet() {
    if (!activeSheet) return;
    if (activeSheet.helper) { addLog('warn', 'Sheet context/helper tidak diimport.'); return; }
    const chosen = activeSheet.rows.filter((_, index) => Object.keys(selectedRows).length ? selectedRows[index] : true);
    if (!chosen.length) { addLog('warn', 'Tidak ada row dipilih.'); return; }
    setBusy(true);
    const parts = chunk(chosen, Math.max(1, Math.min(50, batchSize)));
    addLog('info', `Mulai import ${chosen.length} row dalam ${parts.length} batch.`);
    let created = 0, updated = 0, failed = 0;
    for (let i = 0; i < parts.length; i++) {
      try {
        const data = await callOdoo({ action: 'import_batch', model: activeSheet.model, rows: parts[i] });
        created += Number(data.created || 0); updated += Number(data.updated || 0); failed += Number(data.failed || 0);
        addLog(data.failed ? 'warn' : 'ok', `Batch ${i + 1}/${parts.length}: create ${data.created}, update ${data.updated}, failed ${data.failed}.`, data.results);
        if (data.failed) break;
      } catch (e: any) { failed += parts[i].length; addLog('error', `Batch ${i + 1} gagal: ${e.message}`); break; }
    }
    addLog(failed ? 'warn' : 'ok', `Import selesai. Created ${created}, updated ${updated}, failed ${failed}.`);
    setBusy(false);
  }
  function togglePreset(preset: ModelPreset) {
    setSelectedModelKeys(prev => ({ ...prev, [preset.key]: !prev[preset.key] }));
    setActiveModel(preset.model);
    setExportFields(preset.fields);
    setScanRecords([]); setSelectedIds({}); setScanOffset(0); setScanCount(0);
  }
  function choosePreset(preset: ModelPreset) {
    setSelectedModelKeys(prev => ({ ...prev, [preset.key]: true }));
    setActiveModel(preset.model);
    setExportFields(preset.fields);
    setScanRecords([]); setSelectedIds({}); setScanOffset(0); setScanCount(0);
  }
  function chooseBundle(bundle: BundlePreset) {
    setExportMode('bundle');
    setActiveBundleKey(bundle.key);
    setActiveModel(bundle.primaryModel);
    setExportFields(bundle.fields);
    const defaults: Record<string, boolean> = {};
    (bundle.includes || []).forEach(item => defaults[item.key] = Boolean(item.default));
    setBundleIncludes(defaults);
    setScanRecords([]); setSelectedIds({}); setScanOffset(0); setScanCount(0);
  }
  async function scanModel(reset = true) {
    setBusy(true);
    try {
      let parsedDomain: any[] = [];
      try { parsedDomain = JSON.parse(domain || '[]'); } catch { throw new Error('Domain harus JSON array, contoh: []'); }
      const offset = reset ? 0 : scanOffset;
      const fields = dedupe([...importantFields(detectKind(activeModel, activeModel), activeModel), ...parseCsvFields(exportFields)]).filter(Boolean);
      const data = await callOdoo({ action: 'record_scan', model: activeModel, fields, domain: parsedDomain, offset, limit: 80 });
      setScanCount(data.count || 0);
      setScanOffset(offset + (data.records?.length || 0));
      setScanRecords(prev => reset ? data.records : [...prev, ...data.records]);
      if (reset) setSelectedIds({});
      setExportStep('records');
      addLog('ok', `Scan ${activeModel}: ${data.records.length} dari ${data.count} record.`);
    } catch (e: any) { addLog('error', `Scan gagal: ${e.message}`); }
    finally { setBusy(false); }
  }
  async function exportSelectedRecords() {
    const ids = Object.entries(selectedIds).filter(([, v]) => v).map(([id]) => Number(id));
    if (!ids.length) { addLog('warn', 'Belum ada record dipilih untuk export.'); return; }
    setBusy(true);
    try {
      const data = await callOdoo({ action: 'export_records', model: activeModel, ids, fields: parseCsvFields(exportFields) });
      const sheet = rowsToSheetState(data.sheet || activeModel, data.rows || []);
      setSheets([sheet]);
      setActiveSheetIndex(0);
      setMission('review');
      setImportStep('editor');
      setExportStep('preview');
      addLog('ok', `${ids.length} record ${activeModel} diexport ke editor.`);
    } catch (e: any) { addLog('error', `Export gagal: ${e.message}`); }
    finally { setBusy(false); }
  }
  async function exportSelectedBundle() {
    const ids = Object.entries(selectedIds).filter(([, v]) => v).map(([id]) => Number(id));
    if (!ids.length) { addLog('warn', 'Belum ada record utama dipilih untuk Smart Bundle.'); return; }
    setBusy(true);
    try {
      const data = await callOdoo({ action: 'export_bundle', bundle: activeBundle.key, primary_model: activeBundle.primaryModel, ids, includes: bundleIncludes, fields: parseCsvFields(exportFields) });
      const nextSheets = Object.entries(data.sheets || {}).map(([name, rows]) => rowsToSheetState(name, rows as Row[]));
      setSheets(nextSheets);
      setActiveSheetIndex(0);
      setMission('review');
      setImportStep('editor');
      setExportStep('preview');
      addLog('ok', `${activeBundle.label}: ${ids.length} record utama menghasilkan ${nextSheets.length} sheet.`, data.summary);
    } catch (e: any) { addLog('error', `Smart Bundle gagal: ${e.message}`); }
    finally { setBusy(false); }
  }
  async function exportProject() {
    const id = prompt('Masukkan ID project Odoo, contoh: 73');
    if (!id) return;
    setBusy(true);
    try {
      const data = await callOdoo({ action: 'export_project', project_id: Number(id) });
      const nextSheets = Object.entries(data.sheets || {}).map(([name, rows]) => rowsToSheetState(name, rows as Row[]));
      setSheets(nextSheets); setActiveSheetIndex(0); setMission('review'); setImportStep('editor');
      addLog('ok', `Project ${id} diexport ke ${nextSheets.length} sheet.`);
    } catch (e: any) { addLog('error', `Export project gagal: ${e.message}`); }
    finally { setBusy(false); }
  }
  const connectionReady = Boolean(conn.url && conn.db && conn.username && conn.password);
  const selectedRecordCount = Object.values(selectedIds).filter(Boolean).length;
  const selectedImportCount = Object.values(selectedRows).filter(Boolean).length;
  const activePreset = MODEL_PRESETS.find(p => p.model === activeModel);
  const activeBundle = BUNDLE_PRESETS.find(b => b.key === activeBundleKey) || BUNDLE_PRESETS[0];

  return (
    <main className="min-h-screen pb-24 text-white md:pb-8">
      <div className="orb orb-a" />
      <div className="orb orb-b" />
      <div className="mx-auto flex min-h-screen w-full max-w-6xl flex-col gap-4 px-4 py-4 sm:px-6 lg:px-8">
        <TopBar mission={mission} setMission={setMission} connectionReady={connectionReady} busy={busy} openLogs={() => setLogOpen(true)} />
        {mission === 'home' && <HomeMission connectionReady={connectionReady} setMission={setMission} sheets={sheets} logs={logs} scanCount={scanCount} />}
        {mission === 'settings' && <SettingsScreen conn={conn} setConn={setConn} testConnection={testConnection} busy={busy} connectionReady={connectionReady} />}
        {mission === 'import' && <ImportMission step={importStep} setStep={setImportStep} busy={busy} sheets={sheets} activeSheet={activeSheet} activeSheetIndex={activeSheetIndex} setActiveSheetIndex={setActiveSheetIndex} handleFile={handleFile} issues={issues} selectedImportCount={selectedImportCount} loadSchema={() => loadSchema()} setMission={setMission} />}
        {mission === 'export' && <ExportMission step={exportStep} setStep={setExportStep} busy={busy} presets={MODEL_PRESETS} bundles={BUNDLE_PRESETS} exportMode={exportMode} setExportMode={setExportMode} activeBundle={activeBundle} activeBundleKey={activeBundleKey} chooseBundle={chooseBundle} bundleIncludes={bundleIncludes} setBundleIncludes={setBundleIncludes} selectedModelKeys={selectedModelKeys} togglePreset={togglePreset} choosePreset={choosePreset} activeModel={activeModel} activePreset={activePreset} exportFields={exportFields} setExportFields={setExportFields} domain={domain} setDomain={setDomain} customModel={customModel} setCustomModel={setCustomModel} customFields={customFields} setCustomFields={setCustomFields} scanModel={() => scanModel(true)} loadMore={() => scanModel(false)} exportSelectedRecords={exportSelectedRecords} exportSelectedBundle={exportSelectedBundle} exportProject={exportProject} scanRecords={scanRecords} scanCount={scanCount} scanOffset={scanOffset} selectedIds={selectedIds} setSelectedIds={setSelectedIds} selectedRecordCount={selectedRecordCount} />}
        {mission === 'review' && <ReviewMission activeSheet={activeSheet} activeSheetIndex={activeSheetIndex} sheets={sheets} setActiveSheetIndex={setActiveSheetIndex} issues={issues} columns={visibleColumns} schema={schemaModel === activeSheet?.model ? schema : null} editorMode={editorMode} setEditorMode={setEditorMode} selectedRows={selectedRows} setSelectedRows={setSelectedRows} updateCell={updateCell} addColumn={addColumn} addRow={addRow} loadSchema={() => loadSchema()} download={() => downloadWorkbook('studio2_edited.xlsx')} importActiveSheet={importActiveSheet} batchSize={batchSize} setBatchSize={setBatchSize} busy={busy} />}
      </div>
      <BottomNav mission={mission} setMission={setMission} />
      <LogSheet logs={logs} open={logOpen} onClose={() => setLogOpen(false)} />
    </main>
  );
}

function TopBar({ mission, setMission, connectionReady, busy, openLogs }: { mission: Mission; setMission: (m: Mission) => void; connectionReady: boolean; busy: boolean; openLogs: () => void }) {
  return <header className="top-shell">
    <button className="brand" onClick={() => setMission('home')}>
      <span className="brand-mark">LM</span>
      <span><b>Lokalmart Studio</b><small>Odoo Data Command</small></span>
    </button>
    <div className="top-actions">
      <span className={connectionReady ? 'status ok' : 'status warn'}>{connectionReady ? 'Odoo ready' : 'Setup'}</span>
      {busy && <span className="status pulse">Working</span>}
      <button className={mission === 'settings' ? 'icon-btn active' : 'icon-btn'} onClick={() => setMission('settings')}>⚙</button>
      <button className="icon-btn" onClick={openLogs}>⌁</button>
    </div>
  </header>;
}

function HomeMission({ connectionReady, setMission, sheets, logs, scanCount }: { connectionReady: boolean; setMission: (m: Mission) => void; sheets: SheetState[]; logs: LogItem[]; scanCount: number }) {
  return <section className="mission-home">
    <div className="hero-copy">
      <span className="eyebrow">Studio2 v10.1 · Mission UX Rebuild</span>
      <h1>Kontrol data Odoo tanpa tenggelam di spreadsheet.</h1>
      <p>Mulai dari niat kerja: import data ke Odoo, export record dari Odoo, review hasil, lalu eksekusi aman per batch.</p>
      <div className="hero-stats">
        <Metric label="Target" value={connectionReady ? 'Ready' : 'Belum'} tone={connectionReady ? 'ok' : 'warn'} />
        <Metric label="Dataset" value={sheets.length ? `${sheets.length} sheet` : 'Kosong'} />
        <Metric label="Scan" value={scanCount ? `${scanCount} record` : 'Belum'} />
      </div>
    </div>
    <div className="mission-grid">
      <MissionCard icon="↑" title="Import Mission" desc="Upload XLSX, baca sheet, cek schema, edit row, lalu kirim batch kecil ke Odoo." cta="Mulai import" onClick={() => setMission('import')} />
      <MissionCard icon="↓" title="Export Mission" desc="Pilih model lewat checklist, scan record, pilih field, export ke editor, lalu download." cta="Mulai export" onClick={() => setMission('export')} />
      <MissionCard icon="◇" title="Review Workspace" desc="Buka dataset aktif sebagai object editor, bukan tabel mentah. Cocok untuk validasi sebelum import ulang." cta="Buka review" onClick={() => setMission('review')} muted />
      <MissionCard icon="⚙" title="Koneksi Odoo" desc="Simpan target Odoo di browser. Tidak masuk GitHub dan tidak perlu environment secret." cta="Atur koneksi" onClick={() => setMission('settings')} muted />
    </div>
    <div className="recent-panel">
      <div><b>Activity trail</b><small>Log teknis disembunyikan di sini supaya layar utama tetap tenang.</small></div>
      <div className="mini-log-list">{logs.slice(0, 4).map((log, idx) => <span key={idx} className={`log-pill ${log.level}`}>{log.time} · {compact(log.message, 54)}</span>)}{!logs.length && <span className="empty-note">Belum ada aktivitas.</span>}</div>
    </div>
  </section>;
}
function MissionCard({ icon, title, desc, cta, onClick, muted }: { icon: string; title: string; desc: string; cta: string; onClick: () => void; muted?: boolean }) {
  return <button className={muted ? 'mission-card muted' : 'mission-card'} onClick={onClick}><span className="card-icon">{icon}</span><span><b>{title}</b><small>{desc}</small><em>{cta}</em></span></button>;
}
function Metric({ label, value, tone }: { label: string; value: string; tone?: 'ok' | 'warn' }) { return <div className={`metric ${tone || ''}`}><small>{label}</small><b>{value}</b></div>; }

function SettingsScreen({ conn, setConn, testConnection, busy, connectionReady }: { conn: Conn; setConn: React.Dispatch<React.SetStateAction<Conn>>; testConnection: () => void; busy: boolean; connectionReady: boolean }) {
  return <section className="screen-stack">
    <SectionTitle overline="Connection" title="Target Odoo" desc="Credential disimpan di browser localStorage. Jangan upload password/API key ke GitHub." />
    <div className="settings-card">
      <Field label="Odoo URL" value={conn.url} onChange={v => setConn(c => ({ ...c, url: v }))} placeholder="https://edu-lokalmart.odoo.com" />
      <Field label="Database" value={conn.db} onChange={v => setConn(c => ({ ...c, db: v }))} placeholder="edu-lokalmart" />
      <Field label="Username" value={conn.username} onChange={v => setConn(c => ({ ...c, username: v }))} placeholder="email Odoo" />
      <Field label="Password / API key" type="password" value={conn.password} onChange={v => setConn(c => ({ ...c, password: v }))} placeholder="••••••" />
      <button className="primary-action wide" onClick={testConnection} disabled={busy}>{busy ? 'Mengetes…' : connectionReady ? 'Tes ulang koneksi' : 'Tes koneksi'}</button>
    </div>
  </section>;
}
function Field({ label, value, onChange, placeholder, type = 'text' }: { label: string; value: string; onChange: (v: string) => void; placeholder?: string; type?: string }) { return <label className="field"><span>{label}</span><input type={type} value={value} placeholder={placeholder} onChange={e => onChange(e.target.value)} /></label>; }
function SectionTitle({ overline, title, desc }: { overline: string; title: string; desc: string }) { return <div className="section-title"><span>{overline}</span><h2>{title}</h2><p>{desc}</p></div>; }

function ImportMission(props: { step: ImportStep; setStep: (s: ImportStep) => void; busy: boolean; sheets: SheetState[]; activeSheet?: SheetState; activeSheetIndex: number; setActiveSheetIndex: (i: number) => void; handleFile: (e: ChangeEvent<HTMLInputElement>) => void; issues: Issue[]; selectedImportCount: number; loadSchema: () => void; setMission: (m: Mission) => void }) {
  const { step, setStep, busy, sheets, activeSheet, activeSheetIndex, setActiveSheetIndex, handleFile, issues, loadSchema, setMission } = props;
  return <section className="screen-stack">
    <SectionTitle overline="Import Mission" title="Bawa XLSX ke Odoo dengan aman" desc="Upload dulu, Studio membaca sheet sebagai dataset, bukan langsung mengirim ke Odoo." />
    <StepRail steps={['Upload', 'Review', 'Editor', 'Execute']} active={['upload','review','editor','execute'].indexOf(step)} />
    {!sheets.length && <label className="drop-zone"><input type="file" accept=".xlsx,.xls" onChange={handleFile} disabled={busy} /><span>↑</span><b>Upload XLSX</b><small>File dibaca di browser. Setelah itu kamu pilih sheet dan cek schema.</small></label>}
    {!!sheets.length && <DatasetOverview sheets={sheets} activeSheetIndex={activeSheetIndex} setActiveSheetIndex={setActiveSheetIndex} />}
    {!!activeSheet && <div className="focus-card">
      <div className="object-head"><span className="object-icon">{iconKind(activeSheet.kind)}</span><div><b>{activeSheet.name}</b><small>{activeSheet.model} · {activeSheet.rows.length} rows · {activeSheet.columns.length} fields</small></div><span className="chip">{labelKind(activeSheet.kind)}</span></div>
      <IssueList issues={issues.slice(0, 3)} />
      <div className="action-row"><button className="secondary-action" onClick={loadSchema}>Cek schema</button><button className="primary-action" onClick={() => { setStep('editor'); setMission('review'); }}>Buka editor</button></div>
    </div>}
  </section>;
}
function StepRail({ steps, active }: { steps: string[]; active: number }) { return <div className="step-rail">{steps.map((s, i) => <span key={s} className={i <= active ? 'done' : ''}><b>{i+1}</b>{s}</span>)}</div>; }
function DatasetOverview({ sheets, activeSheetIndex, setActiveSheetIndex }: { sheets: SheetState[]; activeSheetIndex: number; setActiveSheetIndex: (i: number) => void }) {
  return <div className="sheet-carousel">{sheets.map((sheet, index) => <button key={sheet.name} className={index === activeSheetIndex ? 'sheet-card active' : 'sheet-card'} onClick={() => setActiveSheetIndex(index)}><b>{sheet.name}</b><small>{sheet.model}</small><em>{sheet.rows.length} row · {labelKind(sheet.kind)}</em></button>)}</div>;
}
function IssueList({ issues }: { issues: Issue[] }) { return <div className="issue-list">{issues.map((issue, idx) => <div key={idx} className={`issue ${issue.level}`}><b>{issue.title}</b>{issue.detail && <small>{issue.detail}</small>}</div>)}</div>; }

function ExportMission(props: { step: ExportStep; setStep: (s: ExportStep) => void; busy: boolean; presets: ModelPreset[]; bundles: BundlePreset[]; exportMode: ExportMode; setExportMode: (m: ExportMode) => void; activeBundle: BundlePreset; activeBundleKey: string; chooseBundle: (b: BundlePreset) => void; bundleIncludes: Record<string, boolean>; setBundleIncludes: React.Dispatch<React.SetStateAction<Record<string, boolean>>>; selectedModelKeys: Record<string, boolean>; togglePreset: (p: ModelPreset) => void; choosePreset: (p: ModelPreset) => void; activeModel: string; activePreset?: ModelPreset; exportFields: string; setExportFields: (v: string) => void; domain: string; setDomain: (v: string) => void; customModel: string; setCustomModel: (v: string) => void; customFields: string; setCustomFields: (v: string) => void; scanModel: () => void; loadMore: () => void; exportSelectedRecords: () => void; exportSelectedBundle: () => void; exportProject: () => void; scanRecords: Row[]; scanCount: number; scanOffset: number; selectedIds: Record<number, boolean>; setSelectedIds: React.Dispatch<React.SetStateAction<Record<number, boolean>>>; selectedRecordCount: number }) {
  const p = props;
  const smart = p.exportMode === 'bundle';
  const activeTitle = smart ? p.activeBundle.label : (p.activePreset?.label || 'Custom Model');
  const activeIcon = smart ? p.activeBundle.icon : (p.activePreset?.icon || '◇');
  const activeDesc = smart ? `${p.activeBundle.primaryModel} · relasi otomatis` : `${p.activeModel} · ${parseCsvFields(p.exportFields).length} fields`;
  return <section className="screen-stack">
    <SectionTitle overline="Export Mission" title="Pilih objek, bukan cuma tabel" desc="Single Model untuk export sederhana. Smart Bundle untuk membawa record utama beserta data Odoo yang menempel: task, line, kategori, vendor, child contact, dan referensi." />
    <StepRail steps={['Pilih Objek', 'Scan Record', 'Pilih Record', 'Preview']} active={['choose','records','fields','preview'].indexOf(p.step)} />
    <div className="mode-switch"><button className={p.exportMode === 'bundle' ? 'tab active' : 'tab'} onClick={() => { p.setExportMode('bundle'); p.chooseBundle(p.activeBundle); }}>Smart Bundle</button><button className={p.exportMode === 'single' ? 'tab active' : 'tab'} onClick={() => p.setExportMode('single')}>Single Model</button></div>

    {smart ? <>
      <div className="bundle-grid">{p.bundles.map(bundle => <button key={bundle.key} className={p.activeBundleKey === bundle.key ? 'bundle-card active' : 'bundle-card'} onClick={() => p.chooseBundle(bundle)}><span className="model-icon">{bundle.icon}</span><b>{bundle.label}</b><small>{bundle.description}</small><em>{bundle.primaryModel}</em></button>)}</div>
      {!!p.activeBundle.includes?.length && <div className="focus-card compact-card"><b>Relasi opsional</b><small>Default bundle tetap ringan. Centang hanya jika memang perlu dibawa.</small><div className="include-list">{p.activeBundle.includes.map(item => <label key={item.key}><input type="checkbox" checked={Boolean(p.bundleIncludes[item.key])} onChange={e => p.setBundleIncludes(prev => ({ ...prev, [item.key]: e.target.checked }))} />{item.label}</label>)}</div></div>}
    </> : <>
      <div className="model-grid">{p.presets.map(preset => <button key={preset.key} className={p.activeModel === preset.model ? 'model-card active' : 'model-card'} onClick={() => p.choosePreset(preset)}><span className="check" onClick={(e) => { e.stopPropagation(); p.togglePreset(preset); }}>{p.selectedModelKeys[preset.key] ? '✓' : ''}</span><span className="model-icon">{preset.icon}</span><b>{preset.label}</b><small>{preset.description}</small><em>{preset.model}</em><i className={`risk ${preset.risk}`}>{preset.risk}</i></button>)}</div>
    </>}

    <details className="advanced"><summary>Advanced: custom model, domain, dan field</summary><div className="advanced-grid"><Field label="Custom model" value={p.customModel} onChange={v => { p.setCustomModel(v); if (v) { p.setExportMode('single'); p.choosePreset({ key: 'custom', label: 'Custom', model: v, kind: 'dynamic', description: '', fields: p.customFields, risk: 'advanced', icon: '◇' }); } }} placeholder="x_custom_model" /><Field label="Domain JSON" value={p.domain} onChange={p.setDomain} placeholder='[]' /><label className="field wide"><span>Fields untuk record utama</span><textarea value={p.exportFields} onChange={e => p.setExportFields(e.target.value)} /></label></div></details>

    <div className="focus-card">
      <div className="object-head"><span className="object-icon">{activeIcon}</span><div><b>{activeTitle}</b><small>{activeDesc}</small></div><span className="chip">{smart ? 'Bundle aktif' : 'Model aktif'}</span></div>
      <div className="field-pills">{parseCsvFields(p.exportFields).slice(0, 12).map(f => <span key={f}>{f}</span>)}{parseCsvFields(p.exportFields).length > 12 && <span>+{parseCsvFields(p.exportFields).length - 12}</span>}</div>
      <div className="action-row"><button className="secondary-action" onClick={p.exportProject}>Project by ID lama</button><button className="primary-action" onClick={p.scanModel} disabled={p.busy}>{p.busy ? 'Scanning…' : smart ? 'Scan record utama' : 'Scan record'}</button></div>
    </div>
    {!!p.scanRecords.length && <RecordPicker records={p.scanRecords} count={p.scanCount} offset={p.scanOffset} selectedIds={p.selectedIds} setSelectedIds={p.setSelectedIds} loadMore={p.loadMore} exportSelected={smart ? p.exportSelectedBundle : p.exportSelectedRecords} selectedRecordCount={p.selectedRecordCount} exportLabel={smart ? `Export ${p.activeBundle.label}` : 'Export record'} />}
  </section>;
}
function RecordPicker({ records, count, offset, selectedIds, setSelectedIds, loadMore, exportSelected, selectedRecordCount, exportLabel = 'Export record' }: { records: Row[]; count: number; offset: number; selectedIds: Record<number, boolean>; setSelectedIds: React.Dispatch<React.SetStateAction<Record<number, boolean>>>; loadMore: () => void; exportSelected: () => void; selectedRecordCount: number; exportLabel?: string }) {
  const [q, setQ] = useState('');
  const visible = records.filter(r => JSON.stringify(r).toLowerCase().includes(q.toLowerCase()));
  return <div className="record-zone"><div className="record-head"><div><b>{count} record ditemukan</b><small>{selectedRecordCount} dipilih · {offset} sudah dimuat</small></div><input value={q} onChange={e => setQ(e.target.value)} placeholder="Cari record…" /></div><div className="record-list">{visible.map(row => { const id = Number(row.id); const title = row.display_name || row.name || row.email || `Record ${id}`; return <label key={id} className={selectedIds[id] ? 'record-card selected' : 'record-card'}><input type="checkbox" checked={Boolean(selectedIds[id])} onChange={e => setSelectedIds(prev => ({ ...prev, [id]: e.target.checked }))} /><span><b>{compact(title, 60)}</b><small>ID {id} · {compact(row.email || row.phone || row.mobile || row.default_code || row.create_date || '', 80)}</small></span></label>; })}</div><div className="action-row sticky-actions"><button className="secondary-action" onClick={loadMore}>Muat lagi</button><button className="primary-action" onClick={exportSelected}>{exportLabel} {selectedRecordCount || ''}</button></div></div>;
}

function ReviewMission(props: { activeSheet?: SheetState; activeSheetIndex: number; sheets: SheetState[]; setActiveSheetIndex: (i: number) => void; issues: Issue[]; columns: string[]; schema: Record<string, OdooField> | null; editorMode: 'cards' | 'grid'; setEditorMode: (m: 'cards' | 'grid') => void; selectedRows: Record<number, boolean>; setSelectedRows: React.Dispatch<React.SetStateAction<Record<number, boolean>>>; updateCell: (rowIndex: number, column: string, value: any) => void; addColumn: () => void; addRow: () => void; loadSchema: () => void; download: () => void; importActiveSheet: () => void; batchSize: number; setBatchSize: (n: number) => void; busy: boolean }) {
  const p = props;
  if (!p.activeSheet) return <section className="screen-stack"><SectionTitle overline="Review" title="Belum ada dataset aktif" desc="Upload XLSX atau export record dulu supaya editor terbuka." /></section>;
  return <section className="screen-stack review-screen">
    <SectionTitle overline="Review Workspace" title={p.activeSheet.name} desc={`${p.activeSheet.model} · ${p.activeSheet.rows.length} rows · ${p.activeSheet.columns.length} fields`} />
    <DatasetOverview sheets={p.sheets} activeSheetIndex={p.activeSheetIndex} setActiveSheetIndex={p.setActiveSheetIndex} />
    <div className="review-toolbar"><button className={p.editorMode === 'cards' ? 'tab active' : 'tab'} onClick={() => p.setEditorMode('cards')}>Object cards</button><button className={p.editorMode === 'grid' ? 'tab active' : 'tab'} onClick={() => p.setEditorMode('grid')}>Grid</button><button className="tab" onClick={p.loadSchema}>Schema</button><button className="tab" onClick={p.addColumn}>+ Field</button><button className="tab" onClick={p.addRow}>+ Row</button></div>
    <IssueList issues={p.issues} />
    {p.editorMode === 'cards' ? <ObjectEditor sheet={p.activeSheet} columns={p.columns} schema={p.schema} selectedRows={p.selectedRows} setSelectedRows={p.setSelectedRows} updateCell={p.updateCell} /> : <GridEditor sheet={p.activeSheet} columns={p.columns} updateCell={p.updateCell} selectedRows={p.selectedRows} setSelectedRows={p.setSelectedRows} />}
    <div className="execute-bar"><label><span>Batch</span><input type="number" min={1} max={50} value={p.batchSize} onChange={e => p.setBatchSize(Number(e.target.value) || 20)} /></label><button className="secondary-action" onClick={p.download}>Download XLSX</button><button className="primary-action" onClick={p.importActiveSheet} disabled={p.busy}>{p.busy ? 'Importing…' : 'Import sheet aktif'}</button></div>
  </section>;
}
function ObjectEditor({ sheet, columns, schema, selectedRows, setSelectedRows, updateCell }: { sheet: SheetState; columns: string[]; schema: Record<string, OdooField> | null; selectedRows: Record<number, boolean>; setSelectedRows: React.Dispatch<React.SetStateAction<Record<number, boolean>>>; updateCell: (rowIndex: number, column: string, value: any) => void }) {
  const groups = groupFields(sheet.kind, columns);
  return <div className="object-list">{sheet.rows.slice(0, 120).map((row, rowIndex) => <details className="object-row" key={rowIndex} open={rowIndex === 0}><summary><label onClick={e => e.stopPropagation()}><input type="checkbox" checked={Boolean(selectedRows[rowIndex])} onChange={e => setSelectedRows(prev => ({ ...prev, [rowIndex]: e.target.checked }))} /></label><span className="object-icon small">{iconKind(sheet.kind)}</span><span><b>{compact(row.name || row.display_name || row._external_id || `${sheet.name} row ${rowIndex + 1}`, 72)}</b><small>{sheet.model} · row {row.__rownum__ || rowIndex + 2}</small></span></summary><div className="field-groups">{groups.map(group => <div className="field-group" key={group.title}><h3>{group.title}</h3><div className="form-grid">{group.fields.map(col => <SmartInput key={col} col={col} value={row[col] ?? ''} field={schema?.[col]} onChange={v => updateCell(rowIndex, col, v)} />)}</div></div>)}</div></details>)}</div>;
}
function groupFields(kind: EditorKind, columns: string[]) {
  const priority = importantFields(kind, '');
  const basic = columns.filter(c => ['_model','__action','_external_id','x_studio2_odoo_id','id','name','display_name'].includes(c));
  const main = columns.filter(c => priority.includes(c) && !basic.includes(c));
  const technical = columns.filter(c => !basic.includes(c) && !main.includes(c) && c !== '__rownum__');
  return [{ title: 'Identity', fields: basic }, { title: labelKind(kind), fields: main }, { title: 'Technical & others', fields: technical.slice(0, 80) }].filter(g => g.fields.length);
}
function SmartInput({ col, value, field, onChange }: { col: string; value: any; field?: OdooField; onChange: (v: any) => void }) {
  const type = field?.type;
  if (type === 'boolean') return <label className="smart-field check-field"><span>{field?.string || col}<small>{col}</small></span><input type="checkbox" checked={Boolean(value === true || String(value).toLowerCase() === 'true')} onChange={e => onChange(e.target.checked)} /></label>;
  if (type === 'selection' && field?.selection) return <label className="smart-field"><span>{field.string || col}<small>{col}</small></span><select value={String(value ?? '')} onChange={e => onChange(e.target.value)}><option value="">—</option>{field.selection.map(([key, label]) => <option value={key} key={key}>{label}</option>)}</select></label>;
  if (type === 'text' || type === 'html' || String(value || '').length > 100) return <label className="smart-field wide"><span>{field?.string || col}<small>{col}{field?.relation ? ` · ${field.relation}` : ''}</small></span><textarea value={String(value ?? '')} onChange={e => onChange(e.target.value)} /></label>;
  return <label className="smart-field"><span>{field?.string || col}<small>{col}{field?.relation ? ` · ${field.relation}` : field?.type ? ` · ${field.type}` : ''}</small></span><input value={String(value ?? '')} onChange={e => onChange(e.target.value)} /></label>;
}
function GridEditor({ sheet, columns, updateCell, selectedRows, setSelectedRows }: { sheet: SheetState; columns: string[]; updateCell: (rowIndex: number, column: string, value: any) => void; selectedRows: Record<number, boolean>; setSelectedRows: React.Dispatch<React.SetStateAction<Record<number, boolean>>> }) {
  return <div className="grid-wrap"><table><thead><tr><th>✓</th>{columns.map(col => <th key={col}>{col}</th>)}</tr></thead><tbody>{sheet.rows.slice(0, 120).map((row, rowIndex) => <tr key={rowIndex}><td><input type="checkbox" checked={Boolean(selectedRows[rowIndex])} onChange={e => setSelectedRows(prev => ({ ...prev, [rowIndex]: e.target.checked }))} /></td>{columns.map(col => <td key={col}><input value={String(row[col] ?? '')} onChange={e => updateCell(rowIndex, col, e.target.value)} /></td>)}</tr>)}</tbody></table></div>;
}

function BottomNav({ mission, setMission }: { mission: Mission; setMission: (m: Mission) => void }) {
  const items: Array<[Mission, string, string]> = [['home','⌂','Home'], ['import','↑','Import'], ['export','↓','Export'], ['review','◇','Review'], ['settings','⚙','Koneksi']];
  return <nav className="bottom-nav">{items.map(([key, icon, label]) => <button key={key} className={mission === key ? 'active' : ''} onClick={() => setMission(key)}><span>{icon}</span>{label}</button>)}</nav>;
}
function LogSheet({ logs, open, onClose }: { logs: LogItem[]; open: boolean; onClose: () => void }) { return <div className={open ? 'log-backdrop open' : 'log-backdrop'} onClick={onClose}><aside className="log-sheet" onClick={e => e.stopPropagation()}><div className="log-title"><b>Activity trail</b><button onClick={onClose}>×</button></div>{logs.map((log, idx) => <div key={idx} className={`log-line ${log.level}`}><small>{log.time}</small><span>{log.message}</span>{log.detail && <pre>{JSON.stringify(log.detail, null, 2).slice(0, 1200)}</pre>}</div>)}{!logs.length && <p className="empty-note">Belum ada log.</p>}</aside></div>; }
