'use client';

import { useEffect, useMemo, useState } from 'react';
import type { EngineConfig, Job, OdooSchema, OdooTarget, SheetData, WorkbookState } from '../lib/types';
import { addColumn, fileToWorkbookState, localValidate, updateCell, workbookStateToBlob } from '../lib/xlsx';
import { downloadUrl, getSchema, pollJob, recordScan, startExportProject, startExportRecords, startImportXlsx, testOdoo } from '../lib/engine';

const defaultEngine: EngineConfig = {
  baseUrl: process.env.NEXT_PUBLIC_DEFAULT_ENGINE_URL || '',
  apiKey: ''
};
const defaultTarget: OdooTarget = { url: '', db: '', username: '', password: '' };

function loadJson<T>(key: string, fallback: T): T {
  if (typeof window === 'undefined') return fallback;
  try {
    const raw = localStorage.getItem(key);
    return raw ? JSON.parse(raw) : fallback;
  } catch {
    return fallback;
  }
}

function saveJson(key: string, value: any) {
  if (typeof window !== 'undefined') localStorage.setItem(key, JSON.stringify(value));
}

function editorTitle(kind: SheetData['kind'], model: string) {
  if (kind === 'product') return 'Product Editor';
  if (kind === 'contact') return 'Contact Editor';
  if (kind === 'project') return 'Project Editor';
  if (kind === 'knowledge') return 'Knowledge Editor';
  if (kind === 'sales') return 'Sales Editor';
  if (kind === 'context') return 'Context Sheet';
  return model ? 'Dynamic Odoo Editor' : 'XLSX Editor';
}

function importantFields(kind: SheetData['kind'], columns: string[]) {
  const base: Record<string, string[]> = {
    product: ['_model', '__action', '_external_id', 'name', 'default_code', 'barcode', 'list_price', 'standard_price', 'categ_id', 'categ_id_external_id', 'seller_ids', 'x_source_vendor', 'image_url', 'x_source_image_url', 'website_published', 'sale_ok', 'purchase_ok'],
    contact: ['_model', '__action', '_external_id', 'name', 'email', 'phone', 'mobile', 'street', 'city', 'supplier_rank', 'customer_rank', 'company_type', 'category_id_external_ids', 'x_lm_role', 'x_lm_area'],
    project: ['_model', '__action', '_external_id', 'name', 'project_id', 'project_id_external_id', 'parent_id', 'parent_id_external_id', 'stage_id', 'stage_id_external_id', 'user_ids', 'date_deadline', 'description'],
    knowledge: ['_model', '__action', '_external_id', 'name', 'body', 'parent_id', 'parent_id_external_id', 'is_article_visible_by_everyone'],
    sales: ['_model', '__action', '_external_id', 'name', 'partner_id', 'partner_id_external_id', 'date_order', 'order_line', 'amount_total', 'state'],
    dynamic: ['_model', '__action', '_external_id', 'ID', 'id', 'display_name', 'name', 'active', 'write_date'],
    context: columns.slice(0, 12)
  };
  const wanted = base[kind] || base.dynamic;
  const present = wanted.filter(c => columns.includes(c));
  const rest = columns.filter(c => !present.includes(c));
  return [...present, ...rest].slice(0, 28);
}

function defaultFieldSelection(model: string, schema: OdooSchema) {
  const preferred = [
    'id', 'display_name', 'name', 'email', 'phone', 'mobile', 'street', 'street2', 'city', 'zip',
    'default_code', 'barcode', 'list_price', 'standard_price', 'categ_id', 'public_categ_ids',
    'sale_ok', 'purchase_ok', 'website_published', 'project_id', 'parent_id', 'stage_id', 'user_ids',
    'partner_id', 'date_deadline', 'description', 'active', 'write_date', 'create_date'
  ];
  return preferred.filter(f => f === 'id' || schema[f]);
}

export default function Page() {
  const [mode, setMode] = useState<'import' | 'export'>('import');
  const [showSettings, setShowSettings] = useState(false);
  const [engine, setEngine] = useState<EngineConfig>(defaultEngine);
  const [target, setTarget] = useState<OdooTarget>(defaultTarget);
  const [workbook, setWorkbook] = useState<WorkbookState | null>(null);
  const [activeSheet, setActiveSheet] = useState(0);
  const [schema, setSchema] = useState<OdooSchema>({});
  const [schemaModel, setSchemaModel] = useState('');
  const [job, setJob] = useState<Job | null>(null);
  const [toast, setToast] = useState('');

  useEffect(() => {
    setEngine(loadJson('studio2.engine', defaultEngine));
    setTarget(loadJson('studio2.target', defaultTarget));
  }, []);

  useEffect(() => saveJson('studio2.engine', engine), [engine]);
  useEffect(() => saveJson('studio2.target', target), [target]);

  const sheet = workbook?.sheets[activeSheet];
  const validation = sheet ? localValidate(sheet) : { errors: [], warnings: [] };

  async function refreshSchema(model?: string) {
    const m = model || sheet?.model;
    if (!m) return;
    const data = await getSchema(engine, target, m);
    setSchema(data.schema || {});
    setSchemaModel(m);
    setToast(`Schema ${m} dimuat.`);
  }

  async function handleFile(file: File) {
    const state = await fileToWorkbookState(file, file.name);
    setWorkbook(state);
    setActiveSheet(0);
    setSchema({});
    setSchemaModel('');
    setJob(null);
    setToast(`${state.sheets.length} sheet terbaca.`);
  }

  async function handleImport() {
    if (!workbook) return;
    const blob = workbookStateToBlob(workbook);
    const data = await startImportXlsx(engine, target, blob, workbook.fileName.replace(/\.xlsx$/i, '') + '_edited.xlsx', { batchLimit: 80, stopOnError: false });
    setJob(data.job);
    const finalJob = await pollJob(engine, data.job.id, setJob);
    setToast(finalJob.status === 'done' ? 'Import selesai.' : 'Import gagal. Cek log job.');
  }

  function downloadEditedWorkbook() {
    if (!workbook) return;
    const blob = workbookStateToBlob(workbook);
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = workbook.fileName.replace(/\.xlsx$/i, '') + '_edited.xlsx';
    a.click();
    URL.revokeObjectURL(a.href);
  }

  async function loadExportedJob(jobId: string) {
    const res = await fetch(downloadUrl(engine, jobId));
    if (!res.ok) throw new Error(`Download gagal HTTP ${res.status}`);
    const blob = await res.blob();
    const state = await fileToWorkbookState(blob, `studio2_export_${jobId}.xlsx`);
    setWorkbook(state);
    setActiveSheet(0);
    setMode('export');
  }

  return (
    <main className="min-h-screen px-4 py-4 md:px-8">
      <div className="mx-auto flex max-w-7xl flex-col gap-4">
        <Header mode={mode} setMode={setMode} onSettings={() => setShowSettings(true)} target={target} />
        {toast ? <div className="rounded-2xl border border-amber-400/25 bg-amber-500/10 px-4 py-3 text-sm text-amber-100">{toast}</div> : null}

        {mode === 'import' ? (
          <ImportPanel
            workbook={workbook}
            activeSheet={activeSheet}
            setActiveSheet={setActiveSheet}
            onFile={handleFile}
            onImport={handleImport}
            onDownload={downloadEditedWorkbook}
            job={job}
          />
        ) : (
          <ExportPanel
            engine={engine}
            target={target}
            onJob={setJob}
            onToast={setToast}
            onWorkbook={async (jobId) => loadExportedJob(jobId)}
            currentJob={job}
          />
        )}

        {workbook && sheet ? (
          <WorkbookEditor
            workbook={workbook}
            setWorkbook={setWorkbook}
            activeSheet={activeSheet}
            setActiveSheet={setActiveSheet}
            schema={schemaModel === sheet.model ? schema : {}}
            onRefreshSchema={() => refreshSchema(sheet.model)}
            validation={validation}
          />
        ) : null}

        {job ? <JobPanel job={job} /> : null}
      </div>

      {showSettings ? (
        <SettingsModal
          engine={engine}
          target={target}
          setEngine={setEngine}
          setTarget={setTarget}
          onClose={() => setShowSettings(false)}
          onTest={async () => {
            const data = await testOdoo(engine, target);
            setToast(`Login Odoo berhasil. UID: ${data.uid}`);
          }}
        />
      ) : null}
    </main>
  );
}

function Header({ mode, setMode, onSettings, target }: any) {
  return (
    <header className="card flex flex-col gap-3 px-4 py-3 md:flex-row md:items-center md:justify-between">
      <div className="flex items-center gap-3">
        <div className="flex h-11 w-11 items-center justify-center rounded-2xl bg-gradient-to-br from-amber-400 to-orange-500 font-black text-black">LM</div>
        <div>
          <h1 className="text-lg font-black leading-tight">Lokalmart Studio2 v8</h1>
          <p className="text-xs text-orange-100/70">Vercel UI · Render Engine · Native-like Odoo XLSX editor</p>
        </div>
      </div>
      <div className="flex items-center gap-2">
        <button className={`btn ${mode === 'import' ? 'btn-primary' : ''}`} onClick={() => setMode('import')}>⬆ Import</button>
        <button className={`btn ${mode === 'export' ? 'btn-primary' : ''}`} onClick={() => setMode('export')}>⬇ Export</button>
        <span className="badge hidden md:inline-flex">Target {target?.db || 'belum diset'}</span>
        <button className="btn" onClick={onSettings}>⚙</button>
      </div>
    </header>
  );
}

function ImportPanel({ workbook, activeSheet, setActiveSheet, onFile, onImport, onDownload, job }: any) {
  return (
    <section className="grid gap-4 lg:grid-cols-[360px_1fr]">
      <div className="card p-4">
        <h2 className="text-2xl font-black">Import XLSX</h2>
        <p className="mt-1 text-sm text-orange-100/70">Upload, preview, edit, validasi, lalu jalankan import batch di Render.</p>
        <label className="mt-4 flex cursor-pointer flex-col items-center justify-center rounded-3xl border border-dashed border-amber-400/40 bg-black/25 p-8 text-center hover:bg-amber-400/10">
          <span className="text-4xl">📄</span>
          <span className="mt-2 font-bold">Pilih XLSX</span>
          <span className="text-xs text-orange-100/60">File tidak langsung dikirim sampai kamu klik import.</span>
          <input type="file" accept=".xlsx,.xls" className="hidden" onChange={(e) => e.target.files?.[0] && onFile(e.target.files[0])} />
        </label>
        <div className="mt-4 grid grid-cols-2 gap-2">
          <button className="btn" disabled={!workbook} onClick={onDownload}>Download Edited</button>
          <button className="btn btn-primary" disabled={!workbook || job?.status === 'running'} onClick={onImport}>Import ke Odoo</button>
        </div>
      </div>
      <SheetNavigator workbook={workbook} activeSheet={activeSheet} setActiveSheet={setActiveSheet} />
    </section>
  );
}

function ExportPanel({ engine, target, onJob, onToast, onWorkbook, currentJob }: any) {
  const [tab, setTab] = useState<'model' | 'project'>('model');
  const [model, setModel] = useState('res.partner');
  const [limit, setLimit] = useState(120);
  const [records, setRecords] = useState<any[]>([]);
  const [selected, setSelected] = useState<number[]>([]);
  const [schema, setSchema] = useState<OdooSchema>({});
  const [fields, setFields] = useState<string[]>([]);
  const [filter, setFilter] = useState('');
  const [projectId, setProjectId] = useState('');

  const visible = useMemo(() => {
    const q = filter.toLowerCase();
    if (!q) return records;
    return records.filter(r => JSON.stringify(r).toLowerCase().includes(q));
  }, [records, filter]);

  async function scan() {
    const data = await recordScan(engine, target, model, limit);
    setRecords(data.records || []);
    setSchema(data.schema || {});
    const defaults = defaultFieldSelection(model, data.schema || {});
    setFields(defaults.length ? defaults : data.fields || []);
    setSelected((data.records || []).map((r: any) => Number(r.id)).filter(Boolean));
    onToast(`${data.records?.length || 0} record ditemukan dari ${model}.`);
  }

  async function exportSelected() {
    const data = await startExportRecords(engine, target, model, selected, fields);
    onJob(data.job);
    const finalJob = await pollJob(engine, data.job.id, onJob);
    if (finalJob.status === 'done') {
      await onWorkbook(finalJob.id);
      onToast('Export selesai dan masuk ke editor XLSX.');
    } else onToast('Export gagal. Cek log job.');
  }

  async function exportProject() {
    const id = Number(projectId);
    if (!id) return onToast('Isi ID project dulu. Bisa scan project.project untuk melihat ID.');
    const data = await startExportProject(engine, target, id);
    onJob(data.job);
    const finalJob = await pollJob(engine, data.job.id, onJob);
    if (finalJob.status === 'done') {
      await onWorkbook(finalJob.id);
      onToast('Export project selesai dan masuk ke editor XLSX.');
    } else onToast('Export project gagal. Cek log job.');
  }

  return (
    <section className="grid gap-4 lg:grid-cols-[360px_1fr]">
      <div className="card p-4">
        <h2 className="text-2xl font-black">Export XLSX</h2>
        <p className="mt-1 text-sm text-orange-100/70">Scan record Odoo dulu, pilih record dan field, baru export.</p>
        <div className="mt-4 flex rounded-2xl border border-white/10 bg-black/30 p-1">
          <button className={`flex-1 rounded-xl px-3 py-2 text-sm font-bold ${tab === 'model' ? 'bg-amber-500 text-black' : 'text-orange-100'}`} onClick={() => setTab('model')}>Model</button>
          <button className={`flex-1 rounded-xl px-3 py-2 text-sm font-bold ${tab === 'project' ? 'bg-amber-500 text-black' : 'text-orange-100'}`} onClick={() => setTab('project')}>Project</button>
        </div>
        {tab === 'model' ? (
          <div className="mt-4 space-y-3">
            <input className="input" value={model} onChange={e => setModel(e.target.value)} placeholder="res.partner" />
            <input className="input" type="number" value={limit} onChange={e => setLimit(Number(e.target.value))} placeholder="Limit scan" />
            <button className="btn w-full" onClick={scan}>Scan Record</button>
            <button className="btn btn-primary w-full" disabled={!selected.length || currentJob?.status === 'running'} onClick={exportSelected}>Export Record Terpilih</button>
          </div>
        ) : (
          <div className="mt-4 space-y-3">
            <input className="input" value={projectId} onChange={e => setProjectId(e.target.value)} placeholder="ID project.project" />
            <button className="btn btn-primary w-full" disabled={!projectId || currentJob?.status === 'running'} onClick={exportProject}>Export Project</button>
            <p className="text-xs text-orange-100/60">Tips: untuk mencari ID project, pilih tab Model, scan `project.project`, lalu lihat kolom ID.</p>
          </div>
        )}
      </div>

      <div className="card overflow-hidden">
        <div className="border-b border-white/10 p-4">
          <div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
            <div>
              <h3 className="text-xl font-black">Record Picker</h3>
              <p className="text-sm text-orange-100/60">{records.length} record · {selected.length} dipilih · {fields.length} field export</p>
            </div>
            <input className="input md:max-w-xs" placeholder="Cari record..." value={filter} onChange={e => setFilter(e.target.value)} />
          </div>
        </div>
        <div className="grid max-h-[520px] gap-0 overflow-hidden lg:grid-cols-[1fr_320px]">
          <div className="overflow-auto p-4">
            <div className="mb-3 flex gap-2">
              <button className="btn" onClick={() => setSelected(visible.map(r => Number(r.id)).filter(Boolean))}>Pilih terlihat</button>
              <button className="btn" onClick={() => setSelected([])}>Kosongkan</button>
            </div>
            <div className="space-y-2">
              {visible.map(record => (
                <label key={record.id} className="flex cursor-pointer items-start gap-3 rounded-2xl border border-white/10 bg-black/25 p-3 hover:bg-white/10">
                  <input
                    type="checkbox"
                    className="mt-1 h-4 w-4 accent-amber-500"
                    checked={selected.includes(Number(record.id))}
                    onChange={e => setSelected(prev => e.target.checked ? [...new Set([...prev, Number(record.id)])] : prev.filter(id => id !== Number(record.id)))}
                  />
                  <div className="min-w-0">
                    <div className="font-bold">#{record.id} · {record.display_name || record.name || '(tanpa nama)'}</div>
                    <div className="mt-1 text-xs text-orange-100/55">{Object.entries(record).filter(([k]) => !['id', 'display_name', 'name'].includes(k)).slice(0, 6).map(([k, v]) => `${k}: ${Array.isArray(v) ? v.join(' / ') : v}`).join(' · ')}</div>
                  </div>
                </label>
              ))}
              {!records.length ? <div className="rounded-2xl border border-dashed border-white/10 p-8 text-center text-orange-100/60">Belum ada scan record.</div> : null}
            </div>
          </div>
          <div className="border-t border-white/10 p-4 lg:border-l lg:border-t-0">
            <h4 className="font-black">Field Picker</h4>
            <p className="mb-3 text-xs text-orange-100/60">Pilih field supaya XLSX tidak terlalu besar dan tidak membawa field panjang yang tidak perlu.</p>
            <div className="max-h-[420px] space-y-2 overflow-auto pr-1">
              {Object.entries(schema).map(([name, meta]) => (
                <label key={name} className="flex items-start gap-2 rounded-xl border border-white/10 bg-black/25 px-3 py-2 text-xs">
                  <input type="checkbox" className="mt-0.5 accent-amber-500" checked={fields.includes(name)} onChange={e => setFields(prev => e.target.checked ? [...prev, name] : prev.filter(f => f !== name))} />
                  <span><b>{name}</b><br/><span className="text-orange-100/50">{meta.string} · {meta.type}</span></span>
                </label>
              ))}
              {!Object.keys(schema).length ? <div className="text-sm text-orange-100/60">Schema muncul setelah scan.</div> : null}
            </div>
          </div>
        </div>
      </div>
    </section>
  );
}

function SheetNavigator({ workbook, activeSheet, setActiveSheet }: any) {
  return (
    <div className="card overflow-hidden">
      <div className="border-b border-white/10 p-4">
        <h3 className="text-xl font-black">Sheet Preview</h3>
        <p className="text-sm text-orange-100/60">Panel ini menggantikan card mentah. Pilih sheet yang mau difokuskan.</p>
      </div>
      <div className="grid max-h-[260px] gap-2 overflow-auto p-4 md:grid-cols-2 lg:grid-cols-3">
        {workbook?.sheets?.map((s: SheetData, i: number) => (
          <button key={s.name} className={`rounded-2xl border p-4 text-left transition ${activeSheet === i ? 'border-amber-400 bg-amber-400/10' : 'border-white/10 bg-black/25 hover:bg-white/10'}`} onClick={() => setActiveSheet(i)}>
            <div className="font-black">{s.name}</div>
            <div className="mt-1 text-xs text-orange-100/60">{s.model || 'context'} · {s.rows.length} rows · {s.columns.length} kolom</div>
            <div className="mt-2"><span className="badge">{s.kind}</span></div>
          </button>
        )) || <div className="rounded-2xl border border-dashed border-white/10 p-8 text-center text-orange-100/60">Belum ada workbook.</div>}
      </div>
    </div>
  );
}

function WorkbookEditor({ workbook, setWorkbook, activeSheet, setActiveSheet, schema, onRefreshSchema, validation }: any) {
  const sheet: SheetData = workbook.sheets[activeSheet];
  const [newColumn, setNewColumn] = useState('');
  const columns = importantFields(sheet.kind, sheet.columns);
  const rows = sheet.rows.slice(0, 150);

  return (
    <section className="card overflow-hidden">
      <div className="flex flex-col gap-3 border-b border-white/10 p-4 md:flex-row md:items-center md:justify-between">
        <div>
          <h2 className="text-2xl font-black">{sheet.name}</h2>
          <p className="text-sm text-orange-100/60">{sheet.model || 'context/helper'} · {sheet.rows.length} row · {sheet.columns.length} kolom</p>
          <div className="mt-2 flex flex-wrap gap-2">
            <span className="badge">{sheet.kind}</span>
            <span className="badge">{validation.errors.length} error</span>
            <span className="badge">{validation.warnings.length} warning</span>
          </div>
        </div>
        <div className="flex flex-wrap gap-2">
          <button className="btn" onClick={onRefreshSchema} disabled={!sheet.model}>Refresh schema</button>
          <select className="input w-auto" value={activeSheet} onChange={e => setActiveSheet(Number(e.target.value))}>
            {workbook.sheets.map((s: SheetData, i: number) => <option value={i} key={s.name}>{s.name}</option>)}
          </select>
        </div>
      </div>

      <div className="grid gap-4 p-4 lg:grid-cols-[300px_1fr]">
        <aside className="rounded-3xl border border-white/10 bg-black/25 p-4">
          <h3 className="text-lg font-black">{editorTitle(sheet.kind, sheet.model)}</h3>
          <p className="mt-1 text-sm text-orange-100/60">{editorHelp(sheet.kind)}</p>
          <div className="mt-4 space-y-2 text-xs">
            {validation.errors.slice(0, 5).map((e: string) => <div key={e} className="rounded-xl border border-red-400/40 bg-red-500/10 p-2 text-red-100">{e}</div>)}
            {validation.warnings.slice(0, 5).map((w: string) => <div key={w} className="rounded-xl border border-yellow-400/40 bg-yellow-500/10 p-2 text-yellow-100">{w}</div>)}
          </div>
          {sheet.kind === 'product' ? <ProductPhotoPreview sheet={sheet} /> : null}
          <div className="mt-4 flex gap-2">
            <input className="input" placeholder="Tambah kolom" value={newColumn} onChange={e => setNewColumn(e.target.value)} />
            <button className="btn" onClick={() => { setWorkbook(addColumn(workbook, activeSheet, newColumn)); setNewColumn(''); }}>+</button>
          </div>
        </aside>

        <div className="overflow-auto rounded-3xl border border-white/10 bg-black/20">
          <table className="min-w-full border-separate border-spacing-0">
            <thead className="sticky top-0 z-10 bg-[#211a14]">
              <tr>
                <th className="border-b border-white/10 px-3 py-3 text-left text-xs text-orange-100/70">#</th>
                {columns.map(col => (
                  <th key={col} className="border-b border-white/10 px-2 py-3 text-left text-xs text-orange-100/70">
                    <div className="font-bold">{col}</div>
                    <div className="font-normal text-orange-100/40">{schema[col]?.type || 'xlsx'}</div>
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {rows.map((row, rIdx) => (
                <tr key={rIdx} className="hover:bg-white/5">
                  <td className="border-b border-white/5 px-3 py-2 text-xs text-orange-100/40">{rIdx + 2}</td>
                  {columns.map(col => (
                    <td key={col} className="border-b border-white/5 px-2 py-2 align-top">
                      <FieldInput
                        value={row[col] ?? ''}
                        meta={schema[col]}
                        onChange={(value: any) => setWorkbook(updateCell(workbook, activeSheet, rIdx, col, value))}
                      />
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
          {sheet.rows.length > rows.length ? <div className="p-3 text-center text-xs text-orange-100/50">Menampilkan 150 row pertama dari {sheet.rows.length} row demi performa.</div> : null}
        </div>
      </div>
    </section>
  );
}

function FieldInput({ value, meta, onChange }: any) {
  const type = meta?.type;
  if (type === 'boolean') return <input type="checkbox" className="h-5 w-5 accent-amber-500" checked={String(value).toLowerCase() === 'true' || value === true || value === 1 || value === '1'} onChange={e => onChange(e.target.checked)} />;
  if (type === 'selection' && meta?.selection?.length) return <select className="table-input" value={value} onChange={e => onChange(e.target.value)}>{meta.selection.map((x: any) => <option value={x[0]} key={x[0]}>{x[1]}</option>)}</select>;
  if (['text', 'html'].includes(type)) return <textarea className="table-input min-h-[80px]" value={value} onChange={e => onChange(e.target.value)} />;
  if (['integer', 'float', 'monetary'].includes(type)) return <input className="table-input" type="number" value={value} onChange={e => onChange(e.target.value)} />;
  if (['date', 'datetime'].includes(type)) return <input className="table-input" type={type === 'date' ? 'date' : 'datetime-local'} value={String(value).slice(0, type === 'date' ? 10 : 16)} onChange={e => onChange(e.target.value)} />;
  return <input className="table-input" value={value} onChange={e => onChange(e.target.value)} />;
}

function editorHelp(kind: SheetData['kind']) {
  if (kind === 'product') return 'Validasi nama, harga, barcode, kategori, vendor, status jual/publish, dan sumber foto.';
  if (kind === 'contact') return 'Validasi kontak, WA, email, alamat, role, supplier/customer rank.';
  if (kind === 'project') return 'Validasi project, task, parent, stage, deadline, external ID, dan deskripsi.';
  if (kind === 'knowledge') return 'Validasi artikel knowledge, judul, parent, body, dan visibilitas.';
  if (kind === 'context') return 'Sheet konteks dibaca untuk memory/AI, bukan data import utama.';
  return 'Editor mengikuti schema field Odoo jika tersedia.';
}

function ProductPhotoPreview({ sheet }: { sheet: SheetData }) {
  const first = sheet.rows.find(r => r.image_url || r.x_source_image_url || r.image_1920) || {};
  const url = first.image_url || first.x_source_image_url || '';
  return (
    <div className="mt-4 rounded-2xl border border-dashed border-white/10 bg-black/30 p-3 text-center">
      {url ? <img src={url} alt="preview" className="mx-auto max-h-40 rounded-xl object-contain" /> : <div className="py-12 text-sm text-orange-100/50">Belum ada foto</div>}
    </div>
  );
}

function JobPanel({ job }: { job: Job }) {
  return (
    <section className="card p-4">
      <div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
        <div>
          <h3 className="text-xl font-black">Job {job.type}</h3>
          <p className="text-sm text-orange-100/60">{job.status} · {job.progress}% · {job.id}</p>
        </div>
        <div className="h-3 w-full overflow-hidden rounded-full bg-white/10 md:w-80">
          <div className="h-full bg-gradient-to-r from-amber-500 to-orange-500" style={{ width: `${job.progress || 0}%` }} />
        </div>
      </div>
      <div className="mt-4 max-h-48 overflow-auto rounded-2xl bg-black/35 p-3 font-mono text-xs text-orange-100/70">
        {job.logs?.slice(-80).map((l, i) => <div key={i}>[{new Date(l.time).toLocaleTimeString()}] {l.level.toUpperCase()}: {l.message}</div>)}
        {job.errors?.map((e, i) => <div key={`e${i}`} className="text-red-200">ERROR: {e}</div>)}
      </div>
    </section>
  );
}

function SettingsModal({ engine, target, setEngine, setTarget, onClose, onTest }: any) {
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/70 p-4 backdrop-blur-sm">
      <div className="card max-h-[92vh] w-full max-w-3xl overflow-auto p-5">
        <div className="flex items-center justify-between gap-4">
          <div>
            <h2 className="text-2xl font-black">Pengaturan Target</h2>
            <p className="text-sm text-orange-100/60">Disimpan di browser, bukan di GitHub.</p>
          </div>
          <button className="btn" onClick={onClose}>Tutup</button>
        </div>
        <div className="mt-5 grid gap-4 md:grid-cols-2">
          <div className="rounded-3xl border border-white/10 bg-black/25 p-4">
            <h3 className="font-black">Render Engine</h3>
            <label className="mt-3 block text-xs text-orange-100/60">Engine URL</label>
            <input className="input" value={engine.baseUrl} onChange={e => setEngine({ ...engine, baseUrl: e.target.value })} placeholder="https://studio2-engine.onrender.com" />
            <label className="mt-3 block text-xs text-orange-100/60">Engine API Key</label>
            <input className="input" value={engine.apiKey} onChange={e => setEngine({ ...engine, apiKey: e.target.value })} placeholder="STUDIO2_ENGINE_API_KEY" />
          </div>
          <div className="rounded-3xl border border-white/10 bg-black/25 p-4">
            <h3 className="font-black">Odoo Target</h3>
            <label className="mt-3 block text-xs text-orange-100/60">Odoo URL</label>
            <input className="input" value={target.url} onChange={e => setTarget({ ...target, url: e.target.value })} placeholder="https://edu-lokalmart.odoo.com" />
            <label className="mt-3 block text-xs text-orange-100/60">Database</label>
            <input className="input" value={target.db} onChange={e => setTarget({ ...target, db: e.target.value })} />
            <label className="mt-3 block text-xs text-orange-100/60">Username / Email</label>
            <input className="input" value={target.username} onChange={e => setTarget({ ...target, username: e.target.value })} />
            <label className="mt-3 block text-xs text-orange-100/60">Password / API Key</label>
            <input className="input" type="password" value={target.password} onChange={e => setTarget({ ...target, password: e.target.value })} />
          </div>
        </div>
        <div className="mt-5 flex justify-end gap-2">
          <button className="btn" onClick={onTest}>Test Login Odoo</button>
          <button className="btn btn-primary" onClick={onClose}>Simpan</button>
        </div>
      </div>
    </div>
  );
}
