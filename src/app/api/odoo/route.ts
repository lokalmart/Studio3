import { NextRequest, NextResponse } from 'next/server';

export const runtime = 'nodejs';
export const maxDuration = 300;

type Conn = {
  url: string;
  db: string;
  username: string;
  password: string;
};

type AnyRow = Record<string, any>;

type OdooField = {
  string?: string;
  type?: string;
  required?: boolean;
  readonly?: boolean;
  relation?: string;
  selection?: Array<[string, string]>;
};

const EXCEL_CELL_LIMIT = 32767;
const SAFE_CELL_LIMIT = 32000;
const HELPER_COLUMNS = new Set([
  '_model', '__action', '_external_id', 'external_id', 'id', 'x_studio2_odoo_id',
  '_studio2_truncated_fields', '_studio2_note', '_studio2_error'
]);

function cleanUrl(url: string) {
  return String(url || '').replace(/\/+$/, '');
}

function json(data: any, status = 200) {
  return NextResponse.json(data, { status });
}

function assertConn(conn: Conn) {
  if (!conn?.url || !conn?.db || !conn?.username || !conn?.password) {
    throw new Error('Koneksi Odoo belum lengkap: url, db, username, password/api key wajib diisi.');
  }
}

async function rpc(conn: Conn, service: string, method: string, args: any[]) {
  assertConn(conn);
  const res = await fetch(`${cleanUrl(conn.url)}/jsonrpc`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      jsonrpc: '2.0',
      method: 'call',
      params: { service, method, args },
      id: Date.now()
    }),
    cache: 'no-store'
  });
  const text = await res.text();
  let payload: any;
  try {
    payload = JSON.parse(text);
  } catch {
    throw new Error(`Odoo non-JSON HTTP ${res.status}: ${text.slice(0, 500)}`);
  }
  if (!res.ok) {
    throw new Error(`Odoo HTTP ${res.status}: ${text.slice(0, 500)}`);
  }
  if (payload.error) {
    const err = payload.error;
    const detail = err?.data?.message || err?.message || JSON.stringify(err);
    throw new Error(`Odoo JSON-RPC fault: ${detail}`);
  }
  return payload.result;
}

async function login(conn: Conn) {
  const uid = await rpc(conn, 'common', 'login', [conn.db, conn.username, conn.password]);
  if (!uid) throw new Error('Login Odoo gagal. Periksa database, email, password/API key.');
  return uid as number;
}

async function executeKw(conn: Conn, uid: number, model: string, method: string, args: any[] = [], kwargs: AnyRow = {}) {
  return rpc(conn, 'object', 'execute_kw', [conn.db, uid, conn.password, model, method, args, kwargs]);
}

async function fieldsGet(conn: Conn, uid: number, model: string) {
  return executeKw(conn, uid, model, 'fields_get', [], { attributes: ['string', 'type', 'required', 'readonly', 'relation', 'selection'] }) as Promise<Record<string, OdooField>>;
}

function splitXmlId(input: string) {
  const raw = String(input || '').trim();
  if (!raw) return null;
  if (raw.includes('.')) {
    const [module, ...rest] = raw.split('.');
    return { module: safeExternalPart(module), name: safeExternalPart(rest.join('.')) };
  }
  return { module: 'studio2', name: safeExternalPart(raw) };
}

function safeExternalPart(value: string) {
  return String(value || '')
    .trim()
    .replace(/[^A-Za-z0-9_\.\-]/g, '_')
    .replace(/^_+/, '')
    .slice(0, 200) || 'unnamed';
}

function completeXmlId(value: string) {
  const p = splitXmlId(value);
  if (!p) return '';
  return `${p.module}.${p.name}`;
}

async function resolveXmlId(conn: Conn, uid: number, xmlid: string, cache: Map<string, { model: string; id: number } | null>) {
  const key = completeXmlId(xmlid);
  if (!key) return null;
  if (cache.has(key)) return cache.get(key) || null;
  try {
    const rows = await executeKw(conn, uid, 'ir.model.data', 'search_read', [[['complete_name', '=', key]]], {
      fields: ['model', 'res_id'], limit: 1
    });
    if (rows?.length) {
      const found = { model: rows[0].model as string, id: Number(rows[0].res_id) };
      cache.set(key, found);
      return found;
    }
  } catch {
    // Some Odoo permissions are picky; fallback below returns null.
  }
  cache.set(key, null);
  return null;
}

async function mapExternalIdsForModel(conn: Conn, uid: number, model: string, ids: number[]) {
  const map = new Map<number, string>();
  if (!ids.length) return map;
  try {
    const rows = await executeKw(conn, uid, 'ir.model.data', 'search_read', [[['model', '=', model], ['res_id', 'in', ids]]], {
      fields: ['complete_name', 'res_id'],
      limit: Math.max(1000, ids.length * 3)
    });
    for (const row of rows || []) {
      const id = Number(row.res_id);
      if (id && !map.has(id)) map.set(id, row.complete_name);
    }
  } catch {
    // External IDs are helpful, not mandatory.
  }
  return map;
}

async function createExternalId(conn: Conn, uid: number, model: string, resId: number, xmlid: string) {
  const parsed = splitXmlId(xmlid);
  if (!parsed || !resId) return null;
  const complete = `${parsed.module}.${parsed.name}`;
  const existing = await resolveXmlId(conn, uid, complete, new Map());
  if (existing) return existing;
  try {
    await executeKw(conn, uid, 'ir.model.data', 'create', [{
      module: parsed.module,
      name: parsed.name,
      model,
      res_id: resId,
      noupdate: true
    }]);
    return { model, id: resId };
  } catch (err: any) {
    return null;
  }
}

function isBlank(value: any) {
  return value === null || value === undefined || String(value).trim() === '';
}

function parseBoolean(value: any) {
  if (typeof value === 'boolean') return value;
  const s = String(value).trim().toLowerCase();
  if (['true', '1', 'yes', 'y', 'ya', 'iya'].includes(s)) return true;
  if (['false', '0', 'no', 'n', 'tidak', ''].includes(s)) return false;
  return Boolean(value);
}

function parseNumber(value: any) {
  if (typeof value === 'number') return value;
  const cleaned = String(value).replace(/rp/ig, '').replace(/\./g, '').replace(',', '.').replace(/[^0-9.\-]/g, '');
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : value;
}

function convertByFieldType(value: any, field?: OdooField) {
  if (!field) return value;
  if (isBlank(value)) return false;
  switch (field.type) {
    case 'boolean': return parseBoolean(value);
    case 'integer': return Math.round(Number(parseNumber(value)) || 0);
    case 'float':
    case 'monetary': return Number(parseNumber(value)) || 0;
    case 'many2one': return typeof value === 'number' ? value : (String(value).match(/^\d+$/) ? Number(value) : value);
    case 'char':
    case 'text':
    case 'html': return String(value);
    default: return value;
  }
}

function sanitizeCell(value: any) {
  if (value === undefined || value === null) return value;
  if (Array.isArray(value)) return JSON.stringify(value);
  if (typeof value === 'object') return JSON.stringify(value);
  if (typeof value === 'string' && value.length > SAFE_CELL_LIMIT) return value.slice(0, SAFE_CELL_LIMIT);
  return value;
}

function sanitizeExportRow(row: AnyRow) {
  const out: AnyRow = {};
  const truncated: string[] = [];
  for (const [k, v] of Object.entries(row)) {
    if (typeof v === 'string' && v.length > SAFE_CELL_LIMIT) {
      out[k] = v.slice(0, SAFE_CELL_LIMIT);
      truncated.push(k);
    } else if (typeof v === 'string' && v.length > EXCEL_CELL_LIMIT) {
      out[k] = v.slice(0, SAFE_CELL_LIMIT);
      truncated.push(k);
    } else if (Array.isArray(v)) {
      out[k] = JSON.stringify(v).slice(0, SAFE_CELL_LIMIT);
    } else if (v && typeof v === 'object') {
      out[k] = JSON.stringify(v).slice(0, SAFE_CELL_LIMIT);
    } else {
      out[k] = v;
    }
  }
  if (truncated.length) out._studio2_truncated_fields = truncated.join(', ');
  return out;
}

async function buildVals(conn: Conn, uid: number, model: string, row: AnyRow, schema: Record<string, OdooField>, xmlCache: Map<string, { model: string; id: number } | null>) {
  const vals: AnyRow = {};
  const skipped: string[] = [];
  const warnings: string[] = [];

  for (const [key, rawValue] of Object.entries(row)) {
    if (HELPER_COLUMNS.has(key)) continue;
    if (isBlank(rawValue)) continue;

    if (key.endsWith('_external_id')) {
      const fieldName = key.replace(/_external_id$/, '');
      if (!schema[fieldName]) { skipped.push(key); continue; }
      const resolved = await resolveXmlId(conn, uid, String(rawValue), xmlCache);
      if (resolved?.id) vals[fieldName] = resolved.id;
      else warnings.push(`${key}: external ID tidak ditemukan (${rawValue})`);
      continue;
    }

    if (key.endsWith('_external_ids')) {
      const fieldName = key.replace(/_external_ids$/, '');
      if (!schema[fieldName]) { skipped.push(key); continue; }
      const ids: number[] = [];
      for (const piece of String(rawValue).split(/[;,]/).map(x => x.trim()).filter(Boolean)) {
        const resolved = await resolveXmlId(conn, uid, piece, xmlCache);
        if (resolved?.id) ids.push(resolved.id);
        else warnings.push(`${key}: external ID tidak ditemukan (${piece})`);
      }
      if (schema[fieldName].type === 'many2many') vals[fieldName] = [[6, 0, ids]];
      else vals[fieldName] = ids;
      continue;
    }

    if (!schema[key]) { skipped.push(key); continue; }
    if (schema[key].readonly) { skipped.push(key); continue; }
    vals[key] = convertByFieldType(rawValue, schema[key]);
  }

  return { vals, skipped, warnings };
}

async function handleTest(conn: Conn) {
  const uid = await login(conn);
  const count = await executeKw(conn, uid, 'res.partner', 'search_count', [[]]);
  return { ok: true, uid, partner_count: count };
}

async function handleSchema(conn: Conn, body: AnyRow) {
  const uid = await login(conn);
  const model = body.model;
  if (!model) throw new Error('Model kosong.');
  const fields = await fieldsGet(conn, uid, model);
  return { ok: true, model, fields };
}

async function handleRecordScan(conn: Conn, body: AnyRow) {
  const uid = await login(conn);
  const model = body.model;
  if (!model) throw new Error('Model kosong.');
  const limit = Math.min(Number(body.limit || 80), 200);
  const offset = Number(body.offset || 0);
  const domain = Array.isArray(body.domain) ? body.domain : [];
  const requested = Array.isArray(body.fields) && body.fields.length ? body.fields : ['display_name', 'name', 'email', 'phone', 'default_code', 'list_price', 'create_date', 'write_date'];
  const schema = await fieldsGet(conn, uid, model);
  const fields = ['id', ...requested.filter((f: string) => schema[f] && f !== 'id')];
  const records = await executeKw(conn, uid, model, 'search_read', [domain], { fields, limit, offset, order: 'id desc' });
  const count = await executeKw(conn, uid, model, 'search_count', [domain]);
  return { ok: true, model, records: records.map(sanitizeExportRow), count, offset, limit, has_more: offset + records.length < count };
}

async function handleExportRecords(conn: Conn, body: AnyRow) {
  const uid = await login(conn);
  const model = body.model;
  const ids = (body.ids || []).map((x: any) => Number(x)).filter(Boolean);
  if (!model) throw new Error('Model kosong.');
  if (!ids.length) throw new Error('Belum ada record dipilih.');
  const schema = await fieldsGet(conn, uid, model);
  let fields = Array.isArray(body.fields) ? body.fields.filter((f: string) => schema[f]) : [];
  if (!fields.length) {
    fields = Object.keys(schema).filter(k => ['char', 'text', 'html', 'boolean', 'integer', 'float', 'monetary', 'selection', 'date', 'datetime', 'many2one'].includes(schema[k].type || '')).slice(0, 40);
  }
  if (!fields.includes('id')) fields = ['id', ...fields];
  const records = await executeKw(conn, uid, model, 'search_read', [[['id', 'in', ids]]], { fields, limit: ids.length });
  const extMap = await mapExternalIdsForModel(conn, uid, model, ids);
  const rows = (records || []).map((r: AnyRow) => sanitizeExportRow({
    _model: model,
    __action: 'upsert',
    _external_id: extMap.get(Number(r.id)) || '',
    x_studio2_odoo_id: r.id,
    ...r
  }));
  return { ok: true, model, sheet: model, rows, fields };
}

async function handleExportProject(conn: Conn, body: AnyRow) {
  const uid = await login(conn);
  const projectId = Number(body.project_id);
  if (!projectId) throw new Error('project_id kosong.');
  const out: Record<string, AnyRow[]> = {};

  const pushSheet = async (model: string, domain: any[], fields: string[]) => {
    try {
      const schema = await fieldsGet(conn, uid, model);
      const realFields = ['id', ...fields.filter(f => schema[f] && f !== 'id')];
      const records = await executeKw(conn, uid, model, 'search_read', [domain], { fields: realFields, limit: 5000, order: 'id asc' });
      const ids = (records || []).map((r: AnyRow) => Number(r.id)).filter(Boolean);
      const extMap = await mapExternalIdsForModel(conn, uid, model, ids);
      out[model] = (records || []).map((r: AnyRow) => sanitizeExportRow({
        _model: model,
        __action: 'upsert',
        _external_id: extMap.get(Number(r.id)) || '',
        x_studio2_odoo_id: r.id,
        ...r
      }));
    } catch (err: any) {
      out[`${model}_error`] = [{ error: err?.message || String(err) }];
    }
  };

  await pushSheet('project.project', [['id', '=', projectId]], ['name', 'partner_id', 'user_id', 'date_start', 'date', 'description']);
  await pushSheet('project.task', [['project_id', '=', projectId]], ['name', 'project_id', 'parent_id', 'stage_id', 'user_ids', 'partner_id', 'date_deadline', 'description', 'priority', 'sequence']);
  await pushSheet('project.milestone', [['project_id', '=', projectId]], ['name', 'project_id', 'deadline', 'is_reached']);
  await pushSheet('project.update', [['project_id', '=', projectId]], ['name', 'project_id', 'status', 'progress', 'description']);

  return { ok: true, sheets: out };
}


function idsFromRel(value: any): number[] {
  if (value === null || value === undefined || value === false) return [];
  if (typeof value === 'number') return Number.isFinite(value) ? [value] : [];
  if (typeof value === 'string' && /^\d+$/.test(value.trim())) return [Number(value.trim())];
  if (Array.isArray(value)) {
    // many2one from search_read usually comes as [id, display_name]
    if (value.length === 2 && typeof value[0] === 'number' && typeof value[1] === 'string') return [Number(value[0])].filter(Boolean);
    // many2many/one2many usually comes as [1,2,3]
    return value.flatMap(v => idsFromRel(v)).filter(Boolean);
  }
  return [];
}

function uniqNums(values: number[]) {
  return Array.from(new Set(values.map(Number).filter(Boolean)));
}

function collectIds(records: AnyRow[], fields: string[]) {
  return uniqNums(records.flatMap(row => fields.flatMap(field => idsFromRel(row[field]))));
}

async function tryExportSheetAs(conn: Conn, uid: number, out: Record<string, AnyRow[]>, sheetName: string, model: string, domain: any[], fields: string[], limit = 5000) {
  try {
    const schema = await fieldsGet(conn, uid, model);
    const realFields = ['id', ...fields.filter(f => schema[f] && f !== 'id')];
    const records = await executeKw(conn, uid, model, 'search_read', [domain], { fields: realFields, limit, order: 'id asc' });
    const ids = (records || []).map((r: AnyRow) => Number(r.id)).filter(Boolean);
    const extMap = await mapExternalIdsForModel(conn, uid, model, ids);
    out[sheetName] = (records || []).map((r: AnyRow) => sanitizeExportRow({
      _model: model,
      __action: 'upsert',
      _external_id: extMap.get(Number(r.id)) || '',
      x_studio2_odoo_id: r.id,
      ...r
    }));
    return records || [];
  } catch (err: any) {
    out[`${sheetName}_error`] = [{ _model: model, error: err?.message || String(err) }];
    return [];
  }
}

async function tryExportSheet(conn: Conn, uid: number, out: Record<string, AnyRow[]>, model: string, domain: any[], fields: string[], limit = 5000) {
  return tryExportSheetAs(conn, uid, out, model, model, domain, fields, limit);
}

function addRelationship(out: Record<string, AnyRow[]>, fromModel: string, fromIds: number[], relationName: string, toModel: string, toIds: number[], note = '') {
  if (!out.relationship_map) out.relationship_map = [];
  out.relationship_map.push({
    from_model: fromModel,
    from_ids: uniqNums(fromIds).join(','),
    relation: relationName,
    to_model: toModel,
    to_ids: uniqNums(toIds).join(','),
    note
  });
}

function addBundleReadme(out: Record<string, AnyRow[]>, bundle: string, primaryModel: string, ids: number[], summary: AnyRow) {
  out.README_EXPORT = [{
    bundle,
    primary_model: primaryModel,
    primary_ids: ids.join(','),
    exported_at: new Date().toISOString(),
    note: 'Smart Bundle Export Studio2. Sheet utama dan sheet relasi dibuat untuk review/edit/import ulang Odoo. Relasi tetap disimpan sebagai field Odoo mentah plus x_studio2_odoo_id/external id jika tersedia.',
    summary: JSON.stringify(summary).slice(0, SAFE_CELL_LIMIT)
  }];
}

async function handleExportBundle(conn: Conn, body: AnyRow) {
  const uid = await login(conn);
  const bundle = String(body.bundle || '').trim() || 'single';
  const primaryModel = String(body.primary_model || body.model || '').trim();
  const ids = uniqNums((body.ids || []).map((x: any) => Number(x)));
  const includes = body.includes || {};
  if (!primaryModel) throw new Error('Primary model kosong.');
  if (!ids.length) throw new Error('Belum ada record utama dipilih.');

  const out: Record<string, AnyRow[]> = {};
  const summary: AnyRow = { bundle, primary_model: primaryModel, primary_count: ids.length };

  if (bundle === 'project') {
    const projects = await tryExportSheet(conn, uid, out, 'project.project', [['id', 'in', ids]], ['name', 'display_name', 'partner_id', 'user_id', 'date_start', 'date', 'description', 'privacy_visibility', 'stage_id', 'active']);
    const tasks = await tryExportSheet(conn, uid, out, 'project.task', [['project_id', 'in', ids]], ['name', 'display_name', 'project_id', 'parent_id', 'stage_id', 'user_id', 'user_ids', 'partner_id', 'date_deadline', 'description', 'priority', 'sequence', 'active', 'kanban_state']);
    const taskIds = tasks.map((r: AnyRow) => Number(r.id)).filter(Boolean);
    const partnerIds = uniqNums([...collectIds(projects, ['partner_id']), ...collectIds(tasks, ['partner_id'])]);
    const userIds = uniqNums([...collectIds(projects, ['user_id']), ...collectIds(tasks, ['user_id', 'user_ids'])]);
    const stageIds = collectIds(tasks, ['stage_id']);
    await tryExportSheet(conn, uid, out, 'project.milestone', [['project_id', 'in', ids]], ['name', 'display_name', 'project_id', 'deadline', 'is_reached', 'sequence']);
    await tryExportSheet(conn, uid, out, 'project.update', [['project_id', 'in', ids]], ['name', 'display_name', 'project_id', 'status', 'progress', 'description']);
    if (partnerIds.length) await tryExportSheet(conn, uid, out, 'res.partner', [['id', 'in', partnerIds]], ['name', 'display_name', 'email', 'phone', 'mobile', 'street', 'city', 'is_company', 'customer_rank', 'supplier_rank', 'category_id']);
    if (userIds.length) await tryExportSheet(conn, uid, out, 'res.users', [['id', 'in', userIds]], ['name', 'login', 'partner_id', 'active']);
    if (stageIds.length) await tryExportSheet(conn, uid, out, 'project.task.type', [['id', 'in', stageIds]], ['name', 'sequence', 'fold', 'project_ids']);
    out.task_hierarchy = tasks.map((t: AnyRow) => ({ task_id: t.id, task_name: t.name || t.display_name, project_id: idsFromRel(t.project_id).join(','), parent_id: idsFromRel(t.parent_id).join(','), stage_id: idsFromRel(t.stage_id).join(','), sequence: t.sequence, deadline: t.date_deadline }));
    addRelationship(out, 'project.project', ids, 'project_id', 'project.task', taskIds, 'Semua task/subtask dalam project terpilih.');
    addRelationship(out, 'project.project/project.task', ids.concat(taskIds), 'partner_id', 'res.partner', partnerIds, 'Partner/customer/vendor terkait project dan task.');
    addRelationship(out, 'project.project/project.task', ids.concat(taskIds), 'user_id/user_ids', 'res.users', userIds, 'Responsible/user reference.');
    summary.tasks = taskIds.length; summary.partners = partnerIds.length; summary.users = userIds.length;
  } else if (bundle === 'contact') {
    const partners = await tryExportSheet(conn, uid, out, 'res.partner', [['id', 'in', ids]], ['name', 'display_name', 'parent_id', 'child_ids', 'email', 'phone', 'mobile', 'street', 'street2', 'city', 'state_id', 'country_id', 'is_company', 'customer_rank', 'supplier_rank', 'category_id', 'comment', 'active']);
    const children = await tryExportSheetAs(conn, uid, out, 'res.partner.children', 'res.partner', [['parent_id', 'in', ids]], ['name', 'display_name', 'parent_id', 'email', 'phone', 'mobile', 'street', 'city', 'type', 'active'], 2000);
    // Keep import-safe model name in rows even if sheet name is an alias.
    if (out['res.partner.children']) out['res.partner.children'] = out['res.partner.children'].map(r => ({ ...r, _model: 'res.partner' }));
    const allPartners = uniqNums([...ids, ...children.map((r: AnyRow) => Number(r.id)).filter(Boolean)]);
    const tagIds = uniqNums([...collectIds(partners, ['category_id']), ...collectIds(children, ['category_id'])]);
    if (tagIds.length) await tryExportSheet(conn, uid, out, 'res.partner.category', [['id', 'in', tagIds]], ['name', 'display_name', 'parent_id', 'active']);
    if (includes.sales) {
      const orders = await tryExportSheet(conn, uid, out, 'sale.order', [['partner_id', 'in', allPartners]], ['name', 'display_name', 'partner_id', 'date_order', 'state', 'amount_total', 'invoice_status'], 2000);
      addRelationship(out, 'res.partner', allPartners, 'partner_id', 'sale.order', orders.map((r: AnyRow) => Number(r.id)).filter(Boolean), 'Sales order terkait contact.');
    }
    if (includes.projects) {
      const tasks = await tryExportSheet(conn, uid, out, 'project.task', [['partner_id', 'in', allPartners]], ['name', 'display_name', 'project_id', 'partner_id', 'stage_id', 'date_deadline', 'user_ids'], 2000);
      addRelationship(out, 'res.partner', allPartners, 'partner_id', 'project.task', tasks.map((r: AnyRow) => Number(r.id)).filter(Boolean), 'Task terkait contact.');
    }
    addRelationship(out, 'res.partner', ids, 'parent_id/child_ids', 'res.partner', children.map((r: AnyRow) => Number(r.id)).filter(Boolean), 'Alamat/kontak anak dari partner utama.');
    addRelationship(out, 'res.partner', allPartners, 'category_id', 'res.partner.category', tagIds, 'Tag/kategori partner.');
    summary.child_contacts = children.length; summary.tags = tagIds.length;
  } else if (bundle === 'product') {
    const templates = await tryExportSheet(conn, uid, out, 'product.template', [['id', 'in', ids]], ['name', 'display_name', 'default_code', 'barcode', 'list_price', 'standard_price', 'categ_id', 'public_categ_ids', 'uom_id', 'uom_po_id', 'sale_ok', 'purchase_ok', 'website_published', 'description_sale', 'active']);
    const variants = await tryExportSheet(conn, uid, out, 'product.product', [['product_tmpl_id', 'in', ids]], ['name', 'display_name', 'product_tmpl_id', 'default_code', 'barcode', 'lst_price', 'standard_price', 'active'], 5000);
    await tryExportSheet(conn, uid, out, 'product.supplierinfo', [['product_tmpl_id', 'in', ids]], ['partner_id', 'product_tmpl_id', 'product_id', 'price', 'min_qty', 'delay', 'currency_id'], 5000);
    const categIds = collectIds(templates, ['categ_id']);
    const publicCategIds = collectIds(templates, ['public_categ_ids']);
    const uomIds = uniqNums([...collectIds(templates, ['uom_id', 'uom_po_id'])]);
    if (categIds.length) await tryExportSheet(conn, uid, out, 'product.category', [['id', 'in', categIds]], ['name', 'display_name', 'parent_id', 'complete_name']);
    if (publicCategIds.length) await tryExportSheet(conn, uid, out, 'product.public.category', [['id', 'in', publicCategIds]], ['name', 'display_name', 'parent_id']);
    if (uomIds.length) await tryExportSheet(conn, uid, out, 'uom.uom', [['id', 'in', uomIds]], ['name', 'display_name', 'category_id', 'uom_type', 'factor']);
    addRelationship(out, 'product.template', ids, 'product_tmpl_id', 'product.product', variants.map((r: AnyRow) => Number(r.id)).filter(Boolean), 'Variant produk.');
    addRelationship(out, 'product.template', ids, 'categ_id/public_categ_ids', 'product.category/product.public.category', categIds.concat(publicCategIds), 'Kategori teknis dan website.');
    summary.variants = variants.length; summary.categories = categIds.length + publicCategIds.length;
  } else if (bundle === 'sales') {
    const orders = await tryExportSheet(conn, uid, out, 'sale.order', [['id', 'in', ids]], ['name', 'display_name', 'partner_id', 'date_order', 'state', 'amount_total', 'invoice_status', 'user_id', 'team_id', 'validity_date']);
    const lines = await tryExportSheet(conn, uid, out, 'sale.order.line', [['order_id', 'in', ids]], ['order_id', 'product_id', 'product_template_id', 'name', 'product_uom_qty', 'price_unit', 'discount', 'price_subtotal'], 5000);
    const partnerIds = collectIds(orders, ['partner_id']);
    const userIds = collectIds(orders, ['user_id']);
    const productIds = collectIds(lines, ['product_id']);
    const productTemplateIds = collectIds(lines, ['product_template_id']);
    if (partnerIds.length) await tryExportSheet(conn, uid, out, 'res.partner', [['id', 'in', partnerIds]], ['name', 'display_name', 'email', 'phone', 'mobile', 'street', 'city', 'customer_rank', 'supplier_rank']);
    if (userIds.length) await tryExportSheet(conn, uid, out, 'res.users', [['id', 'in', userIds]], ['name', 'login', 'partner_id', 'active']);
    if (productIds.length) await tryExportSheet(conn, uid, out, 'product.product', [['id', 'in', productIds]], ['name', 'display_name', 'product_tmpl_id', 'default_code', 'barcode']);
    if (productTemplateIds.length) await tryExportSheet(conn, uid, out, 'product.template', [['id', 'in', productTemplateIds]], ['name', 'display_name', 'default_code', 'barcode', 'list_price', 'categ_id']);
    addRelationship(out, 'sale.order', ids, 'order_id', 'sale.order.line', lines.map((r: AnyRow) => Number(r.id)).filter(Boolean), 'Order line.');
    summary.lines = lines.length; summary.partners = partnerIds.length; summary.products = productIds.length + productTemplateIds.length;
  } else if (bundle === 'knowledge') {
    const articles = await tryExportSheet(conn, uid, out, 'knowledge.article', [['id', 'in', ids]], ['name', 'display_name', 'parent_id', 'body', 'body_html', 'article_member_ids', 'create_date', 'write_date']);
    const children = await tryExportSheetAs(conn, uid, out, 'knowledge.article.children', 'knowledge.article', [['parent_id', 'in', ids]], ['name', 'display_name', 'parent_id', 'body', 'body_html', 'create_date', 'write_date'], 2000);
    if (out['knowledge.article.children']) out['knowledge.article.children'] = out['knowledge.article.children'].map(r => ({ ...r, _model: 'knowledge.article' }));
    const parentIds = collectIds(articles, ['parent_id']);
    if (parentIds.length) await tryExportSheetAs(conn, uid, out, 'knowledge.article.parents', 'knowledge.article', [['id', 'in', parentIds]], ['name', 'display_name', 'parent_id']);
    if (out['knowledge.article.parents']) out['knowledge.article.parents'] = out['knowledge.article.parents'].map(r => ({ ...r, _model: 'knowledge.article' }));
    addRelationship(out, 'knowledge.article', ids, 'parent_id/children', 'knowledge.article', parentIds.concat(children.map((r: AnyRow) => Number(r.id)).filter(Boolean)), 'Parent dan child article.');
    summary.children = children.length; summary.parents = parentIds.length;
  } else {
    await tryExportSheet(conn, uid, out, primaryModel, [['id', 'in', ids]], Array.isArray(body.fields) ? body.fields : ['name', 'display_name', 'create_date', 'write_date']);
  }

  summary.sheets = Object.keys(out).length;
  addBundleReadme(out, bundle, primaryModel, ids, summary);
  return { ok: true, bundle, primary_model: primaryModel, sheets: out, summary };
}

async function handleImportBatch(conn: Conn, body: AnyRow) {
  const uid = await login(conn);
  const rows: AnyRow[] = Array.isArray(body.rows) ? body.rows : [];
  if (!rows.length) throw new Error('Rows kosong.');
  const defaultModel = body.model || rows[0]?._model;
  if (!defaultModel) throw new Error('Model kosong. Isi _model di XLSX atau pilih model target.');
  const schemas = new Map<string, Record<string, OdooField>>();
  const xmlCache = new Map<string, { model: string; id: number } | null>();
  const results: AnyRow[] = [];
  let created = 0;
  let updated = 0;
  let skipped = 0;
  let failed = 0;

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const rowIndex = row.__rownum__ || i + 1;
    const model = row._model || defaultModel;
    try {
      if (!schemas.has(model)) schemas.set(model, await fieldsGet(conn, uid, model));
      const schema = schemas.get(model)!;
      const action = String(row.__action || 'upsert').trim().toLowerCase();
      if (action === 'skip' || action === 'ignore') {
        skipped++;
        results.push({ row: rowIndex, status: 'skipped', model });
        continue;
      }
      const { vals, skipped: skippedCols, warnings } = await buildVals(conn, uid, model, row, schema, xmlCache);
      const external = row._external_id || row.external_id || '';
      const odooId = Number(row.x_studio2_odoo_id || row.id || 0);
      let existing: { model: string; id: number } | null = null;
      if (external) existing = await resolveXmlId(conn, uid, String(external), xmlCache);
      if (!existing && odooId) existing = { model, id: odooId };

      if (action === 'delete' || action === 'unlink') {
        if (!existing?.id) throw new Error('Tidak bisa delete: record tidak ditemukan dari _external_id / x_studio2_odoo_id.');
        await executeKw(conn, uid, model, 'unlink', [[existing.id]]);
        results.push({ row: rowIndex, status: 'deleted', id: existing.id, model, skippedCols, warnings });
        continue;
      }

      if (existing?.id && (action === 'upsert' || action === 'update' || action === 'write')) {
        if (Object.keys(vals).length) await executeKw(conn, uid, model, 'write', [[existing.id], vals]);
        updated++;
        results.push({ row: rowIndex, status: 'updated', id: existing.id, model, skippedCols, warnings });
      } else if (action === 'upsert' || action === 'create') {
        const newId = await executeKw(conn, uid, model, 'create', [vals]);
        if (external) await createExternalId(conn, uid, model, Number(newId), String(external));
        created++;
        results.push({ row: rowIndex, status: 'created', id: newId, model, skippedCols, warnings });
      } else {
        skipped++;
        results.push({ row: rowIndex, status: 'skipped_unknown_action', action, model });
      }
    } catch (err: any) {
      failed++;
      results.push({ row: rowIndex, status: 'error', model, error: err?.message || String(err) });
    }
  }

  return { ok: failed === 0, processed: rows.length, created, updated, skipped, failed, results };
}

async function handleNameSearch(conn: Conn, body: AnyRow) {
  const uid = await login(conn);
  const model = body.model;
  const name = body.name || '';
  if (!model) throw new Error('Model kosong.');
  const res = await executeKw(conn, uid, model, 'name_search', [name], { limit: Math.min(Number(body.limit || 20), 80) });
  return { ok: true, records: res };
}

export async function GET() {
  return json({
    ok: true,
    app: 'Studio2 v10 Vercel-only Odoo Data Command Studio',
    actions: ['test', 'schema', 'record_scan', 'export_records', 'export_bundle', 'export_project', 'import_batch', 'name_search'],
    note: 'Designed for short serverless calls. Browser handles XLSX preview/editor/batching.'
  });
}

export async function POST(req: NextRequest) {
  try {
    const body = await req.json();
    const conn = body.connection as Conn;
    const action = body.action;
    if (!action) throw new Error('Action kosong.');

    let result: any;
    if (action === 'test') result = await handleTest(conn);
    else if (action === 'schema') result = await handleSchema(conn, body);
    else if (action === 'record_scan') result = await handleRecordScan(conn, body);
    else if (action === 'export_records') result = await handleExportRecords(conn, body);
    else if (action === 'export_bundle') result = await handleExportBundle(conn, body);
    else if (action === 'export_project') result = await handleExportProject(conn, body);
    else if (action === 'import_batch') result = await handleImportBatch(conn, body);
    else if (action === 'name_search') result = await handleNameSearch(conn, body);
    else throw new Error(`Action tidak dikenal: ${action}`);

    return json(result);
  } catch (err: any) {
    return json({ ok: false, error: err?.message || String(err) }, 200);
  }
}
