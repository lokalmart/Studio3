const sessionCache = new Map();
const schemaCache = new Map();
const xmlidCache = new Map();

function normalizeUrl(url) {
  if (!url) throw new Error('Odoo URL kosong.');
  return String(url).replace(/\/+$/, '');
}

function cacheKey(target) {
  return `${normalizeUrl(target.url)}|${target.db}|${target.username}`;
}

export function assertTarget(target) {
  if (!target || !target.url || !target.db || !target.username || !target.password) {
    throw new Error('Target Odoo belum lengkap: url, db, username, password wajib diisi.');
  }
}

export async function jsonRpc(target, service, method, args) {
  const url = `${normalizeUrl(target.url)}/jsonrpc`;
  const payload = {
    jsonrpc: '2.0',
    method: 'call',
    params: { service, method, args },
    id: Math.floor(Math.random() * 1e9)
  };

  const response = await fetch(url, {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify(payload)
  });

  const text = await response.text();
  let data;
  try {
    data = JSON.parse(text);
  } catch {
    throw new Error(`Odoo mengembalikan non-JSON HTTP ${response.status}: ${text.slice(0, 400)}`);
  }
  if (!response.ok || data.error) {
    const err = data.error?.data?.message || data.error?.message || text;
    throw new Error(`Odoo JSON-RPC error: ${err}`);
  }
  return data.result;
}

export async function authenticate(target) {
  assertTarget(target);
  const key = cacheKey(target);
  const cached = sessionCache.get(key);
  if (cached?.uid) return cached.uid;
  const uid = await jsonRpc(target, 'common', 'authenticate', [target.db, target.username, target.password, {}]);
  if (!uid) throw new Error('Login Odoo gagal. Periksa database, email, dan password/API key.');
  sessionCache.set(key, { uid, at: Date.now() });
  return uid;
}

export async function executeKw(target, model, method, args = [], kwargs = {}) {
  const uid = await authenticate(target);
  return jsonRpc(target, 'object', 'execute_kw', [
    target.db,
    uid,
    target.password,
    model,
    method,
    args,
    kwargs
  ]);
}

export async function modelExists(target, model) {
  try {
    const count = await executeKw(target, 'ir.model', 'search_count', [[['model', '=', model]]]);
    return count > 0;
  } catch {
    return false;
  }
}

export async function fieldsGet(target, model) {
  const key = `${cacheKey(target)}|${model}`;
  if (schemaCache.has(key)) return schemaCache.get(key);
  const fields = await executeKw(target, model, 'fields_get', [], {
    attributes: ['string', 'type', 'required', 'readonly', 'relation', 'selection', 'help']
  });
  schemaCache.set(key, fields || {});
  return fields || {};
}

export async function searchRead(target, model, domain = [], fields = [], options = {}) {
  const kwargs = {
    fields,
    limit: options.limit || 80,
    offset: options.offset || 0,
    order: options.order || 'write_date desc, id desc'
  };
  return executeKw(target, model, 'search_read', [domain], kwargs);
}

export async function readRecords(target, model, ids, fields = []) {
  if (!ids?.length) return [];
  return executeKw(target, model, 'read', [ids], { fields });
}

export async function nameSearch(target, model, name = '', limit = 20) {
  return executeKw(target, model, 'name_search', [name], { limit });
}

export async function resolveXmlId(target, xmlid) {
  if (!xmlid) return false;
  const key = `${cacheKey(target)}|xmlid|${xmlid}`;
  if (xmlidCache.has(key)) return xmlidCache.get(key);
  let id = false;
  try {
    id = await executeKw(target, 'ir.model.data', '_xmlid_to_res_id', [xmlid, false]);
  } catch {
    id = false;
  }
  xmlidCache.set(key, id || false);
  return id || false;
}

export async function createExternalId(target, model, resId, xmlid) {
  if (!xmlid || !resId) return null;
  const parts = String(xmlid).split('.');
  const module = parts.length > 1 ? parts[0] : 'studio2';
  const name = parts.length > 1 ? parts.slice(1).join('.') : parts[0];
  const existing = await resolveXmlId(target, `${module}.${name}`);
  if (existing) return existing;
  try {
    await executeKw(target, 'ir.model.data', 'create', [[{
      module,
      name,
      model,
      res_id: resId,
      noupdate: true
    }]]);
  } catch (error) {
    // Jangan gagalkan import hanya karena external ID gagal dibuat.
    return null;
  }
  return resId;
}

export async function getExternalIds(target, model, ids) {
  if (!ids?.length) return {};
  try {
    const result = await executeKw(target, model, 'get_external_id', [ids]);
    return result || {};
  } catch {
    return {};
  }
}
