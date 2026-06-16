import { fieldsGet, executeKw, resolveXmlId, createExternalId } from './odoo.js';
import { inferModelFromRows, isHelperSheet, readWorkbook, sheetToRows } from './xlsx-tools.js';
import { logJob, patchJob } from './jobs.js';

function isEmpty(v) {
  return v === '' || v === null || v === undefined;
}

function boolValue(v) {
  if (typeof v === 'boolean') return v;
  const s = String(v || '').trim().toLowerCase();
  return ['true', '1', 'yes', 'ya', 'y', 'on'].includes(s);
}

function intValue(v) {
  if (isEmpty(v)) return false;
  const n = Number(String(v).replace(/,/g, ''));
  return Number.isFinite(n) ? Math.trunc(n) : false;
}

function floatValue(v) {
  if (isEmpty(v)) return false;
  const n = Number(String(v).replace(/,/g, ''));
  return Number.isFinite(n) ? n : false;
}

function splitList(v) {
  if (Array.isArray(v)) return v;
  return String(v || '').split(/[;,]/).map(x => x.trim()).filter(Boolean);
}

async function resolveMany2one(target, value) {
  if (isEmpty(value)) return false;
  if (/^\d+$/.test(String(value))) return Number(value);
  const id = await resolveXmlId(target, String(value).trim());
  return id || false;
}

async function buildValues(target, model, row, schema) {
  const values = {};
  const warnings = [];
  const keys = Object.keys(row || {});

  for (const key of keys) {
    if (!key || key.startsWith('_') || key === 'ID' || key === 'id' || key === '__action') continue;
    const raw = row[key];
    if (isEmpty(raw)) continue;

    if (key.endsWith('_external_id')) {
      const fieldName = key.replace(/_external_id$/, '');
      const field = schema[fieldName];
      if (!field) {
        warnings.push(`Kolom relation ${key} dilewati karena field ${fieldName} tidak ada di ${model}.`);
        continue;
      }
      if (field.type !== 'many2one') {
        warnings.push(`Kolom ${key} dilewati karena ${fieldName} bukan many2one.`);
        continue;
      }
      values[fieldName] = await resolveMany2one(target, raw);
      continue;
    }

    if (key.endsWith('_external_ids')) {
      const fieldName = key.replace(/_external_ids$/, '');
      const field = schema[fieldName];
      if (!field) {
        warnings.push(`Kolom relation ${key} dilewati karena field ${fieldName} tidak ada di ${model}.`);
        continue;
      }
      const ids = [];
      for (const xmlid of splitList(raw)) {
        const id = await resolveMany2one(target, xmlid);
        if (id) ids.push(id);
      }
      if (field.type === 'many2many') values[fieldName] = [[6, 0, ids]];
      else if (field.type === 'many2one') values[fieldName] = ids[0] || false;
      else warnings.push(`Kolom ${key} dilewati karena ${fieldName} bukan many2one/many2many.`);
      continue;
    }

    const field = schema[key];
    if (!field) {
      warnings.push(`Kolom ${key} dilewati karena bukan field ${model}.`);
      continue;
    }
    if (field.readonly && !['active'].includes(key)) {
      warnings.push(`Kolom ${key} readonly, dilewati.`);
      continue;
    }

    switch (field.type) {
      case 'boolean': values[key] = boolValue(raw); break;
      case 'integer': values[key] = intValue(raw); break;
      case 'float':
      case 'monetary': values[key] = floatValue(raw); break;
      case 'many2one': values[key] = await resolveMany2one(target, raw); break;
      case 'many2many': {
        const ids = splitList(raw).map(x => /^\d+$/.test(x) ? Number(x) : null).filter(Boolean);
        values[key] = [[6, 0, ids]];
        break;
      }
      case 'one2many':
      case 'binary':
        warnings.push(`Kolom ${key} bertipe ${field.type}, dilewati untuk import aman.`);
        break;
      default:
        values[key] = raw;
    }
  }
  return { values, warnings };
}

export async function importWorkbookJob(job, target, filePath, options = {}) {
  const wb = readWorkbook(filePath);
  const batchLimit = Math.max(1, Math.min(Number(options.batchLimit || 80), 250));
  const sheetNames = wb.SheetNames || [];
  let processed = 0;
  let created = 0;
  let updated = 0;
  let skipped = 0;
  const errors = [];
  const warnings = [];

  const totalRows = sheetNames.reduce((n, s) => n + sheetToRows(wb, s).length, 0) || 1;

  for (const sheetName of sheetNames) {
    const rows = sheetToRows(wb, sheetName);
    if (!rows.length) continue;
    if (isHelperSheet(sheetName)) {
      skipped += rows.length;
      logJob(job, `Sheet helper dilewati: ${sheetName}`, 'warn');
      continue;
    }
    const model = inferModelFromRows(sheetName, rows);
    if (!model) {
      skipped += rows.length;
      warnings.push(`${sheetName}: tidak punya _model dan nama sheet bukan model Odoo.`);
      continue;
    }
    const schema = await fieldsGet(target, model);
    logJob(job, `Import sheet ${sheetName} sebagai ${model} (${rows.length} row).`);

    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      const line = i + 2;
      try {
        const rowModel = row._model || model;
        if (rowModel !== model) {
          // Jika model beda per row, ambil schema baru.
        }
        const action = String(row.__action || row._action || 'upsert').toLowerCase();
        const externalId = row._external_id || row.external_id || '';
        const idValue = row.ID || row.id || '';
        const { values, warnings: rowWarnings } = await buildValues(target, model, row, schema);
        rowWarnings.forEach(w => warnings.push(`${sheetName} row ${line}: ${w}`));
        if (!Object.keys(values).length) {
          skipped++;
          continue;
        }

        let resId = false;
        if (externalId) resId = await resolveXmlId(target, externalId);
        if (!resId && /^\d+$/.test(String(idValue))) resId = Number(idValue);

        if ((action === 'update' || action === 'upsert') && resId) {
          await executeKw(target, model, 'write', [[resId], values]);
          updated++;
        } else if (action === 'create' || action === 'upsert' || action === 'insert') {
          const newId = await executeKw(target, model, 'create', [values]);
          if (externalId) await createExternalId(target, model, newId, externalId);
          created++;
        } else {
          skipped++;
        }
      } catch (error) {
        errors.push(`${sheetName} row ${line}: ${error.message}`);
        logJob(job, `${sheetName} row ${line}: ${error.message}`, 'error');
        if (options.stopOnError) throw error;
      }
      processed++;
      if (processed % batchLimit === 0) {
        patchJob(job, { progress: Math.min(95, Math.round((processed / totalRows) * 95)) });
      }
    }
  }

  patchJob(job, { warnings, errors });
  return { processed, created, updated, skipped, errors, warnings };
}
