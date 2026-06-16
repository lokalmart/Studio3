import { fieldsGet, searchRead, readRecords, getExternalIds } from './odoo.js';
import { writeWorkbook, safeExportRow } from './xlsx-tools.js';
import { jobFilePath, logJob, patchJob } from './jobs.js';

export function defaultScanFields(schema) {
  const preferred = [
    'id', 'display_name', 'name', 'email', 'phone', 'mobile', 'default_code', 'barcode',
    'list_price', 'standard_price', 'project_id', 'parent_id', 'stage_id', 'partner_id',
    'active', 'write_date', 'create_date'
  ];
  return preferred.filter(f => f === 'id' || schema[f]);
}

export function defaultExportFields(schema) {
  const preferred = [
    'id', 'display_name', 'name', 'email', 'phone', 'mobile', 'street', 'street2', 'city', 'zip',
    'default_code', 'barcode', 'list_price', 'standard_price', 'categ_id', 'public_categ_ids',
    'sale_ok', 'purchase_ok', 'website_published', 'project_id', 'parent_id', 'stage_id', 'user_ids',
    'partner_id', 'date_deadline', 'description', 'body', 'body_html', 'active', 'write_date', 'create_date'
  ];
  const direct = preferred.filter(f => f === 'id' || schema[f]);
  if (direct.length > 1) return direct;
  return Object.entries(schema)
    .filter(([, meta]) => !['binary', 'one2many'].includes(meta.type))
    .slice(0, 40)
    .map(([name]) => name);
}

export async function scanRecords(target, model, options = {}) {
  const schema = await fieldsGet(target, model);
  const fields = options.fields?.length ? options.fields : defaultScanFields(schema);
  const domain = Array.isArray(options.domain) ? options.domain : [];
  const limit = Math.max(1, Math.min(Number(options.limit || 120), 500));
  const offset = Number(options.offset || 0);
  const records = await searchRead(target, model, domain, fields, { limit, offset, order: options.order });
  return { model, fields, records, schema };
}

function flattenRecordForXlsx(record, model, externalIds = {}) {
  const out = { _model: model, __action: 'update' };
  if (externalIds?.[record.id]) out._external_id = externalIds[record.id];
  for (const [k, v] of Object.entries(record)) {
    if (k === 'id') out.ID = v;
    else out[k] = v;
  }
  return safeExportRow(out);
}

export async function exportRecordsJob(job, target, model, ids, fields = []) {
  if (!ids?.length) throw new Error('Belum ada record yang dipilih untuk export.');
  const schema = await fieldsGet(target, model);
  const exportFields = fields?.length ? fields : defaultExportFields(schema);
  const cleanFields = ['id', ...exportFields.filter(f => f !== 'id')];
  logJob(job, `Membaca ${ids.length} record dari ${model}.`);
  const records = await readRecords(target, model, ids, cleanFields);
  patchJob(job, { progress: 50 });
  const externalIds = await getExternalIds(target, model, ids);
  const rows = records.map(r => flattenRecordForXlsx(r, model, externalIds));
  const outPath = jobFilePath(job.id, `${model.replace(/\./g, '_')}_selected_export.xlsx`);
  writeWorkbook([{ name: model, rows }], outPath);
  patchJob(job, { downloadPath: outPath, progress: 95 });
  return { model, ids: ids.length, fields: cleanFields, file: outPath };
}

export async function exportProjectJob(job, target, projectId, fieldsByModel = {}) {
  if (!projectId) throw new Error('projectId wajib diisi.');
  const sheets = [];

  logJob(job, `Export project.project ID ${projectId}.`);
  const projectSchema = await fieldsGet(target, 'project.project');
  const projectFields = fieldsByModel['project.project']?.length ? fieldsByModel['project.project'] : defaultExportFields(projectSchema);
  const projectRecords = await readRecords(target, 'project.project', [projectId], ['id', ...projectFields.filter(f => f !== 'id')]);
  const projectExt = await getExternalIds(target, 'project.project', [projectId]);
  sheets.push({ name: 'project.project', rows: projectRecords.map(r => flattenRecordForXlsx(r, 'project.project', projectExt)) });
  patchJob(job, { progress: 25 });

  const taskSchema = await fieldsGet(target, 'project.task');
  const taskFields = fieldsByModel['project.task']?.length ? fieldsByModel['project.task'] : defaultExportFields(taskSchema);
  const taskRecords = await searchRead(target, 'project.task', [['project_id', '=', Number(projectId)]], ['id', ...taskFields.filter(f => f !== 'id')], { limit: 2000, order: 'id asc' });
  const taskIds = taskRecords.map(r => r.id);
  const taskExt = await getExternalIds(target, 'project.task', taskIds);
  sheets.push({ name: 'project.task', rows: taskRecords.map(r => flattenRecordForXlsx(r, 'project.task', taskExt)) });
  patchJob(job, { progress: 60 });

  try {
    const milestoneSchema = await fieldsGet(target, 'project.milestone');
    const milestoneFields = fieldsByModel['project.milestone']?.length ? fieldsByModel['project.milestone'] : defaultExportFields(milestoneSchema);
    const milestones = await searchRead(target, 'project.milestone', [['project_id', '=', Number(projectId)]], ['id', ...milestoneFields.filter(f => f !== 'id')], { limit: 500, order: 'id asc' });
    const ids = milestones.map(r => r.id);
    const ext = await getExternalIds(target, 'project.milestone', ids);
    sheets.push({ name: 'project.milestone', rows: milestones.map(r => flattenRecordForXlsx(r, 'project.milestone', ext)) });
  } catch (error) {
    logJob(job, `project.milestone dilewati: ${error.message}`, 'warn');
  }

  try {
    const updateSchema = await fieldsGet(target, 'project.update');
    const updateFields = fieldsByModel['project.update']?.length ? fieldsByModel['project.update'] : defaultExportFields(updateSchema);
    const updates = await searchRead(target, 'project.update', [['project_id', '=', Number(projectId)]], ['id', ...updateFields.filter(f => f !== 'id')], { limit: 500, order: 'id asc' });
    const ids = updates.map(r => r.id);
    const ext = await getExternalIds(target, 'project.update', ids);
    sheets.push({ name: 'project.update', rows: updates.map(r => flattenRecordForXlsx(r, 'project.update', ext)) });
  } catch (error) {
    logJob(job, `project.update dilewati: ${error.message}`, 'warn');
  }

  sheets.push({ name: 'README_EXPORT', rows: [{
    note: 'Export dibuat oleh Studio2 v8. Edit di preview/editor sebelum import ulang.',
    project_id: projectId,
    exported_at: new Date().toISOString()
  }] });

  const outPath = jobFilePath(job.id, `project_${projectId}_export.xlsx`);
  writeWorkbook(sheets, outPath);
  patchJob(job, { downloadPath: outPath, progress: 95 });
  return { projectId, sheets: sheets.map(s => ({ name: s.name, rows: s.rows.length })), file: outPath };
}
