import type { EngineConfig, OdooTarget, Job } from './types';

function cleanBase(baseUrl: string) {
  return (baseUrl || '').replace(/\/+$/, '');
}

export async function engineFetch(config: EngineConfig, path: string, init: RequestInit = {}) {
  if (!config.baseUrl) throw new Error('Engine URL Render belum diisi.');
  const headers: Record<string, string> = { ...(init.headers as any || {}) };
  if (!(init.body instanceof FormData)) headers['content-type'] = headers['content-type'] || 'application/json';
  if (config.apiKey) headers['x-studio2-key'] = config.apiKey;
  const res = await fetch(`${cleanBase(config.baseUrl)}${path}`, { ...init, headers });
  const contentType = res.headers.get('content-type') || '';
  if (contentType.includes('application/json')) {
    const data = await res.json();
    if (!res.ok || data.ok === false) throw new Error(data.error || `HTTP ${res.status}`);
    return data;
  }
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  return res;
}

export async function testOdoo(config: EngineConfig, target: OdooTarget) {
  return engineFetch(config, '/odoo/test', { method: 'POST', body: JSON.stringify({ target }) });
}

export async function getSchema(config: EngineConfig, target: OdooTarget, model: string) {
  return engineFetch(config, '/odoo/schema', { method: 'POST', body: JSON.stringify({ target, model }) });
}

export async function recordScan(config: EngineConfig, target: OdooTarget, model: string, limit = 120, domain: any[] = []) {
  return engineFetch(config, '/odoo/record-scan', {
    method: 'POST',
    body: JSON.stringify({ target, model, limit, domain })
  });
}

export async function startExportRecords(config: EngineConfig, target: OdooTarget, model: string, ids: number[], fields: string[]) {
  return engineFetch(config, '/jobs/export-records', {
    method: 'POST',
    body: JSON.stringify({ target, model, ids, fields })
  });
}

export async function startExportProject(config: EngineConfig, target: OdooTarget, projectId: number) {
  return engineFetch(config, '/jobs/export-project', {
    method: 'POST',
    body: JSON.stringify({ target, projectId })
  });
}

export async function startImportXlsx(config: EngineConfig, target: OdooTarget, file: Blob, fileName: string, options: Record<string, any> = {}) {
  const form = new FormData();
  form.append('file', file, fileName);
  form.append('target', JSON.stringify(target));
  form.append('options', JSON.stringify(options));
  return engineFetch(config, '/jobs/import-xlsx', { method: 'POST', body: form });
}

export async function getJob(config: EngineConfig, id: string): Promise<{ ok: true; job: Job }> {
  return engineFetch(config, `/jobs/${id}`) as any;
}

export function downloadUrl(config: EngineConfig, id: string) {
  const base = cleanBase(config.baseUrl);
  const url = `${base}/jobs/${id}/download`;
  if (config.apiKey) return `${url}?key=${encodeURIComponent(config.apiKey)}`;
  return url;
}

export async function pollJob(config: EngineConfig, id: string, onUpdate: (job: Job) => void) {
  for (;;) {
    const data = await getJob(config, id);
    onUpdate(data.job);
    if (data.job.status === 'done' || data.job.status === 'failed') return data.job;
    await new Promise(resolve => setTimeout(resolve, 1200));
  }
}
