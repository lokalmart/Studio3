import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';

const JOB_DIR = process.env.JOB_DIR || path.join(process.cwd(), 'jobs');
fs.mkdirSync(JOB_DIR, { recursive: true });

const jobs = new Map();

export function createJob(type, meta = {}) {
  const id = crypto.randomUUID();
  const job = {
    id,
    type,
    status: 'queued',
    progress: 0,
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
    meta,
    logs: [],
    warnings: [],
    errors: [],
    result: null,
    downloadPath: null
  };
  jobs.set(id, job);
  return job;
}

export function getJob(id) {
  return jobs.get(id) || null;
}

export function jobFilePath(jobId, filename = 'result.xlsx') {
  const safe = filename.replace(/[^a-zA-Z0-9._-]/g, '_');
  return path.join(JOB_DIR, `${jobId}_${safe}`);
}

export function patchJob(job, patch = {}) {
  Object.assign(job, patch, { updatedAt: new Date().toISOString() });
  jobs.set(job.id, job);
  return job;
}

export function logJob(job, message, level = 'info') {
  const line = { time: new Date().toISOString(), level, message };
  job.logs.push(line);
  if (level === 'warn') job.warnings.push(message);
  if (level === 'error') job.errors.push(message);
  patchJob(job, {});
}

export async function runJob(job, runner) {
  patchJob(job, { status: 'running', progress: 1 });
  try {
    const result = await runner(job);
    patchJob(job, { status: 'done', progress: 100, result: result || job.result });
  } catch (error) {
    const message = error?.stack || error?.message || String(error);
    logJob(job, message, 'error');
    patchJob(job, { status: 'failed', progress: job.progress || 0 });
  }
}
