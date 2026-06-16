import express from 'express';
import cors from 'cors';
import multer from 'multer';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { authenticate, fieldsGet, nameSearch } from './lib/odoo.js';
import { createJob, getJob, jobFilePath, runJob } from './lib/jobs.js';
import { importWorkbookJob } from './lib/importer.js';
import { scanRecords, exportRecordsJob, exportProjectJob } from './lib/exporter.js';

const app = express();
const port = Number(process.env.PORT || 10000);
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 40 * 1024 * 1024 } });
const corsOrigin = process.env.CORS_ORIGIN || '*';
const apiKey = process.env.STUDIO2_ENGINE_API_KEY || '';

app.use(cors({ origin: corsOrigin === '*' ? true : corsOrigin.split(',').map(x => x.trim()), credentials: false }));
app.use(express.json({ limit: '10mb' }));

function requireKey(req, res, next) {
  if (!apiKey) return next();
  const got = req.header('x-studio2-key') || req.query.key;
  if (got !== apiKey) return res.status(401).json({ ok: false, error: 'Engine API key salah atau kosong.' });
  next();
}

app.get('/', (_req, res) => res.json({ ok: true, name: 'Lokalmart Studio2 Engine', version: '8.0.0' }));
app.get('/health', (_req, res) => res.json({ ok: true, status: 'healthy', time: new Date().toISOString() }));

app.post('/odoo/test', requireKey, async (req, res) => {
  try {
    const uid = await authenticate(req.body.target);
    res.json({ ok: true, uid });
  } catch (error) {
    res.status(400).json({ ok: false, error: error.message });
  }
});

app.post('/odoo/schema', requireKey, async (req, res) => {
  try {
    const { target, model } = req.body;
    const schema = await fieldsGet(target, model);
    res.json({ ok: true, model, schema });
  } catch (error) {
    res.status(400).json({ ok: false, error: error.message });
  }
});

app.post('/odoo/name-search', requireKey, async (req, res) => {
  try {
    const { target, model, name, limit } = req.body;
    const records = await nameSearch(target, model, name || '', limit || 20);
    res.json({ ok: true, records });
  } catch (error) {
    res.status(400).json({ ok: false, error: error.message });
  }
});

app.post('/odoo/record-scan', requireKey, async (req, res) => {
  try {
    const { target, model, domain, fields, limit, offset, order } = req.body;
    const result = await scanRecords(target, model, { domain, fields, limit, offset, order });
    res.json({ ok: true, ...result });
  } catch (error) {
    res.status(400).json({ ok: false, error: error.message });
  }
});

app.post('/jobs/export-records', requireKey, async (req, res) => {
  const { target, model, ids, fields } = req.body;
  const job = createJob('export-records', { model, count: ids?.length || 0 });
  res.json({ ok: true, job });
  runJob(job, j => exportRecordsJob(j, target, model, ids, fields));
});

app.post('/jobs/export-project', requireKey, async (req, res) => {
  const { target, projectId, fieldsByModel } = req.body;
  const job = createJob('export-project', { projectId });
  res.json({ ok: true, job });
  runJob(job, j => exportProjectJob(j, target, projectId, fieldsByModel || {}));
});

app.post('/jobs/import-xlsx', requireKey, upload.single('file'), async (req, res) => {
  try {
    const target = JSON.parse(req.body.target || '{}');
    const options = JSON.parse(req.body.options || '{}');
    if (!req.file) throw new Error('File XLSX belum dikirim.');
    const job = createJob('import-xlsx', { filename: req.file.originalname, size: req.file.size });
    const tmp = jobFilePath(job.id, req.file.originalname || 'upload.xlsx');
    fs.writeFileSync(tmp, req.file.buffer);
    res.json({ ok: true, job });
    runJob(job, j => importWorkbookJob(j, target, tmp, options));
  } catch (error) {
    res.status(400).json({ ok: false, error: error.message });
  }
});

app.get('/jobs/:id', requireKey, (req, res) => {
  const job = getJob(req.params.id);
  if (!job) return res.status(404).json({ ok: false, error: 'Job tidak ditemukan.' });
  res.json({ ok: true, job: { ...job, downloadPath: job.downloadPath ? true : null } });
});

app.get('/jobs/:id/download', requireKey, (req, res) => {
  const job = getJob(req.params.id);
  if (!job) return res.status(404).json({ ok: false, error: 'Job tidak ditemukan.' });
  if (!job.downloadPath || !fs.existsSync(job.downloadPath)) {
    return res.status(404).json({ ok: false, error: 'File hasil belum tersedia.' });
  }
  res.download(job.downloadPath, path.basename(job.downloadPath));
});

app.use((err, _req, res, _next) => {
  res.status(500).json({ ok: false, error: err.message || String(err) });
});

app.listen(port, () => {
  console.log(`Studio2 Engine listening on ${port}`);
});
