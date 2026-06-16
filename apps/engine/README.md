# Studio2 Engine

Render backend untuk job berat import/export Odoo.

## Endpoints

- `GET /health`
- `POST /odoo/test`
- `POST /odoo/schema`
- `POST /odoo/record-scan`
- `POST /jobs/import-xlsx`
- `POST /jobs/export-records`
- `POST /jobs/export-project`
- `GET /jobs/:id`
- `GET /jobs/:id/download`

Jika `STUDIO2_ENGINE_API_KEY` diisi, semua endpoint selain `/health` perlu header:

```text
x-studio2-key: <secret>
```
