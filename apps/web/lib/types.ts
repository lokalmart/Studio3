export type OdooTarget = {
  url: string;
  db: string;
  username: string;
  password: string;
};

export type EngineConfig = {
  baseUrl: string;
  apiKey: string;
};

export type SheetData = {
  name: string;
  model: string;
  rows: Record<string, any>[];
  columns: string[];
  kind: 'product' | 'contact' | 'project' | 'knowledge' | 'sales' | 'dynamic' | 'context';
};

export type WorkbookState = {
  fileName: string;
  sheets: SheetData[];
};

export type OdooField = {
  string?: string;
  type?: string;
  required?: boolean;
  readonly?: boolean;
  relation?: string;
  selection?: [string, string][];
  help?: string;
};

export type OdooSchema = Record<string, OdooField>;

export type Job = {
  id: string;
  type: string;
  status: 'queued' | 'running' | 'done' | 'failed';
  progress: number;
  logs: { time: string; level: string; message: string }[];
  warnings: string[];
  errors: string[];
  result?: any;
  downloadPath?: boolean | null;
};
