import fs from 'fs';
import path from 'path';
import * as XLSX from 'xlsx';
import type { SchemaEnrichment, SchemaColumn } from '../model/types';

export class SchemaEnrichmentParser {
  static parse(filePath: string): SchemaEnrichment {
    const ext = path.extname(filePath).toLowerCase();

    if (ext === '.csv') return SchemaEnrichmentParser.parseCsv(filePath);
    if (ext === '.xlsx' || ext === '.xls') return SchemaEnrichmentParser.parseXlsx(filePath);
    if (ext === '.json') return SchemaEnrichmentParser.parseJson(filePath);

    throw new Error(`Unsupported schema file type: ${ext}. Use CSV, XLSX, or JSON.`);
  }

  private static parseCsv(filePath: string): SchemaEnrichment {
    const content = fs.readFileSync(filePath, 'utf-8').replace(/^\uFEFF/, ''); // strip BOM
    const lines = content.split(/\r?\n/).filter(l => l.trim());
    if (lines.length === 0) return { fileName: path.basename(filePath), format: 'csv', columns: [] };

    const headers = lines[0].split(',').map(h => h.trim().replace(/^"|"$/g, ''));
    const sampleRows = lines.slice(1, 4);
    const columns: SchemaColumn[] = headers.map((name, i) => {
      const samples = sampleRows
        .map(row => row.split(',')[i]?.trim().replace(/^"|"$/g, '') ?? '')
        .filter(Boolean);
      return { name, type: SchemaEnrichmentParser.inferType(samples) };
    });

    return {
      fileName: path.basename(filePath),
      format: 'csv',
      columns,
      rowCount: lines.length - 1
    };
  }

  private static parseXlsx(filePath: string): SchemaEnrichment {
    const wb = XLSX.readFile(filePath);
    const ws = wb.Sheets[wb.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1 });

    if (data.length === 0) return { fileName: path.basename(filePath), format: 'xlsx', columns: [] };

    const headers = data[0].map(h => String(h ?? '').trim()).filter(Boolean);
    const sampleRows = data.slice(1, 4) as unknown[][];
    const columns: SchemaColumn[] = headers.map((name, i) => {
      const samples = sampleRows.map(row => String(row[i] ?? '')).filter(Boolean);
      return { name, type: SchemaEnrichmentParser.inferType(samples) };
    });

    return {
      fileName: path.basename(filePath),
      format: 'xlsx',
      columns,
      rowCount: data.length - 1
    };
  }

  private static parseJson(filePath: string): SchemaEnrichment {
    const raw = JSON.parse(fs.readFileSync(filePath, 'utf-8'));
    let columns: SchemaColumn[] = [];

    // Shape 1: array of objects → infer columns from keys
    if (Array.isArray(raw) && raw.length > 0 && typeof raw[0] === 'object') {
      const allKeys = new Set<string>();
      for (const row of raw.slice(0, 5)) Object.keys(row as object).forEach(k => allKeys.add(k));
      columns = [...allKeys].map(name => ({
        name,
        type: SchemaEnrichmentParser.inferType(
          raw.slice(0, 4).map((r: Record<string,unknown>) => String(r[name] ?? '')).filter(Boolean)
        )
      }));
    }
    // Shape 2: { columns: [{ name, type }] } or { fields: [...] }
    else if (typeof raw === 'object' && (raw.columns || raw.fields)) {
      const cols: Record<string,string>[] = raw.columns ?? raw.fields ?? [];
      columns = cols.map(c => ({
        name: String(c.name ?? c.Name ?? c.displayName ?? ''),
        displayName: String(c.displayName ?? c.DisplayName ?? ''),
        type: String(c.type ?? c.Type ?? c.dataType ?? ''),
        required: Boolean(c.required ?? c.Required ?? false)
      }));
    }

    return { fileName: path.basename(filePath), format: 'json', columns };
  }

  private static inferType(samples: string[]): string {
    if (samples.length === 0) return 'text';
    const booleans = samples.filter(s => /^(true|false|yes|no|0|1)$/i.test(s));
    const numbers = samples.filter(s => !isNaN(Number(s)) && s.trim() !== '');
    const dates = samples.filter(s => !isNaN(Date.parse(s)) && s.length > 5);
    if (booleans.length === samples.length) return 'boolean';
    if (numbers.length === samples.length) return 'number';
    if (dates.length === samples.length) return 'datetime';
    return 'text';
  }
}
