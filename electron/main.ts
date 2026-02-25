import { app, BrowserWindow, ipcMain, dialog, shell } from 'electron';
import path from 'path';
import fs from 'fs';
import os from 'os';
import { ParserOrchestrator } from '../src/parsing/ParserOrchestrator';
import { HtmlGenerator } from '../src/generators/HtmlGenerator';
import { DocxGenerator } from '../src/generators/DocxGenerator';
import { MarkdownGenerator } from '../src/generators/MarkdownGenerator';
import { SchemaEnrichmentParser } from '../src/parsing/SchemaEnrichmentParser';
import type { NormalizedAnalysisGraph, Environment, SchemaPromptResponse } from '../src/model/types';

// ─────────────────────────────────────────────────────────────────────────────

let mainWindow: BrowserWindow;
let currentGraph: NormalizedAnalysisGraph | null = null;

const isDev = process.env.NODE_ENV === 'development' || !app.isPackaged;
const RENDERER_URL = 'http://localhost:5173';
const RENDERER_FILE = path.join(__dirname, '../../dist-renderer/index.html');

// ─── Window creation ──────────────────────────────────────────────────────────

function createWindow(): void {
  mainWindow = new BrowserWindow({
    width: 1100,
    height: 760,
    minWidth: 840,
    minHeight: 600,
    title: 'PP Doc Tool',
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      preload: path.join(__dirname, 'preload.js'),
      devTools: isDev
    }
  });

  if (isDev) {
    mainWindow.loadURL(RENDERER_URL);
    mainWindow.webContents.openDevTools({ mode: 'detach' });
  } else {
    mainWindow.loadFile(RENDERER_FILE);
  }

  mainWindow.webContents.setWindowOpenHandler(() => ({ action: 'deny' }));
}

app.whenReady().then(() => {
  createWindow();
  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

// ─── IPC Handlers ─────────────────────────────────────────────────────────────

/** Open native file picker for artifact or schema files */
ipcMain.handle('dialog:open-artifact', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    title: 'Select Power Platform Artifact',
    filters: [
      { name: 'Power Platform Artifacts', extensions: ['zip', 'msapp'] },
      { name: 'All Files', extensions: ['*'] }
    ],
    properties: ['openFile']
  });
  return result.canceled ? null : result.filePaths[0];
});

ipcMain.handle('dialog:open-schema', async (_event, dataSourceName: string) => {
  const result = await dialog.showOpenDialog(mainWindow, {
    title: `Select Schema File for: ${dataSourceName}`,
    filters: [
      { name: 'Schema Files', extensions: ['csv', 'xlsx', 'json'] },
      { name: 'All Files', extensions: ['*'] }
    ],
    properties: ['openFile']
  });
  return result.canceled ? null : result.filePaths[0];
});

ipcMain.handle('dialog:open-output-dir', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    title: 'Select Output Folder',
    properties: ['openDirectory', 'createDirectory']
  });
  return result.canceled ? null : result.filePaths[0];
});

/** Detect artifact type only (fast) */
ipcMain.handle('artifact:detect', async (_event, filePath: string) => {
  try {
    const { ArtifactDetector } = await import('../src/parsing/ArtifactDetector');
    return await ArtifactDetector.detect(filePath);
  } catch (e: unknown) {
    throw new Error(String(e));
  }
});

/** Full analysis pipeline */
ipcMain.handle('artifact:analyze', async (event, filePath: string, environment: Environment) => {
  try {
    currentGraph = await ParserOrchestrator.parse(filePath, environment, (progress) => {
      event.sender.send('progress', progress);
    });

    // Return lightweight summary for UI (not the full graph)
    return {
      success: true,
      summary: {
        artifactType: currentGraph.meta.artifactType,
        canvasApps: currentGraph.canvasApps.length,
        flows: currentGraph.flows.length,
        envVars: currentGraph.envVars.length,
        dataSources: currentGraph.dataSources.length,
        modelDrivenApps: currentGraph.modelDrivenApps.length,
        securityRoles: currentGraph.securityRoles.length,
        riskFlags: currentGraph.executiveSummary?.riskFlags ?? [],
        manualItemCount: currentGraph.manualChecklist.length,
        dataSourcesForSchema: currentGraph.dataSources
          .filter(ds => {
            // Only prompt for schema on sources where it adds documentation value:
            //   • Any SharePoint list (user-managed, always worth documenting)
            //   • Custom Dataverse tables only — identified by a publisher-prefix underscore
            //     in the logical table name (e.g. cr8a3_myentity, new_request).
            //     Standard/system tables (account, contact, task …) have no underscore.
            if (ds.schemaEnrichment || ds.columns?.length) return false;
            if (ds.type === 'sharepoint') return true;
            if (ds.type === 'dataverse') {
              const tbl = (ds.tableName ?? ds.logicalName ?? ds.displayName ?? '').toLowerCase();
              return tbl.includes('_');
            }
            return false;
          })
          .map(ds => ({ id: ds.id, displayName: ds.displayName, type: ds.type }))
      }
    };
  } catch (e: unknown) {
    return { success: false, error: String(e) };
  }
});

/** Enrich a data source with a schema file */
ipcMain.handle('schema:enrich', async (_event, dataSourceId: string, filePath: string) => {
  if (!currentGraph) return { success: false, error: 'No analysis in progress' };
  try {
    const ds = currentGraph.dataSources.find(d => d.id === dataSourceId);
    if (!ds) return { success: false, error: `Data source ${dataSourceId} not found` };
    const enrichment = SchemaEnrichmentParser.parse(filePath);
    ds.schemaEnrichment = enrichment;
    ds.columns = enrichment.columns;
    return { success: true, columnCount: enrichment.columns.length };
  } catch (e: unknown) {
    return { success: false, error: String(e) };
  }
});

/** Generate all outputs */
ipcMain.handle('generate:all', async (event, outputDir: string) => {
  if (!currentGraph) return { success: false, error: 'No analysis loaded' };

  try {
    fs.mkdirSync(outputDir, { recursive: true });
    const base = currentGraph.solution?.uniqueName ?? currentGraph.meta.fileName.replace(/\.[^.]+$/, '');
    const results: Record<string, string> = {};

    event.sender.send('progress', { stage: 'generate', pct: 10, message: 'Generating HTML report...' });
    const htmlPath = path.join(outputDir, `${base}_report.html`);
    fs.writeFileSync(htmlPath, HtmlGenerator.generate(currentGraph));
    results.html = htmlPath;

    event.sender.send('progress', { stage: 'generate', pct: 40, message: 'Generating Markdown...' });
    const mdPath = path.join(outputDir, `${base}_report.md`);
    fs.writeFileSync(mdPath, MarkdownGenerator.generate(currentGraph));
    results.markdown = mdPath;

    event.sender.send('progress', { stage: 'generate', pct: 60, message: 'Generating DOCX...' });
    const docxPath = path.join(outputDir, `${base}_spec.docx`);
    await DocxGenerator.generate(currentGraph, docxPath);
    results.docx = docxPath;

    event.sender.send('progress', { stage: 'generate', pct: 100, message: 'Generation complete' });

    return {
      success: true,
      outputs: results,
      errors: currentGraph.parseErrors
    };
  } catch (e: unknown) {
    return { success: false, error: String(e) };
  }
});

/** Open a file in the OS default app */
ipcMain.handle('shell:open', async (_event, filePath: string) => {
  await shell.openPath(filePath);
});

/** Reveal file in Finder/Explorer */
ipcMain.handle('shell:show-item', async (_event, filePath: string) => {
  shell.showItemInFolder(filePath);
});

/** Get default output directory */
ipcMain.handle('app:get-default-output-dir', () => {
  return path.join(os.homedir(), 'Documents', 'PPDocTool');
});
