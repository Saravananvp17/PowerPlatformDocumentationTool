import { contextBridge, ipcRenderer } from 'electron';
import type { Environment, ProgressEvent } from '../src/model/types';

// Typed API exposed to the renderer via contextBridge
contextBridge.exposeInMainWorld('ppdt', {
  // File dialogs
  openArtifact: (): Promise<string | null> =>
    ipcRenderer.invoke('dialog:open-artifact'),
  openSchema: (dataSourceName: string): Promise<string | null> =>
    ipcRenderer.invoke('dialog:open-schema', dataSourceName),
  openOutputDir: (): Promise<string | null> =>
    ipcRenderer.invoke('dialog:open-output-dir'),

  // Analysis
  detectArtifact: (filePath: string) =>
    ipcRenderer.invoke('artifact:detect', filePath),
  analyze: (filePath: string, environment: Environment) =>
    ipcRenderer.invoke('artifact:analyze', filePath, environment),
  enrichSchema: (dataSourceId: string, filePath: string) =>
    ipcRenderer.invoke('schema:enrich', dataSourceId, filePath),

  // Generation
  generateAll: (outputDir: string) =>
    ipcRenderer.invoke('generate:all', outputDir),

  // Shell
  openFile: (filePath: string) =>
    ipcRenderer.invoke('shell:open', filePath),
  showInFolder: (filePath: string) =>
    ipcRenderer.invoke('shell:show-item', filePath),

  // Defaults
  getDefaultOutputDir: (): Promise<string> =>
    ipcRenderer.invoke('app:get-default-output-dir'),

  // Progress stream
  onProgress: (cb: (evt: ProgressEvent) => void) => {
    const handler = (_: unknown, evt: ProgressEvent) => cb(evt);
    ipcRenderer.on('progress', handler);
    return () => ipcRenderer.removeListener('progress', handler);
  }
});
