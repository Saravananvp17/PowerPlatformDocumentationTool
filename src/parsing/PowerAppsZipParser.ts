import JSZip from 'jszip';
import fs from 'fs';
import type { NormalizedAnalysisGraph, ArtifactMeta, ParseError, ProgressCallback } from '../model/types';
import { MsappParser } from './MsappParser';

export class PowerAppsZipParser {
  static async parse(
    filePath: string,
    embeddedMsappPath: string,
    _meta: ArtifactMeta,
    onProgress?: ProgressCallback
  ): Promise<Partial<NormalizedAnalysisGraph>> {
    const buf = fs.readFileSync(filePath);
    const zip = await JSZip.loadAsync(buf);
    const errors: ParseError[] = [];
    const result: Partial<NormalizedAnalysisGraph> = {
      canvasApps: [], flows: [], modelDrivenApps: [], envVars: [],
      connectors: [], connectionRefs: [], dataSources: [], securityRoles: [],
      dataverseFormulas: [], webResources: [], urls: [],
      missingDeps: [], manualChecklist: [], parseErrors: []
    };

    // Read manifest.json
    let displayName = 'Unknown App';
    try {
      const manifest = JSON.parse(await zip.file('manifest.json')!.async('string'));
      displayName = manifest.displayName ?? manifest.name ?? displayName;
    } catch { /* non-fatal */ }

    onProgress?.({ stage: 'powerapp', pct: 10, message: `Extracting embedded .msapp: ${embeddedMsappPath}` });

    try {
      const msappFile = zip.file(embeddedMsappPath);
      if (!msappFile) throw new Error(`Embedded .msapp not found at ${embeddedMsappPath}`);

      const msappBuf = Buffer.from(await msappFile.async('arraybuffer'));
      const app = await MsappParser.parse(msappBuf, embeddedMsappPath, displayName, onProgress);
      result.canvasApps!.push(app);

      for (const ds of app.dataSources) {
        if (!result.dataSources!.find(d => d.id === ds.id)) result.dataSources!.push(ds);
      }
      for (const conn of app.connectors) {
        conn.usedInApps.push(app.id);
        const existing = result.connectors!.find(c => c.connectorId === conn.connectorId);
        if (existing) {
          if (!existing.usedInApps.includes(app.id)) existing.usedInApps.push(app.id);
        } else {
          result.connectors!.push(conn);
        }
      }
    } catch (e: unknown) {
      errors.push({ archivePath: embeddedMsappPath, error: String(e), partial: false });
    }

    result.parseErrors = errors;
    return result;
  }
}
