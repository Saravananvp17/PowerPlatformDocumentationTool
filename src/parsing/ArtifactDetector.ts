import JSZip from 'jszip';
import fs from 'fs';
import path from 'path';
import type { ArtifactType } from '../model/types';

export interface DetectionResult {
  type: ArtifactType;
  confidence: 'high' | 'low';
  details: string;
  embeddedMsappPath?: string;   // for powerapp-zip
}

export class ArtifactDetector {
  static async detect(filePath: string): Promise<DetectionResult> {
    const ext = path.extname(filePath).toLowerCase();

    if (ext === '.msapp') {
      return { type: 'msapp', confidence: 'high', details: 'Standalone .msapp file' };
    }

    if (ext !== '.zip') {
      throw new Error(`Unsupported file type: ${ext}. Expected .zip or .msapp`);
    }

    const buf = fs.readFileSync(filePath);
    const zip = await JSZip.loadAsync(buf);
    const entries = Object.keys(zip.files);

    // Solution ZIP: must have solution.xml at root
    if (entries.includes('solution.xml')) {
      return {
        type: 'solution-zip',
        confidence: 'high',
        details: `Solution ZIP — solution.xml found. ${entries.filter(e => !zip.files[e].dir).length} total entries.`
      };
    }

    // Power Apps export ZIP: manifest.json + Microsoft.PowerApps/apps/<guid>/<guid>.msapp
    const embeddedMsapp = entries.find(e =>
      /Microsoft\.PowerApps\/apps\/[^/]+\/[^/]+\.msapp$/i.test(e)
    );
    if (entries.includes('manifest.json') && embeddedMsapp) {
      return {
        type: 'powerapp-zip',
        confidence: 'high',
        details: `Power Apps export ZIP — manifest.json + embedded .msapp at ${embeddedMsapp}`,
        embeddedMsappPath: embeddedMsapp
      };
    }

    // Flow export ZIP: workflow JSON(s), no solution.xml
    const flowJsons = entries.filter(e =>
      !zip.files[e].dir &&
      (e.toLowerCase().endsWith('.json')) &&
      !e.toLowerCase().includes('manifest')
    );
    if (flowJsons.length > 0) {
      // Try to confirm by checking first JSON for workflow schema markers
      try {
        const firstJson = JSON.parse(await zip.files[flowJsons[0]].async('string'));
        if (firstJson.properties || firstJson.triggers || firstJson.actions || firstJson.definition) {
          return {
            type: 'flow-zip',
            confidence: 'high',
            details: `Flow export ZIP — ${flowJsons.length} JSON workflow file(s)`
          };
        }
      } catch { /* fall through */ }
    }

    // Last resort: treat any ZIP with JSON as flow-zip
    if (flowJsons.length > 0) {
      return { type: 'flow-zip', confidence: 'low', details: 'Assumed flow export ZIP based on JSON contents' };
    }

    throw new Error('Could not determine artifact type. Expected solution.xml (solution), manifest.json + .msapp (Power Apps export), or workflow JSON(s) (flow export).');
  }
}
