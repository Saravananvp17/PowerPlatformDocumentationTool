import path from 'path';
import type {
  NormalizedAnalysisGraph, ArtifactMeta, Environment,
  ProgressCallback, WhereUsedIndex
} from '../model/types';
import { ArtifactDetector } from './ArtifactDetector';
import { SolutionZipParser } from './SolutionZipParser';
import { FlowDefinitionParser } from './FlowDefinitionParser';
import { PowerAppsZipParser } from './PowerAppsZipParser';
import { MsappParser } from './MsappParser';
import { SecretRedactor } from '../analysis/SecretRedactor';
import { WhereUsedIndexBuilder } from '../analysis/WhereUsedIndexBuilder';
import { MissingDepsAnalyzer } from '../analysis/MissingDepsAnalyzer';
import { ExecutiveSummaryBuilder } from '../analysis/ExecutiveSummaryBuilder';
import fs from 'fs';

export class ParserOrchestrator {
  static async parse(
    filePath: string,
    environment: Environment,
    onProgress?: ProgressCallback
  ): Promise<NormalizedAnalysisGraph> {
    onProgress?.({ stage: 'detect', pct: 0, message: 'Detecting artifact type...' });

    const detection = await ArtifactDetector.detect(filePath);
    const stats = fs.statSync(filePath);

    const meta: ArtifactMeta = {
      fileName: path.basename(filePath),
      artifactType: detection.type,
      environment,
      analysedAt: new Date().toISOString(),
      toolVersion: '1.0.0',
      fileSizeBytes: stats.size
    };

    onProgress?.({ stage: 'detect', pct: 5, message: `Detected: ${detection.type} — ${detection.details}` });

    let partial: Partial<NormalizedAnalysisGraph> = {};

    switch (detection.type) {
      case 'solution-zip':
        partial = await SolutionZipParser.parse(filePath, meta, onProgress);
        break;

      case 'flow-zip': {
        const JSZip = (await import('jszip')).default;
        const buf = fs.readFileSync(filePath);
        const zip = await JSZip.loadAsync(buf);
        const flowFiles = Object.keys(zip.files).filter(e =>
          e.endsWith('.json') && !zip.files[e].dir
        );
        partial = {
          flows: [], canvasApps: [], modelDrivenApps: [], envVars: [],
          connectors: [], connectionRefs: [], dataSources: [], securityRoles: [],
          dataverseFormulas: [], webResources: [], urls: [],
          missingDeps: [], manualChecklist: [], parseErrors: []
        };
        for (let i = 0; i < flowFiles.length; i++) {
          onProgress?.({ stage: 'flows', pct: Math.round((i / flowFiles.length) * 90), message: `Parsing flow ${i + 1}/${flowFiles.length}` });
          try {
            const raw = JSON.parse(await zip.files[flowFiles[i]].async('string'));
            const flow = FlowDefinitionParser.parse(raw, filePath + '/' + flowFiles[i]);
            partial.flows!.push(flow);
          } catch (e: unknown) {
            partial.parseErrors!.push({ archivePath: flowFiles[i], error: String(e), partial: false });
          }
        }
        break;
      }

      case 'powerapp-zip':
        partial = await PowerAppsZipParser.parse(
          filePath, detection.embeddedMsappPath!, meta, onProgress
        );
        break;

      case 'msapp': {
        const buf = fs.readFileSync(filePath);
        const displayName = path.basename(filePath, '.msapp');
        const app = await MsappParser.parse(buf, filePath, displayName, onProgress);
        partial = {
          canvasApps: [app], flows: [], modelDrivenApps: [], copilotAgents: [], aiBuilderModels: [], managedPlans: [], envVars: [],
          connectors: app.connectors, connectionRefs: [], dataSources: app.dataSources,
          securityRoles: [], dataverseFormulas: [], webResources: [], urls: [],
          missingDeps: [], manualChecklist: [], parseErrors: []
        };
        break;
      }
    }

    // Assemble full graph
    const graph: NormalizedAnalysisGraph = {
      meta,
      solution: partial.solution,
      canvasApps: partial.canvasApps ?? [],
      modelDrivenApps: partial.modelDrivenApps ?? [],
      flows: partial.flows ?? [],
      copilotAgents: partial.copilotAgents ?? [],
      aiBuilderModels: partial.aiBuilderModels ?? [],
      managedPlans: partial.managedPlans ?? [],
      envVars: partial.envVars ?? [],
      connectors: partial.connectors ?? [],
      connectionRefs: partial.connectionRefs ?? [],
      dataSources: partial.dataSources ?? [],
      securityRoles: partial.securityRoles ?? [],
      dataverseFormulas: partial.dataverseFormulas ?? [],
      webResources: partial.webResources ?? [],
      urls: partial.urls ?? [],
      missingDeps: partial.missingDeps ?? [],
      parseErrors: partial.parseErrors ?? [],
      manualChecklist: partial.manualChecklist ?? [],
      whereUsed: {
        connectorToFlows: new Map(), connectorToApps: new Map(),
        envVarToFlows: new Map(), envVarToApps: new Map(),
        dataSourceToFlows: new Map(), dataSourceToApps: new Map(),
        variableToScreens: new Map()
      }
    };

    onProgress?.({ stage: 'analysis', pct: 92, message: 'Building where-used index...' });
    WhereUsedIndexBuilder.build(graph);

    onProgress?.({ stage: 'analysis', pct: 94, message: 'Redacting secrets...' });
    SecretRedactor.redact(graph);

    onProgress?.({ stage: 'analysis', pct: 96, message: 'Analysing missing dependencies...' });
    MissingDepsAnalyzer.analyze(graph);

    onProgress?.({ stage: 'analysis', pct: 98, message: 'Building executive summary...' });
    ExecutiveSummaryBuilder.build(graph);

    onProgress?.({ stage: 'analysis', pct: 100, message: 'Analysis complete' });
    return graph;
  }
}
