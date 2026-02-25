import type { NormalizedAnalysisGraph, ExecutiveSummary, RiskFlag } from '../model/types';

export class ExecutiveSummaryBuilder {
  static build(graph: NormalizedAnalysisGraph): ExecutiveSummary {
    const riskFlags: RiskFlag[] = [];

    // Hard-coded localhost URLs
    const localUrls = graph.urls.filter(u => u.category === 'local-dev');
    if (localUrls.length > 0) {
      riskFlags.push({
        severity: 'high',
        message: `${localUrls.length} localhost URL(s) detected — likely hardcoded dev endpoints`,
        location: localUrls[0].url
      });
    }

    // Env vars with no value
    const missingEvVals = graph.envVars.filter(
      ev => !ev.currentValue || ev.currentValue === 'not included in export'
    );
    if (missingEvVals.length > 0) {
      riskFlags.push({
        severity: 'high',
        message: `${missingEvVals.length} environment variable(s) have no current value in the export`,
      });
    }

    // Missing role definitions
    const missingRoles = graph.missingDeps.filter(d => d.type === 'SecurityRole');
    if (missingRoles.length > 0) {
      riskFlags.push({
        severity: 'high',
        message: `${missingRoles.length} security role(s) referenced by apps but not defined in export`
      });
    }

    // AppChecker errors
    const appErrors = graph.canvasApps.reduce((sum, a) =>
      sum + a.appCheckerFindings.filter(f => f.level === 'error').length, 0);
    if (appErrors > 0) {
      riskFlags.push({ severity: 'high', message: `${appErrors} AppChecker error(s) in canvas app(s)` });
    }

    // Env-specific URLs without env vars
    const envUrls = graph.urls.filter(u => u.isEnvSpecific && u.category !== 'sharepoint');
    if (envUrls.length > 5) {
      riskFlags.push({
        severity: 'medium',
        message: `${envUrls.length} environment-specific URLs detected — consider replacing with environment variables`
      });
    }

    // Missing schemas
    const missingSchemas = graph.dataSources.filter(ds => !ds.schemaEnrichment && !ds.columns?.length);
    if (missingSchemas.length > 0) {
      riskFlags.push({
        severity: 'low',
        message: `${missingSchemas.length} data source(s) have no schema documentation`
      });
    }

    // Key dependencies
    const keyDependencies: string[] = [
      ...new Set([
        ...graph.connectors.map(c => c.displayName || c.connectorId),
        ...graph.dataSources.map(ds => `${ds.type}: ${ds.displayName}`)
      ])
    ].slice(0, 10);

    const componentCounts: Record<string, number> = {
      'Canvas Apps': graph.canvasApps.length,
      'Model-Driven Apps': graph.modelDrivenApps.length,
      'Flows': graph.flows.length,
      'Copilot Studio Agents': graph.copilotAgents.length,
      'AI Builder Models': graph.aiBuilderModels.length,
      'Managed Plans': graph.managedPlans.length,
      'Environment Variables': graph.envVars.length,
      'Data Sources': graph.dataSources.length,
      'Connectors': graph.connectors.length,
      'Security Roles': graph.securityRoles.length,
      'Web Resources': graph.webResources.length
    };

    const summary: ExecutiveSummary = {
      artifactName: graph.meta.fileName,
      artifactVersion: graph.solution?.version,
      environment: graph.meta.environment,
      componentCounts,
      keyDependencies,
      riskFlags,
      manualItemCount: graph.manualChecklist.length
    };

    graph.executiveSummary = summary;
    return summary;
  }
}
