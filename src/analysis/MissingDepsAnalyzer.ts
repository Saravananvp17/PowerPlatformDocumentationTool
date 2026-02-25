import type { NormalizedAnalysisGraph, MissingDependency, ChecklistItem } from '../model/types';

export class MissingDepsAnalyzer {
  static analyze(graph: NormalizedAnalysisGraph): void {
    const missing: MissingDependency[] = [];
    const checklist: ChecklistItem[] = [];

    // 1. Role IDs referenced by model-driven apps but not defined in security roles
    const definedRoleIds = new Set(graph.securityRoles.map(r => r.id));
    for (const mda of graph.modelDrivenApps) {
      for (const rid of mda.assignedRoleIds) {
        if (!definedRoleIds.has(rid)) {
          missing.push({
            type: 'SecurityRole',
            identifier: rid,
            displayName: `Role ID: ${rid}`,
            referencedBy: mda.source,
            impact: `Model-driven app '${mda.displayName}' references this role — access control cannot be verified without role definition`,
            checklistItem: `Retrieve security role definition for ID '${rid}' from target environment`
          });
          checklist.push({
            section: 'Roles & Security',
            item: `Retrieve and document security role for ID '${rid}' (referenced by '${mda.displayName}')`,
            priority: 'high'
          });
        }
      }
    }

    // 2. Env var values not in export
    for (const ev of graph.envVars) {
      if (!ev.currentValue || ev.currentValue === 'not included in export') {
        checklist.push({
          section: 'System Specification',
          item: `Set current value for environment variable '${ev.displayName}' (${ev.schemaName}) in target environment`,
          priority: 'high'
        });
      }
    }

    // 3. Data sources with no schema enrichment
    for (const ds of graph.dataSources) {
      if (!ds.schemaEnrichment && !ds.columns?.length) {
        checklist.push({
          section: 'Source Data',
          item: `Provide schema documentation for ${ds.type} data source '${ds.displayName}'`,
          priority: 'medium'
        });
        missing.push({
          type: 'Schema',
          identifier: ds.id,
          displayName: ds.displayName,
          referencedBy: ds.source,
          impact: 'Column-level documentation not available — maintenance staff cannot understand data structure',
          checklistItem: `Export schema for '${ds.displayName}' and re-run analysis`
        });
      }
    }

    // 4. Env-specific hard-coded URLs (recommend env vars)
    for (const url of graph.urls) {
      if (url.isEnvSpecific && url.category !== 'sharepoint') {
        checklist.push({
          section: 'System Specification',
          item: `Consider replacing hard-coded ${url.category} URL with an environment variable: ${url.url.substring(0, 80)}`,
          priority: url.category === 'local-dev' ? 'high' : 'medium'
        });
      }
    }

    // 5. Canvas app AppChecker errors
    for (const app of graph.canvasApps) {
      const errors = app.appCheckerFindings.filter(f => f.level === 'error');
      if (errors.length > 0) {
        checklist.push({
          section: 'Apps',
          item: `Resolve ${errors.length} AppChecker error(s) in canvas app '${app.displayName}'`,
          priority: 'high'
        });
      }
    }

    // 6. Manual content placeholders
    checklist.push({
      section: 'Introduction',
      item: 'Complete Introduction section: document business purpose and stakeholders',
      priority: 'medium'
    });
    checklist.push({
      section: 'Functional Overview',
      item: 'Complete Functional Overview: describe the business process this solution automates',
      priority: 'medium'
    });

    graph.missingDeps = missing;
    graph.manualChecklist = [...graph.manualChecklist, ...checklist];
  }
}
