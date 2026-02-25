import type { NormalizedAnalysisGraph } from '../model/types';

export class MarkdownGenerator {
  static generate(graph: NormalizedAnalysisGraph): string {
    const lines: string[] = [];
    const title = graph.solution?.displayName ?? graph.meta.fileName;
    lines.push(`# ${title} — Documentation Report`);
    lines.push('');
    lines.push(`> **Environment:** ${graph.meta.environment}  `);
    lines.push(`> **Analysed:** ${graph.meta.analysedAt}  `);
    lines.push(`> **Tool Version:** ${graph.meta.toolVersion}`);
    lines.push('');

    // Executive Summary
    const es = graph.executiveSummary;
    if (es) {
      lines.push('## Executive Summary');
      lines.push('');
      lines.push('### Component Inventory');
      for (const [k, v] of Object.entries(es.componentCounts)) {
        if (v > 0) lines.push(`- **${k}:** ${v}`);
      }
      lines.push('');
      if (es.riskFlags.length) {
        lines.push('### Risk Flags');
        for (const r of es.riskFlags) {
          lines.push(`- **[${r.severity.toUpperCase()}]** ${r.message}${r.location ? ` — \`${r.location}\`` : ''}`);
        }
        lines.push('');
      }
    }

    // Solution
    if (graph.solution) {
      const s = graph.solution;
      lines.push('## Solution Information');
      lines.push('');
      lines.push(`| Field | Value |`);
      lines.push(`|-------|-------|`);
      lines.push(`| Unique Name | \`${s.uniqueName}\` |`);
      lines.push(`| Version | ${s.version} |`);
      lines.push(`| Managed | ${s.isManaged ? 'Yes' : 'No'} |`);
      lines.push(`| Publisher | ${s.publisher.displayName} (\`${s.publisher.uniqueName}\`) |`);
      lines.push('');
    }

    // Env Vars
    if (graph.envVars.length) {
      lines.push('## Environment Variables');
      lines.push('');
      lines.push('| Display Name | Schema Name | Type | Default | Current Value |');
      lines.push('|---|---|---|---|---|');
      for (const ev of graph.envVars) {
        lines.push(`| ${ev.displayName} | \`${ev.schemaName}\` | ${ev.type} | ${ev.defaultValue ?? '—'} | ${ev.currentValue ?? '—'} |`);
      }
      lines.push('');
    }

    // Flows
    if (graph.flows.length) {
      lines.push('## Flows');
      lines.push('');
      for (const flow of graph.flows) {
        lines.push(`### ${flow.displayName}`);
        lines.push('');
        const trig = flow.trigger.recurrence
          ? `Recurrence — every ${flow.trigger.recurrence.interval} ${flow.trigger.recurrence.frequency}`
          : flow.triggerType;
        lines.push(`- **Trigger:** ${trig}`);
        lines.push(`- **State:** ${flow.state}`);
        lines.push(`- **Actions:** ${flow.actions.length}`);
        lines.push(`- **Error Handling:** ${flow.errorHandling.hasExplicitHandling ? 'Yes' : 'No'}`);
        lines.push(`- **Source:** \`${flow.source.archivePath}\``);
        lines.push('');
        lines.push('```mermaid');
        lines.push(flow.mermaidDiagram);
        lines.push('```');
        lines.push('');
        lines.push('| Action | Type | Connector | Run After |');
        lines.push('|---|---|---|---|');
        for (const a of flow.actions) {
          lines.push(`| \`${a.id}\` | ${a.type} | ${a.connector ?? '—'} | ${a.runAfter.join(', ') || '—'} |`);
        }
        lines.push('');
      }
    }

    // Canvas Apps
    if (graph.canvasApps.length) {
      lines.push('## Canvas Apps');
      lines.push('');
      for (const app of graph.canvasApps) {
        lines.push(`### ${app.displayName}`);
        lines.push('');
        if (app.appOnStart.redacted) {
          lines.push('**App.OnStart:**');
          lines.push('```');
          lines.push(app.appOnStart.redacted);
          lines.push('```');
          lines.push('');
        }
        lines.push(`**Screens (${app.screens.length}):**`);
        for (const s of app.screens) {
          lines.push(`- **${s.name}**${s.onVisible.redacted ? ` — OnVisible: \`${s.onVisible.redacted.substring(0, 60)}…\`` : ''}`);
        }
        lines.push('');
        if (app.mermaidNavGraph) {
          lines.push(`**Navigation Graph (${app.navConfidence} confidence):**`);
          lines.push('```mermaid');
          lines.push(app.mermaidNavGraph);
          lines.push('```');
          lines.push('');
        }
        if (app.variables.length) {
          lines.push(`**Variables (${app.variables.length}):**`);
          lines.push('| Name | Kind |');
          lines.push('|---|---|');
          for (const v of app.variables) lines.push(`| \`${v.name}\` | ${v.kind} |`);
          lines.push('');
        }
      }
    }

    // Manual Checklist
    if (graph.manualChecklist.length) {
      lines.push('## Manual Completion Checklist');
      lines.push('');
      for (let i = 0; i < graph.manualChecklist.length; i++) {
        const c = graph.manualChecklist[i];
        lines.push(`${i + 1}. **[${c.priority.toUpperCase()}]** [${c.section}] ${c.item}`);
      }
      lines.push('');
    }

    // Parse errors
    if (graph.parseErrors.length) {
      lines.push('## Parse Errors');
      lines.push('');
      for (const e of graph.parseErrors) {
        lines.push(`- \`${e.archivePath}\`: ${e.error}`);
      }
      lines.push('');
    }

    lines.push('---');
    lines.push(`*Generated by PP Doc Tool v${graph.meta.toolVersion}*`);
    return lines.join('\n');
  }
}
