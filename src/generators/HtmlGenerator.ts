import type {
  NormalizedAnalysisGraph, Flow, FlowAction, CanvasApp, Screen,
  ModelDrivenApp, EnvVar, DataSource, SecurityRole, RolePrivilege, ConnectorRef,
  WebResource, ChecklistItem, ParseError, RiskFlag, AppCheckerFinding,
  ButtonAction, ActionStep, CopilotAgent, CopilotTopic, AiBuilderModel, ManagedPlan
} from '../model/types';

// ─────────────────────────────────────────────────────────────────────────────

export class HtmlGenerator {
  static generate(graph: NormalizedAnalysisGraph): string {
    const sections = HtmlGenerator.buildSections(graph);
    const nav = HtmlGenerator.buildNav(sections);
    const content = sections.map(s => s.html).join('\n');
    return HtmlGenerator.wrapPage(graph, nav, content);
  }

  // ── Navigation ─────────────────────────────────────────────────────────────

  private static buildNav(sections: Section[]): string {
    const items = sections.map(s =>
      `<li><a href="#${s.id}" class="nav-link" data-section="${s.id}">${s.title}${s.count !== undefined ? ` <span class="nav-count">${s.count}</span>` : ''}</a>
       ${s.children ? `<ul>${s.children.map(c =>
         `<li><a href="#${c.id}" class="nav-link nav-child">${c.title}${c.count !== undefined ? ` <span class="nav-count">${c.count}</span>` : ''}</a></li>`
       ).join('')}</ul>` : ''}
       </li>`
    ).join('\n');
    return `<ul class="nav-list">${items}</ul>`;
  }

  // ── Sections builder ────────────────────────────────────────────────────────

  private static buildSections(graph: NormalizedAnalysisGraph): Section[] {
    const sections: Section[] = [];
    sections.push(HtmlGenerator.overviewSection(graph));
    if (graph.canvasApps.length) sections.push(HtmlGenerator.canvasAppsSection(graph));
    if (graph.modelDrivenApps.length) sections.push(HtmlGenerator.mdaSection(graph));
    if (graph.flows.length) sections.push(HtmlGenerator.flowsSection(graph));
    if (graph.copilotAgents.length) sections.push(HtmlGenerator.copilotAgentsSection(graph));
    if (graph.aiBuilderModels.length) sections.push(HtmlGenerator.aiBuilderSection(graph));
    if (graph.managedPlans.length) sections.push(HtmlGenerator.managedPlanSection(graph));
    if (graph.envVars.length) sections.push(HtmlGenerator.envVarsSection(graph));
    if (graph.connectors.length) sections.push(HtmlGenerator.connectorsSection(graph));
    if (graph.dataSources.length) sections.push(HtmlGenerator.dataSourcesSection(graph));
    if (graph.securityRoles.length) sections.push(HtmlGenerator.rolesSection(graph));
    if (graph.dataverseFormulas.length) sections.push(HtmlGenerator.formulasSection(graph));
    if (graph.webResources.length) sections.push(HtmlGenerator.webResourcesSection(graph));
    sections.push(HtmlGenerator.missingDepsSection(graph));
    sections.push(HtmlGenerator.parseNotesSection(graph));
    return sections;
  }

  // ── Overview ────────────────────────────────────────────────────────────────

  private static overviewSection(graph: NormalizedAnalysisGraph): Section {
    const es = graph.executiveSummary!;
    const riskHtml = es.riskFlags.map(r => `
      <div class="risk-flag risk-${r.severity}">
        <span class="risk-badge">${r.severity.toUpperCase()}</span> ${esc(r.message)}
        ${r.location ? `<code>${esc(r.location)}</code>` : ''}
      </div>`).join('');

    const countRows = Object.entries(es.componentCounts)
      .filter(([, v]) => v > 0)
      .map(([k, v]) => `<tr><td>${esc(k)}</td><td class="num">${v}</td></tr>`)
      .join('');

    const checklistHtml = graph.manualChecklist.length
      ? `<details><summary>Manual Completion Checklist (${graph.manualChecklist.length} items)</summary>
          <table class="data-table">
            <thead><tr><th>Section</th><th>Item</th><th>Priority</th></tr></thead>
            <tbody>${graph.manualChecklist.map(c =>
              `<tr><td>${esc(c.section)}</td><td>${esc(c.item)}</td><td class="pri pri-${c.priority}">${c.priority}</td></tr>`
            ).join('')}</tbody>
          </table></details>` : '';

    const sol = graph.solution;
    const artifactInfo = sol ? `
      <table class="data-table">
        <tr><td>Solution Name</td><td>${esc(sol.displayName)}</td></tr>
        <tr><td>Unique Name</td><td><code>${esc(sol.uniqueName)}</code></td></tr>
        <tr><td>Version</td><td>${esc(sol.version)}</td></tr>
        <tr><td>Managed</td><td>${sol.isManaged ? '✅ Managed' : '⚠️ Unmanaged'}</td></tr>
        <tr><td>Publisher</td><td>${esc(sol.publisher.displayName)} (<code>${esc(sol.publisher.uniqueName)}</code>)</td></tr>
        <tr><td>Environment</td><td><span class="env-badge env-${graph.meta.environment.toLowerCase()}">${graph.meta.environment}</span></td></tr>
        <tr><td>Analysed</td><td>${graph.meta.analysedAt}</td></tr>
        <tr><td>File</td><td>${esc(graph.meta.fileName)} (${(graph.meta.fileSizeBytes / 1024 / 1024).toFixed(2)} MB)</td></tr>
      </table>` : `
      <table class="data-table">
        <tr><td>File</td><td>${esc(graph.meta.fileName)}</td></tr>
        <tr><td>Artifact Type</td><td>${graph.meta.artifactType}</td></tr>
        <tr><td>Environment</td><td><span class="env-badge env-${graph.meta.environment.toLowerCase()}">${graph.meta.environment}</span></td></tr>
        <tr><td>Analysed</td><td>${graph.meta.analysedAt}</td></tr>
      </table>`;

    return {
      id: 'overview', title: 'Overview',
      html: `
      <section id="overview" class="section">
        <h1>Overview</h1>
        <h2>Artifact Information</h2>${artifactInfo}
        <h2>Component Inventory</h2>
        <table class="data-table"><thead><tr><th>Component</th><th>Count</th></tr></thead>
        <tbody>${countRows}</tbody></table>
        <h2>Risk Flags</h2>${riskHtml || '<p class="muted">No risk flags detected.</p>'}
        <h2>Key Dependencies</h2>
        <ul>${es.keyDependencies.map(d => `<li>${esc(d)}</li>`).join('')}</ul>
        ${checklistHtml}
      </section>`
    };
  }

  // ── Canvas Apps ─────────────────────────────────────────────────────────────

  private static canvasAppsSection(graph: NormalizedAnalysisGraph): Section {
    const children: SectionChild[] = graph.canvasApps.map(app => ({
      id: `app-${slugify(app.id)}`,
      title: app.displayName,
      count: app.screens.length
    }));

    const appsHtml = graph.canvasApps.map(app => {
      const screenDetails = app.screens.map(s => {
        const allFormulas = s.keyFormulas;
        const formulaRows = allFormulas.length
          ? allFormulas.map(kf => `
              <tr>
                <td><code>${esc(kf.controlName)}</code></td>
                <td>${esc(kf.property)}</td>
                <td>${kf.patterns.map(p => `<span class="tag">${esc(p)}</span>`).join(' ')}</td>
                <td><code class="formula">${esc(kf.formula.redacted.substring(0, 120))}${kf.formula.redacted.length > 120 ? '…' : ''}</code></td>
              </tr>`).join('')
          : `<tr><td colspan="4" class="muted">No key formulas on screen</td></tr>`;

        const controlChips = s.controls.length
          ? s.controls.map(c => `<span class="tag">${esc(c.name)} <em class="ctrl-type">${esc(c.type)}</em></span>`).join(' ')
          : '<span class="muted">No controls extracted</span>';

        const onVisibleHtml = s.onVisible.redacted
          ? `<pre class="formula-block">${fxHtml(s.onVisible.redacted)}</pre>`
          : '<p class="muted">—</p>';

        // ── Button / control action cards ──────────────────────────────────
        const buttonActionsHtml = (() => {
          const bas = s.buttonActions ?? [];
          if (!bas.length) return '';

          const actionIcon = (kind: ActionStep['kind']) =>
            kind === 'patch'       ? '💾'
            : kind === 'flow-run'  ? '⚡'
            : kind === 'navigate'  ? '🧭'
            : kind === 'submit-form' ? '📋'
            : kind === 'collect'   ? '📦'
            : kind === 'remove'    ? '🗑️'
            : '▶';

          const actionLabel = (kind: ActionStep['kind']) =>
            kind === 'patch'       ? 'Patch'
            : kind === 'flow-run'  ? 'Flow'
            : kind === 'navigate'  ? 'Navigate'
            : kind === 'submit-form' ? 'SubmitForm'
            : kind === 'collect'   ? 'Collect'
            : kind === 'remove'    ? 'Remove'
            : 'Action';

          const cards = bas.map((ba: ButtonAction) => {
            const chips = ba.actions.map((step: ActionStep) => {
              const payloadTip = step.payload
                ? ` title="${esc(step.payload)}"`
                : '';
              return `<span class="ba-chip ba-${ba.actions.indexOf(step) === 0 ? 'first' : ''} ba-kind-${step.kind}"${payloadTip}>` +
                `${actionIcon(step.kind)} <strong>${actionLabel(step.kind)}</strong> → <code>${esc(step.target)}</code>` +
                (step.payload ? ` <span class="ba-payload-hint" title="${esc(step.payload)}">(${esc(step.payload.substring(0, 60))}${step.payload.length > 60 ? '…' : ''})</span>` : '') +
                `</span>`;
            }).join('\n');

            return `
              <div class="ba-card">
                <div class="ba-header">
                  <code class="ba-ctrl-name">${esc(ba.controlName)}</code>
                  <span class="ba-ctrl-type">${esc(ba.controlType)}</span>
                  <span class="ba-prop">${esc(ba.property)}</span>
                </div>
                <div class="ba-chips">${chips}</div>
                <details class="ba-formula">
                  <summary>Show formula</summary>
                  <pre class="formula-block">${fxHtml(ba.formulaSnippet)}</pre>
                </details>
              </div>`;
          }).join('');

          return `
            <h5>Button / Control Actions (${bas.length})</h5>
            <div class="ba-list">${cards}</div>`;
        })();

        // ── Galleries ──────────────────────────────────────────────────────
        const galleries = s.galleries ?? [];
        const galleriesHtml = (() => {
          if (!galleries.length) return '';
          const cards = galleries.map(g => {
            const itemsText = g.itemsFormula?.redacted || g.itemsFormula?.raw || '—';
            const dsChips = g.inferredDataSources.length
              ? g.inferredDataSources.map(d => '<span class="gal-ds-chip">' + esc(d) + '</span>').join(' ')
              : '<span class="muted">—</span>';
            const fieldChips = g.fieldsUsed.length
              ? g.fieldsUsed.map(f => '<span class="tag">' + esc(f) + '</span>').join(' ')
              : '<span class="muted">none detected</span>';
            const typeLabel = (g.galleryType ?? 'Gallery').replace(/@[\d.]+$/, '');
            return '<div class="gal-card">'
              + '<div class="gal-header">'
              + '<span class="gal-icon">▦</span>'
              + '<code class="gal-name">' + esc(g.controlName) + '</code>'
              + '<span class="gal-type">' + esc(typeLabel) + '</span>'
              + '</div>'
              + '<table class="data-table gal-table">'
              + '<tr><td>Items</td><td><code class="formula">' + esc(itemsText) + '</code></td></tr>'
              + '<tr><td>Data Source(s)</td><td>' + dsChips + '</td></tr>'
              + '<tr><td>Fields Used</td><td>' + fieldChips + '</td></tr>'
              + '</table></div>';
          }).join('');
          return '<h5>Galleries (' + galleries.length + ')</h5><div class="gal-list">' + cards + '</div>';
        })();
        const galBadge = galleries.length
          ? ' &nbsp;·&nbsp; ' + galleries.length + (galleries.length !== 1 ? ' galleries' : ' gallery')
          : '';

        return `
          <details class="screen-detail">
            <summary class="screen-summary">
              <span class="screen-name">${esc(s.name)}</span>
              <span class="screen-meta">${s.controls.length} controls &nbsp;·&nbsp; ${allFormulas.length} key formula${allFormulas.length !== 1 ? 's' : ''}${(s.buttonActions ?? []).length ? ' &nbsp;·&nbsp; ' + s.buttonActions.length + ' action' + (s.buttonActions.length !== 1 ? 's' : '') : ''}${galBadge}</span>
            </summary>
            <div class="screen-body">
              <h5>OnVisible</h5>
              ${onVisibleHtml}
              ${galleriesHtml}
              ${buttonActionsHtml}
              <details class="ctrl-details">
                <summary class="ctrl-summary">Controls (${s.controls.length})</summary>
                <div class="control-chips">${controlChips}</div>
              </details>
              <h5>Key Formulas (${allFormulas.length})</h5>
              <table class="data-table">
                <thead><tr><th>Control</th><th>Property</th><th>Patterns</th><th>Formula</th></tr></thead>
                <tbody>${formulaRows}</tbody>
              </table>
            </div>
          </details>`;
      }).join('');

      // ── Variables (globals + context vars only) ─────────────────────────
      const nonCollectionVars = app.variables.filter(v => v.kind !== 'collection');
      const collections       = app.variables.filter(v => v.kind === 'collection');

      const varRows = nonCollectionVars.map(v => `
        <tr>
          <td><code>${esc(v.name)}</code></td>
          <td>${v.kind}</td>
          <td class="muted" style="font-size:11px">${v.setAt.map(s => esc(s.yamlPath ?? s.archivePath)).join(', ')}</td>
        </tr>`).join('');

      // ── Collections ─────────────────────────────────────────────────────
      const collectionsHtml = (() => {
        if (!collections.length) return '';

        // Supplement with any additional collection names detected by ButtonActionAnalyzer
        // that VariableExtractor may have missed (e.g. used but not explicitly Set)
        const extraNames = new Set<string>();
        for (const screen of app.screens) {
          for (const ba of screen.buttonActions ?? []) {
            for (const step of ba.actions) {
              if (step.kind === 'collect' && !collections.find(c => c.name === step.target)) {
                extraNames.add(step.target);
              }
            }
          }
        }

        const colRows = collections.map(c => {
          const setScreens = [...new Set(c.setAt.map(r => {
            const p = r.yamlPath ?? r.archivePath;
            // Extract just the screen name (first segment before dot)
            return p.split('.')[0].split('/').pop() ?? p;
          }))];
          const usedCount = c.usedAt.length;
          return `<tr>
            <td><code>${esc(c.name)}</code></td>
            <td>${setScreens.map(s => `<span class="tag">${esc(s)}</span>`).join(' ')}</td>
            <td class="num">${usedCount > 0 ? usedCount : '—'}</td>
          </tr>`;
        }).join('');

        const extraRows = [...extraNames].map(name =>
          `<tr>
            <td><code>${esc(name)}</code></td>
            <td class="muted">detected via action formula</td>
            <td class="num">—</td>
          </tr>`
        ).join('');

        const total = collections.length + extraNames.size;
        return `
          <details class="collections-detail">
            <summary>Collections (${total})</summary>
            <table class="data-table" style="margin-top:8px">
              <thead><tr><th>Name</th><th>Populated On</th><th>Used At</th></tr></thead>
              <tbody>${colRows}${extraRows}</tbody>
            </table>
          </details>`;
      })();

      // ── App-level data sources (filtered) ───────────────────────────────
      const relevantAppDs = app.dataSources.filter(isRelevantSource);
      const dsRows = relevantAppDs.map(ds => `
        <tr>
          <td>${esc(ds.displayName)}</td>
          <td>${ds.type === 'sharepoint' ? 'SharePoint' : 'Dataverse'}</td>
          <td><code style="font-size:11px">${esc(ds.siteUrl ?? ds.tableName ?? ds.environmentUrl ?? '—')}</code></td>
        </tr>`).join('');

      const appCheckerHtml = (() => {
        const findings = app.appCheckerFindings;
        if (!findings.length) return '';

        const errorCount = findings.filter(f => f.level === 'error').length;
        const warnCount  = findings.filter(f => f.level === 'warning').length;

        const badgesHtml = [
          errorCount ? `<span class="ck-badge ck-badge-error">${errorCount} error${errorCount !== 1 ? 's' : ''}</span>` : '',
          warnCount  ? `<span class="ck-badge ck-badge-warn">${warnCount} warning${warnCount !== 1 ? 's' : ''}</span>`  : ''
        ].filter(Boolean).join('');

        // Group all findings by ruleId
        const grouped = new Map<string, AppCheckerFinding[]>();
        for (const f of findings) {
          const bucket = grouped.get(f.ruleId) ?? [];
          bucket.push(f);
          grouped.set(f.ruleId, bucket);
        }

        // Sort groups: errors first → warnings → notes; ties broken by count desc
        const levelRank = (fs: AppCheckerFinding[]) =>
          fs.some(f => f.level === 'error') ? 0 : fs.some(f => f.level === 'warning') ? 1 : 2;
        const sortedGroups = Array.from(grouped.entries())
          .sort(([, a], [, b]) => levelRank(a) - levelRank(b) || b.length - a.length);

        const ruleGroupsHtml = sortedGroups.map(([ruleId, groupFindings]) => {
          const worstLevel = groupFindings.some(f => f.level === 'error') ? 'error'
            : groupFindings.some(f => f.level === 'warning') ? 'warning' : 'note';

          const levelBadge = worstLevel === 'error'
            ? `<span class="ck-badge ck-badge-error">error</span>`
            : worstLevel === 'warning'
            ? `<span class="ck-badge ck-badge-warn">warning</span>`
            : `<span class="ck-badge ck-badge-note">note</span>`;

          const rows = groupFindings.map(f =>
            `<tr class="finding-${f.level}">
               <td class="level-cell level-${f.level}">${f.level}</td>
               <td><code class="loc-code">${esc(f.location ?? '—')}</code></td>
               <td>${esc(f.message)}</td>
             </tr>`
          ).join('');

          return `
            <details class="ck-rule-group">
              <summary>
                <code class="ck-rule-id">${esc(ruleId)}</code>
                <span class="ck-rule-count">${groupFindings.length} finding${groupFindings.length !== 1 ? 's' : ''}</span>
                ${levelBadge}
              </summary>
              <table class="data-table ck-table">
                <thead><tr><th>Level</th><th>Location</th><th>Message</th></tr></thead>
                <tbody>${rows}</tbody>
              </table>
            </details>`;
        }).join('');

        return `
          <details class="ck-section">
            <summary class="ck-summary">
              Quality / AppChecker Findings
              <span class="ck-total">(${findings.length})</span>
              ${badgesHtml}
            </summary>
            <div class="ck-body">
              ${ruleGroupsHtml}
            </div>
          </details>`;
      })();

      const navDiagramHtml = app.mermaidNavGraph
        ? `<h4>Navigation Graph <span class="confidence">(${app.navConfidence} confidence)</span></h4>
           <pre class="mermaid">${esc(app.mermaidNavGraph)}</pre>`
        : '';

      const onStartHtml = app.appOnStart.redacted
        ? `<h4>App.OnStart</h4><pre class="formula-block">${fxHtml(app.appOnStart.redacted)}</pre>`
        : '';

      return `
        <div id="app-${slugify(app.id)}" class="subsection">
          <h2>${esc(app.displayName)}</h2>
          <p class="source-ref">Source: <code>${esc(app.msappPath)}</code></p>
          ${onStartHtml}
          <h3>Screens (${app.screens.length})</h3>
          ${app.screens.length ? `<div class="screens-list">${screenDetails}</div>` : '<p class="muted">No screens extracted.</p>'}
          ${navDiagramHtml}
          <h3>Variables (${nonCollectionVars.length})</h3>
          ${nonCollectionVars.length ? `<table class="data-table">
            <thead><tr><th>Name</th><th>Kind</th><th>Set At</th></tr></thead>
            <tbody>${varRows}</tbody></table>` : '<p class="muted">No global/context variables detected.</p>'}
          ${collectionsHtml}
          <h3>Data Sources (${relevantAppDs.length})</h3>
          ${relevantAppDs.length ? `<table class="data-table">
            <thead><tr><th>Name</th><th>Type</th><th>Location / Table</th></tr></thead>
            <tbody>${dsRows}</tbody></table>` : '<p class="muted">No custom data sources detected.</p>'}
          ${(() => {
            const appFlows = app.connectors.filter(c => c.connectorId.startsWith('flow:'));
            if (!appFlows.length) return '';
            const friendlyName = (raw: string) => raw.replace(/_/g, ' ').trim();
            const flowId = (c: ConnectorRef) => c.connectorId.replace('flow:', '');
            const rows = appFlows.map(f => `<tr>
              <td title="${esc(flowId(f))}">${esc(friendlyName(f.displayName))}</td>
              <td class="muted" style="font-size:11px"><code>${esc(f.displayName)}</code></td>
            </tr>`).join('');
            return `<h3>Connected Flows (${appFlows.length})</h3>
              <table class="data-table">
                <thead><tr><th>Flow</th><th>Connector Name</th></tr></thead>
                <tbody>${rows}</tbody>
              </table>`;
          })()}
          ${appCheckerHtml}
        </div>`;
    }).join('\n');

    return {
      id: 'canvas-apps', title: 'Canvas Apps', count: graph.canvasApps.length,
      children,
      html: `<section id="canvas-apps" class="section"><h1>Canvas Apps (${graph.canvasApps.length})</h1>${appsHtml}</section>`
    };
  }

  // ── Model-Driven Apps ───────────────────────────────────────────────────────

  private static mdaSection(graph: NormalizedAnalysisGraph): Section {
    const children: SectionChild[] = graph.modelDrivenApps.map(app => ({
      id: `mda-${slugify(app.uniqueName)}`, title: app.displayName
    }));

    const html = graph.modelDrivenApps.map(app => {
      const roleRows = app.assignedRoleNames.map(r => `<tr><td>${esc(r)}</td></tr>`).join('');
      const tableRows = app.exposedTables.map(t => `<tr><td><code>${esc(t)}</code></td></tr>`).join('');

      return `
        <div id="mda-${slugify(app.uniqueName)}" class="subsection">
          <h2>${esc(app.displayName)}</h2>
          <table class="data-table">
            <tr><td>Unique Name</td><td><code>${esc(app.uniqueName)}</code></td></tr>
            <tr><td>Description</td><td>${esc(app.description ?? '—')}</td></tr>
          </table>
          <h3>Assigned Roles (${app.assignedRoleNames.length})</h3>
          ${app.assignedRoleNames.length ? `<table class="data-table"><tbody>${roleRows}</tbody></table>` : '<p class="muted">No roles assigned.</p>'}
          <h3>Sitemap</h3>
          <pre class="mermaid">${esc(app.mermaidSitemap)}</pre>
          <h3>Exposed Tables (${app.exposedTables.length})</h3>
          ${app.exposedTables.length ? `<table class="data-table"><tbody>${tableRows}</tbody></table>` : '<p class="muted">No tables detected.</p>'}
          <p class="source-ref">Source: <code>${esc(app.source.archivePath)}</code></p>
        </div>`;
    }).join('\n');

    return {
      id: 'model-driven-apps', title: 'Model-Driven Apps', count: graph.modelDrivenApps.length,
      children,
      html: `<section id="model-driven-apps" class="section"><h1>Model-Driven Apps (${graph.modelDrivenApps.length})</h1>${html}</section>`
    };
  }

  // ── Flows ───────────────────────────────────────────────────────────────────

  private static flowsSection(graph: NormalizedAnalysisGraph): Section {
    const children: SectionChild[] = graph.flows.map(f => ({
      id: `flow-${slugify(f.id)}`, title: f.displayName, count: f.actions.length
    }));

    const html = graph.flows.map(flow => {
      const triggerLabel = flow.trigger.recurrence
        ? `Recurrence — every ${flow.trigger.recurrence.interval} ${flow.trigger.recurrence.frequency}` +
          (flow.trigger.recurrence.timeZone ? ` (${flow.trigger.recurrence.timeZone})` : '')
        : flow.trigger.automated
        ? `Automated — ${flow.trigger.automated.connectorId} / ${flow.trigger.automated.operationId}`
        : flow.triggerType;

      const actionRows = flow.actions.map(a => {
        // If this action calls AI Builder, resolve the model name from the graph
        let aiBuilderCell = '';
        if (a.aiBuilderModelId) {
          const pretrainedName = HtmlGenerator.AIB_PRETRAINED_NAMES[a.aiBuilderModelId.toLowerCase()];
          const model = !pretrainedName
            ? graph.aiBuilderModels.find(m => m.modelId === a.aiBuilderModelId)
            : undefined;
          const modelLabel = pretrainedName
            ? `<span class="aib-pretrained-badge">🤖 ${esc(pretrainedName)} <em>(pretrained)</em></span>`
            : model
            ? `<a href="#aib-${slugify(model.modelId)}" class="aib-link">🤖 ${esc(model.name)}</a>`
            : `<span class="aib-link-unresolved">🤖 ${esc(a.aiBuilderModelId)}</span>`;
          aiBuilderCell = `<br>${modelLabel}`;
        }
        return `
        <tr${a.aiBuilderModelId ? ' class="aib-action-row"' : ''}>
          <td><code>${esc(a.id)}</code></td>
          <td>${esc(a.type)}${aiBuilderCell}</td>
          <td>${esc(a.connector ?? '—')}</td>
          <td>${esc(a.operationId ?? '—')}</td>
          <td>${a.runAfter.map(r => `<code>${esc(r)}</code>`).join(', ') || '—'}</td>
          <td>${a.secureInputs ? '🔒' : ''}${a.secureOutputs ? '🔒' : ''}</td>
        </tr>`;
      }).join('');

      const errHtml = flow.errorHandling.hasExplicitHandling
        ? `✅ Explicit error handling detected at: ${flow.errorHandling.handlingLocations.map(l => `<code>${esc(l)}</code>`).join(', ')}`
        : '⚠️ No explicit error handling detected';

      return `
        <div id="flow-${slugify(flow.id)}" class="subsection">
          <h2>${esc(flow.displayName)}</h2>
          <table class="data-table">
            <tr><td>Trigger</td><td>${esc(triggerLabel)}</td></tr>
            <tr><td>State</td><td>${esc(flow.state)}</td></tr>
            <tr><td>Actions</td><td>${flow.actions.length}</td></tr>
            <tr><td>Error Handling</td><td>${errHtml}</td></tr>
          </table>
          <h3>Flow Diagram</h3>
          <pre class="mermaid">${esc(flow.mermaidDiagram)}</pre>
          <details>
            <summary>All Actions (${flow.actions.length})</summary>
            <table class="data-table">
              <thead><tr><th>ID</th><th>Type</th><th>Connector</th><th>Operation</th><th>Run After</th><th>Secure</th></tr></thead>
              <tbody>${actionRows}</tbody>
            </table>
          </details>
          <p class="source-ref">Source: <code>${esc(flow.source.archivePath)}</code></p>
        </div>`;
    }).join('\n');

    return {
      id: 'flows', title: 'Flows', count: graph.flows.length, children,
      html: `<section id="flows" class="section"><h1>Flows (${graph.flows.length})</h1>${html}</section>`
    };
  }

  // ── Env Vars ────────────────────────────────────────────────────────────────

  private static envVarsSection(graph: NormalizedAnalysisGraph): Section {
    const TYPE_ICON: Record<string, string> = {
      string: '📝', number: '#️⃣', boolean: '✅', json: '{ }',
      datasource: '🔗', secret: '🔒'
    };

    const cards = graph.envVars.map(ev => {
      const icon = TYPE_ICON[ev.type] ?? '📋';
      const isMissing = ev.currentValue === 'not included in export';
      return `
        <details class="ev-card">
          <summary class="ev-summary">
            <span class="ev-name">${esc(ev.displayName)}</span>
            <span class="ev-type-badge">${icon} ${esc(ev.type)}</span>
          </summary>
          <div class="ev-body">
            <table class="data-table ev-table">
              <tr><td>Schema Name</td><td><code>${esc(ev.schemaName)}</code></td></tr>
              ${ev.description ? `<tr><td>Description</td><td>${esc(ev.description)}</td></tr>` : ''}
              <tr><td>Default Value</td><td>${ev.defaultValue ? `<code>${esc(ev.defaultValue)}</code>` : '<span class="muted">—</span>'}</td></tr>
              <tr><td>Current Value</td><td class="${isMissing ? 'missing' : ''}">${isMissing ? 'not included in export' : `<code>${esc(ev.currentValue ?? '—')}</code>`}</td></tr>
            </table>
          </div>
        </details>`;
    }).join('');

    return {
      id: 'env-vars', title: 'Environment Variables', count: graph.envVars.length,
      html: `<section id="env-vars" class="section">
        <h1>Environment Variables (${graph.envVars.length})</h1>
        <div class="ev-list">${cards}</div>
      </section>`
    };
  }

  // ── Connectors ──────────────────────────────────────────────────────────────

  private static connectorsSection(graph: NormalizedAnalysisGraph): Section {
    const rows = graph.connectors.map(c => `
      <tr>
        <td><code>${esc(c.connectorId)}</code></td>
        <td>${esc(c.displayName)}</td>
        <td>${c.usedInFlows.length}</td>
        <td>${c.usedInApps.length}</td>
      </tr>`).join('');

    return {
      id: 'connectors', title: 'Connectors', count: graph.connectors.length,
      html: `<section id="connectors" class="section">
        <h1>Connectors (${graph.connectors.length})</h1>
        <table class="data-table">
          <thead><tr><th>Connector ID</th><th>Display Name</th><th>Used in Flows</th><th>Used in Apps</th></tr></thead>
          <tbody>${rows}</tbody>
        </table></section>`
    };
  }

  // ── Data Sources ─────────────────────────────────────────────────────────────

  private static dataSourcesSection(graph: NormalizedAnalysisGraph): Section {
    // Only document SharePoint lists and custom Dataverse tables.
    // System/default tables and other connectors are excluded — they don't need
    // schema documentation and clutter the report.
    const spList = graph.dataSources.filter(d => d.type === 'sharepoint');
    const dvList = graph.dataSources.filter(d => d.type === 'dataverse' && isRelevantSource(d));
    const relevant = spList.length + dvList.length;

    const dsHtml = (list: DataSource[]) => list.map(ds => `
      <div class="ds-card">
        <h4>${esc(ds.displayName)}</h4>
        <table class="data-table">
          <tr><td>Type</td><td>${ds.type === 'sharepoint' ? 'SharePoint List' : 'Dataverse Table (custom)'}</td></tr>
          ${ds.siteUrl ? `<tr><td>Site URL</td><td><code>${esc(ds.siteUrl)}</code></td></tr>` : ''}
          ${ds.listName ? `<tr><td>List / Table</td><td><code>${esc(ds.listName)}</code></td></tr>` : ''}
          ${ds.tableName ? `<tr><td>Logical Name</td><td><code>${esc(ds.tableName)}</code></td></tr>` : ''}
          ${ds.environmentUrl ? `<tr><td>Environment</td><td><code>${esc(ds.environmentUrl)}</code></td></tr>` : ''}
        </table>
        ${ds.columns?.length ? `
          <details open>
            <summary>Schema (${ds.columns.length} columns — auto-extracted)</summary>
            <table class="data-table">
              <thead><tr><th>Display Name</th><th>Internal Name</th><th>Type</th><th>Notes</th></tr></thead>
              <tbody>${ds.columns.map(c =>
                `<tr>
                  <td>${esc(c.displayName ?? c.name)}</td>
                  <td><code>${esc(c.name)}</code></td>
                  <td>${esc(c.type ?? '—')}</td>
                  <td>${esc(c.notes ?? '—')}</td>
                </tr>`
              ).join('')}</tbody>
            </table>
          </details>` : `<p class="muted schema-missing">⚠️ Schema not available — this .msapp may predate embedded schema storage</p>`}
        <p class="source-ref">Source: <code>${esc(ds.source.archivePath)}</code></p>
      </div>`).join('');

    return {
      id: 'data-sources', title: 'Data Sources', count: relevant,
      html: `<section id="data-sources" class="section">
        <h1>Data Sources (${relevant})</h1>
        <p class="muted ds-scope-note">Showing SharePoint lists and custom Dataverse tables only. System tables and default connectors are excluded.</p>
        ${spList.length ? `<h2>SharePoint Lists (${spList.length})</h2>${dsHtml(spList)}` : ''}
        ${dvList.length ? `<h2>Custom Dataverse Tables (${dvList.length})</h2>${dsHtml(dvList)}` : ''}
        ${!relevant ? '<p class="muted">No custom data sources detected.</p>' : ''}
      </section>`
    };
  }

  // ── AI Builder — pretrained model name lookup ────────────────────────────────
  private static readonly AIB_PRETRAINED_NAMES: Record<string, string> = {
    aibuilderpredict_textrecognition:            'Text Recognition (OCR)',
    aibuilderpredict_receiptscanningpretrained:  'Receipt Scanning',
    aibuilderpredict_invoiceprocessingpretrained:'Invoice Processing',
    aibuilderpredict_businesscardreader:         'Business Card Reader',
    aibuilderpredict_idreader:                   'ID Reader',
    aibuilderpredict_objectdetectionpretrained:  'Object Detection',
  };

  // ── Security Roles — view-model constants ───────────────────────────────────
  // Action verbs in preferred display order (AppendTo before Append to avoid partial match)
  private static readonly RV_VERB_ORDER =
    ['Create', 'Read', 'Write', 'Delete', 'Append', 'AppendTo', 'Assign', 'Share'];
  // Numeric rank for scope comparison (higher = more permissive)
  private static readonly RV_SCOPE_RANK: Record<string, number> =
    { None: 0, Basic: 1, Local: 2, Deep: 3, Global: 4 };
  // Background / foreground colours for scope levels — consistent across whole report
  private static readonly RV_SCOPE_BG: Record<string, string> = {
    Basic: '#e8f5e9', Local: '#e3f2fd', Deep: '#fff8e1',
    Global: '#ffebee', None: '#f5f5f5', Unknown: '#f0f0f0'
  };
  private static readonly RV_SCOPE_FG: Record<string, string> = {
    Basic: '#1b5e20', Local: '#0d47a1', Deep: '#bf360c',
    Global: '#b71c1c', None: '#9e9e9e', Unknown: '#666'
  };
  // Verbs that carry meaningful risk when scoped at Global
  private static readonly RV_RISK_VERBS = new Set(['Write', 'Delete', 'Assign', 'Share', 'Create']);
  // Regex — AppendTo must precede Append to avoid short-match
  private static readonly RV_ACTION_RE =
    /^prv(AppendTo|Append|Read|Write|Create|Delete|Assign|Share)/i;
  // System privilege category rules — first match wins; unknown → Other / Unclassified
  private static readonly RV_SYS_CATEGORIES: [RegExp, string][] = [
    [/workflow|asyncoperation|workflowlog|sdkmessage|plugintype|pluginassembly|processsession|processstage/i,
      'Workflow / Process / Automation'],
    [/entity|attribute|relationship|webresource|customization|solution|publisher|savedquery|systemform|importmap/i,
      'Customization / Metadata'],
    [/systemuser|teamtemplate|businessunit|privilege|userentityui|fieldsecurityprofile/i,
      'Users / Teams / Security'],
    [/import|bulkdelete|datamap|importfile|importlog|duplicaterule/i,
      'Import / Data Management'],
    [/knowledgearticle|knowledgebase|feedback|ratingmodel|topic/i,
      'Knowledge / Content'],
    [/sharepointdocument|emailserver|mailbox|exchangesync|activitymime|emailtemplate|socialprofile/i,
      'Integration (SharePoint / Exchange / Mailbox)'],
    [/usersettings|userform|userquery|personalwall|ownermapping|mobileoffline/i,
      'User Personalization'],
    [/connection|connector|environmentvariable|flowsession|catalog|appmodule|canvasapp/i,
      'Platform / Integration Features'],
    [/report|goal|metric|rollup|kpi|service|appointment|serviceappointment|contract|entitlement/i,
      'Reporting / Service'],
  ];
  // Plain-English descriptions for known feature/special privilege names (after stripping "prv")
  private static readonly RV_SPECIAL_DESC: Record<string, string> = {
    GoOffline:              'Access data while offline',
    ExportToExcel:          'Export records to Excel spreadsheets',
    MailMerge:              'Perform mail merge operations',
    WorkflowExecution:      'Trigger and execute workflow processes',
    Flow:                   'Create and manage Power Automate flows',
    UseOfficeApps:          'Integrate with Microsoft Office applications',
    WebMailMerge:           'Perform web-based mail merge',
    SendAsAnotherUser:      'Send email on behalf of another user',
    SendEmail:              'Send email directly from records',
    BulkEmail:              'Send bulk email to multiple recipients',
    RunDataMap:             'Execute data migration maps',
    ImportFiles:            'Import external data files',
    ExportCustomizations:   'Export solution customizations',
    ImportCustomizations:   'Import solution customizations',
    PublishCustomizations:  'Publish Dataverse customizations',
    ExecuteWorkflow:        'Run on-demand workflow processes',
    BulkEditRequest:        'Perform bulk edit operations on records',
    PersonalWall:           'Access the personal activity feed wall',
    ViewAuditHistory:       'View audit history log entries',
    ManageFlows:            'Create, edit and delete Power Automate flows',
  };

  // ── Security Roles — view-model helpers ─────────────────────────────────────

  /** Extract the verb portion from a raw privilege name (e.g. 'prvReadAccount' → 'Read').
   *  Normalizes to canonical casing from RV_VERB_ORDER so 'Appendto' → 'AppendTo'. */
  private static rvExtractAction(privilegeName: string): string | null {
    const m = privilegeName.match(HtmlGenerator.RV_ACTION_RE);
    if (!m) return null;
    const captured = m[1];
    // Normalize: find the canonical form whose lowercase matches the captured string's lowercase
    return HtmlGenerator.RV_VERB_ORDER.find(v => v.toLowerCase() === captured.toLowerCase()) ?? captured;
  }

  /** Classify a privilege as custom-table, system-crud, or special/feature. */
  private static rvClassify(p: { privilegeName: string; entityName: string }): 'custom' | 'system-crud' | 'special' {
    const action = HtmlGenerator.rvExtractAction(p.privilegeName);
    if (!action) return 'special';                        // no known verb → feature privilege
    return p.entityName.includes('_') ? 'custom' : 'system-crud';
  }

  /** Assign a system privilege to the first matching functional category. */
  private static rvCategorize(privilegeName: string): string {
    for (const [re, cat] of HtmlGenerator.RV_SYS_CATEGORIES) {
      if (re.test(privilegeName)) return cat;
    }
    return 'Other / Unclassified';
  }

  /** Return the highest-ranked scope string from an array of depth values. */
  private static rvHighestScope(depths: string[]): string {
    return depths.reduce((best, d) =>
      (HtmlGenerator.RV_SCOPE_RANK[d] ?? -1) > (HtmlGenerator.RV_SCOPE_RANK[best] ?? -1) ? d : best
    , 'None');
  }

  /** Render a coloured scope badge <span>. */
  private static rvScopeBadge(depth: string): string {
    const bg = HtmlGenerator.RV_SCOPE_BG[depth] ?? '#f0f0f0';
    const fg = HtmlGenerator.RV_SCOPE_FG[depth] ?? '#444';
    return `<span class="role-scope-badge" style="background:${bg};color:${fg}">${esc(depth)}</span>`;
  }

  /** Render a matrix cell — coloured when a scope is present, dash when absent. */
  private static rvScopeCell(depth: string | undefined): string {
    if (!depth || depth === 'None') return `<td class="mx-empty">—</td>`;
    const bg = HtmlGenerator.RV_SCOPE_BG[depth] ?? '#f0f0f0';
    const fg = HtmlGenerator.RV_SCOPE_FG[depth] ?? '#444';
    return `<td class="mx-cell" style="background:${bg};color:${fg}">${esc(depth)}</td>`;
  }

  // ── Security Roles — section renderer ───────────────────────────────────────

  private static rolesSection(graph: NormalizedAnalysisGraph): Section {
    const html = graph.securityRoles.map(role => {

      // ── Bucket privileges into three types ────────────────────────────────
      const customPrivs   = role.privileges.filter(p => HtmlGenerator.rvClassify(p) === 'custom');
      const sysCrudPrivs  = role.privileges.filter(p => HtmlGenerator.rvClassify(p) === 'system-crud');
      const specialPrivs  = role.privileges.filter(p => HtmlGenerator.rvClassify(p) === 'special');

      // ── Scope distribution (all privs) ────────────────────────────────────
      const scopeDist: Record<string, number> = {};
      for (const p of role.privileges) {
        const d = p.depth || 'Unknown';
        scopeDist[d] = (scopeDist[d] ?? 0) + 1;
      }

      // ── Global risk counts ────────────────────────────────────────────────
      const riskCounts: Record<string, number> = {};
      for (const p of role.privileges) {
        if (p.depth !== 'Global') continue;
        const verb = HtmlGenerator.rvExtractAction(p.privilegeName);
        if (verb && HtmlGenerator.RV_RISK_VERBS.has(verb)) {
          const key = `Global ${verb}`;
          riskCounts[key] = (riskCounts[key] ?? 0) + 1;
        }
      }
      const totalRisks = Object.values(riskCounts).reduce((s, n) => s + n, 0);

      // ── Custom table: entity → { verb → depth } map ───────────────────────
      const entityMap = new Map<string, Record<string, string>>();
      for (const p of customPrivs) {
        const verb = HtmlGenerator.rvExtractAction(p.privilegeName) ?? p.privilegeName;
        if (!entityMap.has(p.entityName)) entityMap.set(p.entityName, {});
        entityMap.get(p.entityName)![verb] = p.depth;
      }
      // Verb columns: preferred order first, then any unexpected verbs appended
      const customVerbsFound = new Set<string>(customPrivs.map(p =>
        HtmlGenerator.rvExtractAction(p.privilegeName) ?? p.privilegeName));
      const verbCols = [
        ...HtmlGenerator.RV_VERB_ORDER.filter(v => customVerbsFound.has(v)),
        ...[...customVerbsFound].filter(v => !HtmlGenerator.RV_VERB_ORDER.includes(v)),
      ];

      // ── Custom entity risk highlights ─────────────────────────────────────
      interface EntityRisk { entity: string; highest: string; flags: string[] }
      const entityRisks: EntityRisk[] = [];
      for (const [entity, verbs] of entityMap) {
        const depths  = Object.values(verbs);
        const highest = HtmlGenerator.rvHighestScope(depths);
        const flags: string[] = [];
        for (const [v, d] of Object.entries(verbs)) {
          if (d === 'Global' && HtmlGenerator.RV_RISK_VERBS.has(v)) flags.push(`Global ${v}`);
        }
        // Mixed scope: any verb at Global while another verb is below Global
        const hasGlobal = depths.some(d => d === 'Global');
        const hasLower  = depths.some(d => d !== 'Global' && d !== 'None' && (HtmlGenerator.RV_SCOPE_RANK[d] ?? 0) > 0);
        if (hasGlobal && hasLower) flags.push('Mixed scope');
        // Only surface entities with actual risk flags — Read-only Global access is not flagged
        if (flags.length > 0) {
          entityRisks.push({ entity, highest, flags });
        }
      }

      // ── System CRUD: verb × scope count matrix ────────────────────────────
      const sysVerbCounts: Record<string, Record<string, number>> = {}; // verb → scope → count
      const sysVerbsFound = new Set<string>();
      const sysScopesFound = new Set<string>();
      for (const p of sysCrudPrivs) {
        const verb  = HtmlGenerator.rvExtractAction(p.privilegeName) ?? 'Other';
        const scope = p.depth || 'Unknown';
        sysVerbsFound.add(verb); sysScopesFound.add(scope);
        if (!sysVerbCounts[verb]) sysVerbCounts[verb] = {};
        sysVerbCounts[verb][scope] = (sysVerbCounts[verb][scope] ?? 0) + 1;
      }
      const sysVerbs  = [...HtmlGenerator.RV_VERB_ORDER.filter(v => sysVerbsFound.has(v)),
                         ...[...sysVerbsFound].filter(v => !HtmlGenerator.RV_VERB_ORDER.includes(v))];
      const sysScopes = ['Basic', 'Local', 'Deep', 'Global'].filter(s => sysScopesFound.has(s));
      if (sysScopesFound.has('Unknown')) sysScopes.push('Unknown');

      // ── System CRUD: group by functional category ─────────────────────────
      const sysByCategory = new Map<string, RolePrivilege[]>();
      for (const p of sysCrudPrivs) {
        const cat = HtmlGenerator.rvCategorize(p.privilegeName);
        if (!sysByCategory.has(cat)) sysByCategory.set(cat, []);
        sysByCategory.get(cat)!.push(p);
      }

      // ── Parsing notes ─────────────────────────────────────────────────────
      const parsingNotes: string[] = [];
      const unknownCount = role.privileges.filter(p => !p.depth || p.depth === 'Unknown').length;
      const unclassifiedCount = sysCrudPrivs
        .filter(p => HtmlGenerator.rvCategorize(p.privilegeName) === 'Other / Unclassified').length;
      if (unknownCount)
        parsingNotes.push(`${unknownCount} privilege(s) have an Unknown scope level — check source XML @_level attribute`);
      if (unclassifiedCount)
        parsingNotes.push(`${unclassifiedCount} system privilege(s) did not match any category keyword and appear under Other / Unclassified`);

      // ════════════════════════════════════════════════════════════════════════
      // RENDER
      // ════════════════════════════════════════════════════════════════════════

      // 1) Role header meta + scope distribution + risk bar
      const scopeDistHtml = Object.entries(scopeDist)
        .sort(([a], [b]) => (HtmlGenerator.RV_SCOPE_RANK[b] ?? -1) - (HtmlGenerator.RV_SCOPE_RANK[a] ?? -1))
        .map(([s, n]) => `${HtmlGenerator.rvScopeBadge(s)}<span class="role-scope-count">${n}</span>`)
        .join('');

      const riskBarHtml = totalRisks ? `
        <div class="role-risk-bar">
          <span class="role-risk-label">⚠ Global risks:</span>
          ${Object.entries(riskCounts).map(([r, n]) =>
            `<span class="role-risk-chip">${esc(r)}${n > 1 ? ` ×${n}` : ''}</span>`
          ).join('')}
        </div>` : '';

      const metaHtml = `
        <div class="role-meta-block">
          <table class="data-table ev-table" style="margin-bottom:0">
            <tr><td>ID</td><td><code>${esc(role.id || '—')}</code></td></tr>
            <tr><td>Total Privileges</td><td>${role.privilegeCount}</td></tr>
            <tr><td>Custom Tables</td><td>${entityMap.size} ${entityMap.size === 1 ? 'entity' : 'entities'} · ${customPrivs.length} privileges</td></tr>
            <tr><td>System Privileges</td><td>${sysCrudPrivs.length} CRUD · ${specialPrivs.length} feature</td></tr>
            <tr><td>Assigned To</td><td>${role.assignedToApps.length ? role.assignedToApps.map(a => esc(a)).join(', ') : '—'}</td></tr>
          </table>
          <div class="role-scope-dist">${scopeDistHtml}</div>
          ${riskBarHtml}
        </div>`;

      // 2) Custom table section
      const riskTableHtml = entityRisks.length ? `
        <div class="role-section-label">⚠ Elevated Permissions</div>
        <table class="data-table role-risk-table">
          <thead><tr><th>Entity</th><th>Highest Level</th><th>Flags</th></tr></thead>
          <tbody>${entityRisks.map(r => `
            <tr>
              <td><code>${esc(r.entity)}</code></td>
              <td>${HtmlGenerator.rvScopeBadge(r.highest)}</td>
              <td>${r.flags.map(f => `<span class="role-flag-chip">${esc(f)}</span>`).join(' ')}</td>
            </tr>`).join('')}
          </tbody>
        </table>` : '';

      const customMatrixHtml = entityMap.size ? `
        <div class="role-section-block">
          <div class="role-section-label">Custom Table Privileges — ${entityMap.size} ${entityMap.size === 1 ? 'entity' : 'entities'}</div>
          ${riskTableHtml}
          <div class="role-matrix-wrap">
            <table class="role-matrix">
              <thead>
                <tr>
                  <th class="mx-entity-col">Entity</th>
                  ${verbCols.map(v => `<th class="mx-verb-col">${esc(v)}</th>`).join('')}
                </tr>
              </thead>
              <tbody>
                ${[...entityMap.entries()].map(([entity, verbs]) => `
                  <tr>
                    <td class="mx-entity" title="${esc(entity)}">${esc(entity)}</td>
                    ${verbCols.map(v => HtmlGenerator.rvScopeCell(verbs[v])).join('')}
                  </tr>`).join('')}
              </tbody>
            </table>
          </div>
        </div>` : '';

      // 3) System CRUD verb×scope summary matrix
      const sysMatrixHtml = sysVerbs.length ? `
        <div class="role-section-label">System CRUD — ${sysCrudPrivs.length} privileges across ${sysByCategory.size} categories</div>
        <table class="data-table role-sys-matrix">
          <thead>
            <tr>
              <th style="text-align:left">Action</th>
              ${sysScopes.map(s => `<th>${esc(s)}</th>`).join('')}
              <th>Total</th>
            </tr>
          </thead>
          <tbody>
            ${sysVerbs.map(v => {
              const row   = sysVerbCounts[v] ?? {};
              const total = Object.values(row).reduce((s, n) => s + n, 0);
              return `<tr>
                <td style="font-weight:600">${esc(v)}</td>
                ${sysScopes.map(s => {
                  const n = row[s] ?? 0;
                  if (!n) return `<td class="mx-empty">—</td>`;
                  const bg = HtmlGenerator.RV_SCOPE_BG[s] ?? '#f0f0f0';
                  const fg = HtmlGenerator.RV_SCOPE_FG[s] ?? '#444';
                  return `<td style="background:${bg};color:${fg};font-weight:600;text-align:center">${n}</td>`;
                }).join('')}
                <td class="num">${total}</td>
              </tr>`;
            }).join('')}
          </tbody>
        </table>` : '';

      // 4) Special / feature privileges
      const specialHtml = specialPrivs.length ? `
        <details class="role-sys-details">
          <summary class="role-sys-summary">Feature / Special Privileges (${specialPrivs.length})</summary>
          <table class="data-table role-priv-table">
            <thead><tr><th>Feature</th><th>Level</th><th>Description</th></tr></thead>
            <tbody>${specialPrivs.map(p => {
              const featureName = p.privilegeName.replace(/^prv/i, '');
              const desc = HtmlGenerator.RV_SPECIAL_DESC[featureName] ?? '';
              return `<tr>
                <td><code title="${esc(p.privilegeName)}">${esc(featureName)}</code></td>
                <td>${HtmlGenerator.rvScopeBadge(p.depth)}</td>
                <td class="role-desc-cell">${esc(desc)}</td>
              </tr>`;
            }).join('')}
            </tbody>
          </table>
        </details>` : '';

      // 5) System CRUD grouped by category (sorted alphabetically, Other last)
      //    Within each category, further group by entity so each object appears once
      //    with all its actions shown inline — much easier to scan than a flat sorted list.
      const sysCategoryHtml = [...sysByCategory.entries()]
        .sort(([a], [b]) => {
          if (a === 'Other / Unclassified') return 1;
          if (b === 'Other / Unclassified') return -1;
          return a.localeCompare(b);
        })
        .map(([cat, privs]) => {
          // Group by entity name within this category
          const byEntity = new Map<string, { verb: string; depth: string; raw: string }[]>();
          for (const p of privs) {
            const verb = HtmlGenerator.rvExtractAction(p.privilegeName) ?? p.privilegeName;
            if (!byEntity.has(p.entityName)) byEntity.set(p.entityName, []);
            byEntity.get(p.entityName)!.push({ verb, depth: p.depth, raw: p.privilegeName });
          }
          // Sort entities alphabetically within the category
          const sortedEntities = [...byEntity.entries()].sort(([a], [b]) => a.localeCompare(b));
          // Sort each entity's actions in VERB_ORDER order
          const verbRank = Object.fromEntries(HtmlGenerator.RV_VERB_ORDER.map((v, i) => [v, i]));
          for (const [, actions] of sortedEntities) {
            actions.sort((a, b) => (verbRank[a.verb] ?? 99) - (verbRank[b.verb] ?? 99));
          }

          return `
          <details class="role-sys-details">
            <summary class="role-sys-summary">${esc(cat)} — ${sortedEntities.length} ${sortedEntities.length === 1 ? 'object' : 'objects'} (${privs.length} privileges)</summary>
            <table class="data-table role-priv-table">
              <thead><tr><th>Entity / Object</th><th>Permissions</th></tr></thead>
              <tbody>${sortedEntities.map(([entity, actions]) => `
                <tr>
                  <td class="sys-entity-cell"><code>${esc(entity)}</code></td>
                  <td class="sys-actions-cell">${actions.map(a =>
                    `<span class="sys-action-pill" title="${esc(a.raw)}">${esc(a.verb)}<span class="sys-action-scope" style="background:${HtmlGenerator.RV_SCOPE_BG[a.depth] ?? '#eee'};color:${HtmlGenerator.RV_SCOPE_FG[a.depth] ?? '#444'}">${esc(a.depth)}</span></span>`
                  ).join('')}</td>
                </tr>`).join('')}
              </tbody>
            </table>
          </details>`;
        }).join('');

      const sysPrivsHtml = (sysCrudPrivs.length || specialPrivs.length) ? `
        <div class="role-section-block">
          ${sysMatrixHtml}
          ${specialHtml}
          ${sysCategoryHtml ? `<div class="role-section-label" style="margin-top:6px">System CRUD — by Category</div>${sysCategoryHtml}` : ''}
        </div>` : '';

      // 6) Parsing notes
      const parsingHtml = parsingNotes.length ? `
        <details class="role-sys-details">
          <summary class="role-sys-summary role-notes-summary">Parsing Notes (${parsingNotes.length})</summary>
          <ul class="role-notes-list">${parsingNotes.map(n => `<li>${esc(n)}</li>`).join('')}</ul>
        </details>` : '';

      return `
        <details class="role-card">
          <summary class="role-summary">
            <span class="role-name">${esc(role.name)}</span>
            <span class="role-meta">${entityMap.size} custom · ${sysCrudPrivs.length + specialPrivs.length} system</span>
            ${totalRisks ? `<span class="role-risk-pill">⚠ ${totalRisks} risk${totalRisks > 1 ? 's' : ''}</span>` : ''}
          </summary>
          <div class="role-body">
            ${metaHtml}
            ${customMatrixHtml}
            ${sysPrivsHtml}
            ${parsingHtml}
            <p class="source-ref" style="margin-top:6px">Source: <code>${esc(role.source.archivePath)}</code></p>
          </div>
        </details>`;
    }).join('');

    return {
      id: 'security-roles', title: 'Security Roles', count: graph.securityRoles.length,
      html: `<section id="security-roles" class="section">
        <h1>Security Roles (${graph.securityRoles.length})</h1>
        <div class="role-list">${html}</div>
      </section>`
    };
  }

  // ── Managed Plans ────────────────────────────────────────────────────────────

  private static managedPlanSection(graph: NormalizedAnalysisGraph): Section {
    const ARTIFACT_ICON: Record<string, string> = {
      PowerAutomateFlow:   '⚡',
      CopilotStudioAgent:  '🤖',
      PowerAppsModelApp:   '📱',
      PowerAppsCanvasApp:  '🎨',
      PowerBIReport:       '📊',
    };

    const plansHtml = graph.managedPlans.map(plan => {
      const planSlug = `plan-${slugify(plan.planId)}`;

      // ── Original prompt ───────────────────────────────────────────────────
      const promptHtml = plan.originalPrompt
        ? `<details class="mp-prompt-details">
            <summary class="mp-prompt-summary">Original generation prompt</summary>
            <pre class="mp-prompt-pre">${esc(plan.originalPrompt)}</pre>
           </details>`
        : '';

      // ── Personas + user stories ───────────────────────────────────────────
      const personasHtml = plan.personas.map(p => `
        <div class="mp-persona-card">
          <div class="mp-persona-name">👤 ${esc(p.name)}</div>
          ${p.description ? `<p class="mp-persona-desc">${esc(p.description)}</p>` : ''}
          <ul class="mp-story-list">
            ${p.userStories.map(s => `<li>${esc(s)}</li>`).join('')}
          </ul>
        </div>`).join('');

      // ── Artifacts / component proposals ──────────────────────────────────
      const artifactsHtml = plan.artifacts.map(a => {
        const icon = ARTIFACT_ICON[a.type] ?? '📦';
        return `
          <div class="mp-artifact-card">
            <div class="mp-artifact-header">
              <span class="mp-artifact-icon">${icon}</span>
              <span class="mp-artifact-name">${esc(a.name)}</span>
              <span class="mp-artifact-type">${esc(a.type)}</span>
            </div>
            ${a.description ? `<p class="mp-artifact-desc">${esc(a.description)}</p>` : ''}
            ${a.tables.length ? `<div class="mp-artifact-tables">Tables: ${a.tables.map(t => `<code>${esc(t)}</code>`).join(' ')}</div>` : ''}
            ${a.userStories.length ? `<ul class="mp-story-list mp-story-list-sm">${a.userStories.map(s => `<li>${esc(s)}</li>`).join('')}</ul>` : ''}
          </div>`;
      }).join('');

      // ── Entities (data model) ─────────────────────────────────────────────
      const entitiesHtml = plan.entities.length
        ? plan.entities.map(e => `
          <details class="mp-entity-details">
            <summary class="mp-entity-summary">
              <code>${esc(e.schemaName || e.displayName)}</code>
              ${e.displayName && e.displayName !== e.schemaName ? `<span class="mp-entity-display">${esc(e.displayName)}</span>` : ''}
              <span class="mp-entity-attr-count">${e.attributes.length} attributes</span>
            </summary>
            ${e.description ? `<p class="mp-entity-desc">${esc(e.description)}</p>` : ''}
            ${e.attributes.length ? `
              <table class="data-table mp-attr-table">
                <thead><tr><th>Attribute</th><th>Type</th><th>Description</th></tr></thead>
                <tbody>${e.attributes.map(a => `<tr>
                  <td><code>${esc(a.name)}</code></td>
                  <td>${esc(a.type)}</td>
                  <td class="mp-attr-desc">${esc(a.description)}</td>
                </tr>`).join('')}</tbody>
              </table>` : ''}
          </details>`).join('')
        : '<p class="muted">No entity definitions extracted.</p>';

      // ── Process diagrams ──────────────────────────────────────────────────
      const diagramsHtml = plan.processDiagrams.map(pd => `
        <details class="mp-diagram-details">
          <summary class="mp-diagram-summary">${esc(pd.name)}</summary>
          ${pd.description ? `<p class="mp-diagram-desc">${esc(pd.description)}</p>` : ''}
          <pre class="mermaid">${esc(pd.mermaid)}</pre>
        </details>`).join('');

      return `
        <div id="${planSlug}" class="mp-plan-card">
          <div class="mp-plan-header">
            <span class="mp-plan-icon">📋</span>
            <span class="mp-plan-name">${esc(plan.name)}</span>
          </div>
          <div class="mp-plan-body">
            ${plan.description ? `<p class="mp-plan-desc">${esc(plan.description)}</p>` : ''}
            ${promptHtml}

            ${plan.artifacts.length ? `<div class="mp-section-label">Planned Components (${plan.artifacts.length})</div>
            <div class="mp-artifacts">${artifactsHtml}</div>` : ''}

            ${plan.personas.length ? `<div class="mp-section-label">Personas &amp; User Stories</div>
            <div class="mp-personas">${personasHtml}</div>` : ''}

            ${plan.entities.length ? `<div class="mp-section-label">Data Model (${plan.entities.length} entities)</div>
            <div class="mp-entities">${entitiesHtml}</div>` : ''}

            ${plan.processDiagrams.length ? `<div class="mp-section-label">Process Diagrams</div>
            <div class="mp-diagrams">${diagramsHtml}</div>` : ''}

            <div class="source-ref">Source: <code>${esc(plan.source.archivePath)}</code></div>
          </div>
        </div>`;
    }).join('\n');

    const children: SectionChild[] = graph.managedPlans.map(p => ({
      id: `plan-${slugify(p.planId)}`, title: p.name,
      count: p.artifacts.length,
    }));

    return {
      id: 'managed-plans', title: 'Managed Plans', count: graph.managedPlans.length, children,
      html: `<section id="managed-plans" class="section">
        <h1>Managed Plans (${graph.managedPlans.length})</h1>
        <p class="muted" style="margin-bottom:12px">AI-generated solution blueprints from the Power Apps Copilot. These describe the intended architecture, personas, and data model for this solution.</p>
        <div class="mp-list">${plansHtml}</div>
      </section>`
    };
  }

  // ── AI Builder Models ─────────────────────────────────────────────────────────

  private static aiBuilderSection(graph: NormalizedAnalysisGraph): Section {
    const STATUS_BADGE = (status: AiBuilderModel['status']) => {
      const map: Record<string, string> = {
        active:    'aib-status-active',
        published: 'aib-status-active',
        training:  'aib-status-training',
        draft:     'aib-status-draft',
      };
      return `<span class="${map[status] ?? 'aib-status-draft'}">${esc(status)}</span>`;
    };

    const children: SectionChild[] = graph.aiBuilderModels.map(m => ({
      id: `aib-${slugify(m.modelId)}`,
      title: m.name,
    }));

    // Compute which flows use each model (cross-reference)
    const modelFlowMap = new Map<string, string[]>();
    for (const flow of graph.flows) {
      for (const action of flow.actions) {
        if (action.aiBuilderModelId) {
          if (!modelFlowMap.has(action.aiBuilderModelId)) modelFlowMap.set(action.aiBuilderModelId, []);
          modelFlowMap.get(action.aiBuilderModelId)!.push(flow.displayName);
        }
      }
    }

    const modelsHtml = graph.aiBuilderModels.map((model: AiBuilderModel) => {
      const usedIn = modelFlowMap.get(model.modelId) ?? [];
      const usedInHtml = usedIn.length
        ? usedIn.map(f => `<span class="aib-flow-ref">${esc(f)}</span>`).join(' ')
        : '<span class="muted">Not referenced in any parsed flow</span>';

      // Input variables table
      const inputsHtml = model.inputs.length
        ? `<table class="data-table aib-inputs-table">
            <thead><tr><th>Variable</th><th>Display Name</th><th>Type</th></tr></thead>
            <tbody>${model.inputs.map(i => `
              <tr><td><code>${esc(i.id)}</code></td><td>${esc(i.displayName)}</td><td>${esc(i.type)}</td></tr>`
            ).join('')}</tbody>
           </table>`
        : '<p class="muted">No input variables extracted.</p>';

      // Prompt text (collapsible, potentially very long)
      const promptHtml = model.promptText
        ? `<details class="aib-prompt-details">
            <summary class="aib-prompt-summary">System Prompt (${model.promptText.length.toLocaleString()} chars)</summary>
            <pre class="aib-prompt-pre">${esc(model.promptText)}</pre>
           </details>`
        : '';

      const modelId = `aib-${slugify(model.modelId)}`;
      return `
        <div id="${modelId}" class="aib-card">
          <div class="aib-card-header">
            <span class="aib-icon">🤖</span>
            <span class="aib-name">${esc(model.name)}</span>
            <span class="aib-template-badge">${esc(model.templateName)}</span>
            ${STATUS_BADGE(model.status)}
            ${model.modelType ? `<span class="aib-model-type">${esc(model.modelType)}</span>` : ''}
          </div>
          <div class="aib-card-body">
            <table class="data-table aib-meta-table">
              <tr><td>Model ID</td><td><code>${esc(model.modelId)}</code></td></tr>
              <tr><td>Template</td><td>${esc(model.templateName)} <span class="muted">(${esc(model.templateId)})</span></td></tr>
              <tr><td>Output Formats</td><td>${model.outputFormats.join(', ') || '—'}</td></tr>
              <tr><td>Used In Flows</td><td>${usedInHtml}</td></tr>
            </table>
            <div class="aib-section-label">Input Variables</div>
            ${inputsHtml}
            ${promptHtml}
            <div class="source-ref">Source: <code>${esc(model.source.archivePath)}</code></div>
          </div>
        </div>`;
    }).join('\n');

    return {
      id: 'ai-builder', title: 'AI Builder', count: graph.aiBuilderModels.length, children,
      html: `<section id="ai-builder" class="section">
        <h1>AI Builder Models (${graph.aiBuilderModels.length})</h1>
        <div class="aib-list">${modelsHtml}</div>
      </section>`
    };
  }

  // ── Copilot Studio Agents ────────────────────────────────────────────────────

  private static copilotAgentsSection(graph: NormalizedAnalysisGraph): Section {
    const TRIGGER_LABEL: Record<string, string> = {
      OnRecognizedIntent:    'User intent',
      OnSystemRedirect:      'System redirect',
      OnError:               'Error handler',
      OnConversationStart:   'Conversation start',
      OnEndOfConversation:   'End of conversation',
      OnUnknownIntent:       'Fallback / unknown',
    };

    const AI_BADGE = (label: string, active: boolean) =>
      `<span class="cop-ai-badge ${active ? 'cop-ai-on' : 'cop-ai-off'}">${active ? '✓' : '✗'} ${esc(label)}</span>`;

    const CHANNEL_BADGE = (ch: string) =>
      `<span class="cop-channel-badge">${esc(ch)}</span>`;

    const agentsHtml = graph.copilotAgents.map((agent: CopilotAgent) => {
      // ── Channel + auth row ──────────────────────────────────────────────────
      const channelBadges = agent.channels.length
        ? agent.channels.map(CHANNEL_BADGE).join(' ')
        : '<span class="muted">No channels configured</span>';

      // ── AI settings ────────────────────────────────────────────────────────
      const aiHtml = [
        AI_BADGE('Generative Actions', agent.aiSettings.generativeActionsEnabled),
        AI_BADGE('Model Knowledge',     agent.aiSettings.useModelKnowledge),
        AI_BADGE('File Analysis',       agent.aiSettings.fileAnalysisEnabled),
        AI_BADGE('Semantic Search',     agent.aiSettings.semanticSearchEnabled),
      ].join(' ');

      // ── Topics table ───────────────────────────────────────────────────────
      const topicRows = agent.topics.map((t: CopilotTopic) => {
        const trigLabel = TRIGGER_LABEL[t.triggerKind] ?? t.triggerKind;
        const phrases = t.triggerQueries.length
          ? t.triggerQueries.slice(0, 4).map(q => `<li>${esc(q)}</li>`).join('')
            + (t.triggerQueries.length > 4
              ? `<li class="muted">…and ${t.triggerQueries.length - 4} more</li>` : '')
          : '';
        const phrasesHtml = phrases
          ? `<ul class="cop-phrases">${phrases}</ul>`
          : '<span class="muted">—</span>';

        const actionKinds = [...new Set(t.actionKinds)];
        const kindsHtml = actionKinds.length
          ? actionKinds.map(k => `<span class="cop-kind-badge">${esc(k)}</span>`).join(' ')
          : '<span class="muted">—</span>';

        const statusBadge = t.isActive
          ? '<span class="cop-status-active">Active</span>'
          : '<span class="cop-status-inactive">Inactive</span>';

        const triggerLabel = t.triggerDisplayName
          ? `<strong>${esc(t.triggerDisplayName)}</strong><br><span class="cop-trig-kind">${esc(trigLabel)}</span>`
          : `<span class="cop-trig-kind">${esc(trigLabel)}</span>`;

        return `
          <tr>
            <td><code class="cop-topic-name">${esc(t.name || t.schemaName)}</code></td>
            <td class="cop-trig-cell">${triggerLabel}</td>
            <td class="cop-phrases-cell">${phrasesHtml}</td>
            <td class="cop-kinds-cell">${kindsHtml}</td>
            <td class="cop-status-cell">${statusBadge}</td>
          </tr>`;
      }).join('');

      const topicsHtml = topicRows
        ? `<table class="data-table cop-topics-table">
            <thead><tr>
              <th>Topic</th><th>Trigger</th><th>Sample Phrases</th><th>Action Types</th><th>Status</th>
            </tr></thead>
            <tbody>${topicRows}</tbody>
           </table>`
        : '<p class="muted">No topics extracted.</p>';

      // ── GPT instructions ────────────────────────────────────────────────────
      const gptHtml = agent.gptInstructions
        ? `<details class="cop-gpt-details">
            <summary class="cop-gpt-summary">System Prompt / GPT Instructions</summary>
            <pre class="cop-gpt-pre">${esc(agent.gptInstructions)}</pre>
           </details>`
        : '';

      // ── Knowledge files ─────────────────────────────────────────────────────
      const knowledgeHtml = agent.knowledgeFiles.length
        ? `<div class="cop-section-label">Knowledge Files (${agent.knowledgeFiles.length})</div>
           <ul class="cop-file-list">
             ${agent.knowledgeFiles.map(f => `<li><code>${esc(f)}</code></li>`).join('')}
           </ul>`
        : '';

      // ── Active vs inactive topic count ──────────────────────────────────────
      const activeCount = agent.topics.filter(t => t.isActive).length;
      const totalTopics = agent.topics.length;

      const agentId = `copilot-${slugify(agent.schemaName)}`;
      return `
        <details class="cop-card" id="${agentId}">
          <summary class="cop-summary">
            <span class="cop-name">🤖 ${esc(agent.displayName)}</span>
            <span class="cop-meta">${totalTopics} topic${totalTopics !== 1 ? 's' : ''} · ${activeCount} active</span>
            <span class="cop-auth-badge">${esc(agent.authMode)}</span>
          </summary>
          <div class="cop-body">
            <div class="cop-section-label">Channels</div>
            <div class="cop-channels">${channelBadges}</div>

            <div class="cop-section-label">AI Settings</div>
            <div class="cop-ai-settings">${aiHtml}</div>

            <div class="cop-section-label">Topics (${totalTopics})</div>
            ${topicsHtml}
            ${gptHtml}
            ${knowledgeHtml}
            <div class="source-ref">Source: <code>${esc(agent.source.archivePath)}</code></div>
          </div>
        </details>`;
    }).join('\n');

    const children: SectionChild[] = graph.copilotAgents.map(a => ({
      id: `copilot-${slugify(a.schemaName)}`,
      title: a.displayName,
      count: a.topics.length,
    }));

    return {
      id: 'copilot-agents',
      title: 'Copilot Studio Agents',
      count: graph.copilotAgents.length,
      children,
      html: `<section id="copilot-agents" class="section">
        <h1>Copilot Studio Agents (${graph.copilotAgents.length})</h1>
        <div class="cop-list">${agentsHtml}</div>
      </section>`
    };
  }

  // ── Dataverse Formulas ──────────────────────────────────────────────────────

  private static formulasSection(graph: NormalizedAnalysisGraph): Section {
    const rows = graph.dataverseFormulas.map(f => `
      <tr>
        <td><code>${esc(f.tableName)}</code></td>
        <td><code>${esc(f.columnName)}</code></td>
        <td><pre class="formula-inline">${esc(f.expression.substring(0, 200))}</pre></td>
        <td>${f.referencedFields.join(', ')}</td>
      </tr>`).join('');

    return {
      id: 'dataverse-formulas', title: 'Dataverse Formulas', count: graph.dataverseFormulas.length,
      html: `<section id="dataverse-formulas" class="section">
        <h1>Dataverse Formulas (${graph.dataverseFormulas.length})</h1>
        <table class="data-table">
          <thead><tr><th>Table</th><th>Column</th><th>Expression</th><th>Referenced Fields</th></tr></thead>
          <tbody>${rows}</tbody>
        </table></section>`
    };
  }

  // ── Web Resources ───────────────────────────────────────────────────────────

  private static webResourcesSection(graph: NormalizedAnalysisGraph): Section {
    const rows = graph.webResources.map(wr => `
      <tr>
        <td><code>${esc(wr.name)}</code></td>
        <td>${esc(wr.displayName)}</td>
        <td>${esc(wr.type)}</td>
        <td><code>${esc(wr.source.archivePath)}</code></td>
      </tr>`).join('');

    return {
      id: 'web-resources', title: 'Web Resources', count: graph.webResources.length,
      html: `<section id="web-resources" class="section">
        <h1>Web Resources (${graph.webResources.length})</h1>
        <table class="data-table">
          <thead><tr><th>Name</th><th>Display Name</th><th>Type</th><th>Source</th></tr></thead>
          <tbody>${rows}</tbody>
        </table></section>`
    };
  }

  // ── Missing Deps ─────────────────────────────────────────────────────────────

  private static missingDepsSection(graph: NormalizedAnalysisGraph): Section {
    const rows = graph.missingDeps.map(d => `
      <tr>
        <td>${esc(d.type)}</td>
        <td><code>${esc(d.identifier)}</code></td>
        <td>${esc(d.displayName ?? '—')}</td>
        <td>${esc(d.impact)}</td>
        <td><code>${esc(d.referencedBy.archivePath)}</code></td>
      </tr>`).join('');

    const checklistHtml = graph.manualChecklist.length
      ? `<h2 id="manual-checklist">Manual Completion Checklist</h2>
         <table class="data-table">
           <thead><tr><th>#</th><th>Section</th><th>Item</th><th>Priority</th></tr></thead>
           <tbody>${graph.manualChecklist.map((c, i) =>
             `<tr><td>${i + 1}</td><td>${esc(c.section)}</td><td>${esc(c.item)}</td>
              <td class="pri pri-${c.priority}">${c.priority}</td></tr>`
           ).join('')}</tbody>
         </table>` : '';

    return {
      id: 'missing-deps', title: 'Missing Dependencies',
      count: graph.missingDeps.length,
      html: `<section id="missing-deps" class="section">
        <h1>Missing Dependencies (${graph.missingDeps.length})</h1>
        ${rows ? `<table class="data-table">
          <thead><tr><th>Type</th><th>Identifier</th><th>Name</th><th>Impact</th><th>Referenced By</th></tr></thead>
          <tbody>${rows}</tbody>
        </table>` : '<p class="muted">No missing dependencies detected.</p>'}
        ${checklistHtml}
      </section>`
    };
  }

  // ── Parse Notes ─────────────────────────────────────────────────────────────

  private static parseNotesSection(graph: NormalizedAnalysisGraph): Section {
    const rows = graph.parseErrors.map(e => `
      <tr>
        <td><code>${esc(e.archivePath)}</code></td>
        <td>${esc(e.error)}</td>
        <td>${e.partial ? 'Partial' : 'Failed'}</td>
      </tr>`).join('');

    return {
      id: 'parse-notes', title: 'Export Completeness / Parse Notes',
      html: `<section id="parse-notes" class="section">
        <h1>Export Completeness / Parse Notes</h1>
        ${rows ? `<table class="data-table">
          <thead><tr><th>File</th><th>Error</th><th>Status</th></tr></thead>
          <tbody>${rows}</tbody>
        </table>` : '<p class="success">✅ All files parsed successfully.</p>'}
      </section>`
    };
  }

  // ── Page wrapper ────────────────────────────────────────────────────────────

  private static wrapPage(graph: NormalizedAnalysisGraph, nav: string, content: string): string {
    const title = graph.solution?.displayName ?? graph.meta.fileName;
    return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>${esc(title)} — PP Doc Tool</title>
<style>${CSS}</style>
<script src="https://cdn.jsdelivr.net/npm/mermaid@11/dist/mermaid.min.js"></script>
</head>
<body>
<div id="app">
  <aside id="sidebar">
    <div class="sidebar-header">
      <div class="logo">📋 PP Doc Tool</div>
      <div class="artifact-name" title="${esc(graph.meta.fileName)}">${esc(title)}</div>
      <span class="env-badge env-${graph.meta.environment.toLowerCase()}">${graph.meta.environment}</span>
    </div>
    <div class="search-box">
      <input type="search" id="search" placeholder="🔍 Search..." autocomplete="off">
    </div>
    <nav>${nav}</nav>
  </aside>
  <main id="content">
    ${content}
    <footer>Generated by PP Doc Tool v1.0 · ${new Date().toLocaleDateString()}</footer>
  </main>
</div>
<script>${JS}</script>
</body>
</html>`;
  }
}

// ── Helpers ──────────────────────────────────────────────────────────────────

interface Section {
  id: string;
  title: string;
  count?: number;
  html: string;
  children?: SectionChild[];
}

interface SectionChild {
  id: string;
  title: string;
  count?: number;
}

function esc(s: unknown): string {
  return String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

/** Escape a Power Fx formula for HTML, then highlight // comments in green. */
function fxHtml(formula: string): string {
  return esc(formula).replace(/\/\/[^\n]*/g, '<span class="fx-comment">$&</span>');
}

function slugify(s: string): string {
  return s.replace(/[^a-zA-Z0-9]/g, '-').replace(/-+/g, '-').substring(0, 40);
}

/**
 * Returns true for data sources that are worth documenting in detail:
 *   • Any SharePoint list
 *   • Custom Dataverse tables — identified by a publisher-prefix underscore
 *     in the logical name (e.g. cr8a3_requests, new_ticket).
 *     Standard system tables (account, contact, task …) have no underscore.
 */
function isRelevantSource(ds: DataSource): boolean {
  if (ds.type === 'sharepoint') return true;
  if (ds.type === 'dataverse') {
    const tbl = (ds.tableName ?? ds.logicalName ?? ds.displayName ?? '').toLowerCase();
    return tbl.includes('_');
  }
  return false;
}

// ── Embedded CSS ─────────────────────────────────────────────────────────────

const CSS = `
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
:root {
  --blue: #1F497D; --light-blue: #D5E3F0; --grey: #f5f5f5; --border: #ddd;
  --sidebar-w: 240px; --font: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
}
body { font-family: var(--font); font-size: 14px; color: #333; background: #fafafa; }
#app { display: flex; min-height: 100vh; }
#sidebar {
  width: var(--sidebar-w); min-width: var(--sidebar-w); background: #1a2a3e; color: #ccc;
  position: sticky; top: 0; height: 100vh; overflow-y: auto; display: flex; flex-direction: column;
}
.sidebar-header { padding: 16px 12px 8px; border-bottom: 1px solid #2d4060; }
.logo { font-size: 18px; font-weight: bold; color: #fff; margin-bottom: 6px; }
.artifact-name { font-size: 11px; color: #8fa8c0; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; margin-bottom: 6px; }
.search-box { padding: 8px 12px; }
#search { width: 100%; padding: 6px 10px; border-radius: 4px; border: none; font-size: 13px; background: #2d4060; color: #fff; }
#search::placeholder { color: #8fa8c0; }
nav { flex: 1; overflow-y: auto; padding: 8px 0 16px; }
.nav-list { list-style: none; }
.nav-list li { }
.nav-link {
  display: flex; justify-content: space-between; align-items: center;
  padding: 6px 14px; font-size: 13px; color: #aac; text-decoration: none;
  transition: background 0.15s;
}
.nav-link:hover, .nav-link.active { background: #2d4060; color: #fff; }
.nav-child { padding-left: 26px; font-size: 12px; color: #8fa8c0; }
.nav-count { background: #2d4060; border-radius: 10px; padding: 1px 6px; font-size: 11px; }
#content { flex: 1; padding: 24px 32px; max-width: calc(100vw - var(--sidebar-w)); overflow-x: auto; }
.section { margin-bottom: 48px; }
h1 { font-size: 22px; color: var(--blue); border-bottom: 2px solid var(--light-blue); padding-bottom: 8px; margin-bottom: 16px; margin-top: 8px; }
h2 { font-size: 17px; color: #2c3e50; margin: 20px 0 10px; }
h3 { font-size: 15px; color: #34495e; margin: 16px 0 8px; }
h4 { font-size: 14px; margin: 12px 0 6px; color: #555; }
h5 { font-size: 13px; margin: 8px 0 4px; color: #666; }
p { margin-bottom: 8px; line-height: 1.5; }
.subsection { border-left: 3px solid var(--light-blue); padding-left: 16px; margin-bottom: 28px; }
.data-table { border-collapse: collapse; width: 100%; margin-bottom: 12px; font-size: 13px; }
.data-table th { background: var(--blue); color: #fff; text-align: left; padding: 7px 10px; }
.data-table td { padding: 6px 10px; border-bottom: 1px solid var(--border); vertical-align: top; }
.data-table tr:nth-child(even) td { background: var(--grey); }
.data-table tr:hover td { background: #e8f0fb; }
code { background: #f0f0f0; padding: 1px 5px; border-radius: 3px; font-size: 12px; font-family: "Courier New", monospace; }
pre.mermaid { background: #fff; border: 1px solid var(--border); border-radius: 6px; padding: 16px; overflow-x: auto; text-align: center; }
pre.mermaid svg { max-width: 100%; height: auto; }
pre.formula-block { background: #f0f4fa; border: 1px solid #c0d0e8; border-radius: 4px; padding: 10px; overflow-x: auto; font-size: 12px; white-space: pre-wrap; word-break: break-all; }
.fx-comment { color: #3a7d44; font-style: italic; }
.formula-inline { margin: 0; font-size: 11px; }
code.formula { font-size: 11px; }
.risk-flag { display: flex; align-items: flex-start; gap: 8px; padding: 8px 12px; margin-bottom: 8px; border-radius: 4px; font-size: 13px; }
.risk-high { background: #fdecea; border: 1px solid #f5c6cb; }
.risk-medium { background: #fff3cd; border: 1px solid #ffc107; }
.risk-low { background: #d1ecf1; border: 1px solid #bee5eb; }
.risk-badge { font-weight: bold; font-size: 11px; padding: 2px 6px; border-radius: 3px; white-space: nowrap; }
.risk-high .risk-badge { background: #dc3545; color: #fff; }
.risk-medium .risk-badge { background: #ffc107; color: #333; }
.risk-low .risk-badge { background: #17a2b8; color: #fff; }
.env-badge { font-size: 11px; font-weight: bold; padding: 2px 8px; border-radius: 10px; }
.env-dev { background: #d1ecf1; color: #0c5460; }
.env-test { background: #fff3cd; color: #856404; }
.env-prod { background: #f8d7da; color: #721c24; }
.muted { color: #888; font-style: italic; }
.success { color: #155724; font-weight: bold; }
.missing { color: #dc3545; font-style: italic; }
.schema-missing { font-size: 12px; }
.source-ref { font-size: 11px; color: #888; margin-top: 8px; }
.level-error { color: #dc3545; font-weight: bold; }
.level-warning { color: #856404; }
.level-note { color: #0c5460; }
.pri { font-weight: bold; font-size: 11px; }
.pri-high { color: #dc3545; }
.pri-medium { color: #856404; }
.pri-low { color: #17a2b8; }
.num { text-align: center; font-variant-numeric: tabular-nums; }
.confidence { font-weight: normal; font-size: 12px; color: #888; }
details > summary { cursor: pointer; font-weight: bold; padding: 4px 0; color: var(--blue); }
details > summary:hover { text-decoration: underline; }
.screens-list { margin: 8px 0 16px 0; }
details.screen-detail { border: 1px solid var(--border); border-radius: 6px; margin-bottom: 8px; overflow: hidden; }
details.screen-detail > summary.screen-summary { display: flex; align-items: center; gap: 16px; padding: 10px 14px; background: #f5f7fb; font-weight: 600; color: var(--text); list-style: none; }
details.screen-detail > summary.screen-summary::-webkit-details-marker { display: none; }
details.screen-detail > summary.screen-summary::before { content: "▶"; font-size: 10px; color: var(--blue); transition: transform 0.15s; }
details.screen-detail[open] > summary.screen-summary::before { transform: rotate(90deg); }
details.screen-detail > summary.screen-summary:hover { background: #eaeffa; }
.screen-name { font-size: 14px; }
.screen-meta { font-size: 12px; font-weight: normal; color: #666; margin-left: auto; }
.screen-body { padding: 12px 16px 16px; }
.screen-body h5 { font-size: 12px; font-weight: 600; text-transform: uppercase; letter-spacing: 0.05em; color: #888; margin: 12px 0 6px; }
.control-chips { display: flex; flex-wrap: wrap; gap: 4px; margin: 6px 0 0 0; }
.control-chips .tag { background: #edf2ff; border: 1px solid #c5d3f5; border-radius: 4px; padding: 2px 6px; font-size: 11px; }
.ctrl-type { font-style: normal; color: #666; font-weight: normal; }
.ctrl-details { margin: 8px 0 12px 0; }
.ctrl-details > summary.ctrl-summary { cursor: pointer; font-weight: 600; font-size: 13px; color: #1a3a5c; padding: 5px 0; user-select: none; list-style: none; display: flex; align-items: center; gap: 6px; }
.ctrl-details > summary.ctrl-summary::before { content: '▶'; font-size: 10px; color: #888; transition: transform 0.2s; }
.ctrl-details[open] > summary.ctrl-summary::before { transform: rotate(90deg); }
/* ── Button / Control Action cards ─────────────────────────────────────── */
.ba-list { display: flex; flex-direction: column; gap: 8px; margin-bottom: 10px; }
.ba-card { border: 1px solid #d0e4f8; border-radius: 6px; background: #f9fbff; }
.ba-header { display: flex; align-items: center; gap: 8px; padding: 7px 12px; background: #e8f2fc; border-bottom: 1px solid #d0e4f8; flex-wrap: wrap; }
.ev-list { display: flex; flex-direction: column; gap: 6px; margin-top: 12px; }
.ev-card { border: 1px solid #ddd3f0; border-radius: 6px; background: #faf8ff; }
.ev-card > summary.ev-summary { cursor: pointer; list-style: none; display: flex; align-items: center; gap: 10px; padding: 8px 14px; user-select: none; }
.ev-card > summary.ev-summary::before { content: '▶'; font-size: 10px; color: #888; transition: transform 0.2s; flex-shrink: 0; }
.ev-card[open] > summary.ev-summary::before { transform: rotate(90deg); }
.ev-card[open] > summary.ev-summary { border-bottom: 1px solid #ddd3f0; background: #f0ebff; border-radius: 6px 6px 0 0; }
.ev-name { font-weight: 600; font-size: 13px; color: #2d1a5c; flex: 1; }
.ev-type-badge { font-size: 11px; background: #e6dff8; border: 1px solid #c5b8ec; border-radius: 3px; padding: 1px 7px; color: #4a2a8a; white-space: nowrap; }
.ev-body { padding: 10px 14px; }
.ev-table { margin: 0; border: none !important; }
.ev-table td { padding: 4px 10px; font-size: 12px; border-color: #e8e0f8 !important; }
.ev-table td:first-child { width: 120px; font-weight: 500; color: #555; }
/* ── Security Roles ──────────────────────────────────────────────────────────*/
.role-list { display: flex; flex-direction: column; gap: 10px; margin-top: 12px; }
/* Role card collapsible */
.role-card { border: 1px solid #c9d8f0; border-radius: 8px; background: #f8faff; }
.role-card > summary.role-summary { cursor: pointer; list-style: none; display: flex; align-items: center; gap: 10px; padding: 10px 16px; user-select: none; border-radius: 8px; }
.role-card > summary.role-summary::before { content: '▶'; font-size: 10px; color: #888; transition: transform 0.2s; flex-shrink: 0; }
.role-card[open] > summary.role-summary::before { transform: rotate(90deg); }
.role-card[open] > summary.role-summary { border-bottom: 1px solid #c9d8f0; background: #eaf2ff; border-radius: 8px 8px 0 0; }
.role-name { font-weight: 600; font-size: 13px; color: #1a3a6c; flex: 1; }
.role-meta { font-size: 11px; color: #445; background: #e4eeff; border: 1px solid #b8cfee; border-radius: 3px; padding: 2px 8px; white-space: nowrap; }
.role-risk-pill { font-size: 11px; background: #fff0f0; color: #c62828; border: 1px solid #ef9a9a; border-radius: 3px; padding: 2px 8px; white-space: nowrap; font-weight: 600; }
.role-body { padding: 14px 18px; display: flex; flex-direction: column; gap: 14px; }
/* Meta block */
.role-meta-block { display: flex; flex-direction: column; gap: 8px; }
.role-scope-dist { display: flex; flex-wrap: wrap; gap: 6px; align-items: center; padding: 2px 0; }
.role-scope-badge { display: inline-block; border-radius: 4px; padding: 2px 9px; font-size: 11px; font-weight: 600; }
.role-scope-count { font-size: 11px; color: #778; margin-right: 6px; }
/* Risk bar */
.role-risk-bar { display: flex; align-items: center; flex-wrap: wrap; gap: 6px; background: #fff8f8; border: 1px solid #ffcdd2; border-radius: 4px; padding: 6px 10px; }
.role-risk-label { font-size: 12px; font-weight: 700; color: #c62828; white-space: nowrap; }
.role-risk-chip { font-size: 11px; background: #ffebee; color: #b71c1c; border: 1px solid #ef9a9a; border-radius: 3px; padding: 2px 8px; white-space: nowrap; font-weight: 500; }
/* Section labels + blocks */
.role-section-block { display: flex; flex-direction: column; gap: 8px; }
.role-section-label { font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.06em; color: #4a5568; padding: 4px 0 4px 0; border-bottom: 1px solid #e2e8f0; }
/* Risk highlights table */
.role-risk-table td, .role-risk-table th { font-size: 12px; }
.role-flag-chip { display: inline-block; font-size: 10px; background: #fff8e1; color: #bf360c; border: 1px solid #ffe082; border-radius: 3px; padding: 1px 6px; white-space: nowrap; margin: 1px; }
/* Custom table matrix */
.role-matrix-wrap { overflow-x: auto; margin-top: 4px; }
.role-matrix { border-collapse: collapse; font-size: 12px; width: 100%; min-width: 420px; }
.role-matrix th { background: #f0f4f8; font-weight: 600; color: #2d3748; border: 1px solid #dde; padding: 6px 10px; text-align: center; }
.role-matrix td { border: 1px solid #e2e8f0; padding: 5px 8px; text-align: center; }
.mx-entity-col { min-width: 200px; text-align: left !important; }
.mx-verb-col { min-width: 72px; }
.mx-entity { text-align: left !important; font-family: monospace; font-size: 11px; max-width: 240px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; color: #1a3a6c; font-weight: 500; }
.mx-cell { font-size: 11px; font-weight: 700; }
.mx-empty { color: #d0d0d0; font-size: 13px; }
/* System verb×scope summary matrix */
.role-sys-matrix td, .role-sys-matrix th { font-size: 12px; text-align: center; }
.role-sys-matrix td:first-child, .role-sys-matrix th:first-child { text-align: left; }
/* Category collapsibles */
.role-sys-details { border: 1px solid #e2e8f0; border-radius: 5px; background: #fafbfc; }
.role-sys-summary { cursor: pointer; list-style: none; padding: 7px 12px; font-size: 12px; font-weight: 600; color: #4a5568; user-select: none; }
.role-sys-summary::before { content: '▶'; font-size: 9px; color: #a0aec0; margin-right: 7px; transition: transform 0.2s; display: inline-block; vertical-align: middle; }
.role-sys-details[open] > .role-sys-summary::before { transform: rotate(90deg); }
.role-notes-summary { color: #718096; }
.role-notes-list { margin: 6px 12px 8px 28px; font-size: 12px; color: #666; line-height: 1.6; }
.role-priv-table td, .role-priv-table th { font-size: 12px; }
.role-desc-cell { font-size: 11px; color: #555; max-width: 260px; }
/* System CRUD entity grouping */
.sys-entity-cell { white-space: nowrap; vertical-align: middle; }
.sys-actions-cell { display: flex; flex-wrap: wrap; gap: 4px; padding: 4px 0; }
.sys-action-pill { display: inline-flex; align-items: center; font-size: 11px; border: 1px solid #dde; border-radius: 4px; overflow: hidden; white-space: nowrap; background: #f8f9fb; }
.sys-action-pill > span:first-child { padding: 2px 6px; color: #2d3748; font-weight: 500; }
.sys-action-scope { padding: 2px 6px; font-weight: 600; font-size: 10px; border-left: 1px solid rgba(0,0,0,0.08); }
.gal-list { display: flex; flex-direction: column; gap: 8px; margin-bottom: 10px; }
.gal-card { border: 1px solid #c8e0c8; border-radius: 6px; background: #f6fbf6; }
.gal-header { display: flex; align-items: center; gap: 8px; padding: 7px 12px; background: #e0f0e0; border-bottom: 1px solid #c8e0c8; }
.gal-icon { font-size: 16px; }
.gal-name { font-size: 13px; font-weight: 600; color: #1a4a1a; }
.gal-type { font-size: 11px; color: #4a7a4a; background: #d0ebd0; border-radius: 3px; padding: 1px 6px; }
.gal-table { margin: 0; border: none !important; }
.gal-table td { padding: 5px 10px; font-size: 12px; border-color: #d8ead8 !important; }
.gal-table td:first-child { width: 110px; font-weight: 500; color: #555; }
.gal-ds-chip { background: #d0ebd0; border: 1px solid #a0c8a0; border-radius: 3px; padding: 1px 7px; font-size: 11px; font-weight: 600; color: #1a4a1a; margin-right: 4px; }
.ba-ctrl-name { font-size: 12px; font-weight: bold; color: #1a3a5c; }
.ba-ctrl-type { font-size: 11px; color: #555; background: #dce8f5; border-radius: 3px; padding: 1px 6px; }
.ba-prop { font-size: 11px; color: #777; margin-left: auto; font-style: italic; }
.ba-chips { display: flex; flex-wrap: wrap; gap: 6px; padding: 8px 12px; }
.ba-chip { display: inline-flex; align-items: baseline; gap: 4px; padding: 4px 8px; border-radius: 4px; font-size: 12px; line-height: 1.4; border: 1px solid transparent; }
.ba-chip strong { font-weight: 600; }
.ba-chip code { font-size: 11px; }
.ba-payload-hint { font-size: 10px; color: #666; font-style: italic; max-width: 250px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; display: inline-block; vertical-align: middle; }
.ba-kind-patch       { background: #e8f5e9; border-color: #a5d6a7; color: #1b5e20; }
.ba-kind-flow-run    { background: #fff3e0; border-color: #ffcc80; color: #7a3f00; }
.ba-kind-navigate    { background: #e3f2fd; border-color: #90caf9; color: #0d47a1; }
.ba-kind-submit-form { background: #f3e5f5; border-color: #ce93d8; color: #4a148c; }
.ba-kind-collect     { background: #fce4ec; border-color: #f48fb1; color: #880e4f; }
.ba-kind-remove      { background: #fdecea; border-color: #ef9a9a; color: #b71c1c; }
.ba-kind-other       { background: #f5f5f5; border-color: #bdbdbd; color: #333; }
.ba-formula { border-top: 1px solid #d0e4f8; }
.ba-formula > summary { padding: 4px 12px; font-size: 11px; color: #888; cursor: pointer; }
.ba-formula > summary:hover { color: var(--blue); }
.ba-formula .formula-block { margin: 0; border: none; border-radius: 0; font-size: 11px; background: #f0f4fa; max-height: 400px; overflow-y: auto; }
.ds-scope-note { font-size: 12px; margin-bottom: 16px; }
.ds-card { border: 1px solid var(--border); border-radius: 6px; padding: 12px 16px; margin-bottom: 16px; }
/* Collections collapsible */
.collections-detail { margin: 8px 0 16px 0; border: 1px solid #c8e6c9; border-radius: 6px; overflow: hidden; }
.collections-detail > summary { padding: 8px 14px; background: #f1f8f2; cursor: pointer; font-weight: 600; font-size: 14px; color: #1b5e20; list-style: none; }
.collections-detail > summary::-webkit-details-marker { display: none; }
.collections-detail > summary::before { content: "▶"; font-size: 10px; margin-right: 6px; transition: transform 0.15s; }
.collections-detail[open] > summary::before { transform: rotate(90deg); }
.collections-detail > summary:hover { background: #e8f5e9; }
.collections-detail .data-table { margin: 0; }
/* ── Managed Plans ───────────────────────────────────────────────────────── */
.mp-list { display: flex; flex-direction: column; gap: 16px; }
.mp-plan-card { border: 1px solid #c8e0d0; border-radius: 8px; background: #f7fbf9; overflow: hidden; }
.mp-plan-header { display: flex; align-items: center; gap: 10px; padding: 10px 16px; background: #e0f4ea; border-bottom: 1px solid #c8e0d0; }
.mp-plan-icon { font-size: 18px; }
.mp-plan-name { font-weight: 700; font-size: 14px; color: #0d4a2e; flex: 1; }
.mp-plan-body { padding: 14px 18px; display: flex; flex-direction: column; gap: 14px; }
.mp-plan-desc { font-size: 13px; color: #333; line-height: 1.5; }
.mp-section-label { font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.06em; color: #4a5568; border-bottom: 1px solid #c8e0d0; padding-bottom: 3px; }
/* Original prompt */
.mp-prompt-details { border: 1px solid #c8e0d0; border-radius: 5px; }
.mp-prompt-summary { cursor: pointer; list-style: none; padding: 7px 12px; font-size: 12px; font-weight: 600; color: #0d4a2e; user-select: none; }
.mp-prompt-summary::before { content: '▶'; font-size: 9px; color: #888; margin-right: 7px; transition: transform 0.2s; display: inline-block; vertical-align: middle; }
.mp-prompt-details[open] > .mp-prompt-summary::before { transform: rotate(90deg); }
.mp-prompt-pre { padding: 10px 14px; font-size: 12px; font-family: "Courier New", monospace; white-space: pre-wrap; word-break: break-word; max-height: 300px; overflow-y: auto; color: #222; line-height: 1.5; border-top: 1px solid #c8e0d0; }
/* Artifacts */
.mp-artifacts { display: flex; flex-wrap: wrap; gap: 10px; }
.mp-artifact-card { border: 1px solid #d4e8d4; border-radius: 6px; background: #f4fbf4; padding: 10px 14px; min-width: 200px; flex: 1; }
.mp-artifact-header { display: flex; align-items: center; gap: 6px; margin-bottom: 6px; }
.mp-artifact-icon { font-size: 16px; }
.mp-artifact-name { font-weight: 600; font-size: 13px; color: #1a4a1a; flex: 1; }
.mp-artifact-type { font-size: 10px; background: #d4e8d4; border-radius: 3px; padding: 1px 6px; color: #1a4a1a; white-space: nowrap; }
.mp-artifact-desc { font-size: 12px; color: #555; line-height: 1.4; margin-bottom: 6px; }
.mp-artifact-tables { font-size: 11px; color: #666; margin-bottom: 4px; }
.mp-artifact-tables code { font-size: 10px; }
/* Personas */
.mp-personas { display: flex; flex-wrap: wrap; gap: 10px; }
.mp-persona-card { border: 1px solid #dce8f0; border-radius: 6px; background: #f4f9fc; padding: 10px 14px; min-width: 200px; flex: 1; }
.mp-persona-name { font-weight: 600; font-size: 13px; color: #0d3a6e; margin-bottom: 4px; }
.mp-persona-desc { font-size: 12px; color: #555; margin-bottom: 6px; }
.mp-story-list { padding-left: 16px; font-size: 12px; line-height: 1.7; color: #333; }
.mp-story-list-sm { font-size: 11px; line-height: 1.6; }
/* Entities */
.mp-entities { display: flex; flex-direction: column; gap: 4px; }
.mp-entity-details { border: 1px solid #d4e8d4; border-radius: 5px; background: #fafcfa; }
.mp-entity-summary { cursor: pointer; list-style: none; display: flex; align-items: center; gap: 8px; padding: 7px 12px; user-select: none; }
.mp-entity-summary::before { content: '▶'; font-size: 9px; color: #888; transition: transform 0.2s; margin-right: 4px; flex-shrink: 0; }
.mp-entity-details[open] > .mp-entity-summary::before { transform: rotate(90deg); }
.mp-entity-display { font-size: 12px; color: #555; }
.mp-entity-attr-count { font-size: 11px; color: #888; margin-left: auto; white-space: nowrap; }
.mp-entity-desc { font-size: 12px; color: #555; padding: 4px 12px 0; }
.mp-attr-table { margin: 4px 0 0; }
.mp-attr-table td, .mp-attr-table th { font-size: 11px; }
.mp-attr-desc { color: #666; }
/* Process diagrams */
.mp-diagrams { display: flex; flex-direction: column; gap: 6px; }
.mp-diagram-details { border: 1px solid #c8e0d0; border-radius: 5px; }
.mp-diagram-summary { cursor: pointer; list-style: none; padding: 7px 12px; font-size: 12px; font-weight: 600; color: #0d4a2e; user-select: none; }
.mp-diagram-summary::before { content: '▶'; font-size: 9px; color: #888; margin-right: 7px; transition: transform 0.2s; display: inline-block; vertical-align: middle; }
.mp-diagram-details[open] > .mp-diagram-summary::before { transform: rotate(90deg); }
.mp-diagram-desc { font-size: 12px; color: #555; padding: 4px 12px 0; line-height: 1.4; }
/* ── AI Builder ──────────────────────────────────────────────────────────── */
.aib-list { display: flex; flex-direction: column; gap: 16px; margin-top: 12px; }
.aib-card { border: 1px solid #d0c8f8; border-radius: 8px; background: #fcfaff; overflow: hidden; }
.aib-card-header { display: flex; align-items: center; gap: 10px; padding: 10px 16px; background: #ede8ff; border-bottom: 1px solid #d0c8f8; flex-wrap: wrap; }
.aib-icon { font-size: 18px; }
.aib-name { font-weight: 700; font-size: 14px; color: #2c1a5e; flex: 1; }
.aib-template-badge { font-size: 11px; background: #d8d0f8; border: 1px solid #b8a8f0; border-radius: 3px; padding: 2px 8px; color: #2c1a5e; white-space: nowrap; font-weight: 600; }
.aib-model-type { font-size: 11px; background: #e8e0ff; border: 1px solid #c5b8ec; border-radius: 3px; padding: 2px 8px; color: #4a2a8a; white-space: nowrap; font-style: italic; }
.aib-status-active   { font-size: 11px; background: #e6f4ea; border: 1px solid #8ac898; border-radius: 10px; padding: 2px 8px; color: #1b5e20; font-weight: 600; }
.aib-status-training { font-size: 11px; background: #fff8e1; border: 1px solid #ffe082; border-radius: 10px; padding: 2px 8px; color: #7a5300; font-weight: 600; }
.aib-status-draft    { font-size: 11px; background: #f5f5f5; border: 1px solid #ccc; border-radius: 10px; padding: 2px 8px; color: #888; font-weight: 600; }
.aib-card-body { padding: 14px 18px; display: flex; flex-direction: column; gap: 12px; }
.aib-section-label { font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.06em; color: #4a5568; border-bottom: 1px solid #e2e8f0; padding-bottom: 3px; }
.aib-meta-table td:first-child { width: 130px; font-weight: 500; color: #555; font-size: 12px; }
.aib-meta-table td { font-size: 12px; }
.aib-inputs-table { margin: 0; }
.aib-flow-ref { display: inline-block; font-size: 11px; background: #ede8ff; border: 1px solid #c5b8ec; border-radius: 4px; padding: 2px 8px; margin: 1px; color: #2c1a5e; font-weight: 500; }
/* Prompt collapsible */
.aib-prompt-details { border: 1px solid #d0c8f8; border-radius: 5px; background: #f9f7ff; }
.aib-prompt-summary { cursor: pointer; list-style: none; padding: 7px 12px; font-size: 12px; font-weight: 600; color: #2c1a5e; user-select: none; }
.aib-prompt-summary::before { content: '▶'; font-size: 9px; color: #888; margin-right: 7px; transition: transform 0.2s; display: inline-block; vertical-align: middle; }
.aib-prompt-details[open] > .aib-prompt-summary::before { transform: rotate(90deg); }
.aib-prompt-pre { padding: 10px 14px; font-size: 11px; font-family: "Courier New", monospace; white-space: pre-wrap; word-break: break-word; max-height: 500px; overflow-y: auto; color: #222; line-height: 1.5; border-top: 1px solid #d0c8f8; }
/* AI Builder action row highlight in flows table */
.aib-action-row td { background: #f5f0ff !important; }
.aib-link { font-size: 11px; color: #4a2a8a; text-decoration: none; font-weight: 600; }
.aib-link:hover { text-decoration: underline; }
.aib-link-unresolved { font-size: 11px; color: #888; font-style: italic; }
.aib-pretrained-badge { font-size: 11px; color: #4a2a8a; font-style: normal; }
.aib-pretrained-badge em { color: #999; font-size: 10px; }
/* ── Copilot Studio Agents ───────────────────────────────────────────────── */
.cop-list { display: flex; flex-direction: column; gap: 10px; margin-top: 12px; }
.cop-card { border: 1px solid #b8d4f0; border-radius: 8px; background: #f7fbff; }
.cop-summary { cursor: pointer; list-style: none; display: flex; align-items: center; gap: 10px; padding: 10px 16px; user-select: none; border-radius: 8px; }
.cop-summary::before { content: '▶'; font-size: 10px; color: #888; transition: transform 0.2s; flex-shrink: 0; }
.cop-card[open] > .cop-summary::before { transform: rotate(90deg); }
.cop-card[open] > .cop-summary { border-bottom: 1px solid #b8d4f0; background: #e6f2ff; border-radius: 8px 8px 0 0; }
.cop-name { font-weight: 600; font-size: 13px; color: #0d3a6e; flex: 1; }
.cop-meta { font-size: 11px; color: #556; background: #daeaff; border: 1px solid #a8ccee; border-radius: 3px; padding: 2px 8px; white-space: nowrap; }
.cop-auth-badge { font-size: 11px; background: #e8f0fe; border: 1px solid #aac4f8; border-radius: 3px; padding: 2px 8px; white-space: nowrap; color: #1a3a8c; }
.cop-body { padding: 14px 18px; display: flex; flex-direction: column; gap: 12px; }
.cop-section-label { font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.06em; color: #4a5568; border-bottom: 1px solid #e2e8f0; padding-bottom: 3px; }
/* Channels */
.cop-channels { display: flex; flex-wrap: wrap; gap: 6px; }
.cop-channel-badge { font-size: 12px; background: #e0f0e0; border: 1px solid #a0cca0; border-radius: 4px; padding: 2px 10px; color: #1a4a1a; font-weight: 500; }
/* AI Settings */
.cop-ai-settings { display: flex; flex-wrap: wrap; gap: 6px; }
.cop-ai-badge { font-size: 11px; border-radius: 4px; padding: 3px 9px; font-weight: 500; white-space: nowrap; }
.cop-ai-on  { background: #e6f4ea; border: 1px solid #8ac898; color: #1b5e20; }
.cop-ai-off { background: #f5f5f5; border: 1px solid #ccc; color: #888; }
/* Topics table */
.cop-topics-table td { vertical-align: top; }
.cop-topic-name { font-size: 11px; }
.cop-trig-cell { min-width: 140px; font-size: 12px; }
.cop-trig-kind { font-size: 11px; color: #666; }
.cop-phrases-cell { min-width: 180px; }
.cop-phrases { margin: 0; padding-left: 16px; font-size: 12px; }
.cop-kinds-cell { min-width: 130px; }
.cop-kind-badge { display: inline-block; font-size: 10px; background: #e8f0fe; border: 1px solid #aac4f8; border-radius: 3px; padding: 1px 6px; margin: 1px; color: #1a3a8c; white-space: nowrap; }
.cop-status-cell { text-align: center; white-space: nowrap; }
.cop-status-active   { font-size: 11px; background: #e6f4ea; border: 1px solid #8ac898; border-radius: 10px; padding: 2px 8px; color: #1b5e20; font-weight: 600; }
.cop-status-inactive { font-size: 11px; background: #f5f5f5; border: 1px solid #ccc; border-radius: 10px; padding: 2px 8px; color: #888; font-weight: 600; }
/* GPT instructions */
.cop-gpt-details { border: 1px solid #d0e8ff; border-radius: 5px; background: #f7fbff; }
.cop-gpt-summary { cursor: pointer; list-style: none; padding: 7px 12px; font-size: 12px; font-weight: 600; color: #1a3a8c; user-select: none; }
.cop-gpt-summary::before { content: '▶'; font-size: 9px; color: #888; margin-right: 7px; transition: transform 0.2s; display: inline-block; vertical-align: middle; }
.cop-gpt-details[open] > .cop-gpt-summary::before { transform: rotate(90deg); }
.cop-gpt-pre { padding: 10px 14px; font-size: 12px; font-family: "Courier New", monospace; white-space: pre-wrap; word-break: break-word; max-height: 400px; overflow-y: auto; color: #222; line-height: 1.5; }
/* Knowledge files */
.cop-file-list { margin: 4px 0 0 16px; font-size: 12px; line-height: 1.8; }
footer { margin-top: 48px; padding: 16px 0; border-top: 1px solid var(--border); font-size: 12px; color: #888; text-align: center; }
/* ── AppChecker findings ────────────────────────────────────────────────── */
.ck-section { border: 1px solid #e0d0ff; border-radius: 6px; margin: 12px 0 20px 0; overflow: hidden; }
.ck-summary { display: flex; align-items: center; gap: 8px; padding: 10px 14px; background: #f5f0ff; cursor: pointer; font-weight: 600; color: #4a2c8a; list-style: none; }
.ck-summary::-webkit-details-marker { display: none; }
.ck-summary::before { content: "▶"; font-size: 10px; transition: transform 0.15s; }
.ck-section[open] > .ck-summary::before { transform: rotate(90deg); }
.ck-summary:hover { background: #ede8ff; }
.ck-total { color: #888; font-weight: normal; font-size: 13px; }
.ck-badges { display: flex; gap: 4px; margin-left: auto; }
.ck-badge { font-size: 11px; font-weight: bold; padding: 2px 7px; border-radius: 10px; white-space: nowrap; }
.ck-badge-error { background: #fdecea; color: #b71c1c; border: 1px solid #f5c6cb; }
.ck-badge-warn  { background: #fff8e1; color: #7a5300; border: 1px solid #ffe082; }
.ck-body { padding: 12px 14px 14px; }
.ck-table { margin-bottom: 8px; }
.loc-code { font-size: 11px; word-break: break-all; }
/* Row-level tints — error = red, warning = yellow, note = no tint */
tr.finding-error td   { background: #fdecea !important; }
tr.finding-warning td { background: #fff8e1 !important; }
.level-cell { font-weight: bold; white-space: nowrap; }
/* Per-rule collapsible groups */
.ck-rule-group { border: 1px solid #e8e0f5; border-radius: 4px; margin-bottom: 5px; overflow: hidden; }
.ck-rule-group > summary { display: flex; align-items: center; gap: 8px; padding: 7px 12px; background: #faf8ff; cursor: pointer; list-style: none; }
.ck-rule-group > summary::-webkit-details-marker { display: none; }
.ck-rule-group > summary::before { content: "▶"; font-size: 9px; color: #7a5cb8; transition: transform 0.15s; flex-shrink: 0; }
.ck-rule-group[open] > summary::before { transform: rotate(90deg); }
.ck-rule-group > summary:hover { background: #f0e8ff; }
.ck-rule-id { font-size: 12px; font-family: "Courier New", monospace; color: #2c1a5e; }
.ck-rule-count { color: #777; font-size: 12px; font-weight: normal; }
.ck-badge-note { background: #e8f5e9; color: #1b5e20; border: 1px solid #c8e6c9; }
.ck-rule-group .data-table { margin: 0; border-top: 1px solid #e8e0f5; }
.search-hidden { display: none !important; }
@media print {
  #sidebar, #search, footer { display: none; }
  #content { max-width: 100%; }
  details { display: block; }
  details > summary::after { content: " (expanded)"; }
  h1 { page-break-before: always; }
}
`;

// ── Embedded JS ──────────────────────────────────────────────────────────────

const JS = `
// Mermaid rendering
document.addEventListener('DOMContentLoaded', function () {
  if (typeof mermaid !== 'undefined') {
    mermaid.initialize({ startOnLoad: false, theme: 'default', securityLevel: 'loose' });
    mermaid.run({ querySelector: 'pre.mermaid' });
  }
});

// Scrollspy
const observer = new IntersectionObserver(entries => {
  for (const entry of entries) {
    if (entry.isIntersecting) {
      document.querySelectorAll('.nav-link').forEach(l => l.classList.remove('active'));
      const id = entry.target.id;
      const link = document.querySelector(\`.nav-link[href="#\${id}"]\`);
      if (link) link.classList.add('active');
    }
  }
}, { threshold: 0.15 });
document.querySelectorAll('section[id], div[id]').forEach(el => observer.observe(el));

// Pre-expand all details before print
window.addEventListener('beforeprint', () => {
  document.querySelectorAll('details').forEach(d => d.open = true);
});

// Search
const searchInput = document.getElementById('search');
if (searchInput) {
  searchInput.addEventListener('input', function() {
    const q = this.value.toLowerCase().trim();
    if (!q) {
      document.querySelectorAll('[data-searchable]').forEach(el => el.classList.remove('search-hidden'));
      return;
    }
    document.querySelectorAll('[data-searchable]').forEach(el => {
      const text = el.textContent.toLowerCase();
      el.classList.toggle('search-hidden', !text.includes(q));
    });
  });
}
`;
