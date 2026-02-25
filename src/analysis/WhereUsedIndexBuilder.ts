import type { NormalizedAnalysisGraph, WhereUsedIndex } from '../model/types';

export class WhereUsedIndexBuilder {
  static build(graph: NormalizedAnalysisGraph): WhereUsedIndex {
    const idx: WhereUsedIndex = {
      connectorToFlows: new Map(),
      connectorToApps: new Map(),
      envVarToFlows: new Map(),
      envVarToApps: new Map(),
      dataSourceToFlows: new Map(),
      dataSourceToApps: new Map(),
      variableToScreens: new Map()
    };

    // Connectors used in flows
    for (const flow of graph.flows) {
      for (const conn of flow.connectors) {
        const id = conn.connectorId;
        if (!idx.connectorToFlows.has(id)) idx.connectorToFlows.set(id, []);
        if (!idx.connectorToFlows.get(id)!.includes(flow.id))
          idx.connectorToFlows.get(id)!.push(flow.id);
      }
    }

    // Connectors used in apps
    for (const app of graph.canvasApps) {
      for (const conn of app.connectors) {
        const id = conn.connectorId;
        if (!idx.connectorToApps.has(id)) idx.connectorToApps.set(id, []);
        if (!idx.connectorToApps.get(id)!.includes(app.id))
          idx.connectorToApps.get(id)!.push(app.id);
      }
    }

    // Env vars in flows (scan action inputs for schema name patterns)
    for (const flow of graph.flows) {
      for (const ev of graph.envVars) {
        const pattern = ev.schemaName;
        for (const action of flow.actions) {
          const inputStr = JSON.stringify(action.inputs);
          if (inputStr.includes(pattern)) {
            if (!idx.envVarToFlows.has(ev.id)) idx.envVarToFlows.set(ev.id, []);
            if (!idx.envVarToFlows.get(ev.id)!.includes(flow.id))
              idx.envVarToFlows.get(ev.id)!.push(flow.id);
            ev.usedBy.push(action.source);
          }
        }
      }
    }

    // Env vars in apps (scan formula strings)
    for (const app of graph.canvasApps) {
      for (const ev of graph.envVars) {
        const pattern = ev.schemaName;
        let found = false;
        const scanFormula = (f: string) => { if (f.includes(pattern)) found = true; };
        scanFormula(app.appOnStart.raw);
        for (const s of app.screens) {
          scanFormula(s.onVisible.raw);
          for (const kf of s.keyFormulas) scanFormula(kf.formula.raw);
        }
        if (found) {
          if (!idx.envVarToApps.has(ev.id)) idx.envVarToApps.set(ev.id, []);
          if (!idx.envVarToApps.get(ev.id)!.includes(app.id))
            idx.envVarToApps.get(ev.id)!.push(app.id);
        }
      }
    }

    // Data sources in flows
    for (const flow of graph.flows) {
      for (const ds of flow.dataSources) {
        if (!idx.dataSourceToFlows.has(ds.id)) idx.dataSourceToFlows.set(ds.id, []);
        if (!idx.dataSourceToFlows.get(ds.id)!.includes(flow.id))
          idx.dataSourceToFlows.get(ds.id)!.push(flow.id);
      }
    }

    // Data sources in apps
    for (const app of graph.canvasApps) {
      for (const ds of app.dataSources) {
        if (!idx.dataSourceToApps.has(ds.id)) idx.dataSourceToApps.set(ds.id, []);
        if (!idx.dataSourceToApps.get(ds.id)!.includes(app.id))
          idx.dataSourceToApps.get(ds.id)!.push(app.id);
      }
    }

    // Variables to screens
    for (const app of graph.canvasApps) {
      for (const v of app.variables) {
        const screens = new Set<string>();
        for (const ref of [...v.setAt, ...v.usedAt]) {
          const screen = ref.yamlPath?.split('.')?.[0] ?? ref.archivePath;
          if (screen) screens.add(screen);
        }
        if (!idx.variableToScreens.has(v.name)) idx.variableToScreens.set(v.name, []);
        idx.variableToScreens.get(v.name)!.push(...screens);
      }
    }

    graph.whereUsed = idx;
    return idx;
  }
}
