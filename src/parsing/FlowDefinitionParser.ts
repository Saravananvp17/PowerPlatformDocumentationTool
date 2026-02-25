import type {
  Flow, FlowAction, TriggerDetail, TriggerType, ConnectorRef,
  DataSource, UrlReference, ErrorHandlingSummary, RetryPolicy,
  SourceRef, ProgressCallback
} from '../model/types';
import { UrlClassifier } from '../analysis/UrlClassifier';

// ─────────────────────────────────────────────────────────────────────────────

interface RawWorkflow {
  id?: string;
  name?: string;
  properties?: {
    displayName?: string;
    description?: string;
    state?: string;
    definition?: RawDefinition;
  };
  triggers?: Record<string, unknown>;
  actions?: Record<string, unknown>;
  definition?: RawDefinition;
}

interface RawDefinition {
  triggers?: Record<string, RawTrigger>;
  actions?: Record<string, RawAction>;
  contentVersion?: string;
}

interface RawTrigger {
  type?: string;
  recurrence?: { frequency?: string; interval?: number; timeZone?: string; startTime?: string };
  inputs?: Record<string, unknown>;
  metadata?: Record<string, unknown>;
}

interface RawAction {
  type?: string;
  inputs?: Record<string, unknown>;
  actions?: Record<string, RawAction>;
  else?: { actions?: Record<string, RawAction> };
  cases?: Record<string, { actions?: Record<string, RawAction> }>;
  runAfter?: Record<string, string[]>;
  metadata?: { operationMetadataId?: string };
  description?: string;
  retryPolicy?: { type?: string; count?: number; interval?: string };
  runtimeConfiguration?: { secureData?: { properties?: string[] }; concurrency?: unknown; paginationPolicy?: unknown };
}

// ─────────────────────────────────────────────────────────────────────────────

export class FlowDefinitionParser {
  static parse(raw: RawWorkflow, archivePath: string, _onProgress?: ProgressCallback): Flow {
    const src: SourceRef = { archivePath, confidence: 'high' };

    // Normalise: some exports wrap definition under properties.definition
    const definition: RawDefinition =
      raw.definition ??
      raw.properties?.definition ??
      { triggers: raw.triggers as Record<string, RawTrigger>, actions: raw.actions as Record<string, RawAction> };

    const rawTriggers = definition.triggers ?? {};
    const rawActions = definition.actions ?? {};

    const trigger = FlowDefinitionParser.parseTrigger(rawTriggers, archivePath);
    const actions = FlowDefinitionParser.walkActions(rawActions, archivePath);
    const connectors = FlowDefinitionParser.extractConnectors(trigger, actions, archivePath);
    const dataSources = FlowDefinitionParser.extractDataSources(actions, archivePath);
    const urls = FlowDefinitionParser.extractUrls(actions, archivePath);
    const errorHandling = FlowDefinitionParser.analyseErrorHandling(actions);
    const mermaidDiagram = FlowDefinitionParser.buildMermaid(trigger, actions);

    return {
      id: raw.id ?? archivePath,
      name: archivePath,
      displayName: raw.properties?.displayName
        ?? (raw.name && !FlowDefinitionParser.isGuid(String(raw.name)) ? FlowDefinitionParser.stripGuidSuffix(String(raw.name)) : undefined)
        ?? FlowDefinitionParser.fileNameToDisplayName(archivePath),
      description: raw.properties?.description,
      state: raw.properties?.state ?? 'Unknown',
      triggerType: trigger.type,
      trigger,
      actions,
      mermaidDiagram,
      connectors,
      dataSources,
      urls,
      errorHandling,
      source: src
    };
  }

  // ── Trigger parsing ─────────────────────────────────────────────────────────

  private static parseTrigger(
    rawTriggers: Record<string, RawTrigger>,
    archivePath: string
  ): TriggerDetail {
    const src: SourceRef = { archivePath, jsonPath: '$.triggers', confidence: 'high' };
    const entries = Object.entries(rawTriggers);
    if (entries.length === 0) {
      return { type: 'other', raw: {}, source: src };
    }
    const [, trig] = entries[0];
    const type = FlowDefinitionParser.classifyTrigger(trig);
    return {
      type,
      recurrence: trig.recurrence ? {
        frequency: String(trig.recurrence.frequency ?? ''),
        interval: Number(trig.recurrence.interval ?? 1),
        timeZone: trig.recurrence.timeZone,
        startTime: trig.recurrence.startTime
      } : undefined,
      automated: (type === 'automated') ? {
        connectorId: FlowDefinitionParser.extractConnectorId(trig),
        operationId: FlowDefinitionParser.extractOperationId(trig),
      } : undefined,
      raw: trig as Record<string, unknown>,
      source: src
    };
  }

  private static classifyTrigger(t: RawTrigger): TriggerType {
    const tp = (t.type ?? '').toLowerCase();
    if (tp === 'recurrence') return 'recurrence';
    if (tp === 'request') return t.inputs ? 'http' : 'manual';
    if (tp === 'powerapps') return 'powerApps';
    if (tp === 'workflow') return 'child';
    if (tp === 'apiconnection' || tp === 'openapicall') return 'automated';
    return 'other';
  }

  // ── Action walking ──────────────────────────────────────────────────────────

  static walkActions(
    actions: Record<string, RawAction>,
    archivePath: string,
    scopeName?: string,
    depth = 0
  ): FlowAction[] {
    if (depth > 20) return []; // guard infinite recursion
    return Object.entries(actions).flatMap(([key, def]) => {
      const self = FlowDefinitionParser.extractAction(key, def, archivePath, scopeName);
      const children: FlowAction[] = [];
      if (def.actions) children.push(...FlowDefinitionParser.walkActions(def.actions, archivePath, key, depth + 1));
      if (def.else?.actions) children.push(...FlowDefinitionParser.walkActions(def.else.actions, archivePath, key, depth + 1));
      if (def.cases) {
        for (const c of Object.values(def.cases)) {
          if (c.actions) children.push(...FlowDefinitionParser.walkActions(c.actions, archivePath, key, depth + 1));
        }
      }
      return [self, ...children];
    });
  }

  /**
   * Returns the AI Builder model identifier for this action if it calls AI Builder,
   * or undefined otherwise.
   *
   * • Custom models: returns the model GUID from parameters.recordId
   * • Pretrained models: returns the operationId (e.g. "aibuilderpredict_textrecognition")
   *   as a sentinel — these have no recordId but still deserve special treatment.
   */
  private static extractAiBuilderModelId(def: RawAction): string | undefined {
    const inp = def.inputs as Record<string, unknown> | undefined;
    if (!inp) return undefined;
    const host = inp.host as Record<string, unknown> | undefined;
    if (!host) return undefined;
    const apiId = String(host.apiId ?? '').toLowerCase();
    const opId  = String(host.operationId ?? '');
    // AI Builder calls use apiId containing 'aibuilder' or operationId starting with 'aibuilder'
    if (!apiId.includes('aibuilder') && !opId.toLowerCase().startsWith('aibuilder')) return undefined;
    // Custom models have a recordId GUID; pretrained models do not
    const params = inp.parameters as Record<string, unknown> | undefined;
    const recordId = String(params?.recordId ?? '');
    // Return the recordId for custom models, operationId for pretrained (as sentinel)
    return recordId || opId || undefined;
  }

  private static extractAction(
    key: string, def: RawAction, archivePath: string, parentScope?: string
  ): FlowAction {
    const src: SourceRef = { archivePath, jsonPath: `$.actions.${key}`, confidence: 'high' };
    const secureInputs = (def.runtimeConfiguration?.secureData?.properties ?? []).includes('inputs');
    const secureOutputs = (def.runtimeConfiguration?.secureData?.properties ?? []).includes('outputs');

    return {
      id: key,
      displayName: (def as unknown as Record<string, string>).description ?? key,
      type: def.type ?? 'Unknown',
      connector: FlowDefinitionParser.extractConnectorId(def as unknown as RawTrigger),
      operationId: FlowDefinitionParser.extractOperationId(def as unknown as RawTrigger),
      aiBuilderModelId: FlowDefinitionParser.extractAiBuilderModelId(def),
      inputs: (def.inputs ?? {}) as Record<string, unknown>,
      redactedInputs: {},   // filled by SecretRedactor post-pass
      runAfter: Object.keys(def.runAfter ?? {}),
      parentScope: parentScope,
      retryPolicy: def.retryPolicy ? { type: def.retryPolicy.type ?? 'none', count: def.retryPolicy.count, interval: def.retryPolicy.interval } : undefined,
      secureInputs,
      secureOutputs,
      source: src
    };
  }

  // ── Connector extraction ────────────────────────────────────────────────────

  private static extractConnectorId(t: RawTrigger): string {
    if (!t.inputs) return '';
    const inp = t.inputs as Record<string, unknown>;
    if (typeof inp.host === 'object' && inp.host !== null) {
      return String((inp.host as Record<string, unknown>).connection ?? (inp.host as Record<string, unknown>).connectionName ?? '');
    }
    return '';
  }

  private static extractOperationId(t: RawTrigger): string {
    if (!t.inputs) return '';
    const inp = t.inputs as Record<string, unknown>;
    if (typeof inp.host === 'object' && inp.host !== null) {
      return String((inp.host as Record<string, unknown>).operationId ?? (inp.host as Record<string, unknown>).apiId ?? '');
    }
    return '';
  }

  private static extractConnectors(trigger: TriggerDetail, actions: FlowAction[], archivePath: string): ConnectorRef[] {
    const seen = new Map<string, ConnectorRef>();
    const src: SourceRef = { archivePath, confidence: 'medium' };

    const addConnector = (id: string, displayName: string) => {
      if (!id || id === '') return;
      if (!seen.has(id)) {
        seen.set(id, { connectorId: id, displayName, usedInFlows: [], usedInApps: [], source: src });
      }
    };

    if (trigger.automated?.connectorId) addConnector(trigger.automated.connectorId, trigger.automated.connectorId);
    for (const action of actions) {
      if (action.connector) addConnector(action.connector, action.connector);
    }
    return [...seen.values()];
  }

  // ── Data Source extraction ──────────────────────────────────────────────────

  private static extractDataSources(actions: FlowAction[], archivePath: string): DataSource[] {
    const seen = new Map<string, DataSource>();
    for (const action of actions) {
      const inp = action.inputs as Record<string, unknown>;
      if (!inp) continue;

      // SharePoint pattern: inputs.parameters.dataset (site) + inputs.parameters.table (list)
      const params = inp.parameters as Record<string, unknown> | undefined;
      if (params) {
        const dataset = params.dataset as string | undefined;
        const table = params.table as string | undefined;
        if (dataset && dataset.toLowerCase().includes('sharepoint')) {
          const id = `sp:${dataset}:${table ?? ''}`;
          if (!seen.has(id)) {
            seen.set(id, {
              id, type: 'sharepoint',
              displayName: `SharePoint: ${table ?? dataset}`,
              siteUrl: dataset, listName: table,
              usedBy: [action.source], source: action.source
            });
          } else {
            seen.get(id)!.usedBy.push(action.source);
          }
        }
      }

      // Dataverse pattern: entity in host or path
      const host = inp.host as Record<string, unknown> | undefined;
      if (host && typeof host.connection === 'string' && host.connection.includes('dataverse')) {
        const entityName = (params?.entityName ?? params?.entity) as string | undefined;
        if (entityName) {
          const id = `dv:${entityName}`;
          if (!seen.has(id)) {
            seen.set(id, {
              id, type: 'dataverse',
              displayName: `Dataverse: ${entityName}`,
              tableName: entityName,
              usedBy: [action.source], source: action.source
            });
          }
        }
      }
    }
    return [...seen.values()];
  }

  // ── URL extraction ──────────────────────────────────────────────────────────

  private static extractUrls(actions: FlowAction[], archivePath: string): UrlReference[] {
    const urlSet = new Map<string, UrlReference>();
    for (const action of actions) {
      FlowDefinitionParser.scanForUrls(action.inputs, action.source, urlSet);
    }
    return [...urlSet.values()];
  }

  private static scanForUrls(
    obj: unknown,
    src: SourceRef,
    urlSet: Map<string, UrlReference>,
    depth = 0
  ): void {
    if (depth > 10 || !obj) return;
    if (typeof obj === 'string') {
      const matches = obj.match(/https?:\/\/[^\s"'<>{}|\\\^`[\]]+/g) ?? [];
      for (const url of matches) {
        if (!urlSet.has(url)) {
          urlSet.set(url, { url, ...UrlClassifier.classify(url), usedBy: [src] });
        } else {
          urlSet.get(url)!.usedBy.push(src);
        }
      }
    } else if (Array.isArray(obj)) {
      for (const item of obj) FlowDefinitionParser.scanForUrls(item, src, urlSet, depth + 1);
    } else if (typeof obj === 'object') {
      for (const val of Object.values(obj as Record<string, unknown>)) {
        FlowDefinitionParser.scanForUrls(val, src, urlSet, depth + 1);
      }
    }
  }

  // ── Error handling analysis ─────────────────────────────────────────────────

  private static analyseErrorHandling(actions: FlowAction[]): ErrorHandlingSummary {
    const locations: string[] = [];
    const scopes: string[] = [];
    for (const action of actions) {
      const ra = action.runAfter;
      // If any runAfter entry has 'Failed' or 'TimedOut' in its conditions, it's error handling
      // (In the raw format, runAfter values include status arrays)
      const hasFailedRunAfter = Object.keys(action.inputs ?? {}).length === 0 &&
        action.type === 'Terminate';
      if (hasFailedRunAfter) locations.push(action.id);
      if (action.type === 'Scope' && action.displayName.toLowerCase().includes('error')) {
        scopes.push(action.id);
      }
    }
    // Check runAfter fields for 'Failed' status
    for (const action of actions) {
      const ra = action.runAfter;
      for (const dep of ra) {
        // We stored runAfter as just keys; raw analysis would need the value
        // Best effort: check if this action's id suggests error handling
        if (action.displayName.toLowerCase().includes('fail') ||
            action.displayName.toLowerCase().includes('error') ||
            action.displayName.toLowerCase().includes('catch')) {
          if (!locations.includes(action.id)) locations.push(action.id);
        }
      }
    }
    return {
      hasExplicitHandling: locations.length > 0 || scopes.length > 0,
      handlingLocations: locations,
      scopesWithRunAfterFailure: scopes
    };
  }

  // ── Mermaid diagram ─────────────────────────────────────────────────────────

  static buildMermaid(trigger: TriggerDetail, actions: FlowAction[]): string {
    const nodeLines: string[] = [];
    const edgeLines: string[] = [];
    const escape = (s: string) => s.replace(/["\[\](){}|]/g, '_').substring(0, 40);
    const nodeId  = (id: string) => `n_${id.replace(/[^a-zA-Z0-9]/g, '_')}`;

    // Pre-build parent→children map (covers all depths)
    const childMap = new Map<string, FlowAction[]>();
    for (const a of actions) {
      if (a.parentScope) {
        if (!childMap.has(a.parentScope)) childMap.set(a.parentScope, []);
        childMap.get(a.parentScope)!.push(a);
      }
    }

    const CONTAINER_TYPES = new Set(['Scope', 'If', 'Foreach', 'Until', 'Switch']);

    /**
     * Recursively emit node/subgraph definitions, then collect edges.
     * indent grows with nesting depth so Mermaid renders cleanly.
     */
    const renderActions = (list: FlowAction[], indent: string) => {
      for (const action of list) {
        const nid      = nodeId(action.id);
        const label    = escape(action.displayName || action.id);
        const children = childMap.get(action.id) ?? [];
        const isContainer = CONTAINER_TYPES.has(action.type) && children.length > 0;

        if (isContainer) {
          // Icon prefix makes container type obvious at a glance
          const icon = action.type === 'Foreach' || action.type === 'Until' ? '🔁 '
                     : action.type === 'If'      || action.type === 'Switch' ? '🔀 '
                     : '📦 ';
          nodeLines.push(`${indent}subgraph ${nid} ["${icon}${label}"]`);
          renderActions(children, indent + '  ');
          nodeLines.push(`${indent}end`);
        } else if (action.type === 'If') {
          nodeLines.push(`${indent}${nid}{"${label}"}`);
        } else if (action.type === 'Foreach' || action.type === 'Until') {
          nodeLines.push(`${indent}${nid}[["${label}"]]`);
        } else if (action.aiBuilderModelId) {
          // Hexagon for AI Builder calls — visually distinct
          nodeLines.push(`${indent}${nid}{{"🤖 ${label}"}}`);
        } else {
          nodeLines.push(`${indent}${nid}["${label}"]`);
        }

        // Collect edges for this action (defined after all nodes, Mermaid handles cross-subgraph refs)
        for (const dep of action.runAfter) {
          edgeLines.push(`  ${nodeId(dep)} --> ${nid}`);
        }
      }
    };

    // Trigger node
    const trigLabel = trigger.recurrence
      ? `Recurrence: Every ${trigger.recurrence.interval} ${trigger.recurrence.frequency}`
      : `Trigger: ${trigger.type}`;
    nodeLines.push(`  __trigger(["${escape(trigLabel)}"])`);

    const topLevel = actions.filter(a => !a.parentScope);
    renderActions(topLevel, '  ');

    // Trigger → first top-level actions (those with no runAfter dependencies)
    for (const fa of topLevel.filter(a => a.runAfter.length === 0)) {
      edgeLines.push(`  __trigger --> ${nodeId(fa.id)}`);
    }

    return ['flowchart TD', ...nodeLines, ...edgeLines].join('\n');
  }

  /** Returns true if the string looks like a bare UUID/GUID. */
  private static isGuid(s: string): boolean {
    return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(s.trim());
  }

  /**
   * Strips a trailing GUID suffix from a name, handling both separators Power Platform uses:
   *   "IndueDNegativePress_NewRequest-85705400-5BC6-F011-8544-7C1E525DFBD3" → "IndueDNegativePress_NewRequest"
   *   "MyFlow_07595c8a-243c-4b1e-9cfd-5f0d281b4393"                        → "MyFlow"
   */
  private static stripGuidSuffix(s: string): string {
    return s.replace(/[-_][0-9a-f]{8}(?:-[0-9a-f]{4}){3}-[0-9a-f]{12}$/i, '').trim();
  }

  /**
   * Converts a workflow archive path like
   *   "Workflows/IndueDNegativePress_NewRequest-85705400-5BC6-F011-8544-7C1E525DFBD3.json"
   * into a clean display name "IndueDNegativePress_NewRequest".
   */
  private static fileNameToDisplayName(archivePath: string): string {
    return FlowDefinitionParser.stripGuidSuffix(
      archivePath.replace(/.*\//, '').replace(/\.json$/i, '')
    );
  }
}
