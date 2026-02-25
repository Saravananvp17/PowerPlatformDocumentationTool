import JSZip from 'jszip';
import fs from 'fs';
import { XMLParser } from 'fast-xml-parser';
import type {
  NormalizedAnalysisGraph, SolutionInfo, EnvVar, SecurityRole, RolePrivilege,
  ModelDrivenApp, SitemapArea, SitemapGroup, SitemapSubArea,
  WebResource, DataverseFormula, ConnectionReference,
  CopilotAgent, CopilotTopic,
  AiBuilderModel, AiBuilderInputVariable,
  ManagedPlan, ManagedPlanPersona, ManagedPlanArtifact, ManagedPlanEntity, ManagedPlanProcessDiagram,
  SourceRef, ProgressCallback, ParseError, ArtifactMeta
} from '../model/types';
import { FlowDefinitionParser } from './FlowDefinitionParser';
import { MsappParser } from './MsappParser';
import * as yaml from 'js-yaml';

const xmlParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  isArray: (name) => ['RootComponent', 'LocalizedName', 'RolePrivilege',
    'Attribute', 'Entity', 'Role', 'Workflow', 'AppModule', 'WebResource',
    'EnvironmentVariableDefinition', 'EnvironmentVariableValue',
    'Area', 'Group', 'SubArea', 'AppModuleRoleMap', 'AppModuleSiteMap'
  ].includes(name),
  textNodeName: '#text'
});

export class SolutionZipParser {
  static async parse(
    filePath: string,
    meta: ArtifactMeta,
    onProgress?: ProgressCallback
  ): Promise<Partial<NormalizedAnalysisGraph>> {
    const buf = fs.readFileSync(filePath);
    const zip = await JSZip.loadAsync(buf);
    const errors: ParseError[] = [];
    const result: Partial<NormalizedAnalysisGraph> = {
      flows: [], canvasApps: [], modelDrivenApps: [], copilotAgents: [], aiBuilderModels: [], managedPlans: [], envVars: [],
      securityRoles: [], webResources: [], dataverseFormulas: [],
      connectors: [], connectionRefs: [], dataSources: [], urls: [],
      missingDeps: [], manualChecklist: [], parseErrors: []
    };

    onProgress?.({ stage: 'solution', pct: 5, message: 'Reading solution.xml...' });
    try {
      result.solution = await SolutionZipParser.parseSolutionXml(zip, filePath);
    } catch (e: unknown) {
      errors.push({ archivePath: 'solution.xml', error: String(e), partial: false });
    }

    onProgress?.({ stage: 'solution', pct: 20, message: 'Parsing customizations.xml...' });
    try {
      await SolutionZipParser.parseCustomizations(zip, filePath, result, errors);
    } catch (e: unknown) {
      errors.push({ archivePath: 'customizations.xml', error: String(e), partial: true });
    }

    onProgress?.({ stage: 'solution', pct: 55, message: 'Parsing workflow definitions...' });
    await SolutionZipParser.parseWorkflows(zip, filePath, result, errors, onProgress);

    onProgress?.({ stage: 'solution', pct: 75, message: 'Parsing canvas apps...' });
    await SolutionZipParser.parseMsapps(zip, filePath, result, errors, onProgress);

    onProgress?.({ stage: 'solution', pct: 79, message: 'Parsing Managed Plans...' });
    await SolutionZipParser.parseManagedPlans(zip, filePath, result, errors);

    onProgress?.({ stage: 'solution', pct: 81, message: 'Parsing AI Builder models...' });
    await SolutionZipParser.parseAiBuilderModels(zip, filePath, result, errors);

    onProgress?.({ stage: 'solution', pct: 83, message: 'Parsing Copilot Studio agents...' });
    await SolutionZipParser.parseCopilotAgents(zip, filePath, result, errors);

    onProgress?.({ stage: 'solution', pct: 87, message: 'Parsing environment variable definitions...' });
    await SolutionZipParser.parseEnvVarDefinitionFiles(zip, filePath, result, errors);

    onProgress?.({ stage: 'solution', pct: 90, message: 'Parsing Dataverse formulas...' });
    await SolutionZipParser.parseDataverseFormulas(zip, filePath, result, errors);

    result.parseErrors = errors;
    onProgress?.({ stage: 'solution', pct: 100, message: 'Solution parsed' });
    return result;
  }

  // ── solution.xml ────────────────────────────────────────────────────────────

  private static async parseSolutionXml(zip: JSZip, archivePath: string): Promise<SolutionInfo> {
    const file = zip.file('solution.xml');
    if (!file) throw new Error('solution.xml not found');
    const xml = xmlParser.parse(await file.async('string'));
    const imp = xml.ImportExportXml ?? xml;
    const sol = imp.SolutionManifest ?? imp;

    const localizedNames: Record<string,string>[] = sol.LocalizedNames?.LocalizedName ?? [];
    const displayName = localizedNames.find((n: Record<string,string>) =>
      n['@_languagecode'] === '1033')?.['@_description'] ?? sol.UniqueName ?? '';

    const pub = sol.Publisher ?? {};
    const pubNames: Record<string,string>[] = pub.LocalizedNames?.LocalizedName ?? [];
    const pubDisplayName = pubNames.find((n: Record<string,string>) =>
      n['@_languagecode'] === '1033')?.['@_description'] ?? pub.UniqueName ?? '';

    const roots: Record<string,string>[] = sol.RootComponents?.RootComponent ?? [];
    const componentCounts: Record<string, number> = {};
    for (const rc of roots) {
      const type = String(rc['@_type'] ?? 'Unknown');
      componentCounts[type] = (componentCounts[type] ?? 0) + 1;
    }

    return {
      uniqueName: String(sol.UniqueName ?? ''),
      displayName,
      version: String(sol.Version ?? ''),
      isManaged: String(sol.Managed ?? '0') !== '0',
      publisher: {
        uniqueName: String(pub.UniqueName ?? ''),
        displayName: pubDisplayName,
        prefix: String(pub.CustomizationPrefix ?? '')
      },
      description: String(sol.Description ?? ''),
      componentCounts,
      source: { archivePath: archivePath + '/solution.xml', confidence: 'high' }
    };
  }

  // ── customizations.xml ──────────────────────────────────────────────────────

  private static async parseCustomizations(
    zip: JSZip, archivePath: string,
    result: Partial<NormalizedAnalysisGraph>, errors: ParseError[]
  ): Promise<void> {
    const file = zip.file('customizations.xml');
    if (!file) return;
    const content = await file.async('string');
    const xml = xmlParser.parse(content);
    const root = xml.ImportExportXml ?? xml;
    const src: SourceRef = { archivePath: archivePath + '/customizations.xml', confidence: 'high' };

    // Environment variables
    SolutionZipParser.parseEnvVars(root, src, result);

    // Security roles
    SolutionZipParser.parseRoles(root, src, result);

    // Model-driven apps
    SolutionZipParser.parseAppModules(root, src, result);

    // Web resources
    SolutionZipParser.parseWebResources(root, src, result);

    // Connection references
    SolutionZipParser.parseConnectionReferences(root, src, result);
  }

  private static parseEnvVars(
    root: Record<string, unknown>, src: SourceRef, result: Partial<NormalizedAnalysisGraph>
  ): void {
    const defs: Record<string, unknown>[] =
      (root as Record<string, Record<string, Record<string, unknown>[]>>)
        ?.EnvironmentVariableDefinitions?.EnvironmentVariableDefinition ?? [];
    const vals: Record<string, unknown>[] =
      (root as Record<string, Record<string, Record<string, unknown>[]>>)
        ?.EnvironmentVariableValues?.EnvironmentVariableValue ?? [];

    const valMap = new Map<string, string>();
    for (const v of vals) {
      valMap.set(String(v.SchemaName ?? ''), String(v.Value ?? v['#text'] ?? ''));
    }

    for (const def of defs) {
      const schemaName = String(def.SchemaName ?? def['@_schemaname'] ?? '');
      const localNames: Record<string,string>[] = (def as Record<string,Record<string,Record<string,string>[]>>)
        .DisplayNames?.DisplayName ?? [];
      const displayName = localNames.find(n => n['@_languagecode'] === '1033')?.['@_description'] ?? schemaName;

      result.envVars!.push({
        id: schemaName,
        schemaName,
        displayName,
        description: String(def.Description ?? ''),
        type: String(def.Type ?? def['@_type'] ?? 'string'),
        defaultValue: String(def.DefaultValue ?? ''),
        currentValue: valMap.has(schemaName) ? valMap.get(schemaName) : 'not included in export',
        usedBy: [],
        source: { ...src, xmlXPath: `//EnvironmentVariableDefinition[@schemaname='${schemaName}']` }
      });
    }
  }

  // ── Individual environmentvariabledefinitions/*.xml files ───────────────────
  // Modern solution exports store each component in its own folder/file.
  // Structure: environmentvariabledefinitions/<schemaname>/environmentvariabledefinition.xml
  // These are merged with (and take priority over) anything already parsed from customizations.xml.

  private static async parseEnvVarDefinitionFiles(
    zip: JSZip, archivePath: string,
    result: Partial<NormalizedAnalysisGraph>, errors: import('../model/types').ParseError[]
  ): Promise<void> {
    // Numeric type codes used in the XML → human-readable labels
    const TYPE_LABELS: Record<string, string> = {
      '100000000': 'string',
      '100000001': 'number',
      '100000002': 'boolean',
      '100000003': 'json',
      '100000004': 'datasource',
      '100000005': 'secret',
    };

    // Find every XML file anywhere under environmentvariabledefinitions/
    const evFiles = Object.keys(zip.files).filter(p => {
      const lower = p.replace(/\\/g, '/').toLowerCase();
      return lower.includes('environmentvariabledefinitions/') && lower.endsWith('.xml');
    });

    if (!evFiles.length) return;

    // Build a set of schema names already found (from customizations.xml) so we can fill gaps
    const existing = new Map<string, import('../model/types').EnvVar>();
    for (const ev of result.envVars ?? []) existing.set(ev.schemaName.toLowerCase(), ev);

    for (const filePath of evFiles) {
      const file = zip.file(filePath);
      if (!file) continue;
      try {
        const xml = xmlParser.parse(await file.async('string'));
        // Root element may be named 'environmentvariabledefinition' (lowercase) or camelCase
        const def: Record<string, unknown> =
          xml.environmentvariabledefinition ??
          xml.EnvironmentVariableDefinition ??
          xml[Object.keys(xml)[0]] ??
          {};

        // Schema name — attribute on root element, fall back to folder name
        const folderName = filePath.replace(/\\/g, '/').split('/').slice(-2)[0] ?? '';
        const schemaName = String(
          def['@_schemaname'] ?? def['@_uniquename'] ?? def['@_SchemaName'] ??
          folderName
        ).trim();
        if (!schemaName) continue;

        const key = schemaName.toLowerCase();
        const src: SourceRef = { archivePath: `${archivePath}/${filePath}`, confidence: 'high' };

        // Helper: resolve a localised-label element which has this shape:
        //   <displayname default="Human Name">
        //     <label description="Human Name" languagecode="1033" />
        //   </displayname>
        // Returns the 'default' attribute, or the English label's description, or fallback.
        const resolveLabel = (node: unknown, fallback: string): string => {
          if (!node || typeof node !== 'object') return fallback;
          const n = node as Record<string, unknown>;
          // 'default' attribute — the most direct value
          if (typeof n['@_default'] === 'string' && n['@_default']) return n['@_default'];
          // child <label> nodes — may be single object or array
          const labels: Record<string, string>[] = Array.isArray(n.label)
            ? n.label as Record<string, string>[]
            : n.label && typeof n.label === 'object' ? [n.label as Record<string, string>] : [];
          const en = labels.find(l => l['@_languagecode'] === '1033');
          return en?.['@_description'] ?? labels[0]?.['@_description'] ?? fallback;
        };

        const displayName = resolveLabel(def.displayname ?? def.DisplayName, schemaName);
        const description = resolveLabel(def.description ?? def.Description, '');

        // Type — child text element containing a numeric code
        const rawType = String(def.type ?? def.Type ?? def['@_type'] ?? '');
        const typeLabel = TYPE_LABELS[rawType] ?? (rawType || 'string');

        // Default value — child text element or attribute
        const rawDefault = def.defaultvalue ?? def.DefaultValue ??
          def['@_defaultvalue'] ?? def['@_DefaultValue'];
        const defaultValue = typeof rawDefault === 'string' ? rawDefault : '';

        // If we already have this var from customizations.xml, enrich it; otherwise add it fresh
        if (existing.has(key)) {
          const ev = existing.get(key)!;
          if (!ev.displayName || ev.displayName === ev.schemaName) ev.displayName = displayName;
          if (!ev.description)  ev.description  = description;
          if (!ev.defaultValue) ev.defaultValue  = defaultValue || undefined;
          if (!ev.type || ev.type === 'string') ev.type = typeLabel;
          ev.source = src;   // upgrade to the more specific file reference
        } else {
          result.envVars!.push({
            id: schemaName,
            schemaName,
            displayName,
            description,
            type: typeLabel,
            defaultValue: defaultValue || undefined,
            currentValue: 'not included in export',
            usedBy: [],
            source: src
          });
          existing.set(key, result.envVars![result.envVars!.length - 1]);
        }
      } catch (e) {
        errors.push({ archivePath: filePath, error: String(e), partial: true });
      }
    }
  }

  private static parseRoles(
    root: Record<string, unknown>, src: SourceRef, result: Partial<NormalizedAnalysisGraph>
  ): void {
    const roles: Record<string, unknown>[] =
      (root as Record<string, Record<string, Record<string, unknown>[]>>)?.Roles?.Role ?? [];

    for (const role of roles) {
      const privs: Record<string,string>[] =
        (role as Record<string,Record<string,Record<string,string>[]>>).RolePrivileges?.RolePrivilege ?? [];

      // Action verbs — AppendTo must precede Append to avoid partial match
      const ACTION_RE = /^prv(Read|Write|Create|Delete|AppendTo|Append|Assign|Share)/;
      const privileges: RolePrivilege[] = privs.map(p => ({
        privilegeName: String(p['@_name'] ?? ''),
        entityName: String(p['@_name'] ?? '').replace(ACTION_RE, ''),
        depth: String(p['@_level'] ?? p['@_permissiontype'] ?? 'Unknown')
      }));

      result.securityRoles!.push({
        id: String(role['@_id'] ?? role.RoleId ?? ''),
        name: String(role['@_name'] ?? role.Name ?? ''),
        templateId: String(role['@_templateId'] ?? ''),
        privilegeCount: privileges.length,
        privileges,
        assignedToApps: [],
        source: { ...src, xmlXPath: `//Role[@name='${role['@_name']}']` }
      });
    }
  }

  private static parseAppModules(
    root: Record<string, unknown>, src: SourceRef, result: Partial<NormalizedAnalysisGraph>
  ): void {
    const appModules: Record<string,unknown>[] =
      (root as Record<string,Record<string,Record<string,unknown>[]>>)?.AppModules?.AppModule ?? [];

    for (const am of appModules) {
      const uniqueName = String(am['@_UniqueName'] ?? am.UniqueName ?? '');
      const localNames: Record<string,string>[] =
        (am as Record<string,Record<string,Record<string,string>[]>>).LocalizedNames?.LocalizedName ?? [];
      const displayName = localNames.find(n => n['@_languagecode'] === '1033')?.['@_description'] ?? uniqueName;

      // Role maps
      const roleMaps: Record<string,string>[] =
        (am as Record<string,Record<string,Record<string,string>[]>>).AppModuleRoleMaps?.AppModuleRoleMap ?? [];
      const roleIds = roleMaps.map(r => String(r['@_RoleId'] ?? r.RoleId ?? ''));

      // Assign to roles
      for (const rid of roleIds) {
        const role = result.securityRoles!.find(r => r.id === rid || r.templateId === rid);
        if (role && !role.assignedToApps.includes(displayName)) role.assignedToApps.push(displayName);
      }

      // Sitemap
      const sitemaps: Record<string,unknown>[] =
        (am as Record<string,Record<string,Record<string,unknown>[]>>).AppModuleSiteMaps?.AppModuleSiteMap ?? [];
      const areas: SitemapArea[] = [];
      for (const sm of sitemaps) {
        const smAreas: Record<string,unknown>[] =
          (sm as Record<string,Record<string,Record<string,unknown>[]>>).SiteMap?.Area ??
          (sm as Record<string,Record<string,Record<string,unknown>[]>>).Area ?? [];
        for (const area of smAreas) {
          const areaTitle = String((area as Record<string,Record<string,Record<string,string>>>).Titles?.Title?.['@_title'] ?? area['@_Title'] ?? '');
          const groups: SitemapGroup[] = [];
          const rawGroups: Record<string,unknown>[] =
            (area as Record<string, Record<string,unknown>[]>).Group ?? [];
          for (const grp of rawGroups) {
            const grpTitle = String((grp as Record<string,Record<string,Record<string,string>>>).Titles?.Title?.['@_title'] ?? grp['@_Title'] ?? '');
            const subAreas: SitemapSubArea[] = [];
            const rawSubAreas: Record<string,unknown>[] =
              (grp as Record<string, Record<string,unknown>[]>).SubArea ?? [];
            for (const sa of rawSubAreas) {
              subAreas.push({
                id: String(sa['@_Id'] ?? ''),
                type: sa['@_Entity'] ? 'Entity' : sa['@_Url'] ? 'URL' : 'Other',
                title: String((sa as Record<string,Record<string,Record<string,string>>>).Titles?.Title?.['@_title'] ?? sa['@_Id'] ?? ''),
                entityName: String(sa['@_Entity'] ?? ''),
                url: String(sa['@_Url'] ?? '')
              });
            }
            groups.push({ id: String(grp['@_Id'] ?? ''), title: grpTitle, subAreas });
          }
          areas.push({ id: String(area['@_Id'] ?? ''), title: areaTitle, groups });
        }
      }

      const exposedTables = areas
        .flatMap(a => a.groups)
        .flatMap(g => g.subAreas)
        .filter(s => s.type === 'Entity' && s.entityName)
        .map(s => s.entityName!);

      result.modelDrivenApps!.push({
        id: String(am['@_AppModuleId'] ?? ''),
        uniqueName,
        displayName,
        description: String(am['@_description'] ?? ''),
        assignedRoleIds: roleIds,
        assignedRoleNames: roleIds.map(rid => {
          const role = result.securityRoles!.find(r => r.id === rid);
          return role?.name ?? rid;
        }),
        sitemapAreas: areas,
        exposedTables: [...new Set(exposedTables)],
        mermaidSitemap: SolutionZipParser.buildSitemapMermaid(displayName, areas),
        source: { ...src, xmlXPath: `//AppModule[@UniqueName='${uniqueName}']` }
      });
    }
  }

  private static buildSitemapMermaid(appName: string, areas: SitemapArea[]): string {
    const lines = ['flowchart TD'];
    const esc = (s: string) => s.replace(/[^a-zA-Z0-9_]/g, '_').substring(0, 30) || '_';
    lines.push(`  root["${appName}"]`);
    for (const area of areas) {
      const aId = `a_${esc(area.id || area.title)}`;
      lines.push(`  ${aId}["📁 ${area.title || area.id}"]`);
      lines.push(`  root --> ${aId}`);
      for (const grp of area.groups) {
        const gId = `g_${esc(grp.id || grp.title)}`;
        lines.push(`  ${gId}["📂 ${grp.title || grp.id}"]`);
        lines.push(`  ${aId} --> ${gId}`);
        for (const sa of grp.subAreas) {
          const sId = `s_${esc(sa.id)}`;
          const icon = sa.type === 'Entity' ? '📋' : sa.type === 'URL' ? '🔗' : '📄';
          lines.push(`  ${sId}["${icon} ${sa.title || sa.id}"]`);
          lines.push(`  ${gId} --> ${sId}`);
        }
      }
    }
    return lines.join('\n');
  }

  private static parseWebResources(
    root: Record<string, unknown>, src: SourceRef, result: Partial<NormalizedAnalysisGraph>
  ): void {
    const wrs: Record<string,string>[] =
      (root as Record<string,Record<string,Record<string,string>[]>>)?.WebResources?.WebResource ?? [];
    const typeMap: Record<string,string> = {
      '1':'HTML', '2':'CSS', '3':'JS', '4':'XML', '5':'PNG', '6':'JPG',
      '7':'GIF', '10':'ICO', '11':'SVG', '12':'RESX'
    };
    for (const wr of wrs) {
      result.webResources!.push({
        name: String(wr['@_Name'] ?? wr.Name ?? ''),
        displayName: String(wr['@_DisplayName'] ?? wr.DisplayName ?? ''),
        type: typeMap[String(wr['@_WebResourceType'] ?? '')] ?? String(wr['@_WebResourceType'] ?? 'Unknown'),
        usedBy: [],
        source: src
      });
    }
  }

  private static parseConnectionReferences(
    root: Record<string, unknown>, src: SourceRef, result: Partial<NormalizedAnalysisGraph>
  ): void {
    const refs: Record<string,string>[] =
      (root as Record<string,Record<string,Record<string,string>[]>>)?.ConnectionReferences?.ConnectionReference ?? [];
    for (const ref of refs) {
      result.connectionRefs!.push({
        logicalName: String(ref['@_ConnectionReferenceLogicalName'] ?? ''),
        displayName: String(ref['@_ConnectionReferenceDisplayName'] ?? ''),
        connectorId: String(ref['@_ConnectorId'] ?? ''),
        connectionReferenceType: String(ref['@_Type'] ?? ''),
        source: src
      });
    }
  }

  // ── Workflows ────────────────────────────────────────────────────────────────

  private static async parseWorkflows(
    zip: JSZip, archivePath: string,
    result: Partial<NormalizedAnalysisGraph>, errors: ParseError[],
    onProgress?: ProgressCallback
  ): Promise<void> {
    const flowFiles = Object.keys(zip.files).filter(e =>
      e.startsWith('Workflows/') && e.endsWith('.json') && !zip.files[e].dir
    );

    for (let i = 0; i < flowFiles.length; i++) {
      const filePath = flowFiles[i];
      onProgress?.({ stage: 'flows', pct: Math.round((i / flowFiles.length) * 100), message: `Parsing flow: ${filePath}` });
      try {
        const content = await zip.files[filePath].async('string');
        const raw = JSON.parse(content);
        const flow = FlowDefinitionParser.parse(raw, archivePath + '/' + filePath);
        result.flows!.push(flow);

        // Merge data sources & connectors
        for (const ds of flow.dataSources) {
          if (!result.dataSources!.find(d => d.id === ds.id)) result.dataSources!.push(ds);
        }
        for (const conn of flow.connectors) {
          conn.usedInFlows.push(flow.id);
          const existing = result.connectors!.find(c => c.connectorId === conn.connectorId);
          if (existing) {
            if (!existing.usedInFlows.includes(flow.id)) existing.usedInFlows.push(flow.id);
          } else {
            result.connectors!.push(conn);
          }
        }
        for (const url of flow.urls) {
          const existing = result.urls!.find(u => u.url === url.url);
          if (existing) {
            existing.usedBy.push(...url.usedBy);
          } else {
            result.urls!.push(url);
          }
        }
      } catch (e: unknown) {
        errors.push({ archivePath: archivePath + '/' + filePath, error: String(e), partial: false });
      }
    }
  }

  // ── Canvas apps (embedded .msapp files in solution) ──────────────────────────

  private static async parseMsapps(
    zip: JSZip, archivePath: string,
    result: Partial<NormalizedAnalysisGraph>, errors: ParseError[],
    onProgress?: ProgressCallback
  ): Promise<void> {
    const msappFiles = Object.keys(zip.files).filter(e =>
      e.endsWith('.msapp') && !zip.files[e].dir
    );

    for (const filePath of msappFiles) {
      onProgress?.({ stage: 'apps', pct: 0, message: `Parsing canvas app: ${filePath}` });
      try {
        const buf = Buffer.from(await zip.files[filePath].async('arraybuffer'));
        const displayName = filePath.replace(/.*\//, '').replace('.msapp', '');
        const app = await MsappParser.parse(buf, archivePath + '/' + filePath, displayName, onProgress);
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
        errors.push({ archivePath: archivePath + '/' + filePath, error: String(e), partial: false });
      }
    }
  }

  // ── Copilot Studio agents ─────────────────────────────────────────────────────

  private static async parseCopilotAgents(
    zip: JSZip, archivePath: string,
    result: Partial<NormalizedAnalysisGraph>, errors: ParseError[]
  ): Promise<void> {
    // Discover all bot.xml files — one per agent
    const botXmlPaths = Object.keys(zip.files)
      .filter(p => /^bots\/[^/]+\/bot\.xml$/.test(p));
    if (!botXmlPaths.length) return;

    const AUTH_MODES: Record<string, string> = {
      '0': 'No authentication',
      '1': 'Integrated (Teams / AAD)',
      '2': 'Manual / OAuth',
    };

    // Trigger-kind → human label
    const TRIGGER_LABELS: Record<string, string> = {
      OnRecognizedIntent:    'User intent',
      OnSystemRedirect:      'System redirect',
      OnError:               'Error handler',
      OnConversationStart:   'Conversation start',
      OnEndOfConversation:   'End of conversation',
      OnUnknownIntent:       'Unknown intent / fallback',
    };

    // Recursively collect unique action kinds from a YAML actions array
    const collectActionKinds = (actions: unknown[]): string[] => {
      if (!Array.isArray(actions)) return [];
      const kinds: string[] = [];
      for (const action of actions) {
        const a = action as Record<string, unknown>;
        if (a?.kind) kinds.push(String(a.kind));
        if (Array.isArray(a?.actions)) kinds.push(...collectActionKinds(a.actions as unknown[]));
        if (Array.isArray(a?.conditions)) {
          for (const cond of a.conditions as Record<string, unknown>[]) {
            if (Array.isArray(cond?.actions)) kinds.push(...collectActionKinds(cond.actions as unknown[]));
          }
        }
      }
      return [...new Set(kinds)];
    };

    for (const botXmlPath of botXmlPaths) {
      const botFolder   = botXmlPath.replace('/bot.xml', '');   // e.g. "bots/ndmlab01_..."
      const schemaName  = botFolder.replace('bots/', '');
      const src: SourceRef = { archivePath: `${archivePath}/${botXmlPath}`, confidence: 'high' };

      try {
        // ── Parse bot.xml ──────────────────────────────────────────────────────
        const botFile = zip.file(botXmlPath);
        if (!botFile) continue;
        const botXmlParsed = xmlParser.parse(await botFile.async('string'));
        const bot = (botXmlParsed.bot ?? {}) as Record<string, unknown>;
        const displayName = String(bot.name ?? schemaName);
        const authMode = AUTH_MODES[String(bot.authenticationmode ?? '0')] ?? 'Unknown';

        // ── Parse configuration.json ───────────────────────────────────────────
        let channels: string[] = [];
        const aiSettings = {
          generativeActionsEnabled: false,
          useModelKnowledge: false,
          fileAnalysisEnabled: false,
          semanticSearchEnabled: false,
        };
        const configFile = zip.file(`${botFolder}/configuration.json`);
        if (configFile) {
          try {
            const cfg = JSON.parse(await configFile.async('string')) as Record<string, unknown>;
            channels = ((cfg.channels ?? []) as Record<string, unknown>[])
              .map(c => String(c.channelId ?? ''));
            const s = (cfg.settings ?? {}) as Record<string, unknown>;
            const ai = (cfg.aISettings ?? {}) as Record<string, unknown>;
            aiSettings.generativeActionsEnabled = Boolean(s.GenerativeActionsEnabled);
            aiSettings.useModelKnowledge        = Boolean(ai.useModelKnowledge);
            aiSettings.fileAnalysisEnabled      = Boolean(ai.isFileAnalysisEnabled);
            aiSettings.semanticSearchEnabled    = Boolean(ai.isSemanticSearchEnabled);
          } catch (e) {
            errors.push({ archivePath: `${archivePath}/${botFolder}/configuration.json`, error: String(e), partial: true });
          }
        }

        // ── Parse botcomponents for this agent ────────────────────────────────
        const topics: CopilotTopic[] = [];
        let gptInstructions = '';
        const knowledgeFiles: string[] = [];

        const compXmlPaths = Object.keys(zip.files)
          .filter(p => p.startsWith('botcomponents/') && p.endsWith('/botcomponent.xml'));

        for (const compXmlPath of compXmlPaths) {
          try {
            const compFile = zip.file(compXmlPath);
            if (!compFile) continue;
            const compParsed = xmlParser.parse(await compFile.async('string'));
            const comp = (compParsed.botcomponent ?? {}) as Record<string, unknown>;

            // Only include components that belong to this agent
            const parentSchema = (comp.parentbotid as Record<string, unknown> | undefined)?.schemaname;
            if (String(parentSchema ?? '') !== schemaName) continue;

            const compSchema  = String(comp['@_schemaname'] ?? '');
            const compName    = String(comp.name ?? compSchema);
            const compDesc    = String(comp.description ?? '');
            const compType    = Number(comp.componenttype ?? 0);
            const isActive    = String(comp.statecode ?? '0') === '0';
            const compFolder  = compXmlPath.replace('/botcomponent.xml', '');
            const dataPath    = `${compFolder}/data`;

            if (compType === 14) {
              // Knowledge / file attachment
              const fileEl = comp.filedata as Record<string, unknown> | undefined;
              knowledgeFiles.push(String(fileEl?.['#text'] ?? comp.name ?? compSchema));
              continue;
            }

            if (compType === 15) {
              // GPT instructions component
              const dataFile = zip.file(dataPath);
              if (dataFile) {
                try {
                  const dataYaml = yaml.load(await dataFile.async('string')) as Record<string, unknown>;
                  gptInstructions = String(dataYaml?.instructions ?? '');
                } catch (e) {
                  errors.push({ archivePath: `${archivePath}/${dataPath}`, error: String(e), partial: true });
                }
              }
              continue;
            }

            if (compType === 9) {
              // Topic
              let triggerKind = '', triggerDisplayName = '';
              let triggerQueries: string[] = [], actionKinds: string[] = [];

              const dataFile = zip.file(dataPath);
              if (dataFile) {
                try {
                  const dataYaml = yaml.load(await dataFile.async('string')) as Record<string, unknown>;
                  const bd = (dataYaml?.beginDialog ?? {}) as Record<string, unknown>;
                  triggerKind        = String(bd.kind ?? '');
                  const intent       = (bd.intent ?? {}) as Record<string, unknown>;
                  triggerDisplayName = String(intent.displayName ?? '');
                  triggerQueries     = Array.isArray(intent.triggerQueries)
                    ? (intent.triggerQueries as unknown[]).map(q => String(q))
                    : [];
                  actionKinds = collectActionKinds((bd.actions ?? []) as unknown[]);
                } catch (e) {
                  errors.push({ archivePath: `${archivePath}/${dataPath}`, error: String(e), partial: true });
                }
              }

              topics.push({
                schemaName: compSchema,
                name: compName,
                description: compDesc,
                triggerKind,
                triggerDisplayName: triggerDisplayName || compName,
                triggerQueries,
                actionKinds,
                isActive,
                source: { archivePath: `${archivePath}/${compXmlPath}`, confidence: 'high' },
              } as CopilotTopic);
            }
          } catch (e) {
            errors.push({ archivePath: compXmlPath, error: String(e), partial: true });
          }
        }

        // Sort topics: user-triggered first, then system/other
        topics.sort((a, b) => {
          const aUser = a.triggerKind === 'OnRecognizedIntent' ? 0 : 1;
          const bUser = b.triggerKind === 'OnRecognizedIntent' ? 0 : 1;
          return aUser - bUser || a.name.localeCompare(b.name);
        });

        result.copilotAgents!.push({
          schemaName, displayName, channels, authMode,
          aiSettings, topics, gptInstructions, knowledgeFiles, source: src,
        } as CopilotAgent);

      } catch (e) {
        errors.push({ archivePath: botXmlPath, error: String(e), partial: true });
      }
    }
  }

  // ── Managed Plans (msdyn_plans) ───────────────────────────────────────────────

  private static buildProcessDiagramMermaid(
    nodes: Record<string, unknown>[],
    edges: Record<string, unknown>[]
  ): string {
    const escape = (s: string) => s.replace(/["\[\](){}|]/g, '_').substring(0, 50);
    const nid = (id: string) => `n_${id.replace(/[^a-zA-Z0-9]/g, '_')}`;

    const lines: string[] = ['flowchart TD'];
    for (const node of nodes) {
      const id   = String(node.id ?? '');
      const ntype = String(node.type ?? 'activity');
      const data  = (node.data ?? {}) as Record<string, unknown>;
      const cfg   = (node.cfgNode ?? {}) as Record<string, unknown>;
      const label = escape(String(data.label ?? cfg.label ?? data.name ?? cfg.name ?? id));
      const shape =
        ntype === 'event'   ? `(["${label}"])` :
        ntype === 'gateway' ? `{"${label}"}` :
        `["${label}"]`;
      lines.push(`  ${nid(id)}${shape}`);
    }
    for (const edge of edges) {
      const src  = String(edge.sourceNodeId ?? '');
      const tgt  = String(edge.targetNodeId ?? '');
      if (src && tgt) lines.push(`  ${nid(src)} --> ${nid(tgt)}`);
    }
    return lines.join('\n');
  }

  private static async parseManagedPlans(
    zip: JSZip, archivePath: string,
    result: Partial<NormalizedAnalysisGraph>, errors: ParseError[]
  ): Promise<void> {
    // Each plan lives at msdyn_plans/<planId>/msdyn_plan.xml
    const planXmlPaths = Object.keys(zip.files)
      .filter(p => /^msdyn_plans\/[^/]+\/msdyn_plan\.xml$/.test(p));
    if (!planXmlPaths.length) return;

    for (const planXmlPath of planXmlPaths) {
      try {
        const planXml = xmlParser.parse(await zip.files[planXmlPath].async('string')) as Record<string, unknown>;
        const plan = (planXml.msdyn_plan ?? planXml) as Record<string, unknown>;
        const planId = String(plan.msdyn_planid ?? '').replace(/[{}]/g, '');
        const name   = String(plan.msdyn_name ?? 'Unnamed Plan');
        const desc   = String(plan.msdyn_description ?? '');
        const prompt = String(plan.msdyn_prompt ?? '');

        const planBase = planXmlPath.replace('msdyn_plan.xml', '');
        const src: SourceRef = { archivePath: archivePath + '/' + planXmlPath, confidence: 'high' };

        // ── Main content JSON ────────────────────────────────────────────────
        let personas: ManagedPlanPersona[] = [];
        let entities: ManagedPlanEntity[] = [];
        let processDiagrams: ManagedPlanProcessDiagram[] = [];
        let artifactIds: string[] = [];

        const mainContentPath = planBase + 'msdyn_content/content';
        if (zip.files[mainContentPath]) {
          try {
            const mainContent = JSON.parse(await zip.files[mainContentPath].async('string')) as Record<string, unknown>;

            // Personas + user stories
            const pas = mainContent.personasAndUserStories as Array<Record<string, unknown>> | undefined;
            if (Array.isArray(pas)) {
              for (const pa of pas) {
                const user = pa.user as Record<string, unknown> | undefined;
                const stories = pa.userStories as Array<Record<string, unknown>> | undefined;
                personas.push({
                  id:          String(user?.id ?? ''),
                  name:        String(user?.name ?? ''),
                  description: String(user?.description ?? ''),
                  userStories: Array.isArray(stories)
                    ? stories.map(s => String(s.description ?? '')).filter(Boolean)
                    : [],
                });
              }
            }

            // Entities collection
            const ec = mainContent.entitiesCollection as Array<Record<string, unknown>> | undefined;
            if (Array.isArray(ec)) {
              for (const e of ec) {
                const def = (e.EntityDefinition ?? e) as Record<string, unknown>;
                const schemaName   = String(def.SchemaName ?? e.schemaName ?? '');
                const displayName  = String(
                  (((def.DisplayName as Record<string, unknown>)?.LocalizedLabels as Array<Record<string, unknown>>)
                    ?.[0]?.Label) ?? e.name ?? schemaName
                );
                const entityDesc = String(
                  (((def.Description as Record<string, unknown>)?.LocalizedLabels as Array<Record<string, unknown>>)
                    ?.[0]?.Label) ?? e.description ?? ''
                );
                const attrs = (def.Attributes ?? []) as Array<Record<string, unknown>>;
                entities.push({
                  schemaName, displayName, description: entityDesc,
                  attributes: attrs.slice(0, 20).map(a => ({
                    name:        String(a.SchemaName ?? a.name ?? ''),
                    type:        String(a.AttributeType ?? a.type ?? ''),
                    description: String(
                      (((a.Description as Record<string, unknown>)?.LocalizedLabels as Array<Record<string, unknown>>)
                        ?.[0]?.Label) ?? a.description ?? ''
                    ),
                  })).filter(a => a.name),
                });
              }
            }

            // Process diagrams
            const pds = mainContent.planProcessDiagrams as Array<Record<string, unknown>> | undefined;
            if (Array.isArray(pds)) {
              for (const pd of pds) {
                const nodes = (pd.nodes ?? []) as Record<string, unknown>[];
                const edges = (pd.edges ?? []) as Record<string, unknown>[];
                processDiagrams.push({
                  name:        String(pd.name ?? ''),
                  description: String(pd.description ?? ''),
                  mermaid:     SolutionZipParser.buildProcessDiagramMermaid(nodes, edges),
                });
              }
            }

            artifactIds = ((mainContent.artifactIds ?? []) as string[]);
          } catch (e) {
            errors.push({ archivePath: mainContentPath, error: `Plan content parse: ${String(e)}`, partial: true });
          }
        }

        // ── Artifact proposals ────────────────────────────────────────────────
        const artifacts: ManagedPlanArtifact[] = [];
        const artifactDirs = Object.keys(zip.files)
          .filter(p => p.startsWith(planBase + 'msdyn_planartifacts/') && p.endsWith('msdyn_planartifact.xml'));

        for (const artifactXmlPath of artifactDirs) {
          try {
            const aXml = xmlParser.parse(await zip.files[artifactXmlPath].async('string')) as Record<string, unknown>;
            const aRoot = (aXml.msdyn_planartifact ?? aXml) as Record<string, unknown>;
            const aId   = String(aRoot.msdyn_planartifactid ?? '').replace(/[{}]/g, '');
            const aName = String(aRoot.msdyn_name ?? '');
            const aType = String(aRoot.msdyn_type ?? '');

            // Read proposal JSON
            const proposalPath = artifactXmlPath.replace('msdyn_planartifact.xml', 'msdyn_proposal/content');
            let aDesc = '';
            let aTables: string[] = [];
            let aStories: string[] = [];
            if (zip.files[proposalPath]) {
              try {
                const proposal = JSON.parse(await zip.files[proposalPath].async('string')) as Record<string, unknown>;
                aDesc   = String(proposal.description ?? '');
                aTables = (proposal.tableSchemaNames ?? []) as string[];
                const us = (proposal.userStories ?? []) as Array<Record<string, unknown>>;
                aStories = us.map(s => String(s.description ?? '')).filter(Boolean);
              } catch { /* ignore */ }
            }
            artifacts.push({ id: aId, name: aName, type: aType, description: aDesc, tables: aTables, userStories: aStories });
          } catch (e) {
            errors.push({ archivePath: artifactXmlPath, error: String(e), partial: true });
          }
        }

        result.managedPlans!.push({
          planId, name, description: desc, originalPrompt: prompt,
          personas, artifacts, entities, processDiagrams, source: src,
        } as ManagedPlan);

      } catch (e) {
        errors.push({ archivePath: planXmlPath, error: String(e), partial: true });
      }
    }
  }

  // ── AI Builder models ──────────────────────────────────────────────────────────

  /** Map known template GUIDs to human-readable names. */
  private static readonly AIB_TEMPLATE_NAMES: Record<string, string> = {
    'edfdb190-3791-45d8-9a6c-8f90a37c278a': 'Custom Prompt',
    '08cddd26-7b2f-4eed-8ef4-3b0f43e4534d': 'Entity Extraction',
    '2c1d8025-e3d5-4b65-b0e6-14e0793a4e45': 'Binary Classification',
    'babc79f2-31b6-4fa1-a116-ab9ca79d3ecf': 'Category Classification',
    'f5fa92d3-06e0-452b-9b1b-6bb32e7b5ddc': 'Object Detection',
    '3b7e3cad-bfb8-49b3-830c-00fc2cb3e8ca': 'Document Processing',
    '89f9c29b-8da7-47e3-a6e4-e74b85e38543': 'Prediction',
  };

  private static async parseAiBuilderModels(
    zip: JSZip, archivePath: string,
    result: Partial<NormalizedAnalysisGraph>, errors: ParseError[]
  ): Promise<void> {
    // AI Builder models live in customizations.xml under <AIModels>
    const custFile = zip.file('customizations.xml');
    if (!custFile) return;

    let xml: Record<string, unknown>;
    try {
      const content = await custFile.async('string');
      xml = xmlParser.parse(content) as Record<string, unknown>;
    } catch (e) {
      errors.push({ archivePath: 'customizations.xml', error: `AIBuilder parse: ${String(e)}`, partial: true });
      return;
    }

    const root = (xml.ImportExportXml ?? xml) as Record<string, unknown>;
    const aiModelsNode = root.AIModels as Record<string, unknown> | undefined;
    if (!aiModelsNode) return;

    const modelsRaw = aiModelsNode.AIModel;
    if (!modelsRaw) return;
    const models: Record<string, unknown>[] = Array.isArray(modelsRaw) ? modelsRaw : [modelsRaw];

    for (const m of models) {
      try {
        const modelId = String(m.msdyn_aimodelid ?? '').replace(/[{}]/g, '');
        const name    = String(m.msdyn_name ?? 'Unnamed AI Model');
        const templateId = String(m.msdyn_templateid ?? '').replace(/[{}]/g, '');
        const templateName = SolutionZipParser.AIB_TEMPLATE_NAMES[templateId.toLowerCase()] ?? 'Unknown Template';

        // Status: statecode 1 = active
        const stateCode = Number(m.statecode ?? 0);
        const status: AiBuilderModel['status'] = stateCode === 1 ? 'active' : 'draft';

        // Parse the active configuration to extract prompt text and inputs
        let promptText = '';
        let inputs: AiBuilderInputVariable[] = [];
        let outputFormats: string[] = [];
        let modelType = '';

        const aiConfigs = (m as Record<string, unknown>).AIConfigurations as Record<string, unknown> | undefined;
        const configsRaw = aiConfigs?.AIConfiguration;
        if (configsRaw) {
          const configs: Record<string, unknown>[] = Array.isArray(configsRaw) ? configsRaw : [configsRaw];
          // Active config = statecode 2 AND statuscode 7 (published) → match activerunconfigurationid
          const activeConfigId = String(m.msdyn_activerunconfigurationid ?? '').replace(/[{}]/g, '');
          const activeConfig = configs.find(c =>
            String(c.msdyn_aiconfigurationid ?? '').replace(/[{}]/g, '') === activeConfigId.toLowerCase()
            || String(c.msdyn_aiconfigurationid ?? '').replace(/[{}]/g, '') === activeConfigId
          ) ?? configs[0];

          if (activeConfig?.msdyn_customconfiguration) {
            try {
              const cfg = JSON.parse(String(activeConfig.msdyn_customconfiguration)) as Record<string, unknown>;

              // Extract prompt text from GptDynamicPrompt prompt array
              const promptArr = cfg.prompt as Array<Record<string, unknown>> | undefined;
              if (Array.isArray(promptArr)) {
                promptText = promptArr
                  .filter(p => p.type === 'literal')
                  .map(p => String(p.text ?? ''))
                  .join('')
                  .trim();
              }

              // Inputs
              const defs = cfg.definitions as Record<string, unknown> | undefined;
              const inputsArr = defs?.inputs as Array<Record<string, unknown>> | undefined;
              if (Array.isArray(inputsArr)) {
                inputs = inputsArr.map(i => ({
                  id:          String(i.id ?? ''),
                  displayName: String(i.text ?? i.id ?? ''),
                  type:        String(i.type ?? 'text'),
                }));
              }

              // Output formats
              const output = defs?.output as Record<string, unknown> | undefined;
              const fmts = output?.formats as string[] | undefined;
              if (Array.isArray(fmts)) outputFormats = fmts;

              // Model type
              const mp = cfg.modelParameters as Record<string, unknown> | undefined;
              modelType = String(mp?.modelType ?? '');

            } catch { /* ignore parse errors in config JSON */ }
          }
        }

        result.aiBuilderModels!.push({
          modelId,
          name,
          templateId,
          templateName,
          status,
          promptText,
          inputs,
          outputFormats,
          modelType,
          source: { archivePath: archivePath + '/customizations.xml', confidence: 'high' },
        } as AiBuilderModel);

      } catch (e) {
        errors.push({ archivePath: 'customizations.xml', error: `AIModel parse: ${String(e)}`, partial: true });
      }
    }
  }

  // ── Dataverse formulas ────────────────────────────────────────────────────────

  private static async parseDataverseFormulas(
    zip: JSZip, archivePath: string,
    result: Partial<NormalizedAnalysisGraph>, errors: ParseError[]
  ): Promise<void> {
    const formulaFiles = Object.keys(zip.files).filter(e =>
      e.startsWith('Formulas/') && e.endsWith('.yaml') && !zip.files[e].dir
    );

    for (const filePath of formulaFiles) {
      try {
        const content = await zip.files[filePath].async('string');
        const doc = yaml.load(content) as Record<string, unknown>;
        if (!doc) continue;

        const tableName = filePath.replace('Formulas/', '').replace('.yaml', '');
        for (const [colName, expr] of Object.entries(doc)) {
          if (typeof expr !== 'string') continue;
          const fieldRefs = [...expr.matchAll(/\b([A-Za-z_][A-Za-z0-9_]*)\./g)].map(m => m[1]);
          result.dataverseFormulas!.push({
            tableName,
            columnName: colName,
            expression: expr,
            referencedFields: [...new Set(fieldRefs)],
            source: { archivePath: archivePath + '/' + filePath, confidence: 'high' }
          });
        }
      } catch (e: unknown) {
        errors.push({ archivePath: archivePath + '/' + filePath, error: String(e), partial: false });
      }
    }
  }
}
