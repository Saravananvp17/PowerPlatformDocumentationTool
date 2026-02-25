import JSZip from 'jszip';
import * as yaml from 'js-yaml';
import type {
  CanvasApp, Screen, Control, KeyFormula, FormulaExtract,
  Variable, Gallery, DataSource, ConnectorRef, AppCheckerFinding,
  SchemaColumn, SourceRef, ProgressCallback
} from '../model/types';
import { VariableExtractor } from '../analysis/VariableExtractor';
import { NavigationInferencer } from '../analysis/NavigationInferencer';
import { ButtonActionAnalyzer } from '../analysis/ButtonActionAnalyzer';

// ─────────────────────────────────────────────────────────────────────────────

const KEY_PATTERNS = ['Navigate', 'Patch', 'SubmitForm', 'Collect', 'ClearCollect',
  'Remove', 'ForAll', 'Run', 'UpdateContext', 'Set', 'Launch'];

// Property names that should ALWAYS be captured regardless of formula content.
// Items is the critical one: gallery data source formulas like "Filter(SP_List,...)"
// don't contain any KEY_PATTERNS entries but are essential for documentation.
const KEY_PROPERTY_NAMES = new Set(['Items', 'OnSelect', 'Default', 'OnChange']);

export class MsappParser {
  static async parse(
    msappBuffer: Buffer,
    msappPath: string,
    displayName: string,
    onProgress?: ProgressCallback
  ): Promise<CanvasApp> {
    const src: SourceRef = { archivePath: msappPath, confidence: 'high' };
    onProgress?.({ stage: 'msapp', pct: 0, message: `Parsing canvas app: ${displayName}` });

    const zip = await JSZip.loadAsync(msappBuffer);
    // Power Apps on Windows stores paths with backslashes inside the ZIP.
    // Build a normalised (forward-slash) → original key map so lookups work on both platforms.
    const pathMap = new Map<string, string>();
    for (const key of Object.keys(zip.files)) {
      pathMap.set(key.replace(/\\/g, '/'), key);
    }
    const entries = Array.from(pathMap.keys());

    // Try to get the friendly display name from Properties.json inside the .msapp
    let resolvedDisplayName = displayName;
    try {
      const propsFile = MsappParser.zipFile(zip, pathMap, 'Properties.json');
      if (propsFile) {
        const props = JSON.parse(await propsFile.async('string')) as Record<string, unknown>;
        const friendly = String(props.Name ?? props.DisplayName ?? props.displayName ?? '').trim();
        if (friendly) resolvedDisplayName = friendly;
      }
    } catch { /* non-fatal — keep the filename-derived name */ }

    // Detect format: PA YAML (modern) vs legacy JSON
    // Power Apps uses .pa.yaml, .fx.yaml, or plain .yaml depending on version
    const hasPaYaml = entries.some(e => {
      const lower = e.toLowerCase();
      return lower.startsWith('src/') && (
        lower.endsWith('.pa.yaml') || lower.endsWith('.fx.yaml') || lower.endsWith('.yaml')
      );
    });

    const app: CanvasApp = {
      id: msappPath,
      displayName: resolvedDisplayName,
      msappPath,
      appOnStart: { raw: '', redacted: '', source: { archivePath: msappPath + '/Src/App.pa.yaml', confidence: 'high' } },
      screens: [],
      variables: [],
      galleries: [],
      dataSources: [],
      connectors: [],
      navigationEdges: [],
      navConfidence: 'low',
      mermaidNavGraph: '',
      appCheckerFindings: [],
      source: src
    };

    onProgress?.({ stage: 'msapp', pct: 20, message: 'Parsing data sources...' });
    await MsappParser.parseDataSources(zip, pathMap, msappPath, app);

    onProgress?.({ stage: 'msapp', pct: 40, message: hasPaYaml ? 'Parsing PA YAML source...' : 'Parsing legacy JSON source...' });
    if (hasPaYaml) {
      await MsappParser.parsePaYaml(zip, pathMap, msappPath, app);
    } else {
      await MsappParser.parseLegacyJson(zip, pathMap, msappPath, app);
    }

    // Collect galleries from all screens into the app-level list
    app.galleries = app.screens.flatMap(s => s.galleries ?? []);

    onProgress?.({ stage: 'msapp', pct: 72, message: 'Analyzing button actions...' });
    ButtonActionAnalyzer.analyze(app);

    onProgress?.({ stage: 'msapp', pct: 75, message: 'Extracting variables...' });
    app.variables = VariableExtractor.extract(app);

    onProgress?.({ stage: 'msapp', pct: 85, message: 'Inferring navigation...' });
    const navResult = NavigationInferencer.infer(app);
    app.navigationEdges = navResult.edges;
    app.navConfidence = navResult.confidence;
    app.mermaidNavGraph = navResult.mermaid;

    onProgress?.({ stage: 'msapp', pct: 90, message: 'Parsing AppChecker results...' });
    await MsappParser.parseSarif(zip, pathMap, msappPath, app);

    onProgress?.({ stage: 'msapp', pct: 100, message: 'Canvas app parsed' });
    return app;
  }

  // ── DataSources.json ────────────────────────────────────────────────────────

  private static zipFile(zip: JSZip, pathMap: Map<string, string>, normalPath: string) {
    return zip.file(pathMap.get(normalPath) ?? normalPath);
  }

  private static async parseDataSources(zip: JSZip, pathMap: Map<string, string>, msappPath: string, app: CanvasApp): Promise<void> {
    const dsFile = MsappParser.zipFile(zip, pathMap, 'References/DataSources.json');
    if (!dsFile) return;

    try {
      const raw = JSON.parse(await dsFile.async('string'));
      const dataSources: Record<string, unknown>[] = Array.isArray(raw) ? raw : (raw.DataSources ?? raw.dataSources ?? []);
      const src: SourceRef = { archivePath: msappPath + '/References/DataSources.json', confidence: 'high' };

      for (const ds of dataSources) {
        const name   = String(ds.Name ?? ds.name ?? '');
        const dsType = String(ds.Type ?? ds.type ?? '');
        const apiId  = String(ds.ApiId ?? ds.apiId ?? '').toLowerCase();

        // ── SharePoint list (ConnectedDataSourceInfo via sharepointonline) ─────
        if (dsType === 'ConnectedDataSourceInfo' && apiId.includes('sharepointonline')) {
          const siteUrl  = String(ds.DatasetName ?? '');
          const tableGuid = String(ds.TableName ?? '');
          const columns  = MsappParser.extractSpColumns(ds);
          // Use the list's real title from schema if available, fall back to app-name
          const listTitle = columns.length > 0
            ? MsappParser.extractSpListTitle(ds) ?? name
            : name;

          app.dataSources.push({
            id: `sp:${name}`,
            type: 'sharepoint',
            displayName: name,
            listName: listTitle,
            siteUrl,
            tableName: tableGuid,
            columns: columns.length > 0 ? columns : undefined,
            usedBy: [], source: src
          });

        // ── Native Dataverse table (NativeCDSDataSourceInfo) ──────────────────
        } else if (dsType === 'NativeCDSDataSourceInfo') {
          const td = MsappParser.safeJsonParse(ds.TableDefinition);
          const tableName    = String(td?.TableName ?? ds.LogicalName ?? ds.EntitySetName ?? name);
          const environmentUrl = String(ds.DatasetName ?? '');
          const columns = MsappParser.extractCdsColumns(td);

          app.dataSources.push({
            id: `dv:${name}`,
            type: 'dataverse',
            displayName: name,
            logicalName: tableName,
            tableName,
            environmentUrl,
            columns: columns.length > 0 ? columns : undefined,
            usedBy: [], source: src
          });

        // ── Power Automate flow (ServiceInfo with FlowNameId) ─────────────────
        } else if (dsType === 'ServiceInfo' && ds.FlowNameId) {
          const flowId = String(ds.FlowNameId);
          // Register as a connector so ButtonActionAnalyzer can match by name
          const connId = `flow:${flowId}`;
          if (!app.connectors.find(c => c.connectorId === connId)) {
            app.connectors.push({
              connectorId: connId,
              displayName: name,
              usedInFlows: [],
              usedInApps: [],
              source: src
            });
          }

        // ── Standard connector (ServiceInfo without FlowNameId) ───────────────
        } else if (dsType === 'ServiceInfo') {
          // e.g. shared_office365users, shared_sharepointonline parent def
          const connId = apiId.split('/').pop() ?? name;
          if (connId && !app.connectors.find(c => c.connectorId === connId)) {
            app.connectors.push({
              connectorId: connId,
              displayName: name,
              usedInFlows: [],
              usedInApps: [],
              source: src
            });
          }

        // ── Legacy / unknown format fallback (older .msapp files use Kind) ────
        } else {
          const kind = String(ds.Kind ?? ds.kind ?? '').toLowerCase();
          if (kind.includes('sharepoint')) {
            app.dataSources.push({
              id: `sp:${name}`,
              type: 'sharepoint',
              displayName: name,
              siteUrl: String(ds.DatasetName ?? ''),
              listName: String(ds.TableName ?? ''),
              usedBy: [], source: src
            });
          } else if (kind.includes('dataverse') || kind.includes('common') || kind.includes('commondataservice')) {
            app.dataSources.push({
              id: `dv:${name}`,
              type: 'dataverse',
              displayName: name,
              tableName: String(ds.TableName ?? ''),
              environmentUrl: String(ds.DatasetName ?? ''),
              usedBy: [], source: src
            });
          }
        }
      }
    } catch (e) {
      // non-fatal
    }
  }

  /** Parse DataEntityMetadataJson and extract SchemaColumns for a SP list. */
  private static extractSpColumns(ds: Record<string, unknown>): SchemaColumn[] {
    try {
      const outer = ds.DataEntityMetadataJson as Record<string, string> | undefined;
      if (!outer || typeof outer !== 'object') return [];
      // Keyed by list GUID — take first (there's normally exactly one)
      const metaStr = Object.values(outer)[0];
      if (!metaStr) return [];
      const meta = JSON.parse(metaStr) as Record<string, unknown>;
      const props = (meta?.schema as Record<string, unknown>)
        ?.items as Record<string, unknown>;
      const properties = (props?.properties ?? {}) as Record<string, Record<string, unknown>>;

      const cols: SchemaColumn[] = [];
      for (const [colName, colDef] of Object.entries(properties)) {
        // Skip internal/virtual columns like {Identifier}, {IsFolder}, etc.
        if (colName.startsWith('{')) continue;
        // Skip pure-system read-only cols unlikely to be referenced in formulas
        const perm = String(colDef['x-ms-permission'] ?? 'read-write');
        const isKey = String(colDef['x-ms-keyType'] ?? '');
        // Keep ID (PK) and all read-write cols; skip other read-only system cols
        // unless they have a meaningful title that might appear in formulas
        if (perm === 'read-only' && !isKey) continue;

        const fmt = String(colDef['format'] ?? '');
        const typ = String(colDef['type'] ?? '');
        cols.push({
          name: colName,
          displayName: String(colDef['title'] ?? colName),
          type: fmt || typ,
          required: false,
          notes: colDef['description'] ? String(colDef['description']).replace(/\n/g, ' ').trim() : undefined
        });
      }
      return cols;
    } catch {
      return [];
    }
  }

  /** Extract the real SharePoint list title from DataEntityMetadataJson. */
  private static extractSpListTitle(ds: Record<string, unknown>): string | undefined {
    try {
      const outer = ds.DataEntityMetadataJson as Record<string, string> | undefined;
      if (!outer || typeof outer !== 'object') return undefined;
      const metaStr = Object.values(outer)[0];
      if (!metaStr) return undefined;
      const meta = JSON.parse(metaStr) as Record<string, unknown>;
      return meta?.title ? String(meta.title) : undefined;
    } catch {
      return undefined;
    }
  }

  /** Parse NativeCDS TableDefinition → EntityMetadata → Attributes → SchemaColumns. */
  private static extractCdsColumns(td: Record<string, unknown> | null): SchemaColumn[] {
    try {
      if (!td) return [];
      const entityMeta = MsappParser.safeJsonParse(td.EntityMetadata);
      if (!entityMeta) return [];
      const attrs = Array.isArray(entityMeta.Attributes) ? entityMeta.Attributes as Record<string, unknown>[] : [];
      const cols: SchemaColumn[] = [];
      for (const attr of attrs) {
        const logicalName = String(attr.LogicalName ?? '');
        if (!logicalName) continue;
        if (!attr.IsValidForRead) continue;
        // Skip internal/relationship attributes that clutter the schema
        if (String(attr.AttributeType ?? '').toLowerCase() === 'virtual') continue;

        const displayName = (attr.DisplayName as Record<string, unknown>)
          ?.UserLocalizedLabel
          ? String(((attr.DisplayName as Record<string, unknown>).UserLocalizedLabel as Record<string, unknown>).Label ?? logicalName)
          : logicalName;
        const attrType = String(attr.AttributeType ?? '');
        cols.push({
          name: logicalName,
          displayName,
          type: attrType,
          required: attr.RequiredLevel
            ? String((attr.RequiredLevel as Record<string, unknown>).Value ?? '') === 'ApplicationRequired'
            : false
        });
      }
      return cols;
    } catch {
      return [];
    }
  }

  /** Safely parse a JSON string or return null. */
  private static safeJsonParse(val: unknown): Record<string, unknown> | null {
    if (!val) return null;
    if (typeof val === 'object') return val as Record<string, unknown>;
    try { return JSON.parse(String(val)) as Record<string, unknown>; }
    catch { return null; }
  }

  // ── PA YAML parsing ─────────────────────────────────────────────────────────

  private static async parsePaYaml(zip: JSZip, pathMap: Map<string, string>, msappPath: string, app: CanvasApp): Promise<void> {
    // Accept .pa.yaml, .fx.yaml, and .yaml — Power Apps uses all three across versions
    const yamlFiles = Array.from(pathMap.keys()).filter(e => {
      const lower = e.toLowerCase();
      return lower.startsWith('src/') && (
        lower.endsWith('.pa.yaml') || lower.endsWith('.fx.yaml') || lower.endsWith('.yaml')
      );
    });

    for (const filePath of yamlFiles) {
      const file = MsappParser.zipFile(zip, pathMap, filePath);
      if (!file) continue;
      try {
        const content = await file.async('string');
        const src: SourceRef = { archivePath: msappPath + '/' + filePath, confidence: 'high' };
        const fileName = filePath.split('/').pop() ?? '';
        const fileBase = fileName.replace(/\.(pa|fx)?\.yaml$/i, '');

        let doc: unknown;
        try { doc = yaml.load(content); } catch { continue; }
        if (!doc || typeof doc !== 'object') continue;
        const docObj = doc as Record<string, unknown>;

        // App-level file (App.pa.yaml, App.fx.yaml, App.yaml)
        if (fileBase.toLowerCase() === 'app') {
          const appNode = (docObj.App ?? docObj) as Record<string, unknown>;
          const props = (appNode.Properties ?? appNode) as Record<string, unknown>;
          const onStartRaw = String(props.OnStart ?? '');
          if (onStartRaw) {
            app.appOnStart = { raw: onStartRaw, redacted: onStartRaw, source: src };
          }
          continue;
        }

        // ── Primary format (confirmed from real solution exports) ──────────────
        // Top-level key is "Screens", value is a dict: { screenName: { Properties, Children } }
        // e.g. SCR_Home.pa.yaml contains: Screens: { SCR_Home: { Properties: ..., Children: [...] } }
        if ('Screens' in docObj && docObj.Screens && typeof docObj.Screens === 'object') {
          const screens = docObj.Screens as Record<string, unknown>;
          for (const [screenName, screenData] of Object.entries(screens)) {
            if (!screenData || typeof screenData !== 'object') continue;
            if (!app.screens.find(s => s.name === screenName)) {
              app.screens.push(
                MsappParser.parseScreenYaml(screenData as Record<string, unknown>, screenName, src, msappPath)
              );
            }
          }
          continue;
        }

        // ── Fallback: "ScreenName As Screen.xxx" top-level key ─────────────────
        const topKey = Object.keys(docObj)[0] ?? '';
        const asMatch = topKey.match(/^(.+?)\s+[Aa]s\s+[Ss]creen/i);
        if (asMatch) {
          const screenName = asMatch[1].trim();
          const screenData = (docObj[topKey] ?? {}) as Record<string, unknown>;
          if (!app.screens.find(s => s.name === screenName)) {
            app.screens.push(MsappParser.parseScreenYaml(screenData, screenName, src, msappPath));
          }
          continue;
        }

        // ── Fallback: array of control objects ─────────────────────────────────
        if (Array.isArray(doc)) {
          for (const item of doc as Record<string, unknown>[]) {
            const asVal = String(item.As ?? item.type ?? '').toLowerCase();
            if (asVal.includes('screen')) {
              const screenName = String(item.Name ?? item.name ?? fileBase);
              if (!app.screens.find(s => s.name === screenName)) {
                app.screens.push(MsappParser.parseScreenYaml(item, screenName, src, msappPath));
              }
            }
          }
          continue;
        }

        // ── Last resort: single object with Properties/Children ────────────────
        const screenData = (docObj[topKey] ?? docObj) as Record<string, unknown>;
        if (screenData.Properties || screenData.Children) {
          const screenName = String(screenData.Name ?? screenData.name ?? (topKey || fileBase));
          if (!app.screens.find(s => s.name === screenName)) {
            app.screens.push(MsappParser.parseScreenYaml(screenData, screenName, src, msappPath));
          }
        }
      } catch { /* parsing error — skip this file gracefully */ }
    }
  }

  private static parseScreenYaml(
    doc: Record<string, unknown>, name: string, src: SourceRef, msappPath: string
  ): Screen {
    const props = (doc.Properties ?? {}) as Record<string, unknown>;
    const onVisibleRaw = String(props.OnVisible ?? '');
    const screenSrc: SourceRef = { ...src, yamlPath: `Screens.${name}.OnVisible` };

    const keyFormulas: KeyFormula[] = [];
    for (const [prop, val] of Object.entries(props)) {
      if (typeof val === 'string' && val.trim()) {
        const patterns = KEY_PATTERNS.filter(p => val.includes(p));
        if (patterns.length > 0 || ['OnVisible', 'OnHidden', 'OnStart'].includes(prop)) {
          keyFormulas.push({
            controlName: name,
            property: prop,
            formula: { raw: val, redacted: val, source: { ...src, yamlPath: `Screens.${name}.${prop}` } },
            patterns
          });
        }
      }
    }

    const controls: Control[] = [];
    const children = (doc.Children ?? doc.Tree ?? []) as unknown[];
    if (Array.isArray(children)) {
      for (const child of children) {
        const c = child as Record<string, unknown>;
        // Pass screenName as the top-level parent; extractControls will track nested parents
        controls.push(...MsappParser.extractControls(c, name, src, msappPath, name));
      }
    }

    // Build parent → direct-children map so we can find all gallery descendants
    const childrenByParent = new Map<string, Control[]>();
    for (const ctrl of controls) {
      const p = ctrl.parent ?? name;
      if (!childrenByParent.has(p)) childrenByParent.set(p, []);
      childrenByParent.get(p)!.push(ctrl);
    }
    // Recursively collect all descendants of a given control name
    const getDescendants = (ctrlName: string): Control[] => {
      const direct = childrenByParent.get(ctrlName) ?? [];
      return direct.flatMap(ch => [ch, ...getDescendants(ch.name)]);
    };

    // Extract galleries — populate Items formula, inferred data sources, and ThisItem field refs
    const galleries: Gallery[] = controls
      .filter(c => c.type.toLowerCase().includes('gallery'))
      .map(c => {
        const itemsKf = c.keyFormulas.find((f: KeyFormula) => f.property === 'Items');
        const itemsFormula: FormulaExtract = itemsKf?.formula ?? { raw: '', redacted: '', source: src };

        // Infer primary data source: first identifier in the Items formula.
        // Handles bare names, single-quoted names, and nested wrappers:
        //   "MySPList"
        //   "'My SP List'"
        //   "Filter(MySPList, IsActive)"
        //   "Sort(Filter('My SP List', ...), ...)"
        //   "=Sort(Filter(MySPList, ...), ...)"  ← PA YAML prefix
        const inferredDataSources: string[] = [];
        // PA YAML formulas often start with '=' — strip it first, then normalise whitespace
        const formulaText = (itemsFormula.redacted || itemsFormula.raw)
          .replace(/^[\s=]+/, '').trim();
        if (formulaText) {
          // Known PA wrapper functions whose first argument IS the data source
          const WRAPPER_RE = /^(?:Sort|SortByColumns|Filter|Search|Distinct|FirstN|LastN|ShowColumns|DropColumns|AddColumns|RenameColumns|CountRows|CountIf|LookUp|If|Coalesce|Table)\s*\(\s*/i;
          // PA built-in functions to skip when scanning for identifiers
          const BUILTIN_FNS = new Set([
            'Sort','SortByColumns','Filter','Search','Distinct','FirstN','LastN',
            'ShowColumns','DropColumns','AddColumns','RenameColumns','CountRows',
            'CountIf','LookUp','If','Coalesce','Table','And','Or','Not','IsBlank',
            'IsEmpty','Value','Text','Date','DateDiff','Now','Today','Concatenate',
            'Left','Right','Mid','Len','Lower','Upper','Trim','Replace','Substitute',
          ]);

          // Strategy 1: iteratively strip leading wrappers to reach the DS identifier
          let s = formulaText;
          for (let i = 0; i < 6; i++) {
            const next = s.replace(WRAPPER_RE, '');
            if (next === s) break;
            s = next.trim();
          }
          // Single-quoted identifier (can contain spaces): 'My SP List'
          const quotedMatch = s.match(/^'([^']+)'/);
          // Plain identifier (no spaces, not a known function): MySPList
          const plainMatch = s.match(/^([A-Za-z_\u00C0-\u017E][A-Za-z0-9_\u00C0-\u017E]*)/);
          const candidate = quotedMatch?.[1]?.trim() ?? plainMatch?.[1]?.trim();
          if (candidate && !BUILTIN_FNS.has(candidate)) {
            inferredDataSources.push(candidate);
          } else {
            // Strategy 2 fallback: scan entire formula for first quoted or plain identifier
            // that isn't a known built-in function — catches unusual formula shapes
            const quotedFallback = formulaText.match(/'([^']+)'/);
            if (quotedFallback?.[1]) {
              inferredDataSources.push(quotedFallback[1].trim());
            } else {
              const allIdents = formulaText.matchAll(/\b([A-Za-z_\u00C0-\u017E][A-Za-z0-9_\u00C0-\u017E]*)\b/g);
              for (const m of allIdents) {
                if (!BUILTIN_FNS.has(m[1]) && m[1] !== 'ThisItem') {
                  inferredDataSources.push(m[1]);
                  break;
                }
              }
            }
          }
        }

        // Scan the gallery control + all its descendant controls for ThisItem.FieldName refs.
        // Template item child controls (labels, images, etc.) are where these refs live.
        const galleryDescendants = getDescendants(c.name);
        const allFormulas = [c, ...galleryDescendants]
          .flatMap(ctrl => ctrl.keyFormulas)
          .map(kf => kf.formula.redacted || kf.formula.raw)
          .join('\n');
        const thisItemRe = /\bThisItem\s*\.\s*'?([A-Za-z_\u00C0-\u017E][A-Za-z0-9_ \u00C0-\u017E()\-]*)'?/g;
        const fieldsUsed: string[] = [];
        let tm: RegExpExecArray | null;
        while ((tm = thisItemRe.exec(allFormulas)) !== null) {
          const field = tm[1].trim();
          if (!fieldsUsed.includes(field)) fieldsUsed.push(field);
        }

        return {
          controlName: c.name,
          screenName: name,
          galleryType: c.type,
          itemsFormula,
          inferredDataSources,
          fieldsUsed,
          keyFormulas: c.keyFormulas,
          navEdges: [],
          source: src
        };
      });

    return {
      name,
      onVisible: { raw: onVisibleRaw, redacted: onVisibleRaw, source: screenSrc },
      keyFormulas,
      controls,
      galleries,
      buttonActions: [],   // populated later by ButtonActionAnalyzer
      source: src
    };
  }

  private static extractControls(
    node: Record<string, unknown>, screenName: string, src: SourceRef, _msappPath: string,
    directParent?: string,  // actual parent control name (undefined = direct child of screen)
    captureAll?: boolean    // true when inside a gallery template — capture every property
  ): Control[] {
    // Two child formats exist in Power Apps PA YAML:
    //
    // Format A – legacy/fallback:  { Name: "ctrl", As: "type", Properties: {...}, Children: [...] }
    //
    // Format B – real solution exports (confirmed):
    //   { ControlName: { Control: "Rectangle@2.3.0", Properties: {...}, Children: [...] } }
    //   The control name is the sole dict key; all data lives under it.

    let name: string;
    let type: string;
    let props: Record<string, unknown>;
    let childrenArr: unknown[];

    if ('Name' in node || 'As' in node || 'Properties' in node) {
      // Format A
      name = String(node.Name ?? node.name ?? '');
      type = String(node.As ?? node.Type ?? node.type ?? 'Unknown');
      props = (node.Properties ?? {}) as Record<string, unknown>;
      childrenArr = (node.Children ?? []) as unknown[];
    } else {
      // Format B: single top-level key = control name
      const ctrlName = Object.keys(node)[0] ?? '';
      const ctrlData = (node[ctrlName] ?? {}) as Record<string, unknown>;
      name = ctrlName;
      // "Control" field holds "Type@version" — strip the version suffix for readability
      type = String(ctrlData.Control ?? ctrlData.Type ?? ctrlData.type ?? 'Unknown')
        .replace(/@[0-9.]+$/, '');
      props = (ctrlData.Properties ?? {}) as Record<string, unknown>;
      childrenArr = (ctrlData.Children ?? []) as unknown[];
    }

    const isGallery = type.toLowerCase().includes('gallery');

    const keyFormulas: KeyFormula[] = [];
    for (const [prop, val] of Object.entries(props)) {
      if (typeof val === 'string' && val.trim()) {
        const patterns = KEY_PATTERNS.filter(p => val.includes(p));
        // captureAll = we're inside a gallery template; capture everything so ThisItem refs
        // in Text, Image, Color, Visible, etc. are all available for field scanning.
        if (patterns.length > 0 || KEY_PROPERTY_NAMES.has(prop) || captureAll) {
          keyFormulas.push({
            controlName: name,
            property: prop,
            formula: { raw: val, redacted: val, source: { ...src, yamlPath: `${screenName}.${name}.${prop}` } },
            patterns
          });
        }
      }
    }

    const controls: Control[] = [{ name, type, parent: directParent ?? screenName, keyFormulas }];
    if (Array.isArray(childrenArr)) {
      for (const child of childrenArr) {
        // Pass this control's name as the parent so descendants can be located.
        // If this is a gallery (or we're already inside one), children get captureAll=true.
        controls.push(...MsappParser.extractControls(child as Record<string, unknown>, screenName, src, _msappPath, name, captureAll || isGallery));
      }
    }
    return controls;
  }

  // ── Legacy JSON parsing ─────────────────────────────────────────────────────

  private static async parseLegacyJson(zip: JSZip, pathMap: Map<string, string>, msappPath: string, app: CanvasApp): Promise<void> {
    const srcFiles = Array.from(pathMap.keys()).filter(e =>
      e.startsWith('Controls/') && e.endsWith('.json')
    );

    for (const filePath of srcFiles) {
      const file = MsappParser.zipFile(zip, pathMap, filePath);
      if (!file) continue;
      try {
        const raw = JSON.parse(await file.async('string'));
        const src: SourceRef = { archivePath: msappPath + '/' + filePath, confidence: 'medium' };

        // Top-level control could be a screen
        const topControl = raw.TopParent ?? raw;
        if (!topControl) continue;

        const screenName = String(topControl.Name ?? filePath.replace(/.*\//, '').replace('.json', ''));
        const template = topControl.Template?.Name ?? '';
        const isScreen = template.toLowerCase() === 'screen' || !template;

        if (isScreen) {
          const props = topControl.Rules ?? {};
          const onVisibleRule = Array.isArray(props) ? props.find((r: Record<string,unknown>) => r.Property === 'OnVisible') : null;
          const onVisibleRaw = String(onVisibleRule?.InvariantScript ?? '');

          const keyFormulas: KeyFormula[] = [];
          const rules: Record<string,unknown>[] = Array.isArray(props) ? props : [];
          for (const rule of rules) {
            const script = String(rule.InvariantScript ?? '');
            const patterns = KEY_PATTERNS.filter(p => script.includes(p));
            if (patterns.length > 0) {
              keyFormulas.push({
                controlName: screenName,
                property: String(rule.Property ?? ''),
                formula: { raw: script, redacted: script, source: src },
                patterns
              });
            }
          }

          app.screens.push({
            name: screenName,
            onVisible: { raw: onVisibleRaw, redacted: onVisibleRaw, source: src },
            keyFormulas,
            controls: [],
            galleries: [],
            buttonActions: [],   // populated later by ButtonActionAnalyzer
            source: src
          });
        }
      } catch { /* skip */ }
    }
  }

  // ── SARIF / AppChecker ──────────────────────────────────────────────────────

  private static async parseSarif(zip: JSZip, pathMap: Map<string, string>, msappPath: string, app: CanvasApp): Promise<void> {
    const sarifFile = Array.from(pathMap.keys()).find(e => e.endsWith('.sarif'));
    if (!sarifFile) return;

    const file = MsappParser.zipFile(zip, pathMap, sarifFile);
    if (!file) return;
    try {
      const sarif = JSON.parse(await file.async('string'));
      const runs: Record<string, unknown>[] = sarif.runs ?? [];

      // Build a rule description map from tool.driver.rules for template-based messages
      const ruleDescriptions: Map<string, string> = new Map();
      for (const run of runs) {
        const rules: Record<string, unknown>[] =
          ((run.tool as Record<string, unknown>)?.driver as Record<string, unknown>)?.rules as Record<string, unknown>[] ?? [];
        for (const rule of rules) {
          const id = String(rule.id ?? '');
          const desc = String(
            (rule.shortDescription as Record<string, unknown>)?.text ??
            (rule.fullDescription as Record<string, unknown>)?.text ??
            (rule.messageStrings as Record<string, unknown>)?.default ??
            ''
          );
          if (id && desc) ruleDescriptions.set(id, desc);
        }
      }

      for (const run of runs) {
        for (const result of (run.results as Record<string, unknown>[] ?? [])) {
          const locs = result.locations as Record<string, unknown>[] | undefined;
          const loc = locs?.[0] as Record<string, unknown> | undefined;

          // ── Location ─────────────────────────────────────────────────────────
          // Power Apps AppChecker stores the control path in logicalLocations
          // e.g. { fullyQualifiedName: "SCR_Home.IMG_English.AccessibleLabel" }
          // Fall back to physicalLocation.artifactLocation.uri for standard SARIF.
          const logicalLocs = loc?.logicalLocations as Record<string, unknown>[] | undefined;
          const logicalLoc  = logicalLocs?.[0] as Record<string, unknown> | undefined;
          const logicalName = String(logicalLoc?.fullyQualifiedName ?? logicalLoc?.name ?? '');

          const physLoc    = loc?.physicalLocation as Record<string, unknown> | undefined;
          const artifactLoc = physLoc?.artifactLocation as Record<string, unknown> | undefined;
          const region      = physLoc?.region as Record<string, unknown> | undefined;
          const physUri     = String(artifactLoc?.uri ?? artifactLoc?.uriBaseId ?? '');
          const physLine    = region?.startLine ? `:${String(region.startLine)}` : '';
          const physRef     = physUri ? physUri + physLine : '';

          const location = logicalName || physRef || undefined;

          // ── Level ────────────────────────────────────────────────────────────
          // Per SARIF spec, absent level defaults to "warning".
          // Handle mixed case (some tools emit "Warning" / "Error").
          const rawLevel = String(result.level ?? 'warning').toLowerCase();
          const level: 'error' | 'warning' | 'note' =
            rawLevel === 'error' ? 'error' : rawLevel === 'warning' ? 'warning' : 'note';

          // ── Message ──────────────────────────────────────────────────────────
          // message.text is standard; fall back to rule description for template messages
          const msgObj = result.message as Record<string, unknown> | string | undefined;
          const msgText = typeof msgObj === 'string'
            ? msgObj
            : String(
                (msgObj as Record<string, unknown>)?.text ??
                (msgObj as Record<string, unknown>)?.markdown ??
                ruleDescriptions.get(String(result.ruleId ?? '')) ??
                ''
              );

          app.appCheckerFindings.push({
            ruleId: String(result.ruleId ?? ''),
            level,
            message: msgText,
            location
          });
        }
      }
    } catch { /* non-fatal */ }
  }
}
