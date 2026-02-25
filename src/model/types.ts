// ─────────────────────────────────────────────────────────────────────────────
// Core types for the Power Platform Documentation Tool
// ─────────────────────────────────────────────────────────────────────────────

export type Environment = 'DEV' | 'TEST' | 'PROD';
export type ArtifactType = 'solution-zip' | 'flow-zip' | 'powerapp-zip' | 'msapp';
export type Confidence = 'high' | 'medium' | 'low';

// ── Traceability ─────────────────────────────────────────────────────────────

export interface SourceRef {
  archivePath: string;       // e.g. "Workflows/MyFlow.json"
  jsonPath?: string;          // e.g. "$.actions.Send_email.inputs.to"
  yamlPath?: string;          // e.g. "Screens.Home.OnVisible"
  xmlXPath?: string;          // e.g. "/ImportExportXml/Entities/Entity[1]"
  lineNumber?: number;
  confidence: Confidence;
}

// ── Meta ─────────────────────────────────────────────────────────────────────

export interface ArtifactMeta {
  fileName: string;
  artifactType: ArtifactType;
  environment: Environment;
  analysedAt: string;   // ISO 8601
  toolVersion: string;
  fileSizeBytes: number;
}

// ── Solution ──────────────────────────────────────────────────────────────────

export interface Publisher {
  uniqueName: string;
  displayName: string;
  prefix: string;
}

export interface SolutionInfo {
  uniqueName: string;
  displayName: string;
  version: string;
  isManaged: boolean;
  publisher: Publisher;
  description: string;
  source: SourceRef;
  componentCounts: Record<string, number>;
}

// ── Environment Variables ─────────────────────────────────────────────────────

export interface EnvVar {
  id: string;
  schemaName: string;
  displayName: string;
  description: string;
  type: string;            // string | number | boolean | json | datasource | secret
  defaultValue?: string;
  currentValue?: string;   // "not included in export" if absent
  usedBy: SourceRef[];
  source: SourceRef;
}

// ── Connectors ────────────────────────────────────────────────────────────────

export interface ConnectorRef {
  connectorId: string;     // e.g. "shared_sharepointonline"
  displayName: string;
  usedInFlows: string[];
  usedInApps: string[];
  source: SourceRef;
}

export interface ConnectionReference {
  logicalName: string;
  displayName: string;
  connectorId: string;
  connectionReferenceType: string;
  source: SourceRef;
}

// ── Data Sources ──────────────────────────────────────────────────────────────

export type DataSourceType = 'sharepoint' | 'dataverse' | 'excel' | 'sql' | 'other';

export interface SchemaColumn {
  name: string;
  displayName?: string;
  type?: string;
  required?: boolean;
  notes?: string;
}

export interface SchemaEnrichment {
  fileName: string;
  format: 'csv' | 'xlsx' | 'json';
  columns: SchemaColumn[];
  rowCount?: number;
}

export interface DataSource {
  id: string;
  type: DataSourceType;
  displayName: string;
  logicalName?: string;
  siteUrl?: string;       // SharePoint
  listName?: string;      // SharePoint
  tableName?: string;     // Dataverse / Excel
  environmentUrl?: string;
  schemaEnrichment?: SchemaEnrichment;
  columns?: SchemaColumn[];
  usedBy: SourceRef[];
  source: SourceRef;
}

// ── Security Roles ────────────────────────────────────────────────────────────

export interface RolePrivilege {
  entityName: string;
  privilegeName: string;
  depth: string;  // 'None' | 'User' | 'BusinessUnit' | 'ParentChildBusinessUnit' | 'Organization'
}

export interface SecurityRole {
  id: string;
  name: string;
  templateId?: string;
  privilegeCount: number;
  privileges: RolePrivilege[];
  assignedToApps: string[];   // model-driven app names
  source: SourceRef;
}

// ── Managed Plans (msdyn_plans — AI-generated solution blueprints) ────────────

export interface ManagedPlanPersona {
  id: string;
  name: string;
  description: string;
  userStories: string[];    // story description strings for this persona
}

export interface ManagedPlanArtifact {
  id: string;
  name: string;
  type: string;             // 'PowerAutomateFlow', 'CopilotStudioAgent', 'PowerAppsModelApp', 'PowerBIReport', etc.
  description: string;
  tables: string[];         // tableSchemaNames referenced
  userStories: string[];    // user story descriptions
}

export interface ManagedPlanEntity {
  schemaName: string;
  displayName: string;
  description: string;
  attributes: { name: string; type: string; description: string }[];
}

export interface ManagedPlanProcessDiagram {
  name: string;
  description: string;
  mermaid: string;          // auto-generated from nodes/edges
}

export interface ManagedPlan {
  planId: string;
  name: string;
  description: string;
  originalPrompt: string;   // the natural-language prompt that generated this plan
  personas: ManagedPlanPersona[];
  artifacts: ManagedPlanArtifact[];
  entities: ManagedPlanEntity[];
  processDiagrams: ManagedPlanProcessDiagram[];
  source: SourceRef;
}

// ── Copilot Studio Agents ─────────────────────────────────────────────────────

export interface CopilotTopic {
  schemaName: string;
  name: string;
  description: string;
  triggerKind: string;         // e.g. 'OnRecognizedIntent', 'OnSystemRedirect', 'OnError'
  triggerDisplayName: string;  // from intent.displayName in the YAML
  triggerQueries: string[];    // sample trigger phrases
  actionKinds: string[];       // unique action kinds used in this topic
  isActive: boolean;
  source: SourceRef;
}

export interface CopilotAgent {
  schemaName: string;
  displayName: string;
  channels: string[];          // e.g. ['MsTeams', 'Microsoft365Copilot']
  authMode: string;            // 'None' | 'Integrated (Teams/AAD)' | 'Manual / OAuth'
  aiSettings: {
    generativeActionsEnabled: boolean;
    useModelKnowledge: boolean;
    fileAnalysisEnabled: boolean;
    semanticSearchEnabled: boolean;
  };
  topics: CopilotTopic[];
  gptInstructions: string;     // system-prompt / GPT component instructions
  knowledgeFiles: string[];    // names of attached knowledge documents
  source: SourceRef;
}

// ── Web Resources ─────────────────────────────────────────────────────────────

export interface WebResource {
  name: string;
  displayName: string;
  type: string;   // 1=HTML, 2=CSS, 3=JS, 4=XML, 5=PNG, 6=JPG, 7=GIF, 8=XAP, 9=XSL, 10=ICO, 11=SVG, 12=RESX
  usedBy: string[];
  source: SourceRef;
}

// ── Dataverse Formulas ────────────────────────────────────────────────────────

export interface DataverseFormula {
  tableName: string;
  columnName: string;
  expression: string;
  referencedFields: string[];
  source: SourceRef;
}

// ── URLs ──────────────────────────────────────────────────────────────────────

export type UrlCategory = 'sharepoint' | 'dataverse' | 'graph-teams' | 'external-http' | 'local-dev' | 'other';

export interface UrlReference {
  url: string;
  category: UrlCategory;
  isEnvSpecific: boolean;
  flagReason?: string;
  usedBy: SourceRef[];
}

// ── Canvas App — Controls & Formulas ─────────────────────────────────────────

export interface FormulaExtract {
  raw: string;
  redacted: string;
  source: SourceRef;
}

export interface KeyFormula {
  controlName: string;
  property: string;    // OnSelect, OnVisible, Items, etc.
  formula: FormulaExtract;
  patterns: string[]; // e.g. ["Navigate", "Patch"]
}

export interface Variable {
  name: string;
  kind: 'global' | 'context' | 'collection';
  setAt: SourceRef[];
  usedAt: SourceRef[];
}

export interface Gallery {
  controlName: string;
  screenName: string;
  galleryType?: string;
  itemsFormula: FormulaExtract;
  inferredDataSources: string[];
  templateSize?: string;
  wrapCount?: string;
  fieldsUsed: string[];          // ThisItem.FieldName references
  keyFormulas: KeyFormula[];
  navEdges: NavEdge[];
  source: SourceRef;
}

export interface Control {
  name: string;
  type: string;
  parent?: string;
  keyFormulas: KeyFormula[];
}

// ── Button / Control Action Analysis ─────────────────────────────────────────

export type ActionKind = 'patch' | 'flow-run' | 'navigate' | 'submit-form' | 'collect' | 'remove' | 'other';

export interface ActionStep {
  kind: ActionKind;
  target: string;       // data source name, screen, flow name, form, collection …
  payload?: string;     // truncated argument list — shows what data is passed
}

export interface ButtonAction {
  controlName: string;
  controlType: string;
  property: string;         // OnSelect, OnChange, OnSuccess …
  screenName: string;
  actions: ActionStep[];
  formulaSnippet: string;   // first 300 chars of the raw formula
  source: SourceRef;
}

export interface Screen {
  name: string;
  onVisible: FormulaExtract;
  keyFormulas: KeyFormula[];
  controls: Control[];
  buttonActions: ButtonAction[];  // extracted action patterns from control formulas
  galleries: Gallery[];
  source: SourceRef;
}

export interface NavEdge {
  from: string;       // screen or control name
  to: string;
  confidence: Confidence;
  triggeredBy?: string;
}

export interface AppCheckerFinding {
  ruleId: string;
  level: 'error' | 'warning' | 'note';
  message: string;
  location?: string;
}

export interface CanvasApp {
  id: string;
  displayName: string;
  uniqueName?: string;
  msappPath: string;
  appOnStart: FormulaExtract;
  screens: Screen[];
  variables: Variable[];
  galleries: Gallery[];
  dataSources: DataSource[];
  connectors: ConnectorRef[];
  navigationEdges: NavEdge[];
  navConfidence: Confidence;
  mermaidNavGraph: string;
  appCheckerFindings: AppCheckerFinding[];
  source: SourceRef;
}

// ── Model-Driven Apps ─────────────────────────────────────────────────────────

export interface SitemapSubArea {
  id: string;
  type: string;     // 'Entity' | 'Dashboard' | 'URL' | 'WebResource' | 'Custom'
  title: string;
  entityName?: string;
  url?: string;
}

export interface SitemapGroup {
  id: string;
  title: string;
  subAreas: SitemapSubArea[];
}

export interface SitemapArea {
  id: string;
  title: string;
  groups: SitemapGroup[];
}

export interface ModelDrivenApp {
  id: string;
  uniqueName: string;
  displayName: string;
  description?: string;
  assignedRoleIds: string[];
  assignedRoleNames: string[];    // resolved if available
  sitemapAreas: SitemapArea[];
  exposedTables: string[];
  mermaidSitemap: string;
  source: SourceRef;
}

// ── Power Automate Flows ──────────────────────────────────────────────────────

export type TriggerType = 'recurrence' | 'automated' | 'manual' | 'http' | 'child' | 'powerApps' | 'other';

export interface RecurrenceTrigger {
  frequency: string;
  interval: number;
  timeZone?: string;
  startTime?: string;
}

export interface AutomatedTrigger {
  connectorId: string;
  operationId: string;
  event?: string;
}

export interface TriggerDetail {
  type: TriggerType;
  recurrence?: RecurrenceTrigger;
  automated?: AutomatedTrigger;
  raw: Record<string, unknown>;
  source: SourceRef;
}

export interface RetryPolicy {
  type: string;
  count?: number;
  interval?: string;
}

// ── AI Builder ────────────────────────────────────────────────────────────────

export type AiBuilderModelStatus = 'active' | 'training' | 'draft' | 'published';

export interface AiBuilderInputVariable {
  id: string;
  displayName: string;
  type: string;   // 'text', 'number', etc.
}

export interface AiBuilderModel {
  modelId: string;
  name: string;
  templateId: string;        // GUID → lookup to human label
  templateName: string;      // e.g. 'Custom Prompt', 'Entity Extraction', etc.
  status: AiBuilderModelStatus;
  promptText: string;        // decoded from msdyn_customconfiguration
  inputs: AiBuilderInputVariable[];
  outputFormats: string[];   // e.g. ['text', 'json']
  modelType: string;         // e.g. 'reasoning', 'gpt4o'
  source: SourceRef;
}

export interface FlowAction {
  id: string;
  displayName: string;
  type: string;
  connector?: string;
  operationId?: string;
  aiBuilderModelId?: string;   // set when action calls AI Builder (references AiBuilderModel.modelId)
  inputs: Record<string, unknown>;
  redactedInputs: Record<string, unknown>;
  runAfter: string[];
  parentScope?: string;
  retryPolicy?: RetryPolicy;
  secureInputs: boolean;
  secureOutputs: boolean;
  source: SourceRef;
}

export interface ErrorHandlingSummary {
  hasExplicitHandling: boolean;
  handlingLocations: string[];
  scopesWithRunAfterFailure: string[];
}

export interface Flow {
  id: string;
  name: string;
  displayName: string;
  description?: string;
  state: string;
  triggerType: TriggerType;
  trigger: TriggerDetail;
  actions: FlowAction[];
  mermaidDiagram: string;
  connectors: ConnectorRef[];
  dataSources: DataSource[];
  urls: UrlReference[];
  errorHandling: ErrorHandlingSummary;
  source: SourceRef;
}

// ── Missing Dependencies ──────────────────────────────────────────────────────

export interface MissingDependency {
  type: string;
  identifier: string;
  displayName?: string;
  referencedBy: SourceRef;
  impact: string;
  checklistItem: string;
}

// ── Manual Checklist ──────────────────────────────────────────────────────────

export interface ChecklistItem {
  section: string;
  item: string;
  priority: 'high' | 'medium' | 'low';
}

// ── Parse Errors ──────────────────────────────────────────────────────────────

export interface ParseError {
  archivePath: string;
  error: string;
  partial: boolean;
}

// ── Executive Summary ─────────────────────────────────────────────────────────

export interface RiskFlag {
  severity: 'high' | 'medium' | 'low';
  message: string;
  location?: string;
}

export interface ExecutiveSummary {
  artifactName: string;
  artifactVersion?: string;
  environment: Environment;
  componentCounts: Record<string, number>;
  keyDependencies: string[];
  riskFlags: RiskFlag[];
  manualItemCount: number;
}

// ── Where Used Index ──────────────────────────────────────────────────────────

export interface WhereUsedIndex {
  connectorToFlows: Map<string, string[]>;
  connectorToApps: Map<string, string[]>;
  envVarToFlows: Map<string, string[]>;
  envVarToApps: Map<string, string[]>;
  dataSourceToFlows: Map<string, string[]>;
  dataSourceToApps: Map<string, string[]>;
  variableToScreens: Map<string, string[]>;
}

// ── Root Graph ────────────────────────────────────────────────────────────────

export interface NormalizedAnalysisGraph {
  meta: ArtifactMeta;
  solution?: SolutionInfo;
  canvasApps: CanvasApp[];
  modelDrivenApps: ModelDrivenApp[];
  flows: Flow[];
  copilotAgents: CopilotAgent[];
  aiBuilderModels: AiBuilderModel[];
  managedPlans: ManagedPlan[];
  envVars: EnvVar[];
  connectors: ConnectorRef[];
  connectionRefs: ConnectionReference[];
  dataSources: DataSource[];
  securityRoles: SecurityRole[];
  dataverseFormulas: DataverseFormula[];
  webResources: WebResource[];
  urls: UrlReference[];
  missingDeps: MissingDependency[];
  parseErrors: ParseError[];
  manualChecklist: ChecklistItem[];
  executiveSummary?: ExecutiveSummary;
  whereUsed: WhereUsedIndex;
}

// ── Progress ──────────────────────────────────────────────────────────────────

export interface ProgressEvent {
  stage: string;
  pct: number;
  message: string;
}

export type ProgressCallback = (evt: ProgressEvent) => void;

// ── IPC Channel types ────────────────────────────────────────────────────────

export interface SessionConfig {
  environment: Environment;
  artifactPath: string;
  outputDir: string;
}

export interface SchemaPromptRequest {
  dataSourceId: string;
  displayName: string;
  type: DataSourceType;
}

export interface SchemaPromptResponse {
  dataSourceId: string;
  action: 'upload' | 'skip';
  filePath?: string;
}

export interface GenerationResult {
  htmlPath: string;
  docxPath: string;
  markdownPath: string;
  pdfPath?: string;
  errors: ParseError[];
}
