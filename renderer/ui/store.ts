import { create } from 'zustand';
import type { Environment } from '../../src/model/types';

export type WizardStep = 'env' | 'upload' | 'analyzing' | 'schema' | 'generating' | 'done' | 'error';

export interface DataSourceForSchema {
  id: string;
  displayName: string;
  type: string;
}

export interface ProgressItem {
  stage: string;
  pct: number;
  message: string;
}

export interface AnalysisSummary {
  artifactType: string;
  canvasApps: number;
  flows: number;
  envVars: number;
  dataSources: number;
  modelDrivenApps: number;
  securityRoles: number;
  riskFlags: Array<{ severity: string; message: string }>;
  manualItemCount: number;
  dataSourcesForSchema: DataSourceForSchema[];
}

export interface GenerationOutputs {
  html?: string;
  docx?: string;
  markdown?: string;
}

interface AppState {
  step: WizardStep;
  environment: Environment | null;
  artifactPath: string | null;
  artifactName: string | null;
  outputDir: string | null;
  progress: ProgressItem;
  summary: AnalysisSummary | null;
  schemaQueue: DataSourceForSchema[];
  currentSchemaIndex: number;
  outputs: GenerationOutputs;
  errorMessage: string | null;

  setEnvironment: (env: Environment) => void;
  setArtifactPath: (p: string, name: string) => void;
  setOutputDir: (d: string) => void;
  setProgress: (p: ProgressItem) => void;
  setSummary: (s: AnalysisSummary) => void;
  setSchemaQueue: (q: DataSourceForSchema[]) => void;
  advanceSchemaQueue: () => void;
  setOutputs: (o: GenerationOutputs) => void;
  setStep: (s: WizardStep) => void;
  setError: (msg: string) => void;
  reset: () => void;
}

const initialState = {
  step: 'env' as WizardStep,
  environment: null,
  artifactPath: null,
  artifactName: null,
  outputDir: null,
  progress: { stage: '', pct: 0, message: '' },
  summary: null,
  schemaQueue: [],
  currentSchemaIndex: 0,
  outputs: {},
  errorMessage: null
};

export const useStore = create<AppState>((set) => ({
  ...initialState,

  setEnvironment: (environment) => set({ environment, step: 'upload' }),
  setArtifactPath: (artifactPath, artifactName) => set({ artifactPath, artifactName }),
  setOutputDir: (outputDir) => set({ outputDir }),
  setProgress: (progress) => set({ progress }),
  setSummary: (summary) => set({ summary }),
  setSchemaQueue: (schemaQueue) => set({ schemaQueue, currentSchemaIndex: 0 }),
  advanceSchemaQueue: () => set((state) => ({ currentSchemaIndex: state.currentSchemaIndex + 1 })),
  setOutputs: (outputs) => set({ outputs }),
  setStep: (step) => set({ step }),
  setError: (errorMessage) => set({ errorMessage, step: 'error' }),
  reset: () => set(initialState)
}));
