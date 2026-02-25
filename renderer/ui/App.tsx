import React, { useEffect } from 'react';
import { useStore } from './store';
import { StepEnv } from './StepEnv';
import { StepUpload } from './StepUpload';
import { StepAnalyzing } from './StepAnalyzing';
import { StepGenerating } from './StepGenerating';
import { StepDone } from './StepDone';
import { StepError } from './StepError';

declare global {
  interface Window {
    ppdt: {
      openArtifact: () => Promise<string | null>;
      openSchema: (name: string) => Promise<string | null>;
      openOutputDir: () => Promise<string | null>;
      detectArtifact: (p: string) => Promise<unknown>;
      analyze: (p: string, env: string) => Promise<unknown>;
      enrichSchema: (id: string, p: string) => Promise<unknown>;
      generateAll: (dir: string) => Promise<unknown>;
      openFile: (p: string) => Promise<void>;
      showInFolder: (p: string) => Promise<void>;
      getDefaultOutputDir: () => Promise<string>;
      onProgress: (cb: (evt: { stage: string; pct: number; message: string }) => void) => () => void;
    };
  }
}

const STEPS = ['Environment', 'Upload', 'Analyse', 'Generate', 'Done'];

function stepIndex(step: string): number {
  const map: Record<string, number> = {
    env: 0, upload: 1, analyzing: 2, generating: 3, done: 4, error: -1
  };
  return map[step] ?? 0;
}

export default function App() {
  const { step, reset } = useStore();

  // Register progress listener
  useEffect(() => {
    const { setProgress } = useStore.getState();
    const off = window.ppdt.onProgress((evt) => setProgress(evt));
    return off;
  }, []);

  const idx = stepIndex(step);

  return (
    <div className="app-shell">
      {/* Header */}
      <header className="app-header">
        <div className="app-logo">📋 PP Doc Tool</div>
        <div className="app-subtitle">Power Platform Documentation Generator</div>
        {step !== 'env' && (
          <button className="btn-ghost" onClick={reset} title="Start over">↩ Start Over</button>
        )}
      </header>

      {/* Progress bar */}
      {step !== 'error' && (
        <div className="step-bar">
          {STEPS.map((label, i) => (
            <div key={label} className={`step-item ${i < idx ? 'done' : i === idx ? 'active' : ''}`}>
              <div className="step-dot">{i < idx ? '✓' : i + 1}</div>
              <div className="step-label">{label}</div>
            </div>
          ))}
        </div>
      )}

      {/* Step content */}
      <main className="step-content">
        {step === 'env'        && <StepEnv />}
        {step === 'upload'     && <StepUpload />}
        {step === 'analyzing'  && <StepAnalyzing />}
        {step === 'generating' && <StepGenerating />}
        {step === 'done'       && <StepDone />}
        {step === 'error'      && <StepError />}
      </main>
    </div>
  );
}
