import React, { useState, useEffect } from 'react';
import { useStore } from './store';

// Minimum time (ms) to show the analyzing screen — gives the animation time to play
const MIN_ANALYZE_MS = 7200;

export function StepUpload() {
  const {
    environment, artifactPath, artifactName,
    outputDir, setOutputDir,
    setArtifactPath, setStep, setProgress, setSummary, setError
  } = useStore();
  const [isAnalyzing, setIsAnalyzing] = useState(false);

  // Load the default output directory once on mount
  useEffect(() => {
    if (!outputDir) {
      window.ppdt.getDefaultOutputDir().then(dir => {
        if (dir) setOutputDir(dir);
      });
    }
  }, []);

  async function handlePick() {
    const filePath = await window.ppdt.openArtifact();
    if (!filePath) return;
    const name = filePath.split(/[\\/]/).pop() ?? filePath;
    setArtifactPath(filePath, name);
  }

  async function handleChangeOutputDir() {
    const dir = await window.ppdt.openOutputDir();
    if (dir) setOutputDir(dir);
  }

  async function handleAnalyze() {
    if (!artifactPath || !environment) return;
    setIsAnalyzing(true);
    setStep('analyzing');
    setProgress({ stage: 'detect', pct: 0, message: 'Starting…' });

    const startTime = Date.now();

    const result = await window.ppdt.analyze(artifactPath, environment) as {
      success: boolean; error?: string; summary?: {
        dataSourcesForSchema: Array<{ id: string; displayName: string; type: string }>;
        [key: string]: unknown;
      };
    };

    // Ensure the analyzing screen shows for at least MIN_ANALYZE_MS
    const elapsed = Date.now() - startTime;
    if (elapsed < MIN_ANALYZE_MS) {
      await new Promise(resolve => setTimeout(resolve, MIN_ANALYZE_MS - elapsed));
    }

    if (!result.success) {
      setError(result.error ?? 'Analysis failed');
      return;
    }

    const summary = result.summary as unknown as import('./store').AnalysisSummary;
    setSummary(summary);

    // Schema is now auto-detected from the artifact — go straight to generating
    await handleGenerate();
  }

  async function handleGenerate() {
    setStep('generating');
    const dir = outputDir ?? await window.ppdt.getDefaultOutputDir();
    const genResult = await window.ppdt.generateAll(dir) as {
      success: boolean; error?: string;
      outputs?: { html?: string; docx?: string; markdown?: string };
    };
    if (!genResult.success) {
      setError(genResult.error ?? 'Generation failed');
      return;
    }
    const { setOutputs } = useStore.getState();
    setOutputs(genResult.outputs ?? {});
    setStep('done');
  }

  // Shorten a long path for display: show last 2 segments with ellipsis
  function shortPath(p: string): string {
    const parts = p.replace(/\\/g, '/').split('/').filter(Boolean);
    if (parts.length <= 3) return p;
    return '…/' + parts.slice(-2).join('/');
  }

  return (
    <div className="step-panel">
      <h2>Step 2: Upload Artifact</h2>
      <div className="env-pill">Environment: <strong>{environment}</strong></div>
      <p className="step-desc">Upload one Power Platform artifact: a Solution export (.zip), Flow export (.zip), Power Apps export (.zip), or a standalone Canvas App (.msapp).</p>

      <div className={`drop-zone ${artifactPath ? 'has-file' : ''}`} onClick={handlePick}>
        {artifactPath ? (
          <>
            <div className="drop-icon">📦</div>
            <div className="drop-filename">{artifactName}</div>
            <div className="drop-hint">Click to change file</div>
          </>
        ) : (
          <>
            <div className="drop-icon">📂</div>
            <div className="drop-hint">Click to select .zip or .msapp file</div>
          </>
        )}
      </div>

      <div className="format-list">
        <strong>Supported formats:</strong>
        <span className="tag">Solution .zip</span>
        <span className="tag">Flow export .zip</span>
        <span className="tag">Power Apps .zip</span>
        <span className="tag">Standalone .msapp</span>
      </div>

      <div className="upload-notice">
        ⚠️ <strong>Power Pages</strong> components have not been validated with this tool. Solutions containing Power Pages artifacts will still process, but those components may be omitted from the report.
      </div>

      {/* Output directory picker */}
      <div className="output-dir-row">
        <span className="output-dir-label">📁 Output folder</span>
        <span className="output-dir-path" title={outputDir ?? ''}>
          {outputDir ? shortPath(outputDir) : 'Loading…'}
        </span>
        <button className="btn-small btn-ghost" onClick={handleChangeOutputDir} disabled={isAnalyzing}>
          Change
        </button>
      </div>

      <button
        className="btn-primary"
        onClick={handleAnalyze}
        disabled={!artifactPath || isAnalyzing}
      >
        {isAnalyzing ? 'Analysing…' : '🔍 Analyse Artifact'}
      </button>
    </div>
  );
}
