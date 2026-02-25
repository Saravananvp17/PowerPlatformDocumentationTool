import React from 'react';
import { useStore } from './store';

interface OutputRowProps {
  icon: string;
  label: string;
  filePath?: string;
  ext: string;
}

function OutputRow({ icon, label, filePath, ext }: OutputRowProps) {
  if (!filePath) return null;
  return (
    <div className="output-row">
      <span className="output-icon">{icon}</span>
      <div className="output-info">
        <div className="output-label">{label}</div>
        <div className="output-path">{filePath}</div>
      </div>
      <div className="output-actions">
        <button className="btn-small" onClick={() => window.ppdt.openFile(filePath)}>
          Open
        </button>
        <button className="btn-small btn-ghost" onClick={() => window.ppdt.showInFolder(filePath)}>
          Show in Folder
        </button>
      </div>
    </div>
  );
}

export function StepDone() {
  const { summary, outputs, environment, artifactName, reset } = useStore();

  const riskCount = (summary?.riskFlags ?? []).filter(r => r.severity === 'high').length;

  return (
    <div className="step-panel">
      <div className="done-header">
        <div className="done-icon">✅</div>
        <h2>Documentation Generated</h2>
      </div>

      <div className="summary-chips">
        {summary && <>
          <div className="chip chip-blue">📱 {summary.canvasApps} Canvas App{summary.canvasApps !== 1 ? 's' : ''}</div>
          <div className="chip chip-blue">⚡ {summary.flows} Flow{summary.flows !== 1 ? 's' : ''}</div>
          <div className="chip chip-blue">🔧 {summary.envVars} Env Var{summary.envVars !== 1 ? 's' : ''}</div>
          <div className="chip chip-blue">🗃️ {summary.dataSources} Data Source{summary.dataSources !== 1 ? 's' : ''}</div>
          {riskCount > 0 && <div className="chip chip-red">⚠️ {riskCount} High Risk Flag{riskCount !== 1 ? 's' : ''}</div>}
          {summary.manualItemCount > 0 && <div className="chip chip-yellow">📝 {summary.manualItemCount} Manual Item{summary.manualItemCount !== 1 ? 's' : ''}</div>}
        </>}
      </div>

      <h3>Output Files</h3>
      <div className="output-list">
        <OutputRow icon="🌐" label="HTML Report (interactive)" filePath={outputs.html} ext="html" />
        <OutputRow icon="📄" label="Word Document (CoE Spec)" filePath={outputs.docx} ext="docx" />
        <OutputRow icon="📝" label="Markdown Report" filePath={outputs.markdown} ext="md" />
      </div>

      {summary && summary.riskFlags.length > 0 && (
        <div className="risk-summary">
          <h3>⚠️ Risk Flags</h3>
          {summary.riskFlags.map((r, i) => (
            <div key={i} className={`risk-item risk-${r.severity}`}>
              <span className="risk-level">{r.severity.toUpperCase()}</span> {r.message}
            </div>
          ))}
        </div>
      )}

      <button className="btn-primary" style={{ marginTop: '24px' }} onClick={reset}>
        ↩ Analyse Another Artifact
      </button>
    </div>
  );
}
