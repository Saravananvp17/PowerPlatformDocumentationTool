import React from 'react';
import { useStore } from './store';

export function StepGenerating() {
  const { progress } = useStore();
  return (
    <div className="step-panel centered">
      <div className="spinner" />
      <h2>Generating Documentation</h2>
      <div className="progress-bar-wrap">
        <div className="progress-bar" style={{ width: `${progress.pct}%` }} />
      </div>
      <div className="progress-stage">{progress.message}</div>
      <p className="step-desc">Building HTML report, DOCX specification, and Markdown export…</p>
    </div>
  );
}
