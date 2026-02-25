import React from 'react';
import { useStore } from './store';

const ENVS = [
  { id: 'DEV',  label: 'Development',  desc: 'For development and testing environments', color: '#0c5460', bg: '#d1ecf1' },
  { id: 'TEST', label: 'Test / UAT',   desc: 'For test, UAT, or staging environments',  color: '#856404', bg: '#fff3cd' },
  { id: 'PROD', label: 'Production',   desc: 'For live / production environments',       color: '#721c24', bg: '#f8d7da' },
] as const;

export function StepEnv() {
  const { setEnvironment } = useStore();

  return (
    <div className="step-panel">
      <h2>Step 1: Select Target Environment</h2>
      <p className="step-desc">This helps the tool flag environment-specific URLs and configuration.</p>
      <div className="env-grid">
        {ENVS.map(env => (
          <button
            key={env.id}
            className="env-card"
            style={{ '--env-color': env.color, '--env-bg': env.bg } as React.CSSProperties}
            onClick={() => setEnvironment(env.id)}
          >
            <div className="env-badge-large">{env.id}</div>
            <div className="env-card-label">{env.label}</div>
            <div className="env-card-desc">{env.desc}</div>
          </button>
        ))}
      </div>
    </div>
  );
}
