import React, { useEffect, useState } from 'react';
import { useStore } from './store';

const STAGED_MESSAGES = [
  { pct: 8,  text: 'Unpacking files…' },
  { pct: 22, text: 'Reading manifest…' },
  { pct: 38, text: 'Parsing canvas apps…' },
  { pct: 54, text: 'Analysing flows…' },
  { pct: 68, text: 'Mapping data connections…' },
  { pct: 80, text: 'Extracting schema…' },
  { pct: 90, text: 'Compiling results…' },
  { pct: 97, text: 'Finalising…' },
];

// Time (ms) each staged message stays visible before advancing
const STAGE_INTERVAL = 900;

export function StepAnalyzing() {
  const { progress, artifactName } = useStore();

  // Scripted animation state — runs independently of the real backend progress
  const [stageIdx, setStageIdx] = useState(0);
  const [displayPct, setDisplayPct] = useState(0);
  const [displayMsg, setDisplayMsg] = useState(STAGED_MESSAGES[0].text);

  useEffect(() => {
    setStageIdx(0);
    setDisplayPct(STAGED_MESSAGES[0].pct);
    setDisplayMsg(STAGED_MESSAGES[0].text);

    const interval = setInterval(() => {
      setStageIdx(prev => {
        const next = Math.min(prev + 1, STAGED_MESSAGES.length - 1);
        setDisplayPct(STAGED_MESSAGES[next].pct);
        setDisplayMsg(STAGED_MESSAGES[next].text);
        return next;
      });
    }, STAGE_INTERVAL);

    return () => clearInterval(interval);
  }, []);

  // Once the backend reports real progress past ~50%, let the real message show
  const showReal = progress.pct > 50 && progress.message;
  const barPct = showReal ? Math.max(progress.pct, displayPct) : displayPct;
  const msg = showReal ? progress.message : displayMsg;

  return (
    <div className="step-panel centered">
      <div className="spinner" />
      <h2>Analysing Artifact</h2>
      <div className="artifact-pill">📦 {artifactName}</div>

      <div className="progress-bar-wrap">
        <div
          className="progress-bar"
          style={{ width: `${barPct}%`, transition: 'width 0.6s ease' }}
        />
      </div>

      <div className="progress-stage">{msg}</div>
      <p className="step-desc muted">All processing is local — no data is sent off your device.</p>
    </div>
  );
}
