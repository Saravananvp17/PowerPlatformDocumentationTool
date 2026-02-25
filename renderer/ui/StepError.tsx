import React from 'react';
import { useStore } from './store';

export function StepError() {
  const { errorMessage, reset } = useStore();
  return (
    <div className="step-panel centered">
      <div className="error-icon">❌</div>
      <h2>Analysis Error</h2>
      <div className="error-box">{errorMessage}</div>
      <p className="step-desc">The artifact could not be fully parsed. Try with a different file or check that it is a valid Power Platform export.</p>
      <button className="btn-primary" onClick={reset}>↩ Try Again</button>
    </div>
  );
}
