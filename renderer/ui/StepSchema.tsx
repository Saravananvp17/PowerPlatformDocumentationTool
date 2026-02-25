import React, { useState } from 'react';
import { useStore } from './store';

export function StepSchema() {
  const { schemaQueue, currentSchemaIndex, advanceSchemaQueue, setStep, setOutputs, setError } = useStore();
  const [isProcessing, setIsProcessing] = useState(false);

  const ds = schemaQueue[currentSchemaIndex];
  const isLast = currentSchemaIndex >= schemaQueue.length - 1;

  async function advance() {
    if (isLast) {
      await generateDocs();
    } else {
      advanceSchemaQueue();
    }
  }

  async function handleUpload() {
    if (!ds) return;
    const filePath = await window.ppdt.openSchema(ds.displayName);
    if (!filePath) return;
    setIsProcessing(true);
    const result = await window.ppdt.enrichSchema(ds.id, filePath) as { success: boolean; error?: string; columnCount?: number };
    setIsProcessing(false);
    if (!result.success) {
      alert(`Failed to parse schema: ${result.error}`);
      return;
    }
    alert(`✅ Schema loaded: ${result.columnCount} column(s) detected`);
    await advance();
  }

  async function handleSkip() {
    await advance();
  }

  async function generateDocs() {
    setStep('generating');
    const outputDir = await window.ppdt.getDefaultOutputDir();
    const genResult = await window.ppdt.generateAll(outputDir) as {
      success: boolean; error?: string;
      outputs?: { html?: string; docx?: string; markdown?: string };
    };
    if (!genResult.success) {
      setError(genResult.error ?? 'Generation failed');
      return;
    }
    setOutputs(genResult.outputs ?? {});
    setStep('done');
  }

  if (!ds) return null;

  return (
    <div className="step-panel">
      <h2>Step 4: Schema Enrichment</h2>
      <div className="schema-progress">
        Data source {currentSchemaIndex + 1} of {schemaQueue.length}
      </div>
      <div className="schema-card">
        <div className="schema-icon">{ds.type === 'sharepoint' ? '📋' : '🗃️'}</div>
        <div className="schema-info">
          <div className="schema-name">{ds.displayName}</div>
          <div className="schema-type">{ds.type}</div>
        </div>
      </div>
      <p className="step-desc">
        Upload a schema or export file so this tool can document the column structure of <strong>{ds.displayName}</strong>.
        Accepted formats: CSV, XLSX, JSON.
      </p>
      <div className="schema-actions">
        <button className="btn-primary" onClick={handleUpload} disabled={isProcessing}>
          📄 Upload Schema File
        </button>
        <button className="btn-ghost" onClick={handleSkip} disabled={isProcessing}>
          Skip (document manually later)
        </button>
      </div>
      <p className="muted">Skipped sources will be listed in the Manual Completion Checklist.</p>
    </div>
  );
}
