import type { CanvasApp, NavEdge, Confidence } from '../model/types';

const NAV_RE = /\bNavigate\s*\(\s*([A-Za-z_][A-Za-z0-9_]*)/g;
const BACK_RE = /\bBack\s*\(/g;

export interface NavInferenceResult {
  edges: NavEdge[];
  confidence: Confidence;
  mermaid: string;
}

export class NavigationInferencer {
  static infer(app: CanvasApp): NavInferenceResult {
    const screenNames = new Set(app.screens.map(s => s.name));
    const edges: NavEdge[] = [];

    const scanFormula = (formula: string, sourceScreen: string) => {
      for (const m of formula.matchAll(NAV_RE)) {
        const target = m[1];
        const conf: Confidence = screenNames.has(target) ? 'high' : 'medium';
        edges.push({ from: sourceScreen, to: target, confidence: conf });
      }
      if (BACK_RE.test(formula)) {
        edges.push({ from: sourceScreen, to: '__back__', confidence: 'high', triggeredBy: 'Back()' });
        BACK_RE.lastIndex = 0;
      }
    };

    for (const screen of app.screens) {
      if (screen.onVisible.raw) scanFormula(screen.onVisible.raw, screen.name);
      for (const kf of screen.keyFormulas) scanFormula(kf.formula.raw, screen.name);
      for (const ctrl of screen.controls) {
        for (const kf of ctrl.keyFormulas) scanFormula(kf.formula.raw, screen.name);
      }
    }

    // Dedup
    const seen = new Set<string>();
    const unique: NavEdge[] = [];
    for (const e of edges) {
      const key = `${e.from}->${e.to}`;
      if (!seen.has(key)) { seen.add(key); unique.push(e); }
    }

    const allHigh = unique.every(e => e.confidence === 'high');
    const allLow = unique.length === 0;
    const confidence: Confidence = allLow ? 'low' : allHigh ? 'high' : 'medium';

    const mermaid = NavigationInferencer.buildMermaid(app.screens.map(s => s.name), unique);
    return { edges: unique, confidence, mermaid };
  }

  private static buildMermaid(screenNames: string[], edges: NavEdge[]): string {
    if (screenNames.length === 0) return '';
    const lines = ['flowchart LR'];
    const escape = (s: string) => s.replace(/[^a-zA-Z0-9_]/g, '_');

    for (const name of screenNames) {
      lines.push(`  ${escape(name)}["${name}"]`);
    }
    for (const e of edges) {
      if (e.to === '__back__') continue;
      const conf = e.confidence === 'high' ? '' : e.confidence === 'medium' ? ' -.- ' : '-.->';
      if (e.confidence === 'high') {
        lines.push(`  ${escape(e.from)} --> ${escape(e.to)}`);
      } else {
        lines.push(`  ${escape(e.from)} -.-> ${escape(e.to)}`);
      }
    }
    return lines.join('\n');
  }
}
