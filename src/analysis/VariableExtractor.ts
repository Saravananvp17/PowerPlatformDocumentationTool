import type { CanvasApp, Variable, SourceRef } from '../model/types';

// Regex patterns for PowerFx function calls
const SET_RE = /\bSet\s*\(\s*([A-Za-z_][A-Za-z0-9_]*)\s*,/g;
const UPDATE_CTX_RE = /\bUpdateContext\s*\(\s*\{([^}]+)\}/g;
const COLLECT_RE = /\b(?:ClearCollect|Collect)\s*\(\s*([A-Za-z_][A-Za-z0-9_]*)\s*,/g;

export class VariableExtractor {
  static extract(app: CanvasApp): Variable[] {
    const vars = new Map<string, Variable>();

    const addSet = (name: string, kind: Variable['kind'], src: SourceRef) => {
      if (!vars.has(name)) {
        vars.set(name, { name, kind, setAt: [], usedAt: [] });
      }
      const v = vars.get(name)!;
      if (!v.setAt.find(s => s.archivePath === src.archivePath && s.yamlPath === src.yamlPath)) {
        v.setAt.push(src);
      }
    };

    const scanFormula = (formula: string, src: SourceRef) => {
      // Set() calls → global variable
      for (const m of formula.matchAll(SET_RE)) {
        addSet(m[1], 'global', src);
      }
      // UpdateContext() calls → context variable
      for (const m of formula.matchAll(UPDATE_CTX_RE)) {
        // Extract key names from { key: val, key2: val2 }
        const pairs = m[1].split(',');
        for (const pair of pairs) {
          const keyMatch = pair.match(/^\s*([A-Za-z_][A-Za-z0-9_]*)\s*:/);
          if (keyMatch) addSet(keyMatch[1], 'context', src);
        }
      }
      // ClearCollect / Collect → collection
      for (const m of formula.matchAll(COLLECT_RE)) {
        addSet(m[1], 'collection', src);
      }
    };

    // Scan App.OnStart
    if (app.appOnStart.raw) {
      scanFormula(app.appOnStart.raw, app.appOnStart.source);
    }

    // Scan all screen formulas
    for (const screen of app.screens) {
      if (screen.onVisible.raw) scanFormula(screen.onVisible.raw, screen.onVisible.source);
      for (const kf of screen.keyFormulas) {
        if (kf.formula.raw) scanFormula(kf.formula.raw, kf.formula.source);
      }
      for (const ctrl of screen.controls) {
        for (const kf of ctrl.keyFormulas) {
          if (kf.formula.raw) scanFormula(kf.formula.raw, kf.formula.source);
        }
      }
    }

    return [...vars.values()];
  }
}
