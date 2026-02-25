/**
 * ButtonActionAnalyzer
 *
 * Parses Power Fx formulas from control properties (OnSelect, OnChange, OnSuccess …)
 * and extracts structured "action steps" — Patch calls, flow Run calls, Navigate
 * calls, form submissions, and collection operations.
 *
 * Power Fx is not standard JS, so this uses heuristic regex + a bracket-depth tracker
 * rather than a full AST parser.  It handles the most common patterns found in
 * production canvas apps and degrades gracefully on complex or nested formulas.
 */

import type {
  CanvasApp, ButtonAction, ActionStep, ActionKind
} from '../model/types';

// Properties that carry user-initiated actions worth surfacing
const ACTION_PROPERTIES = new Set([
  'OnSelect', 'OnChange', 'OnSuccess', 'OnFailure', 'OnCheck', 'OnUncheck',
  'OnReset', 'OnTimeout', 'OnHidden'
]);

// Power Fx built-in functions — used to exclude false-positive `.Run()` matches
// (e.g. `ForAll(...).Run(...)` doesn't exist, but Filter, Sort, etc. should never
// appear before a real `.Run()` anyway.  This list is conservative.)
const BUILTIN_PREFIXES = new Set([
  'Filter', 'Sort', 'SortByColumns', 'Search', 'LookUp', 'First', 'Last',
  'Index', 'ForAll', 'With', 'If', 'Switch', 'IfError', 'Coalesce',
  'Table', 'Record', 'Sequence', 'GroupBy', 'Ungroup', 'Distinct',
  'AddColumns', 'DropColumns', 'ShowColumns', 'RenameColumns',
  'Sum', 'Average', 'Min', 'Max', 'Count', 'CountRows', 'CountIf',
  'Text', 'Value', 'DateAdd', 'DateDiff', 'Now', 'Today', 'TimeZoneOffset',
  'Left', 'Right', 'Mid', 'Len', 'Trim', 'Upper', 'Lower', 'Substitute',
  'Split', 'Concat', 'Concatenate', 'Find', 'Replace', 'StartsWith', 'EndsWith',
  'IsBlank', 'IsEmpty', 'IsError', 'IsNumeric',
  'JSON', 'ParseJSON', 'Error', 'Assert',
]);

// ─────────────────────────────────────────────────────────────────────────────

export class ButtonActionAnalyzer {
  /** Mutates app.screens — populates buttonActions on every screen. */
  static analyze(app: CanvasApp): void {
    for (const screen of app.screens) {
      screen.buttonActions = [];

      for (const control of screen.controls) {
        for (const kf of control.keyFormulas) {
          if (!ACTION_PROPERTIES.has(kf.property)) continue;
          const formula = kf.formula.raw.trim();
          if (!formula) continue;

          const steps = extractSteps(formula);
          if (steps.length === 0) continue;

          screen.buttonActions.push({
            controlName: control.name,
            controlType: control.type,
            property: kf.property,
            screenName: screen.name,
            actions: steps,
            formulaSnippet: formula.length > 5000 ? formula.substring(0, 4997) + '…' : formula,
            source: kf.formula.source
          });
        }
      }

      // Also check screen-level key formulas (OnVisible, OnHidden etc.)
      // These are less likely to have Patch/Run, but include them for completeness.
      for (const kf of screen.keyFormulas) {
        if (!ACTION_PROPERTIES.has(kf.property)) continue;
        const formula = kf.formula.raw.trim();
        if (!formula) continue;

        const steps = extractSteps(formula);
        if (steps.length === 0) continue;

        screen.buttonActions.push({
          controlName: screen.name,
          controlType: 'Screen',
          property: kf.property,
          screenName: screen.name,
          actions: steps,
          formulaSnippet: formula.length > 5000 ? formula.substring(0, 4997) + '…' : formula,
          source: kf.formula.source
        });
      }
    }
  }
}

// ─── Extraction helpers ───────────────────────────────────────────────────────

/**
 * Extract all action steps from a Power Fx formula string.
 * Returns deduplicated steps in order of appearance.
 */
function extractSteps(formula: string): ActionStep[] {
  const steps: ActionStep[] = [];
  const seen = new Set<string>();

  const add = (kind: ActionKind, target: string, payload?: string) => {
    const key = `${kind}:${target}`;
    if (!seen.has(key)) {
      seen.add(key);
      steps.push({ kind, target, payload: payload?.trim() || undefined });
    }
  };

  // ── Patch(DataSource, record …) ──────────────────────────────────────────
  const patchRe = /\bPatch\s*\(/gi;
  let m: RegExpExecArray | null;
  while ((m = patchRe.exec(formula)) !== null) {
    const args = parenContent(formula, m.index + m[0].length - 1);
    const ds   = firstArg(args);
    if (ds) {
      const payload = restArgs(args);
      add('patch', ds, payload.length > 150 ? payload.substring(0, 147) + '…' : payload || undefined);
    }
  }

  // ── 'Flow Name'.Run(…) or ConnectorName.Run(…) ──────────────────────────
  // Match: (optional-quote)(name)(optional-quote) whitespace? . whitespace? Run whitespace? (
  const flowRe = /(?:'([^']+)'|(\b[A-Za-z_]\w*))\s*\.\s*Run\s*\(/gi;
  // Power Apps appends the WorkflowEntityId GUID to the connector name in formulas,
  // e.g. 'MyFlow-85705400-5BC6-F011-8544-7C1E525DFBD3'.Run(...)  Strip it off.
  const stripFlowGuid = (n: string) =>
    n.replace(/-[0-9A-Fa-f]{8}(?:-[0-9A-Fa-f]{4}){3}-[0-9A-Fa-f]{12}$/i, '').trim();
  while ((m = flowRe.exec(formula)) !== null) {
    const rawName = stripFlowGuid((m[1] ?? m[2] ?? '').trim());
    if (!rawName || BUILTIN_PREFIXES.has(rawName)) continue;
    const args    = parenContent(formula, m.index + m[0].length - 1);
    const preview = args.length > 200 ? args.substring(0, 197) + '…' : args;
    add('flow-run', rawName, preview || undefined);
  }

  // ── Navigate(Screen, transition?) ───────────────────────────────────────
  const navRe = /\bNavigate\s*\(/gi;
  while ((m = navRe.exec(formula)) !== null) {
    const args   = parenContent(formula, m.index + m[0].length - 1);
    const screen = firstArg(args);
    if (screen) add('navigate', screen);
  }

  // ── SubmitForm(FormControl) ──────────────────────────────────────────────
  const sfRe = /\bSubmitForm\s*\(/gi;
  while ((m = sfRe.exec(formula)) !== null) {
    const args = parenContent(formula, m.index + m[0].length - 1);
    const form = firstArg(args);
    if (form) add('submit-form', form);
  }

  // ── Collect / ClearCollect(Collection, …) ───────────────────────────────
  const colRe = /\b(?:ClearCollect|Collect)\s*\(/gi;
  while ((m = colRe.exec(formula)) !== null) {
    const args = parenContent(formula, m.index + m[0].length - 1);
    const col  = firstArg(args);
    if (col) add('collect', col);
  }

  // ── Remove / RemoveIf(DataSource, …) ────────────────────────────────────
  const removeRe = /\bRemove(?:If)?\s*\(/gi;
  while ((m = removeRe.exec(formula)) !== null) {
    const args = parenContent(formula, m.index + m[0].length - 1);
    const ds   = firstArg(args);
    if (ds) add('remove', ds);
  }

  return steps;
}

/**
 * Extract the text inside the outermost parentheses starting at openParenIdx.
 * Handles nested parens, curly braces, and square brackets.
 */
function parenContent(formula: string, openParenIdx: number): string {
  // Scan forward to find the actual '('
  let i = openParenIdx;
  while (i < formula.length && formula[i] !== '(') i++;
  if (i >= formula.length) return '';

  let depth = 0;
  const start = i + 1;
  for (; i < formula.length; i++) {
    const ch = formula[i];
    if (ch === '(' || ch === '{' || ch === '[') depth++;
    else if (ch === ')' || ch === '}' || ch === ']') {
      depth--;
      if (depth === 0) return formula.substring(start, i).trim();
    }
  }
  // Unclosed — return what we have
  return formula.substring(start).trim();
}

/**
 * Extract the first comma-separated argument from a parenthesised argument list.
 * Handles:
 *   - Power Fx single-quoted identifiers: 'My List'
 *   - Nested function calls and records: { field: value }
 */
function firstArg(argContent: string): string {
  const s = argContent.trim();

  // Single-quoted Power Fx identifier
  if (s.startsWith("'")) {
    const end = s.indexOf("'", 1);
    return end > 0 ? s.substring(1, end) : s.replace(/^'/, '');
  }

  // Walk to first comma at depth 0
  let depth = 0;
  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    if (ch === '(' || ch === '{' || ch === '[') depth++;
    else if (ch === ')' || ch === '}' || ch === ']') depth--;
    else if (ch === ',' && depth === 0) return s.substring(0, i).trim();
  }
  return s;
}

/**
 * Everything after the first argument (the "payload" / record / remaining args).
 * Used to show what data is being written in a Patch call.
 */
function restArgs(argContent: string): string {
  const s = argContent.trim();
  if (s.startsWith("'")) {
    const end = s.indexOf("'", 1);
    const afterQuote = end > 0 ? s.substring(end + 1).trimStart() : '';
    return afterQuote.startsWith(',') ? afterQuote.substring(1).trim() : '';
  }
  let depth = 0;
  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    if (ch === '(' || ch === '{' || ch === '[') depth++;
    else if (ch === ')' || ch === '}' || ch === ']') depth--;
    else if (ch === ',' && depth === 0) return s.substring(i + 1).trim();
  }
  return '';
}
