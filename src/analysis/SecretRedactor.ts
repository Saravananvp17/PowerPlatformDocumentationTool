import type { NormalizedAnalysisGraph, FlowAction } from '../model/types';

// Patterns and their replacement labels
const REDACTION_RULES: Array<{ pattern: RegExp; label: string }> = [
  // JWT / Bearer tokens
  { pattern: /Bearer\s+[A-Za-z0-9\-_]+\.[A-Za-z0-9\-_]+\.[A-Za-z0-9\-_]+/g,
    label: '[REDACTED:BEARER_TOKEN]' },
  // Azure SAS signature
  { pattern: /sig=[A-Za-z0-9%+/=]{20,}/g,
    label: 'sig=[REDACTED:SAS_SIGNATURE]' },
  // Connection strings with AccountKey
  { pattern: /AccountKey=[A-Za-z0-9+/=]{10,}/g,
    label: 'AccountKey=[REDACTED:ACCOUNT_KEY]' },
  // Shared key patterns
  { pattern: /SharedAccessKey=[A-Za-z0-9+/=]{10,}/g,
    label: 'SharedAccessKey=[REDACTED:SHARED_KEY]' },
  // PEM private keys
  { pattern: /-----BEGIN (?:RSA |EC |ENCRYPTED )?PRIVATE KEY-----[\s\S]+?-----END (?:RSA |EC |ENCRYPTED )?PRIVATE KEY-----/g,
    label: '[REDACTED:PRIVATE_KEY]' },
  // OAuth tokens in URLs
  { pattern: /[?&](?:access_token|code|id_token)=[A-Za-z0-9\-_.~+/]+(&|$)/g,
    label: '?[REDACTED:OAUTH_TOKEN]$1' },
  // Password in connection strings
  { pattern: /(?:Password|password|pwd|Pwd)=[^;,"'\s]{3,}/g,
    label: 'password=[REDACTED:PASSWORD]' },
];

// JSON keys that should have their values redacted
const SECRET_KEYS = new Set([
  'password', 'Password', 'secret', 'Secret', 'clientSecret', 'ClientSecret',
  'apiKey', 'api_key', 'ApiKey', 'token', 'Token', 'accessToken', 'access_token',
  'sharedKey', 'connectionString', 'credentials', 'Credentials'
]);

export class SecretRedactor {
  /** Redact a plain string in place */
  static redactString(value: string): string {
    let out = value;
    for (const rule of REDACTION_RULES) {
      rule.pattern.lastIndex = 0;
      out = out.replace(rule.pattern, rule.label);
    }
    return out;
  }

  /** Redact an arbitrary JSON-like object recursively */
  static redactObject(obj: unknown, depth = 0): unknown {
    if (depth > 20) return obj;
    if (typeof obj === 'string') return SecretRedactor.redactString(obj);
    if (Array.isArray(obj)) return obj.map(v => SecretRedactor.redactObject(v, depth + 1));
    if (obj !== null && typeof obj === 'object') {
      const result: Record<string, unknown> = {};
      for (const [k, v] of Object.entries(obj as Record<string, unknown>)) {
        if (SECRET_KEYS.has(k)) {
          result[k] = '[REDACTED:SECRET]';
        } else {
          result[k] = SecretRedactor.redactObject(v, depth + 1);
        }
      }
      return result;
    }
    return obj;
  }

  /** Full pass over the NormalizedAnalysisGraph */
  static redact(graph: NormalizedAnalysisGraph): void {
    // Redact flow action inputs
    for (const flow of graph.flows) {
      for (const action of flow.actions) {
        action.redactedInputs = SecretRedactor.redactObject(action.inputs) as Record<string, unknown>;
      }
    }

    // Redact env var default/current values that look like secrets
    for (const ev of graph.envVars) {
      if (ev.type?.toLowerCase() === 'secret') {
        ev.defaultValue = ev.defaultValue ? '[REDACTED:SECRET_ENV_VAR]' : ev.defaultValue;
        ev.currentValue = ev.currentValue && ev.currentValue !== 'not included in export'
          ? '[REDACTED:SECRET_ENV_VAR]' : ev.currentValue;
      }
    }

    // Redact canvas app formulas (in-place on redacted field only)
    for (const app of graph.canvasApps) {
      app.appOnStart.redacted = SecretRedactor.redactString(app.appOnStart.raw);
      for (const screen of app.screens) {
        screen.onVisible.redacted = SecretRedactor.redactString(screen.onVisible.raw);
        for (const kf of screen.keyFormulas) {
          kf.formula.redacted = SecretRedactor.redactString(kf.formula.raw);
        }
      }
    }
  }
}
