# Power Platform Documentation Tool

**v0.1.0** — A local-only, cross-platform desktop application (Windows + macOS) that parses Power
Platform export artifacts and generates complete maintenance documentation — with zero data leaving
the machine.

---

## Quick Start (Development)

### Prerequisites

| Tool | Version |
|------|---------|
| Node.js | 20 LTS or later |
| npm | 10+ (bundled with Node 20) |
| Electron | downloaded automatically by `npm install` |

### Install & run

```bash
cd app
npm install          # downloads all deps including the Electron binary
npm run dev          # starts Vite dev server + Electron window in watch mode
```

> **First run tip:** `npm install` triggers a background download of the Electron binary
> (~120 MB). Run it once on a machine with internet access; subsequent `npm install` runs
> use the local cache.

---

## Supported Artifact Types

| File | Detected as |
|------|-------------|
| `*.msapp` | Standalone canvas app |
| `*.zip` containing `solution.xml` | Dataverse solution export |
| `*.zip` containing `manifest.json` + embedded `.msapp` | Power Apps environment export |
| `*.zip` containing workflow JSON(s) | Power Automate flow export |

> **Note:** Power Pages components have not been validated with this tool. Solutions containing
> Power Pages artifacts will still process, but those components may be omitted from the report.

---

## What It Analyses

| Component | Details extracted |
|-----------|-------------------|
| **Canvas Apps** | Screens, controls, variables, collections, formulas, navigation, App Checker findings, risk flags |
| **Model-Driven Apps** | Forms, views, site map components |
| **Cloud Flows** | All action types, triggers, run-after logic, Mermaid flow diagrams, error-handling gaps |
| **Copilot Studio Agents** | Topics (trigger phrases, action kinds, status), GPT instructions, knowledge files, channel configuration, AI feature settings |
| **AI Builder Models** | Custom Prompt models (prompt text, input variables, output formats) and pretrained models (Text Recognition, Receipt Scanning, Invoice Processing, etc.) with cross-links to flows that call them |
| **Managed Plans** | Power Apps Copilot solution blueprints — personas, user stories, planned artifacts, entity definitions, process diagrams (rendered as Mermaid flowcharts) |
| **Environment Variables** | Name, type, default value, current value, usage across flows and apps |
| **Connectors & Connection References** | Connector IDs, display names, where-used index |
| **Data Sources** | SharePoint lists, Dataverse tables (custom and standard), other connectors |
| **Security Roles** | Role names, privilege counts, privilege detail tables |
| **Web Resources** | Name, type, content summary |
| **Missing Dependencies** | Cross-references between components; flags unresolved connector and env var references |
| **Secret Redaction** | Bearer tokens, SAS signatures, account keys, OAuth secrets, and password-like values are replaced with `[REDACTED]` before any output is written |

---

## Output Files

For each analysis session the tool writes three files to the chosen output folder:

| File | Description |
|------|-------------|
| `<name>_report.html` | Single-page HTML report with sidebar navigation, collapsible sections, and Mermaid diagrams |
| `<name>_report.md` | GitHub-flavored Markdown with Mermaid fences |
| `<name>_spec.docx` | CoE Technical Specification Word document |

`<name>` is the solution's unique name (from `solution.xml`) or the source filename if no solution
name is available.

---

## npm Scripts

| Script | Purpose |
|--------|---------|
| `npm run dev` | Launch dev server + Electron (hot-reload) |
| `npm run build` | Compile renderer (Vite) + main process (tsc) |
| `npm run dist` | Full production build + electron-builder package |
| `npm run dist:win` | Windows NSIS installer (run on Windows) |
| `npm run dist:mac` | macOS DMG (run on macOS) |
| `npm run typecheck` | Renderer type-check only (`tsc --noEmit`) |
| `npx tsc -p tsconfig.electron.json --noEmit` | Main process type-check |

---

## Project Layout

```
app/
├── electron/
│   ├── main.ts          # Electron main process, all IPC handlers
│   └── preload.ts       # contextBridge API (window.ppdt)
├── renderer/
│   ├── index.html
│   ├── main.tsx         # React entry point
│   └── ui/
│       ├── store.ts     # Zustand wizard state
│       ├── App.tsx      # Root + step-bar
│       ├── StepEnv.tsx
│       ├── StepUpload.tsx
│       ├── StepAnalyzing.tsx
│       ├── StepGenerating.tsx
│       ├── StepDone.tsx
│       ├── StepError.tsx
│       └── styles.css
├── src/
│   ├── model/types.ts           # All shared TypeScript interfaces
│   ├── parsing/
│   │   ├── ArtifactDetector.ts  # ZIP/MSAPP type detection
│   │   ├── FlowDefinitionParser.ts
│   │   ├── MsappParser.ts
│   │   ├── SolutionZipParser.ts
│   │   ├── PowerAppsZipParser.ts
│   │   ├── SchemaEnrichmentParser.ts
│   │   └── ParserOrchestrator.ts
│   ├── analysis/
│   │   ├── SecretRedactor.ts
│   │   ├── UrlClassifier.ts
│   │   ├── VariableExtractor.ts
│   │   ├── NavigationInferencer.ts
│   │   ├── WhereUsedIndexBuilder.ts
│   │   ├── MissingDepsAnalyzer.ts
│   │   └── ExecutiveSummaryBuilder.ts
│   └── generators/
│       ├── HtmlGenerator.ts
│       ├── MarkdownGenerator.ts
│       └── DocxGenerator.ts
├── build/                       # Place icon.icns (macOS) and icon.ico (Windows) here
├── package.json
├── tsconfig.json                # renderer (ESNext / bundler resolution)
├── tsconfig.electron.json       # main process (CommonJS / node resolution)
└── vite.renderer.config.ts
```

---

## Security Notes

- **No network calls** — all processing is local. The Electron `webPreferences` have
  `nodeIntegration: false` and `contextIsolation: true`.
- **Secret redaction** — before any output is written, `SecretRedactor` strips Bearer tokens,
  Azure SAS signatures, account keys, PEM blocks, OAuth secrets, and password-like JSON keys.
  Redacted values are replaced with `[REDACTED]`; raw values are never written to disk.
- **IPC surface** — the renderer can only call the methods exposed via `contextBridge`
  in `preload.ts`. All file I/O runs exclusively in the main process.

---

## Adding a New Parser

1. Add types to `src/model/types.ts`.
2. Create `src/parsing/MyParser.ts` — implement `static async parse(zip: JSZip, ...): Promise<MyType[]>`.
3. Call it from `SolutionZipParser.ts` (or the appropriate parser) and wire the result into `ParserOrchestrator.ts`.
4. Add a section to `HtmlGenerator.ts`, `MarkdownGenerator.ts`, and `DocxGenerator.ts`.
5. Update `ExecutiveSummaryBuilder.ts` to include the new component count.

---

## Known Limitations

- Electron binary must be downloaded from GitHub on first `npm install` (requires internet).
- Parsers run synchronously in the main process; very large solution ZIPs (>500 MB) may
  block the UI briefly. Migrating to `worker_threads` is the recommended next step.
- Power Pages components have not been validated and may not appear in the report.
- Vitest unit test fixtures are scaffolded but test files are not yet written.
