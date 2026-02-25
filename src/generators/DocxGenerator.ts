import {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, PageNumber, PageBreak, LevelFormat, TableOfContents,
  TabStopType, TabStopPosition
} from 'docx';
import fs from 'fs';
import type {
  NormalizedAnalysisGraph, Flow, CanvasApp, ModelDrivenApp,
  EnvVar, DataSource, SecurityRole
} from '../model/types';

// ─── Constants ────────────────────────────────────────────────────────────────
const BLUE = '1F497D';
const CONTENT_W = 9360;
const COL1 = 2520, COL2 = 6840;
const bdr = { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' };
const borders = { top: bdr, bottom: bdr, left: bdr, right: bdr };

// ─── Helpers ──────────────────────────────────────────────────────────────────
function cell(text: string, w: number, header = false): TableCell {
  return new TableCell({
    width: { size: w, type: WidthType.DXA }, borders,
    shading: { fill: header ? BLUE : 'FFFFFF', type: ShadingType.CLEAR },
    margins: { top: 60, bottom: 60, left: 120, right: 120 },
    children: [new Paragraph({
      children: [new TextRun({ text, font: 'Arial', size: 20, bold: header, color: header ? 'FFFFFF' : '000000' })]
    })]
  });
}
function shadeCell(text: string, w: number): TableCell {
  return new TableCell({
    width: { size: w, type: WidthType.DXA }, borders,
    shading: { fill: 'F2F2F2', type: ShadingType.CLEAR },
    margins: { top: 60, bottom: 60, left: 120, right: 120 },
    children: [new Paragraph({ children: [new TextRun({ text, font: 'Arial', size: 20 })] })]
  });
}
function codeCell(text: string, w: number): TableCell {
  return new TableCell({
    width: { size: w, type: WidthType.DXA }, borders,
    shading: { fill: 'F5F5F5', type: ShadingType.CLEAR },
    margins: { top: 60, bottom: 60, left: 120, right: 120 },
    children: [new Paragraph({ children: [new TextRun({ text, font: 'Courier New', size: 18 })] })]
  });
}
function h1(text: string) {
  return new Paragraph({ heading: HeadingLevel.HEADING_1, pageBreakBefore: true,
    children: [new TextRun({ text, font: 'Arial', size: 32, bold: true, color: BLUE })] });
}
function h2(text: string) {
  return new Paragraph({ heading: HeadingLevel.HEADING_2,
    children: [new TextRun({ text, font: 'Arial', size: 26, bold: true, color: BLUE })] });
}
function h3(text: string) {
  return new Paragraph({ heading: HeadingLevel.HEADING_3,
    children: [new TextRun({ text, font: 'Arial', size: 24, bold: true })] });
}
function para(text: string) {
  return new Paragraph({ spacing: { after: 120 },
    children: [new TextRun({ text, font: 'Arial', size: 22 })] });
}
function placeholder(text: string) {
  return new Paragraph({ spacing: { after: 120 },
    shading: { fill: 'FFF3CD', type: ShadingType.CLEAR },
    children: [new TextRun({ text: `[TODO: ${text}]`, font: 'Arial', size: 22, color: '856404', italics: true })] });
}
function bullet(text: string) {
  return new Paragraph({ numbering: { reference: 'bullets', level: 0 },
    children: [new TextRun({ text, font: 'Arial', size: 22 })] });
}
function gap() { return new Paragraph({ children: [new TextRun('')] }); }
function twoColRow(label: string, value: string, shade = false) {
  return new TableRow({ children: [shadeCell(label, COL1), (shade ? shadeCell : (t: string, w: number) =>
    new TableCell({ width: { size: w, type: WidthType.DXA }, borders,
      shading: { fill: 'FFFFFF', type: ShadingType.CLEAR },
      margins: { top: 60, bottom: 60, left: 120, right: 120 },
      children: [new Paragraph({ children: [new TextRun({ text: t, font: 'Arial', size: 20 })] })]
    }))(value, COL2)] });
}
function twoColTable(rows: [string, string][]): Table {
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: [COL1, COL2],
    rows: rows.map(([l, v]) => new TableRow({ children: [shadeCell(l, COL1),
      new TableCell({ width: { size: COL2, type: WidthType.DXA }, borders,
        shading: { fill: 'FFFFFF', type: ShadingType.CLEAR },
        margins: { top: 60, bottom: 60, left: 120, right: 120 },
        children: [new Paragraph({ children: [new TextRun({ text: v, font: 'Arial', size: 20 })] })]
      })
    ]}))
  });
}

// ─────────────────────────────────────────────────────────────────────────────

export class DocxGenerator {
  static async generate(graph: NormalizedAnalysisGraph, outputPath: string): Promise<void> {
    const numbering = { config: [
      { reference: 'bullets', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•',
          alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: 'numbers', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.',
          alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] }
    ]};

    const styles = {
      default: { document: { run: { font: 'Arial', size: 22 } } },
      paragraphStyles: [
        { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { size: 32, bold: true, font: 'Arial', color: BLUE },
          paragraph: { spacing: { before: 360, after: 240 }, outlineLevel: 0 } },
        { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { size: 26, bold: true, font: 'Arial', color: BLUE },
          paragraph: { spacing: { before: 240, after: 160 }, outlineLevel: 1 } },
        { id: 'Heading3', name: 'Heading 3', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { size: 24, bold: true, font: 'Arial' },
          paragraph: { spacing: { before: 200, after: 120 }, outlineLevel: 2 } }
      ]
    };

    const doc = new Document({ styles, numbering, sections: [{
      properties: {
        page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } }
      },
      headers: { default: new Header({ children: [new Paragraph({
        border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: BLUE, space: 2 } },
        tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
        children: [
          new TextRun({ text: `${graph.solution?.displayName ?? graph.meta.fileName} — Technical Specification`, font: 'Arial', size: 18, color: '595959' }),
          new TextRun({ text: '\t', font: 'Arial', size: 18 }),
          new TextRun({ text: graph.meta.environment, font: 'Arial', size: 18, bold: true, color: BLUE })
        ]
      })]})},
      footers: { default: new Footer({ children: [new Paragraph({
        border: { top: { style: BorderStyle.SINGLE, size: 4, color: BLUE, space: 2 } },
        tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
        children: [
          new TextRun({ text: `PP Doc Tool v${graph.meta.toolVersion} · ${new Date().toLocaleDateString()}`, font: 'Arial', size: 18, color: '595959' }),
          new TextRun({ text: '\t', font: 'Arial', size: 18 }),
          new TextRun({ text: 'Page ', font: 'Arial', size: 18, color: '595959' }),
          new TextRun({ children: [PageNumber.CURRENT], font: 'Arial', size: 18, color: '595959' }),
          new TextRun({ text: ' of ', font: 'Arial', size: 18, color: '595959' }),
          new TextRun({ children: [PageNumber.TOTAL_PAGES], font: 'Arial', size: 18, color: '595959' })
        ]
      })]})},
      children: [
        // Cover
        ...DocxGenerator.coverPage(graph),
        new Paragraph({ children: [new PageBreak()] }),
        // TOC
        new TableOfContents('Table of Contents', { hyperlink: true, headingStyleRange: '1-3' }),
        new Paragraph({ children: [new PageBreak()] }),
        // Sections
        ...DocxGenerator.section1(graph),
        ...DocxGenerator.section2(graph),
        ...DocxGenerator.section3(graph),
        ...DocxGenerator.section4(graph),
        ...DocxGenerator.section5(graph),
        ...DocxGenerator.section6(graph),
        ...DocxGenerator.section7(graph),
        ...DocxGenerator.appendixA(graph),
        ...DocxGenerator.appendixB(graph),
      ]
    }]});

    const buf = await Packer.toBuffer(doc);
    fs.writeFileSync(outputPath, buf);
  }

  // ── Cover Page ──────────────────────────────────────────────────────────────
  private static coverPage(graph: NormalizedAnalysisGraph): Paragraph[] {
    const title = graph.solution?.displayName ?? graph.meta.fileName;
    return [
      gap(), gap(), gap(), gap(),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 240 },
        children: [new TextRun({ text: title, font: 'Arial', size: 56, bold: true, color: BLUE })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 120 },
        children: [new TextRun({ text: 'Technical Specification', font: 'Arial', size: 36, color: '595959' })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 120 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: BLUE, space: 2 } },
        children: [new TextRun('')] }),
      gap(),
      new Paragraph({ alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: `Environment: ${graph.meta.environment}  ·  Version: ${graph.solution?.version ?? 'N/A'}  ·  ${graph.meta.analysedAt.substring(0, 10)}`, font: 'Arial', size: 22, color: '595959' })] }),
    ];
  }

  // ── Section 1: Introduction ─────────────────────────────────────────────────
  private static section1(graph: NormalizedAnalysisGraph): (Paragraph | Table)[] {
    const s = graph.solution;
    return [
      h1('1. Introduction & Context'), gap(),
      h2('1.1 Document Purpose'),
      para('This document provides a complete technical specification of the Power Platform artifact analysed by PP Doc Tool. It is intended for maintenance and support staff.'),
      gap(),
      h2('1.2 Artifact Information'),
      twoColTable([
        ['File Name', graph.meta.fileName],
        ['Artifact Type', graph.meta.artifactType],
        ['Environment', graph.meta.environment],
        ['Solution Name', s?.displayName ?? 'N/A'],
        ['Unique Name', s?.uniqueName ?? 'N/A'],
        ['Version', s?.version ?? 'N/A'],
        ['Managed', s ? (s.isManaged ? 'Yes (Managed)' : 'No (Unmanaged)') : 'N/A'],
        ['Publisher', s ? `${s.publisher.displayName} (${s.publisher.uniqueName})` : 'N/A'],
        ['Analysed', graph.meta.analysedAt],
        ['Tool Version', graph.meta.toolVersion]
      ]),
      gap(), h2('1.3 Business Context'), placeholder('Describe the business purpose of this solution and the problem it solves'),
      gap(), h2('1.4 Stakeholders'), placeholder('List the key stakeholders, business owners, and technical contacts'),
    ];
  }

  // ── Section 2: Functional Overview ─────────────────────────────────────────
  private static section2(graph: NormalizedAnalysisGraph): (Paragraph | Table)[] {
    const es = graph.executiveSummary!;
    const countRows = Object.entries(es.componentCounts).filter(([, v]) => v > 0)
      .map(([k, v]) => new TableRow({ children: [shadeCell(k, COL1),
        new TableCell({ width: { size: COL2, type: WidthType.DXA }, borders,
          shading: { fill: 'FFFFFF', type: ShadingType.CLEAR },
          margins: { top: 60, bottom: 60, left: 120, right: 120 },
          children: [new Paragraph({ children: [new TextRun({ text: String(v), font: 'Arial', size: 20 })] })] })
      ]}));

    return [
      h1('2. Functional Overview'), gap(),
      h2('2.1 Solution / Artifact Summary'),
      new Table({ width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: [COL1, COL2],
        rows: [new TableRow({ children: [cell('Component', COL1, true), cell('Count', COL2, true)] }), ...countRows] }),
      gap(), h2('2.2 Business Process Description'),
      placeholder('Describe the end-to-end business process this solution automates or supports'),
      gap(), h2('2.3 Key User Personas'),
      placeholder('Describe the primary users (e.g. Field engineers, Administrators, Managers)'),
    ];
  }

  // ── Section 3: System Specification ─────────────────────────────────────────
  private static section3(graph: NormalizedAnalysisGraph): (Paragraph | Table)[] {
    const items: (Paragraph | Table)[] = [h1('3. System Specification'), gap()];

    // Env Vars
    items.push(h2('3.1 Environment Variables'));
    if (graph.envVars.length === 0) {
      items.push(para('No environment variables detected in this export.'));
    } else {
      const evRows = graph.envVars.map(ev => new TableRow({ children: [
        shadeCell(ev.displayName, 2000),
        codeCell(ev.schemaName, 1800),
        new TableCell({ width: { size: 1200, type: WidthType.DXA }, borders, margins: { top: 60, bottom: 60, left: 120, right: 120 },
          children: [new Paragraph({ children: [new TextRun({ text: ev.type, font: 'Arial', size: 20 })] })] }),
        new TableCell({ width: { size: 2160, type: WidthType.DXA }, borders, margins: { top: 60, bottom: 60, left: 120, right: 120 },
          shading: { fill: ev.currentValue === 'not included in export' ? 'FFF3CD' : 'FFFFFF', type: ShadingType.CLEAR },
          children: [new Paragraph({ children: [new TextRun({ text: ev.currentValue ?? '—', font: 'Arial', size: 20, italics: ev.currentValue === 'not included in export' })] })] }),
        new TableCell({ width: { size: 2200, type: WidthType.DXA }, borders, margins: { top: 60, bottom: 60, left: 120, right: 120 },
          children: [new Paragraph({ children: [new TextRun({ text: ev.description || '—', font: 'Arial', size: 20 })] })] }),
      ]}));
      items.push(new Table({ width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: [2000, 1800, 1200, 2160, 2200],
        rows: [new TableRow({ children: [cell('Display Name', 2000, true), cell('Schema Name', 1800, true), cell('Type', 1200, true), cell('Current Value', 2160, true), cell('Description', 2200, true)] }), ...evRows] }));
    }

    // Connectors
    items.push(gap(), h2('3.2 Connections & Connectors'));
    if (graph.connectors.length === 0) {
      items.push(para('No connectors detected.'));
    } else {
      for (const c of graph.connectors) {
        items.push(bullet(`${c.displayName || c.connectorId} — used in ${c.usedInFlows.length} flow(s) and ${c.usedInApps.length} app(s)`));
      }
    }

    // URLs
    items.push(gap(), h2('3.3 URL / Endpoint Inventory'));
    const flaggedUrls = graph.urls.filter(u => u.isEnvSpecific || u.category === 'local-dev');
    if (flaggedUrls.length === 0) {
      items.push(para('No environment-specific or flagged URLs detected.'));
    } else {
      const urlRows = flaggedUrls.map(u => new TableRow({ children: [
        codeCell(u.url.substring(0, 80), 5200),
        new TableCell({ width: { size: 1800, type: WidthType.DXA }, borders, margins: { top: 60, bottom: 60, left: 120, right: 120 },
          children: [new Paragraph({ children: [new TextRun({ text: u.category, font: 'Arial', size: 20 })] })] }),
        new TableCell({ width: { size: 2360, type: WidthType.DXA }, borders, margins: { top: 60, bottom: 60, left: 120, right: 120 },
          shading: { fill: 'FFF3CD', type: ShadingType.CLEAR },
          children: [new Paragraph({ children: [new TextRun({ text: u.flagReason ?? '—', font: 'Arial', size: 18, italics: true })] })] }),
      ]}));
      items.push(new Table({ width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: [5200, 1800, 2360],
        rows: [new TableRow({ children: [cell('URL', 5200, true), cell('Category', 1800, true), cell('Flag Reason', 2360, true)] }), ...urlRows] }));
    }
    return items;
  }

  // ── Section 4: Roles & Security ─────────────────────────────────────────────
  private static section4(graph: NormalizedAnalysisGraph): (Paragraph | Table)[] {
    const items: (Paragraph | Table)[] = [h1('4. Roles & Security'), gap()];
    if (graph.securityRoles.length === 0) {
      items.push(para('No security roles found in this export.'), placeholder('Document security roles and permission levels for this solution'));
    } else {
      for (const role of graph.securityRoles) {
        items.push(h2(role.name),
          twoColTable([
            ['Role ID', role.id],
            ['Assigned To Apps', role.assignedToApps.join(', ') || '—'],
            ['Total Privileges', String(role.privilegeCount)]
          ]), gap());
      }
    }
    if (graph.missingDeps.filter(d => d.type === 'SecurityRole').length > 0) {
      items.push(h2('4.1 Missing Role Definitions'));
      for (const dep of graph.missingDeps.filter(d => d.type === 'SecurityRole')) {
        items.push(bullet(`${dep.displayName ?? dep.identifier} — ${dep.impact}`));
      }
    }
    return items;
  }

  // ── Section 5: Source Data ──────────────────────────────────────────────────
  private static section5(graph: NormalizedAnalysisGraph): (Paragraph | Table)[] {
    const items: (Paragraph | Table)[] = [h1('5. Source Data'), gap()];
    if (graph.dataSources.length === 0) {
      items.push(para('No data sources detected.'));
      return items;
    }
    for (const ds of graph.dataSources) {
      items.push(h2(ds.displayName),
        twoColTable([
          ['Type', ds.type],
          ['Site / Environment URL', ds.siteUrl ?? ds.environmentUrl ?? '—'],
          ['List / Table', ds.listName ?? ds.tableName ?? '—'],
        ]));
      if (ds.columns?.length) {
        const colRows = ds.columns.map(c => new TableRow({ children: [
          shadeCell(c.name, 3000),
          new TableCell({ width: { size: 2000, type: WidthType.DXA }, borders, margins: { top: 60, bottom: 60, left: 120, right: 120 },
            children: [new Paragraph({ children: [new TextRun({ text: c.type ?? '—', font: 'Arial', size: 20 })] })] }),
          new TableCell({ width: { size: 4360, type: WidthType.DXA }, borders, margins: { top: 60, bottom: 60, left: 120, right: 120 },
            children: [new Paragraph({ children: [new TextRun({ text: c.notes ?? '—', font: 'Arial', size: 20 })] })] }),
        ]}));
        items.push(h3('Schema'), new Table({ width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: [3000, 2000, 4360],
          rows: [new TableRow({ children: [cell('Column', 3000, true), cell('Type', 2000, true), cell('Notes', 4360, true)] }), ...colRows] }));
      } else {
        items.push(placeholder(`Provide schema for ${ds.type} data source '${ds.displayName}'`));
      }
      items.push(gap());
    }
    return items;
  }

  // ── Section 6: Flows ────────────────────────────────────────────────────────
  private static section6(graph: NormalizedAnalysisGraph): (Paragraph | Table)[] {
    const items: (Paragraph | Table)[] = [h1('6. Flows'), gap()];
    if (graph.flows.length === 0) {
      items.push(para('No flows detected in this export.'));
      return items;
    }
    for (const flow of graph.flows) {
      const trig = flow.trigger.recurrence
        ? `Recurrence — every ${flow.trigger.recurrence.interval} ${flow.trigger.recurrence.frequency}`
        : flow.trigger.automated
        ? `${flow.trigger.automated.connectorId} / ${flow.trigger.automated.operationId}`
        : flow.triggerType;

      items.push(
        h2(flow.displayName),
        twoColTable([
          ['Trigger', trig],
          ['State', flow.state],
          ['Total Actions', String(flow.actions.length)],
          ['Error Handling', flow.errorHandling.hasExplicitHandling ? 'Yes — explicit handling detected' : 'No — no explicit error handling'],
          ['Source', flow.source.archivePath]
        ]),
        h3('Actions'),
        new Table({
          width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: [2200, 1600, 1600, 1200, 2760],
          rows: [
            new TableRow({ children: [cell('Action ID', 2200, true), cell('Type', 1600, true), cell('Connector', 1600, true), cell('Secure', 1200, true), cell('Run After', 2760, true)] }),
            ...flow.actions.map(a => new TableRow({ children: [
              codeCell(a.id.substring(0, 30), 2200),
              new TableCell({ width: { size: 1600, type: WidthType.DXA }, borders, margins: { top: 60, bottom: 60, left: 120, right: 120 },
                children: [new Paragraph({ children: [new TextRun({ text: a.type, font: 'Arial', size: 18 })] })] }),
              new TableCell({ width: { size: 1600, type: WidthType.DXA }, borders, margins: { top: 60, bottom: 60, left: 120, right: 120 },
                children: [new Paragraph({ children: [new TextRun({ text: a.connector ?? '—', font: 'Arial', size: 18 })] })] }),
              new TableCell({ width: { size: 1200, type: WidthType.DXA }, borders, margins: { top: 60, bottom: 60, left: 120, right: 120 },
                children: [new Paragraph({ children: [new TextRun({ text: (a.secureInputs ? 'In ' : '') + (a.secureOutputs ? 'Out' : '') || '—', font: 'Arial', size: 18 })] })] }),
              codeCell(a.runAfter.join(', ') || '—', 2760)
            ]}))
          ]
        }),
        h3('Flow Diagram (Mermaid Source)'),
        new Paragraph({ children: [new TextRun({ text: flow.mermaidDiagram, font: 'Courier New', size: 16 })] }),
        gap()
      );
    }
    return items;
  }

  // ── Section 7: Apps ─────────────────────────────────────────────────────────
  private static section7(graph: NormalizedAnalysisGraph): (Paragraph | Table)[] {
    const items: (Paragraph | Table)[] = [h1('7. Apps'), gap()];

    for (const app of graph.canvasApps) {
      items.push(h2(`Canvas App: ${app.displayName}`),
        twoColTable([
          ['Source', app.msappPath],
          ['Screens', String(app.screens.length)],
          ['Variables', String(app.variables.length)],
          ['Data Sources', String(app.dataSources.length)],
          ['Nav Confidence', app.navConfidence]
        ]));

      if (app.appOnStart.redacted) {
        items.push(h3('App.OnStart'));
        items.push(new Paragraph({ children: [new TextRun({ text: app.appOnStart.redacted.substring(0, 500), font: 'Courier New', size: 16 })] }));
      }

      if (app.screens.length) {
        items.push(h3(`Screens (${app.screens.length})`));
        const screenRows = app.screens.map(s => new TableRow({ children: [
          shadeCell(s.name, 2500),
          new TableCell({ width: { size: 6860, type: WidthType.DXA }, borders, margins: { top: 60, bottom: 60, left: 120, right: 120 },
            children: [new Paragraph({ children: [new TextRun({ text: s.onVisible.redacted.substring(0, 100) || '—', font: 'Courier New', size: 16 })] })] })
        ]}));
        items.push(new Table({ width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: [2500, 6860],
          rows: [new TableRow({ children: [cell('Screen', 2500, true), cell('OnVisible', 6860, true)] }), ...screenRows] }));
      }

      if (app.variables.length) {
        items.push(h3(`Variables (${app.variables.length})`));
        for (const v of app.variables.slice(0, 20)) {
          items.push(bullet(`${v.name} (${v.kind})`));
        }
      }
      items.push(gap());
    }

    for (const mda of graph.modelDrivenApps) {
      items.push(h2(`Model-Driven App: ${mda.displayName}`),
        twoColTable([
          ['Unique Name', mda.uniqueName],
          ['Assigned Roles', mda.assignedRoleNames.join(', ') || '—'],
          ['Exposed Tables', mda.exposedTables.join(', ') || '—']
        ]),
        h3('Sitemap Hierarchy (Mermaid Source)'),
        new Paragraph({ children: [new TextRun({ text: mda.mermaidSitemap, font: 'Courier New', size: 16 })] }),
        gap()
      );
    }
    return items;
  }

  // ── Appendix A: Checklist ────────────────────────────────────────────────────
  private static appendixA(graph: NormalizedAnalysisGraph): (Paragraph | Table)[] {
    if (graph.manualChecklist.length === 0) return [];
    const items: (Paragraph | Table)[] = [h1('Appendix A — Manual Completion Checklist')];
    const rows = graph.manualChecklist.map((c, i) => new TableRow({ children: [
      new TableCell({ width: { size: 600, type: WidthType.DXA }, borders, margins: { top: 60, bottom: 60, left: 120, right: 120 },
        children: [new Paragraph({ children: [new TextRun({ text: String(i + 1), font: 'Arial', size: 20 })] })] }),
      shadeCell(c.section, 1800),
      new TableCell({ width: { size: 5160, type: WidthType.DXA }, borders, margins: { top: 60, bottom: 60, left: 120, right: 120 },
        children: [new Paragraph({ children: [new TextRun({ text: c.item, font: 'Arial', size: 20 })] })] }),
      new TableCell({ width: { size: 1800, type: WidthType.DXA }, borders, margins: { top: 60, bottom: 60, left: 120, right: 120 },
        shading: { fill: c.priority === 'high' ? 'F8D7DA' : c.priority === 'medium' ? 'FFF3CD' : 'D1ECF1', type: ShadingType.CLEAR },
        children: [new Paragraph({ children: [new TextRun({ text: c.priority.toUpperCase(), font: 'Arial', size: 20, bold: true })] })] })
    ]}));
    items.push(new Table({ width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: [600, 1800, 5160, 1800],
      rows: [new TableRow({ children: [cell('#', 600, true), cell('Section', 1800, true), cell('Item', 5160, true), cell('Priority', 1800, true)] }), ...rows] }));
    return items;
  }

  // ── Appendix B: Parse Errors ─────────────────────────────────────────────────
  private static appendixB(graph: NormalizedAnalysisGraph): (Paragraph | Table)[] {
    if (graph.parseErrors.length === 0) return [];
    const items: (Paragraph | Table)[] = [h1('Appendix B — Parse Error Report')];
    for (const e of graph.parseErrors) {
      items.push(bullet(`[${e.partial ? 'Partial' : 'Failed'}] ${e.archivePath}: ${e.error}`));
    }
    return items;
  }
}
