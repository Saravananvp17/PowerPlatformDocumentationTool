import type { UrlCategory } from '../model/types';

export interface ClassifiedUrl {
  category: UrlCategory;
  isEnvSpecific: boolean;
  flagReason?: string;
}

export class UrlClassifier {
  static classify(url: string): ClassifiedUrl {
    const lower = url.toLowerCase();

    if (lower.includes('sharepoint.com') || lower.includes('sharepoint.office.com')) {
      const envSpecific = /\/(dev|test|staging|uat|prod|production)\//i.test(url) ||
                          /(dev|test|staging|uat)-/.test(lower);
      return {
        category: 'sharepoint',
        isEnvSpecific: envSpecific,
        flagReason: envSpecific ? 'SharePoint URL appears environment-specific' : undefined
      };
    }

    if (lower.includes('.dynamics.com') || lower.includes('.crm.dynamics.com') ||
        lower.includes('api.crm') || lower.includes('.dataverse.com')) {
      return {
        category: 'dataverse',
        isEnvSpecific: true,
        flagReason: 'Dataverse environment URL is always environment-specific — recommend env var'
      };
    }

    if (lower.includes('graph.microsoft.com') || lower.includes('teams.microsoft.com') ||
        lower.includes('microsoftgraph')) {
      return { category: 'graph-teams', isEnvSpecific: false };
    }

    if (lower.startsWith('http://localhost') || lower.startsWith('http://127.') ||
        lower.startsWith('http://0.0.0.0')) {
      return {
        category: 'local-dev',
        isEnvSpecific: true,
        flagReason: 'Localhost URL detected — likely hardcoded dev endpoint, HIGH RISK'
      };
    }

    const envSpecific = /(dev|test|staging|uat|prod|production)[\.\-_]/i.test(url);
    return {
      category: 'external-http',
      isEnvSpecific: envSpecific,
      flagReason: envSpecific ? 'URL appears environment-specific' : undefined
    };
  }
}
