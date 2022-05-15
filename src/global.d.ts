
export {};

declare global {
  interface Window {
    formulaCount: number;
    O365: {
      env: 'dev' | 'alpha' | 'beta' | 'prod';
      stsBaseURL: string;
      appDomain: string;
      redirectURI: string;
      clientId: string;
      loginPublicDomain: string;
      loginProtectedDomains: string[];
      appVersion: string;
      datacenter: string;
      piUrl: string;
      piToken: string;
      piFlushPeriod: number;
      formulaLimit: number;
    };
    auditIdObject: {
      [formula: string]: any;
    };
    opsSession: any;
  }
}
