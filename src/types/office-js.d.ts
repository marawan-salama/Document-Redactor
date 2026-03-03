// Type definitions for Office.js

declare namespace Office {
  interface Context {
    requirements: {
      isSetSupported: (name: string, minVersion?: string) => boolean;
    };
  }

  const context: Context;
}

declare module '@microsoft/office-js' {
  export * from 'office-js';
}

// Add global Office object
declare const Office: Office.Context;
