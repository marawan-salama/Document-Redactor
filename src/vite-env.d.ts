/// <reference types="vite/client" />

interface ImportMetaEnv {
  readonly VITE_APP_TITLE: string;
  // Add other environment variables here
}

interface ImportMeta {
  readonly env: ImportMetaEnv;
  readonly hot: {
    accept: (callback?: (mod: any) => void) => void;
    dispose: (callback: () => void) => void;
    // Add other HMR methods as needed
  };
}
