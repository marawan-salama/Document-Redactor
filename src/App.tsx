// src/App.tsx
import React, { useEffect, useState } from "react";
import { redactDocument } from "./utils/redaction";
import "./styles/main.css";

declare global {
  interface Window {
    Office?: any;
    Word?: any;
  }
}

/**
 * Root application component for the Word add-in UI.
 * Responsible for:
 * - Detecting Office/Word readiness
 * - Determining Tracking Changes API availability
 * - Orchestrating the redaction workflow
 * - Presenting status and error feedback to the user
 */
const App: React.FC = () => {
  /** Indicates whether Office.js and Word are fully initialized */
  const [officeReady, setOfficeReady] = useState(false);

  /** Prevents concurrent redaction runs and disables the UI while active */
  const [running, setRunning] = useState(false);

  /** Human-readable progress messages emitted during redaction */
  const [status, setStatus] = useState<string | null>(null);

  /** Fatal or recoverable error messages shown to the user */
  const [error, setError] = useState<string | null>(null);

  /**
   * Reflects availability of the *Tracking Changes API* (WordApi 1.5),
   * not the current document toggle state.
   */
  const [trackingApiAvailable, setTrackingApiAvailable] = useState(false);

  /**
   * Initializes Office.js and checks runtime capabilities.
   * This runs once on mount and never blocks rendering.
   */
  useEffect(() => {
    const init = async () => {
      if (!window.Office || !window.Word) {
        setError("Open this add-in inside Microsoft Word.");
        return;
      }

      // Ensure Office.js is fully ready before accessing context APIs
      await new Promise<void>((resolve) => {
        try {
          window.Office.onReady(() => resolve());
        } catch {
          resolve();
        }
      });

      setOfficeReady(true);

      // Requirement-set gate for Tracking Changes support
      const reqSupported =
        !!window.Office?.context?.requirements?.isSetSupported &&
        window.Office.context.requirements.isSetSupported("WordApi", "1.5");

      if (!reqSupported) {
        setTrackingApiAvailable(false);
        return;
      }

      // Runtime capability check (differs across Word Web / Desktop)
      try {
        const hasTrackChangesProp = await window.Word.run(async (context: any) => {
          return typeof context.document.trackChanges === "boolean";
        });

        setTrackingApiAvailable(!!hasTrackChangesProp);
      } catch {
        setTrackingApiAvailable(false);
      }
    };

    init();
  }, []);

  /**
   * Executes the document redaction pipeline.
   * Handles UI state transitions and error reporting.
   */
  const onRedact = async () => {
    if (!officeReady) return;

    setRunning(true);
    setError(null);
    setStatus("Redacting document…");

    try {
      await redactDocument(
        (msg) => setStatus(msg),
        (err) => setError(typeof err === "string" ? err : "Unexpected error")
      );

      setStatus("Redaction complete.");
    } catch {
      setError("Something went wrong while redacting the document.");
    } finally {
      setRunning(false);
    }
  };

  return (
    <div className="app-root">
      <div className="card">
        <div className="header">
          <img
            src="/icon-80.png"
            alt="Document Redactor"
            className="app-icon"
          />
          <div>
            <h1>Document Redactor</h1>
            <p className="subtitle">
              Redact emails, phone numbers, and SSNs in one click.
            </p>
          </div>
        </div>

        <div className="divider" />

        <button
          className="primary-button"
          onClick={onRedact}
          disabled={running || !officeReady}
        >
          {running ? "Redacting…" : "Redact Document"}
        </button>

        {/* Display capability status only when the API is genuinely available */}
        {trackingApiAvailable && (
          <div className="tracking-pill">
            <span>Tracking Changes</span>
            <strong>Available</strong>
          </div>
        )}

        {status && <div className="status">{status}</div>}
        {error && <div className="error">{error}</div>}

        <div className="footer-note">
          Redaction replaces detected values with <code>[REDACTED]</code>.
        </div>
      </div>
    </div>
  );
};

export default App;