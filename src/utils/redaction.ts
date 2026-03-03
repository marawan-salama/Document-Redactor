// src/utils/redaction.ts
declare const Office: any;
declare const Word: any;

type StatusCallback = (message: string) => void;
type ErrorCallback = (error: any) => void;

const REDACTION_TEXT = "[REDACTED]";

const EMAIL_RE = /\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b/g;
const PHONE_DETECT_RE =
  /(?:\+?\d{1,3}[-.\s]?)?(?:\(\d{3}\)|\d{3})[-.\s]?\d{3}[-.\s]?\d{4}/g;
const SSN_FULL_RE = /\b\d{3}[- ]?\d{2}[- ]?\d{4}\b/g;

const LAST4_RE = /\b\d{4}\b/g;
const SSN_CONTEXT_RE =
  /\b(ssn|social security|social-security|last\s*4|last\s*four|ending\s*in|last four digits)\b/i;

/**
 * Normalizes Unicode dash variants to ASCII hyphen to avoid “same-looking” characters
 * breaking matching logic.
 */
function normalizeDashes(s: string): string {
  return s.replace(/[\u2010\u2011\u2012\u2013\u2014\u2212]/g, "-");
}

/** Extracts digits only (used for normalization and “last-10” phone identity). */
function digitsOnly(s: string): string {
  return s.replace(/\D/g, "");
}

/** Deduplicates case-insensitively while preserving a stable output order. */
function uniqueLower(list: string[]): string[] {
  const seen = new Set<string>();
  for (const v of list) seen.add(v.toLowerCase());
  return Array.from(seen);
}

/**
 * Formats Office.js errors with debug details when available.
 * This helps diagnose host-specific API failures (Word Web vs Desktop).
 */
function formatOfficeError(e: any): string {
  const base = e?.message ? String(e.message) : "Unknown error";
  const code = e?.code ? ` (code: ${e.code})` : "";
  const debug = e?.debugInfo ? `\nDEBUG INFO:\n${JSON.stringify(e.debugInfo, null, 2)}` : "";
  return `${base}${code}${debug}`;
}

/** Reads the full document body as plain text (used for detection only). */
async function getDocumentText(context: any): Promise<string> {
  const body = context.document.body;
  body.load("text");
  await context.sync();
  return typeof body.text === "string" ? body.text : "";
}

/**
 * Enables Track Changes when the Word API supports it.
 * If not supported, we proceed with redaction (requirement: “only use when available”).
 */
async function tryEnableTrackChanges(context: any, status: StatusCallback): Promise<boolean> {
  const supported =
    !!Office?.context?.requirements?.isSetSupported &&
    Office.context.requirements.isSetSupported("WordApi", "1.5");

  if (!supported) {
    status("Track Changes isn’t available in this Word environment. Redaction will still run.");
    return false;
  }

  if (typeof context.document.trackChanges !== "boolean") {
    status("Track Changes can’t be controlled here. Redaction will still run.");
    return false;
  }

  try {
    context.document.trackChanges = true;
    await context.sync();
    status("Track Changes: requested.");
    return true;
  } catch {
    status("Couldn’t enable Track Changes here. Redaction will still run.");
    return false;
  }
}

/**
 * Inserts "CONFIDENTIAL DOCUMENT" in the document header when supported.
 * Falls back to a top-of-document banner to avoid host/version-specific failures.
 */
async function addConfidentialHeader(context: any, status: StatusCallback): Promise<void> {
  status('Adding "CONFIDENTIAL DOCUMENT" header...');

  try {
    if (typeof context.document.body.insertHeader === "function") {
      const header = context.document.body.insertHeader(Word.HeaderFooterType.primary);

      try {
        header.load("text");
        await context.sync();
        if (typeof header.text === "string" && header.text.includes("CONFIDENTIAL DOCUMENT")) {
          status("Header already present. Skipping.");
          return;
        }
      } catch {
        // If header text isn’t readable in this host, we still attempt insertion idempotently below.
      }

      if (typeof header.insertParagraph === "function") {
        const para = header.insertParagraph("CONFIDENTIAL DOCUMENT", Word.InsertLocation.start);
        try {
          const r = para.getRange();
          if (r?.font) {
            r.font.bold = true;
            r.font.size = 12;
            r.font.color = "#c50f1f";
          }
        } catch {
          // Styling is best-effort; insertion is the functional requirement.
        }
      } else if (typeof header.insertText === "function") {
        header.insertText("CONFIDENTIAL DOCUMENT\n", Word.InsertLocation.start);
      }

      await context.sync();
      status("Header added.");
      return;
    }
  } catch {
    // Header insertion is host-dependent; fall back to body banner.
  }

  try {
    const paras = context.document.body.paragraphs;
    paras.load("items/text");
    await context.sync();

    const firstText =
      paras?.items?.length && typeof paras.items[0].text === "string" ? paras.items[0].text : "";

    if (firstText.includes("CONFIDENTIAL DOCUMENT")) {
      status("Top banner already present. Skipping.");
      return;
    }

    const p = context.document.body.insertParagraph("CONFIDENTIAL DOCUMENT", Word.InsertLocation.start);
    try {
      const r = p.getRange();
      if (r?.font) {
        r.font.bold = true;
        r.font.size = 12;
        r.font.color = "#c50f1f";
      }
    } catch {
      // Styling is best-effort.
    }

    await context.sync();
    status("Header not supported here; added a top-of-document banner instead.");
  } catch {
    status("Could not add header/banner. Continuing.");
  }
}

/**
 * Replaces matching ranges with the redaction marker.
 * This is intentionally range-based (not full-body rewrite) to keep Track Changes readable.
 * Matches that already contain the marker are skipped for idempotency.
 */
async function searchAndReplace(context: any, searchText: string, options: any): Promise<number> {
  const results = context.document.body.search(searchText, options);
  results.load("items");
  await context.sync();

  const items = results.items ?? [];
  if (!items.length) return 0;

  results.load("items/text");
  await context.sync();

  let replaced = 0;
  for (const item of items) {
    const t = typeof item.text === "string" ? item.text : "";
    if (t.includes(REDACTION_TEXT)) continue;
    item.insertText(REDACTION_TEXT, Word.InsertLocation.replace);
    replaced++;
  }

  await context.sync();
  return replaced;
}

/**
 * Redacts phone numbers by:
 * 1) Extracting phone candidates from plain text
 * 2) Normalizing them to a unique identity (last 10 digits)
 * 3) Searching for common display variants with punctuation-tolerant options
 * 4) Falling back to digits-only tolerant search for split runs / odd separators
 *
 * We stop after the first successful strategy per unique phone to avoid repeat replacements.
 */
async function redactPhones(context: any, text: string, status: StatusCallback): Promise<number> {
  const detectedRaw = text.match(PHONE_DETECT_RE) ?? [];
  status(`Detected ${detectedRaw.length} phone candidate(s). Redacting...`);

  const detected = detectedRaw.map(normalizeDashes);

  const digits10Set = new Set<string>();
  for (const raw of detected) {
    const d = digitsOnly(raw);
    if (d.length >= 10) digits10Set.add(d.slice(d.length - 10));
  }

  const punctTolerant = {
    matchCase: false,
    matchWholeWord: false,
    matchWildcards: false,
    ignorePunct: true,
    ignoreSpace: false,
  };

  const tolerant = {
    matchCase: false,
    matchWholeWord: false,
    matchWildcards: false,
    ignorePunct: true,
    ignoreSpace: true,
  };

  let total = 0;

  for (const d of digits10Set) {
    const area = d.slice(0, 3);
    const mid = d.slice(3, 6);
    const last = d.slice(6);

    const variants = [
      normalizeDashes(`(${area}) ${mid}-${last}`),
      normalizeDashes(`${area}-${mid}-${last}`),
      normalizeDashes(`${area} ${mid} ${last}`),
      normalizeDashes(`${area}.${mid}.${last}`),
      normalizeDashes(`+1 ${area}-${mid}-${last}`),
      normalizeDashes(`+20 ${area} ${mid} ${last}`),
    ];

    let replacedForThisPhone = 0;

    for (const v of variants) {
      replacedForThisPhone = await searchAndReplace(context, v, punctTolerant);
      if (replacedForThisPhone > 0) {
        total += replacedForThisPhone;
        break;
      }
    }

    if (replacedForThisPhone > 0) continue;

    replacedForThisPhone = await searchAndReplace(context, d, tolerant);
    total += replacedForThisPhone;
  }

  return total;
}

/**
 * Main entrypoint: performs detection on body text, then redacts via targeted range replacements.
 * Order is chosen to keep logic simple and results stable.
 */
export async function redactDocument(status: StatusCallback, error: ErrorCallback): Promise<void> {
  try {
    if (!Office || !Word) {
      throw new Error("Office.js is not available. Open this add-in in Microsoft Word.");
    }

    await Word.run(async (context: any) => {
      status("Reading document content...");
      const text = await getDocumentText(context);

      status("Checking Track Changes support...");
      const trackAttempted = await tryEnableTrackChanges(context, status);

      await addConfidentialHeader(context, status);

      let total = 0;

      const emails = uniqueLower(text.match(EMAIL_RE) ?? []);
      status(`Detected ${emails.length} email(s). Redacting...`);

      const exactEmailOptions = {
        matchCase: false,
        matchWholeWord: false,
        matchWildcards: false,
        ignorePunct: false,
        ignoreSpace: false,
      };

      for (const em of emails) {
        total += await searchAndReplace(context, em, exactEmailOptions);
      }

      total += await redactPhones(context, text, status);

      const ssns = text.match(SSN_FULL_RE) ?? [];
      status(`Detected ${ssns.length} SSN(s). Redacting...`);

      const tolerant = {
        matchCase: false,
        matchWholeWord: false,
        matchWildcards: false,
        ignorePunct: true,
        ignoreSpace: true,
      };

      const ssnDigitsSet = new Set<string>();
      for (const raw of ssns) {
        const d = digitsOnly(raw);
        if (d.length === 9) ssnDigitsSet.add(d);
      }

      for (const d of ssnDigitsSet) {
        const v1 = `${d.slice(0, 3)}-${d.slice(3, 5)}-${d.slice(5)}`;
        const v2 = `${d.slice(0, 3)} ${d.slice(3, 5)} ${d.slice(5)}`;
        total += await searchAndReplace(context, v1, tolerant);
        total += await searchAndReplace(context, v2, tolerant);
        total += await searchAndReplace(context, d, tolerant);
      }

      let last4Accepted = 0;
      const windowSize = 70;

      for (const m of text.matchAll(LAST4_RE)) {
        const last4 = m[0];
        const idx = m.index ?? -1;
        if (idx < 0) continue;

        const start = Math.max(0, idx - windowSize);
        const end = Math.min(text.length, idx + last4.length + windowSize);
        const neighborhood = text.slice(start, end);

        if (!SSN_CONTEXT_RE.test(neighborhood)) continue;

        total += await searchAndReplace(context, last4, tolerant);
        last4Accepted++;
      }

      status(`Detected ${last4Accepted} SSN last-4 digit(s) by context.`);

      status(`Redaction complete. ${total} replacement(s) made.`);
      void trackAttempted;
    });
  } catch (e: any) {
    const msg = formatOfficeError(e);
    error(msg);
    throw e;
  }
}