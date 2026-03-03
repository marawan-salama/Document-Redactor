# Document Redaction Add-in

## Overview

This project implements a Microsoft Word add-in that redacts sensitive information from a document, inserts a confidentiality header, and enables Track Changes when supported.

The add-in identifies and replaces:
- Email addresses  
- Phone numbers  
- Social Security Numbers (full SSNs and contextual last-4 digits)

All replacements are made directly in the document using the Word JavaScript API and are idempotent. Already redacted content is not processed again.

The solution is written in TypeScript, uses React for the UI, and applies custom CSS without external UI libraries.

---

## Features

### Sensitive Information Redaction
- Detects and replaces:
  - Emails
  - Phone numbers in common formats
  - Full Social Security Numbers
  - SSN last-4 digits only when SSN context is present
- Works in:
  - Body text
  - Tables
- Avoids duplicate or partial redactions

### Confidentiality Header
- Inserts “CONFIDENTIAL DOCUMENT” at the top of the document
- Prevents duplicate insertion
- Falls back to a top-of-document banner if headers are not supported

### Track Changes
- Attempts to enable Track Changes using the Word JavaScript API (Requirement Set 1.5)
- Gracefully continues when the API is unavailable (for example, Word on the web)

---

## Prerequisites

The following must be installed on the machine before running the project:

- Node.js (v18 or newer recommended)
- npm (included with Node.js)
- Microsoft Word (Word on the Web or Word Desktop)

The add-in runs on HTTPS locally. If prompted, allow or trust the local development certificate.

On a fresh machine, HTTPS or certificate errors are common. Installing the Office Add-in development certificate tools is recommended:

```bash
npm install -g office-addin-dev-certs
office-addin-dev-certs install
````

---

## Running the Project

1. Install dependencies:

```bash
npm install
```

2. Start the development server:

```bash
npm start
```

This will:

* Start a local server on `https://localhost:3000`
* Compile the TypeScript source
* Serve the add-in for sideloading

---

## Uploading / Sideloading the Add-in in Word

### Word on the Web (Recommended for Testing)

1. Open Word on the Web
2. Open any document
3. Go to **Insert → Add-ins → My Add-ins**
4. Click **Upload My Add-in**
5. Upload the provided `manifest.xml`
6. The add-in will appear in the ribbon under **Home → Document Redaction**

This method was used for all functional testing.

---

## Testing

The add-in was tested using the provided `Document-To-Be-Redacted.docx` file in Word on the Web.

Test coverage includes:

* Emails, phone numbers, and SSNs in body text and tables
* Context-aware SSN last-4 redaction
* Idempotent behavior (already redacted values are skipped)
* Confidential header insertion
* Track Changes API detection

---

## Platform Notes

### Word on the Web

* Track Changes API support can be inconsistent depending on the environment
* The add-in attempts to enable tracking and continues gracefully if unsupported

### Word Desktop (Windows / macOS)

* The same code path enables Track Changes when supported
* Microsoft Word on macOS does not allow sideloading Office add-ins without enterprise deployment
* Due to this limitation and lack of access to a Windows machine, desktop testing was not possible
* All functional testing was performed using Word on the Web

The code includes defensive checks to ensure compatibility across environments.
