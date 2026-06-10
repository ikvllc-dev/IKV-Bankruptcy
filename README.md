# IKV Bankruptcy

Bankruptcy case management app for IKV, built with Google Apps Script and a static Armenian-language HTML frontend.

## Files

- `Code.gs` - Google Apps Script backend for Sheets, Drive folders, document uploads, audit logs, and document generation.
- `ikv_bankruptcy.html` - Frontend dashboard, case management UI, intake questionnaire, document tracker, and generated text templates.

## Setup

1. Open Google Apps Script.
2. Create or open the project for the backend.
3. Paste/update `Code.gs`.
4. Deploy as a Web App.
5. Copy the Web App URL into `SCRIPT_URL` in `ikv_bankruptcy.html`.
6. Open `ikv_bankruptcy.html` in a browser or host it as a static page.

## Important

This app stores sensitive client and legal data. Before production use, restrict access to the Apps Script web app and avoid deploying it publicly without authentication.
