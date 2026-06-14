# IKV Case Management

Armenian-first legal case management for preparation, court process, and physical archive stages. The system uses Google Apps Script, Google Sheets, Drive, Calendar, and a static HTML frontend.

## Capabilities

- Editable case types and three-stage case lifecycle.
- Stable internal numbers and separate official court numbers.
- Incoming paper-mail lookup, receipt register, and printable labels.
- Physical archive room/shelf/box locator.
- Hearings, procedural deadlines, email reminders, and shared Calendar sync.
- Existing bankruptcy questionnaire, documents, debts, certificates, uploads, templates, folders, and audit history.
- Schema migration, scheduled backups, and restore-friendly spreadsheet copies.

## Production Deployment

1. Create an Apps Script project and replace its backend with `Code.gs`.
2. Deploy as a Web App, execute as the owner, and set access to **Anyone**.
3. `PUBLIC_ACCESS_ENABLED` is currently set to `true` in `Code.gs`.
5. Put the deployment URL in `SCRIPT_URL` inside `ikv_bankruptcy.html`.
6. Open **Կարգավորումներ** in the app:
   - save the shared Calendar ID;
   - save the backup Drive folder ID;
   - set default reminder days;
   - run **Միգրացնել առկա գործերը** once;
   - run **Տեղադրել ավտոմատ գործարկումները** once.

The migration creates hidden sheet backups, appends new columns without removing existing columns, assigns stable `IKV-000001` style numbers, and defaults legacy cases to `bankruptcy` / `preparation`.

## Backup And Restore

`backupSpreadsheet()` creates a complete timestamped spreadsheet copy every day after triggers are installed.

Restore procedure:

1. Stop editing and retain the current spreadsheet as evidence.
2. Open the required `IKV_Backup_YYYY-MM-DD_HHMMSS` copy.
3. Verify record counts for `Cases`, `Documents`, `Debts`, `BankCertificates`, `CaseFiles`, and `AuditLog`.
4. Replace `SPREADSHEET_ID` in `Code.gs` with the verified backup ID.
5. Redeploy Apps Script and run `repairAllSheetSchemas()`.
6. Confirm Drive folder links and run the acceptance checks below.

## Acceptance Smoke Test

1. Create a bankruptcy case and verify the original questionnaire and generated documents.
2. Create a non-bankruptcy case and confirm bankruptcy-only screens are hidden.
3. Move a case `preparation -> process -> archive`; confirm court and archive fields are required.
4. Confirm the internal number and Drive folder do not change.
5. Find the case from **Մուտքային փոստ**, log receipt, and print a label.
6. Add a hearing and deadline; confirm dashboard, email reminder, and shared Calendar event.
7. Verify every action in `AuditLog` and every new value directly in Sheets.

## Կարճ օգտագործողի ուղեցույց

- **Գործեր** բաժնում ընտրեք փուլը և սեղմեք **Ավելացնել գործ**։
- Սնանկության նախապատրաստման գործի համար լրացրեք ամբողջական հարցաթերթը։
- Գործի ներսում **Փուլ և արխիվ** ներդիրից տեղափոխեք գործը հաջորդ կամ նախորդ փուլ։
- Ընթացիկ փուլ տեղափոխելիս լրացրեք դատական համարը, դատարանը, ատյանը և դատավորին։
- Արխիվ տեղափոխելիս պարտադիր լրացրեք սենյակը, դարակը և արկղը։
- **Նիստեր և ժամկետներ** ներդիրում ավելացրեք նիստերը, վերջնաժամկետները և հիշեցման հասցեները։
- **Մուտքային փոստ** բաժնում գտեք գործը, նամակի վրա գրեք ներքին համարը, անհրաժեշտության դեպքում գրանցեք ստացումը և տպեք պիտակը։
- **Կարգավորումներ** բաժնում կառավարեք գործերի տեսակները, Calendar-ը, հիշեցումները և backup-ները։

## Technical Specification Build

```powershell
npm install
npm run build:spec
```

This regenerates `IKV_Technical_Specification.docx` from `build_tz.js`.
