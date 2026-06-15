// Generates the bilingual (EN + ARM) Technical Specification for the IKV
// case-management system expansion. Run: node build_tz.js
const path = require('path');
const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, LevelFormat, TableOfContents, HeadingLevel,
  BorderStyle, WidthType, ShadingType, PageNumber, PageBreak, VerticalAlign
} = require('docx');

const GOLD = '8a6d1f';
const NAVY = '1a1a2e';
const HY = '444444'; // muted colour for Armenian secondary text
const RED = '9c2a2a';

// ---------- helpers ----------
const border = { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' };
const borders = { top: border, bottom: border, left: border, right: border };
const cellMargins = { top: 80, bottom: 80, left: 120, right: 120 };

function en(text, opts = {}) {
  return new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text, ...opts })] });
}
function hy(text) {
  return new Paragraph({ spacing: { after: 160 }, children: [new TextRun({ text, color: HY, italics: true })] });
}
function bp(text, opts = {}) { // bilingual paragraph: EN then HY
  return [en(text.en, opts), hy(text.hy)];
}
function h1(enText, hyText) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 280, after: 120 },
    children: [
      new TextRun({ text: enText, bold: true, color: NAVY }),
      new TextRun({ text: '   /  ' + hyText, bold: true, color: GOLD, size: 24 })
    ]
  });
}
function h2(enText, hyText) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 200, after: 80 },
    children: [
      new TextRun({ text: enText, bold: true, color: NAVY }),
      new TextRun({ text: '  /  ' + hyText, bold: true, color: GOLD, size: 22 })
    ]
  });
}
function bullet(enText, hyText) {
  return [
    new Paragraph({ numbering: { reference: 'b', level: 0 }, spacing: { after: 20 },
      children: [new TextRun({ text: enText })] }),
    new Paragraph({ numbering: { reference: 'b', level: 0 }, spacing: { after: 120 },
      children: [new TextRun({ text: hyText, color: HY, italics: true })] })
  ];
}
function num(enText, hyText) {
  return [
    new Paragraph({ numbering: { reference: 'n', level: 0 }, spacing: { after: 20 },
      children: [new TextRun({ text: enText })] }),
    new Paragraph({ numbering: { reference: 'n', level: 0 }, spacing: { after: 120 },
      children: [new TextRun({ text: hyText, color: HY, italics: true })] })
  ];
}

// 3-col table: field | EN desc | HY desc
function dictTable(rows) {
  const head = new TableRow({ tableHeader: true, children: [
    ['Field / Դաշտ', 2200], ['Description (EN)', 3580], ['Նկարագրություն (ՀՅ)', 3580]
  ].map(([t, w]) => new TableCell({ borders, width: { size: w, type: WidthType.DXA }, margins: cellMargins,
    shading: { fill: NAVY, type: ShadingType.CLEAR },
    children: [new Paragraph({ children: [new TextRun({ text: t, bold: true, color: 'FFFFFF', size: 18 })] })] }))
  });
  const body = rows.map(r => new TableRow({ children: [
    new TableCell({ borders, width: { size: 2200, type: WidthType.DXA }, margins: cellMargins,
      children: [new Paragraph({ children: [new TextRun({ text: r[0], bold: true, font: 'Consolas', size: 16 })] })] }),
    new TableCell({ borders, width: { size: 3580, type: WidthType.DXA }, margins: cellMargins,
      children: [new Paragraph({ children: [new TextRun({ text: r[1], size: 18 })] })] }),
    new TableCell({ borders, width: { size: 3580, type: WidthType.DXA }, margins: cellMargins,
      children: [new Paragraph({ children: [new TextRun({ text: r[2], size: 18, color: HY })] })] })
  ] }));
  return new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [2200, 3580, 3580], rows: [head, ...body] });
}

const children = [];

// ---------- TITLE PAGE ----------
children.push(
  new Paragraph({ spacing: { before: 1800, after: 0 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: 'ԱՅ ՔԻ ՎԻ  /  IKV', bold: true, size: 56, color: NAVY })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 600 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: GOLD, space: 6 } },
    children: [new TextRun({ text: 'Law Firm Case-Management System', size: 24, color: GOLD })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 },
    children: [new TextRun({ text: 'TECHNICAL SPECIFICATION', bold: true, size: 40, color: NAVY })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 },
    children: [new TextRun({ text: 'ՏԵԽՆԻԿԱԿԱՆ ԱՌԱՋԱԴՐԱՆՔ (ՏԱ)', bold: true, size: 32, color: GOLD })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 600 },
    children: [new TextRun({ text: 'Expansion: Case Types, Stages, Mail Sorting & Physical Archive',
      italics: true, size: 22, color: HY })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 },
    children: [new TextRun({ text: 'Փուլերի, գործերի տեսակների, փոստի դասակարգման և ֆիզիկական արխիվի ընդլայնում',
      italics: true, size: 20, color: HY })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 1200, after: 20 },
    children: [new TextRun({ text: 'Version 1.0 (Draft for tender)  /  Տարբերակ 1.0 (նախագիծ՝ մրցույթի համար)', size: 18 })] }),
  new Paragraph({ alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: 'Date / Ամսաթիվ: 2026-06-14', size: 18 })] }),
  new Paragraph({ children: [new PageBreak()] })
);

// ---------- TOC ----------
children.push(
  new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: 'Table of Contents  /  Բովանդակություն', bold: true, color: NAVY })] }),
  new TableOfContents('Table of Contents', { hyperlink: true, headingStyleRange: '1-2' }),
  new Paragraph({ children: [new PageBreak()] })
);

// ---------- 1. PURPOSE ----------
children.push(h1('1. Purpose & Background', '1. Նպատակ և նախապատմություն'));
children.push(...bp({
  en: 'IKV is a law firm in the Republic of Armenia. It already operates a working in-house application that manages bankruptcy cases at the preparation stage. The firm now also handles cases that are at the trial (process) stage and cases that are finished and archived, and not all cases are bankruptcy cases. The purpose of this project is to expand the existing system so that it covers the full life-cycle of every case — preparation, process (trial), and archive — across all case types, and to add two practical tools the firm needs daily: a fast incoming-mail (paper letter) sorting aid and a physical-archive locator.',
  hy: 'ԱՅ ՔԻ ՎԻ-ն Հայաստանի Հանրապետությունում գործող իրավաբանական ընկերություն է։ Այն արդեն ունի աշխատող ներքին ծրագիր, որը կառավարում է սնանկության գործերը՝ նախապատրաստման փուլում։ Ընկերությունն այժմ վարում է նաև դատական (ընթացիկ) փուլում գտնվող գործեր և ավարտված, արխիվացված գործեր, ընդ որում՝ ոչ բոլոր գործերն են սնանկության։ Սույն նախագծի նպատակն է ընդլայնել գործող համակարգը, որպեսզի այն ընդգրկի յուրաքանչյուր գործի ողջ կյանքի ցիկլը՝ նախապատրաստում, ընթացիկ (դատական) և արխիվ, բոլոր տեսակի գործերի համար, և ավելացնի ընկերությանն ամենօրյա անհրաժեշտ երկու գործիք՝ մուտքային փոստի (թղթային նամակների) արագ դասակարգման օգնական և ֆիզիկական արխիվի տեղորոշիչ։'
}));
children.push(...bp({
  en: 'This document is a tender-ready technical specification (ТЗ). It tells the contractor exactly what to build, what must not be broken, and how the result will be accepted. It is written so that competing bidders can quote against the same scope.',
  hy: 'Սույն փաստաթուղթը մրցույթի համար պատրաստ տեխնիկական առաջադրանք է (ՏԱ)։ Այն կապալառուին հստակ ասում է՝ ինչ կառուցել, ինչը չպետք է խախտվի և ինչպես է ընդունվելու արդյունքը։ Այն գրված է այնպես, որ մրցակից առաջարկ ներկայացնողները կարողանան գնահատել նույն ծավալը։'
}));

// ---------- 2. CURRENT SYSTEM ----------
children.push(h1('2. Current System (As-Is)', '2. Գործող համակարգը (ինչպես կա)'));
children.push(...bp({
  en: 'The existing application is built on Google Apps Script (backend) with a single-page Armenian-language HTML frontend. Data is stored in Google Sheets and case documents in Google Drive. The following capabilities already exist and are in production use:',
  hy: 'Գործող ծրագիրը կառուցված է Google Apps Script-ի վրա (հետնամաս)՝ հայերեն մեկ-էջանոց HTML ճակատամասով։ Տվյալները պահվում են Google Sheets-ում, իսկ գործերի փաստաթղթերը՝ Google Drive-ում։ Ստորև բերված հնարավորություններն արդեն գոյություն ունեն և օգտագործվում են աշխատանքում.'
}));
[
  ['Case intake form and bankruptcy questionnaire/assessment.', 'Գործի ընդունման ձևաթուղթ և սնանկության հարցաթերթ/գնահատում։'],
  ['Per-case Google Drive folder tree (01 Personal Documents … 08 Other Documents), created automatically.', 'Յուրաքանչյուր գործի համար Google Drive թղթապանակների ծառ (01 Անձնական փաստաթղթեր … 08 Այլ փաստաթղթեր)՝ ստեղծվող ինքնաշխատ։'],
  ['Document tracker, debts register, and bank-certificate tracker.', 'Փաստաթղթերի հետևման գործիք, պարտքերի ռեգիստր և բանկային տեղեկանքների հետևում։'],
  ['Automatic generation of powers of attorney (POA) and state/registry request letters from templates.', 'Լիազորագրերի (POA) և պետական/ռեգիստրի հարցումների նամակների ինքնաշխատ գեներացում ձևանմուշներից։'],
  ['File uploads attached to cases, and an audit log of actions.', 'Գործերին կցվող ֆայլերի վերբեռնում և գործողությունների աուդիտի մատյան։'],
  ['Sheets used today: Cases, Documents, Debts, BankCertificates, CaseFiles, AuditLog.', 'Այսօր օգտագործվող աղյուսակները՝ Cases, Documents, Debts, BankCertificates, CaseFiles, AuditLog։']
].forEach(([e, a]) => children.push(...bullet(e, a)));
children.push(...bp({
  en: 'Gaps relevant to this project: a case has no "stage" field (everything is implicitly preparation); no "case type" field (all assumed bankruptcy); no official court-case number; and nothing about physical paper mail or physical archive storage.',
  hy: 'Սույն նախագծին վերաբերող բացերը. գործը չունի «փուլ» դաշտ (ամեն ինչ ենթադրաբար նախապատրաստում է), չունի «գործի տեսակ» դաշտ (բոլորը ենթադրվում են սնանկություն), չունի պաշտոնական դատական գործի համար, և ոչինչ նախատեսված չէ ֆիզիկական թղթային փոստի կամ ֆիզիկական արխիվի պահպանման համար։'
}));

// ---------- 3. SCOPE ----------
children.push(h1('3. Scope of Work', '3. Աշխատանքի ծավալ'));
children.push(...bp({
  en: 'The contractor shall deliver the following, as a single combined delivery (the firm has chosen one-time delivery rather than phased):',
  hy: 'Կապալառուն պետք է մատուցի ստորև բերվածը՝ որպես մեկ միասնական առաքում (ընկերությունն ընտրել է միանվագ առաքում, ոչ թե փուլային).'
}));
[
  ['Case TYPES — classify every case (bankruptcy, civil litigation, criminal defence, administrative, advisory/contracts, other). The type list must be editable by the firm without a developer.',
   'Գործերի ՏԵՍԱԿՆԵՐ — դասակարգել յուրաքանչյուր գործ (սնանկություն, քաղաքացիական դատավարություն, քրեական պաշտպանություն, վարչական, խորհրդատվություն/պայմանագրեր, այլ)։ Տեսակների ցանկը պետք է խմբագրելի լինի ընկերության կողմից՝ առանց ծրագրավորողի։'],
  ['Case STAGES — Preparation → Process (Trial) → Archive, with an "Add" button on each stage and a "Migrate" button that moves a case from one stage to the next, recording who and when.',
   'Գործերի ՓՈՒԼԵՐ — Նախապատրաստում → Ընթացիկ (Դատական) → Արխիվ, յուրաքանչյուր փուլում «Ավելացնել» կոճակով և «Տեղափոխել» կոճակով, որը գործը տեղափոխում է մեկ փուլից հաջորդ՝ գրանցելով ով և երբ։'],
  ['NUMBERING — one internal number kept for the whole life of the case; a separate official court-case number added when trial begins.',
   'ՀԱՄԱՐԱԿԱԼՈՒՄ — մեկ ներքին համար՝ պահվող գործի ողջ կյանքի ընթացքում, և առանձին պաշտոնական դատական գործի համար՝ ավելացվող դատաքննության սկզբում։'],
  ['Incoming-MAIL sorting aid — a fast lookup so a junior can find the correct internal number and physical location in seconds, mark the paper letter, and file it; with an optional one-click "letter received" log.',
   'Մուտքային ՓՈՍՏԻ դասակարգման օգնական — արագ որոնում, որպեսզի կրտսեր աշխատակիցը վայրկյանների ընթացքում գտնի ճիշտ ներքին համարը և ֆիզիկական տեղը, նշում կատարի թղթային նամակի վրա և տեղադրի այն. ընտրովի «նամակը ստացված է» մեկ-սեղմումով գրանցմամբ։'],
  ['Physical ARCHIVE locator — impose a clean room → shelf → box location scheme; store archival date (timestamp), storage number and location hints; make any archived file findable from both the front-end and the backend sheets.',
   'Ֆիզիկական ԱՐԽԻՎԻ տեղորոշիչ — սահմանել մաքուր սենյակ → դարակ → արկղ տեղադրման սխեմա, պահել արխիվացման ամսաթիվը (ժամանակային դրոշմ), պահեստի համարը և տեղի հուշումներ, դարձնել ցանկացած արխիվային ֆայլ գտանելի և՛ ճակատամասից, և՛ հետնամասի աղյուսակներից։'],
  ['TRIAL tracking — court, instance, judge, hearing dates, procedural deadlines, current status and final outcome.',
   'ԴԱՏԱԿԱՆ հետևում — դատարան, ատյան, դատավոր, դատական նիստերի ամսաթվեր, դատավարական ժամկետներ, ընթացիկ կարգավիճակ և վերջնական արդյունք։'],
  ['REMINDERS — an in-app dashboard list of upcoming hearings/deadlines, automatic email reminders, and Google Calendar integration.',
   'ՀԻՇԵՑՈՒՄՆԵՐ — ծրագրի ներսում առաջիկա նիստերի/ժամկետների ցանկ, ինքնաշխատ էլ. փոստի հիշեցումներ և Google Calendar ինտեգրում։'],
  ['Backend visibility & SEARCH — all new data lives in worksheets/database and is searchable (find files from the back end, see archival date, storage number, location hints).',
   'Հետնամասի տեսանելիություն և ՈՐՈՆՈՒՄ — բոլոր նոր տվյալները պահվում են աղյուսակներում/տվյալների բազայում և որոնելի են (ֆայլերը գտնել նաև հետնամասից՝ տեսնելով արխիվացման ամսաթիվ, պահեստի համար, տեղի հուշումներ)։']
].forEach(([e, a]) => children.push(...bullet(e, a)));

// ---------- 4. CONSTRAINTS ----------
children.push(h1('4. Critical Constraints — Preserve What Exists', '4. Կարևորագույն սահմանափակումներ — պահպանել եղածը'));
children.push(new Paragraph({ spacing: { after: 120 },
  children: [new TextRun({ text: 'These constraints are non-negotiable and are pass/fail at acceptance.  /  Այս սահմանափակումները չեն քննարկվում և ընդունման ժամանակ pass/fail են։', bold: true, color: RED })] }));
[
  ['No loss of functionality. Every feature working today must keep working: intake, bankruptcy questionnaire, Drive folder tree, document/debt/bank-certificate trackers, POA and state/registry letter generation, file uploads, audit log.',
   'Ֆունկցիոնալության կորուստ չի թույլատրվում։ Այսօր աշխատող յուրաքանչյուր գործառույթ պետք է շարունակի աշխատել՝ ընդունում, սնանկության հարցաթերթ, Drive թղթապանակների ծառ, փաստաթղթերի/պարտքերի/բանկային տեղեկանքների հետևում, լիազորագրերի և պետական հարցումների գեներացում, ֆայլերի վերբեռնում, աուդիտի մատյան։'],
  ['No loss of logic or structure. Existing business logic, templates (POA, state requests), folder naming and field meanings must be preserved exactly.',
   'Տրամաբանության կամ կառուցվածքի կորուստ չի թույլատրվում։ Գործող բիզնես-տրամաբանությունը, ձևանմուշները (լիազորագրեր, պետական հարցումներ), թղթապանակների անվանումները և դաշտերի իմաստները պետք է պահպանվեն ճշգրիտ։'],
  ['Full 1:1 data migration. All existing cases, documents, debts, bank certificates, files and audit records must be carried over with no data loss; existing internal numbers and Drive folders must remain valid.',
   'Ամբողջական 1:1 տվյալների միգրացիա։ Բոլոր առկա գործերը, փաստաթղթերը, պարտքերը, բանկային տեղեկանքները, ֆայլերը և աուդիտի գրառումները պետք է փոխանցվեն առանց կորստի, առկա ներքին համարներն ու Drive թղթապանակները պետք է մնան վավեր։'],
  ['Rebuild is permitted but not required. The contractor may re-implement internally (e.g. move to a database) if it improves reliability, but only if every point above is fully satisfied and proven. Given the current scale (under 1,000 cases) extending the present Google Apps Script + Sheets + Drive stack is acceptable and expected to be sufficient.',
   'Վերակառուցումը թույլատրվում է, բայց պարտադիր չէ։ Կապալառուն կարող է ներքին վերաիրականացում անել (օր.՝ անցում տվյալների բազայի), եթե դա բարելավում է հուսալիությունը, բայց միայն եթե վերը նշված յուրաքանչյուր կետը լիովին բավարարված և ապացուցված է։ Ներկայիս ծավալի դեպքում (1000-ից պակաս գործ) գործող Google Apps Script + Sheets + Drive կառուցվածքի ընդլայնումն ընդունելի է և ակնկալվում է բավարար։'],
  ['Armenian-first UI. All user-facing text remains in Armenian; new screens follow the existing visual style.',
   'Հայերեն-առաջնային ինտերֆեյս։ Օգտատիրոջն ուղղված ողջ տեքստը մնում է հայերեն, նոր էկրանները հետևում են առկա ոճին։']
].forEach(([e, a]) => children.push(...num(e, a)));

// ---------- 5. FUNCTIONAL REQUIREMENTS ----------
children.push(h1('5. Functional Requirements', '5. Ֆունկցիոնալ պահանջներ'));

children.push(h2('5.1 Case Types', '5.1 Գործերի տեսակներ'));
[
  ['Every case carries a type. Initial list: Bankruptcy, Civil litigation, Criminal defence, Administrative, Advisory/Contracts, Other.',
   'Յուրաքանչյուր գործ ունի տեսակ։ Սկզբնական ցանկ՝ Սնանկություն, Քաղաքացիական դատավարություն, Քրեական պաշտպանություն, Վարչական, Խորհրդատվություն/Պայմանագրեր, Այլ։'],
  ['Existing cases are all set to "Bankruptcy" during migration.',
   'Առկա բոլոր գործերը միգրացիայի ընթացքում սահմանվում են «Սնանկություն»։'],
  ['Bankruptcy-specific features (questionnaire, POA/state-letter generation) appear only for bankruptcy cases; other types show a general case form. Nothing bankruptcy-specific is removed.',
   'Սնանկությանը հատուկ գործառույթները (հարցաթերթ, լիազորագիր/պետական նամակների գեներացում) հայտնվում են միայն սնանկության գործերի համար, մյուս տեսակները ցույց են տալիս ընդհանուր ձև։ Սնանկությանը հատուկ ոչինչ չի հեռացվում։'],
  ['The type list is editable by the firm (add/rename/disable) without code changes.',
   'Տեսակների ցանկը խմբագրելի է ընկերության կողմից (ավելացնել/վերանվանել/անջատել)՝ առանց կոդի փոփոխման։']
].forEach(([e, a]) => children.push(...bullet(e, a)));

children.push(h2('5.2 Case Stages & Migration', '5.2 Գործերի փուլեր և տեղափոխում'));
[
  ['Three stages: Preparation, Process (Trial), Archive. The home screen has a section/tab per stage with its own "Add case" button and list.',
   'Երեք փուլ՝ Նախապատրաստում, Ընթացիկ (Դատական), Արխիվ։ Գլխավոր էկրանն ունի յուրաքանչյուր փուլի համար բաժին/ներդիր՝ իր «Ավելացնել գործ» կոճակով և ցանկով։'],
  ['A "Migrate" button moves a case forward (Preparation→Process→Archive). The system records the migration timestamp and the user who performed it, and writes it to the audit log.',
   '«Տեղափոխել» կոճակը գործը տանում է առաջ (Նախապատրաստում→Ընթացիկ→Արխիվ)։ Համակարգը գրանցում է տեղափոխման ժամանակը և կատարող օգտատիրոջը և գրառում աուդիտի մատյանում։'],
  ['Migrating to Process prompts for the official court-case number, court, instance and judge (Section 5.6).',
   'Ընթացիկ տեղափոխելիս համակարգը պահանջում է պաշտոնական դատական համար, դատարան, ատյան և դատավոր (բաժին 5.6)։'],
  ['Migrating to Archive prompts for the physical location (room/shelf/box) and sets the archival date automatically (Section 5.5).',
   'Արխիվ տեղափոխելիս համակարգը պահանջում է ֆիզիկական տեղը (սենյակ/դարակ/արկղ) և ինքնաշխատ սահմանում արխիվացման ամսաթիվը (բաժին 5.5)։'],
  ['Backward migration (e.g. archive re-opened) is allowed for authorised users and is also logged.',
   'Հետընթաց տեղափոխումը (օր.՝ արխիվից վերաբացում) թույլատրվում է և նույնպես գրանցվում է մատյանում։'],
  ['The internal number never changes across stages; the Drive folder is preserved (no re-creation, no re-numbering).',
   'Ներքին համարը երբեք չի փոխվում փուլերի միջև, Drive թղթապանակը պահպանվում է (առանց վերստեղծման, առանց վերահամարակալման)։']
].forEach(([e, a]) => children.push(...bullet(e, a)));

children.push(h2('5.3 Numbering', '5.3 Համարակալում'));
[
  ['Internal number: one consecutive number assigned at creation and kept for the entire life of the case, through every stage including archive. It is never reused or renumbered.',
   'Ներքին համար՝ մեկ հաջորդական համար, որը տրվում է ստեղծման պահին և պահվում գործի ողջ կյանքի ընթացքում՝ բոլոր փուլերում, ներառյալ արխիվը։ Այն երբեք չի վերաօգտագործվում կամ վերահամարակալվում։'],
  ['Official court-case number: a separate field, entered when the trial begins; multiple instances (first/appeal/cassation) may each have their own official number.',
   'Պաշտոնական դատական համար՝ առանձին դաշտ, լրացվում է դատաքննության սկզբում, տարբեր ատյանները (առաջին/վերաքննիչ/վճռաբեկ) կարող են ունենալ իրենց համարը։'],
  ['Both numbers are searchable; a letter or file can be located by either.',
   'Երկու համարներն էլ որոնելի են, նամակը կամ ֆայլը կարելի է գտնել ցանկացածով։']
].forEach(([e, a]) => children.push(...bullet(e, a)));

children.push(h2('5.4 Incoming-Mail Sorting Aid', '5.4 Մուտքային փոստի դասակարգման օգնական'));
children.push(...bp({
  en: 'Mail is paper, not electronic, and is sorted manually. The tool exists only to let a junior staff member identify, in seconds and without a senior lawyer, where each letter belongs and mark it for filing.',
  hy: 'Փոստը թղթային է, ոչ էլեկտրոնային, և դասակարգվում է ձեռքով։ Գործիքը գոյություն ունի միայն այն բանի համար, որ կրտսեր աշխատակիցը վայրկյանների ընթացքում և առանց ավագ իրավաբանի որոշի, թե յուրաքանչյուր նամակ ուր է պատկանում, և նշում կատարի տեղադրման համար։'
}));
[
  ['A simple "Incoming mail" screen with one search box. The junior types any of: client name, internal number, official court number, or company/sender name.',
   'Պարզ «Մուտքային փոստ» էկրան՝ մեկ որոնման դաշտով։ Կրտսերը մուտքագրում է հետևյալներից ցանկացածը՝ վստահորդի անուն, ներքին համար, պաշտոնական դատական համար կամ ընկերության/ուղարկողի անուն։'],
  ['Results instantly show: client name, case type, current stage, the internal number to write on the letter, and — if archived — the exact physical location (room/shelf/box).',
   'Արդյունքներն անմիջապես ցույց են տալիս՝ վստահորդի անունը, գործի տեսակը, ընթացիկ փուլը, նամակի վրա գրվող ներքին համարը և, եթե արխիվացված է՝ ճշգրիտ ֆիզիկական տեղը (սենյակ/դարակ/արկղ)։'],
  ['Optional one-click "Letter received" action logs date received, the selected case and the user — building a searchable mail register without slowing the junior down. This step must be skippable.',
   'Ընտրովի «Նամակն ստացված է» մեկ-սեղմումով գործողությունը գրանցում է ստացման ամսաթիվը, ընտրված գործը և օգտատիրոջը՝ ստեղծելով որոնելի փոստի ռեգիստր՝ առանց կրտսերին դանդաղեցնելու։ Այս քայլը պետք է լինի բացթողնելի։'],
  ['Optional: a printable label (internal number + client + location) the junior can stick on the letter or folder.',
   'Ընտրովի՝ տպվող պիտակ (ներքին համար + վստահորդ + տեղ), որը կրտսերը կարող է փակցնել նամակին կամ թղթապանակին։']
].forEach(([e, a]) => children.push(...bullet(e, a)));

children.push(h2('5.5 Physical Archive Locator', '5.5 Ֆիզիկական արխիվի տեղորոշիչ'));
children.push(...bp({
  en: 'There is no fixed physical storage system today. The contractor shall impose a clean, hierarchical location scheme and let the firm relabel storage to match it over time.',
  hy: 'Այսօր ֆիզիկական պահպանման հաստատված համակարգ չկա։ Կապալառուն պետք է սահմանի մաքուր, հիերարխիկ տեղադրման սխեմա և թույլ տա ընկերությանը ժամանակի ընթացքում վերապիտակել պահեստը՝ դրան համապատասխան։'
}));
[
  ['Location model: Room/Site → Shelf/Rack → Box → Case. Each level is a simple editable value; a case stores its full location path.',
   'Տեղադրման մոդել՝ Սենյակ/Տարածք → Դարակ → Արկղ → Գործ։ Յուրաքանչյուր մակարդակ պարզ խմբագրելի արժեք է, գործը պահում է իր ամբողջական տեղադրման ուղին։'],
  ['On archival the system records: archival date/time (timestamp), storage number, location path, and free-text location hints.',
   'Արխիվացման պահին համակարգը գրանցում է՝ արխիվացման ամսաթիվ/ժամ (ժամանակային դրոշմ), պահեստի համար, տեղադրման ուղի և ազատ տեքստով տեղի հուշումներ։'],
  ['An "Archive" view lists archived cases sortable/filterable by internal number, client, type, archival date and location, so a specific file is found without searching shelves blindly.',
   '«Արխիվ» տեսքը ցուցադրում է արխիվացված գործերը՝ դասակարգելի/զտելի ըստ ներքին համարի, վստահորդի, տեսակի, արխիվացման ամսաթվի և տեղի, որպեսզի կոնկրետ ֆայլը գտնվի առանց դարակները կուրորեն փնտրելու։'],
  ['Optional but recommended: a printable box manifest listing all cases inside a given box.',
   'Ընտրովի, բայց խորհուրդ է տրվում՝ տպվող արկղի ցուցակ՝ տվյալ արկղում գտնվող բոլոր գործերով։'],
  ['All of the above is visible in the backend worksheet/table as well as the UI, per the firm’s requirement to find files from the back end.',
   'Վերը նշված ամեն ինչ տեսանելի է թե՛ հետնամասի աղյուսակում, թե՛ ինտերֆեյսում՝ ըստ ընկերության՝ հետնամասից ֆայլեր գտնելու պահանջի։']
].forEach(([e, a]) => children.push(...bullet(e, a)));

children.push(h2('5.6 Trial / Process Tracking', '5.6 Դատական / ընթացիկ հետևում'));
[
  ['Court fields: court name, instance (first / appeal / cassation), assigned judge.',
   'Դատական դաշտեր՝ դատարանի անվանում, ատյան (առաջին / վերաքննիչ / վճռաբեկ), նշանակված դատավոր։'],
  ['Hearing dates: each hearing has date/time and optional note; upcoming hearings appear in the dashboard.',
   'Նիստերի ամսաթվեր՝ յուրաքանչյուր նիստ ունի ամսաթիվ/ժամ և ընտրովի նշում, առաջիկա նիստերը հայտնվում են վահանակում։'],
  ['Procedural deadlines: each deadline has a due date and description; a "due soon" view highlights approaching ones.',
   'Դատավարական ժամկետներ՝ յուրաքանչյուրն ունի վերջնաժամկետ և նկարագրություն, «մոտալուտ» տեսքը ընդգծում է մոտեցողները։'],
  ['Status & outcome: current procedural status, and the final outcome/verdict recorded when the case closes.',
   'Կարգավիճակ և արդյունք՝ ընթացիկ դատավարական կարգավիճակ և վերջնական արդյունք/վճիռ՝ գրանցվող գործի փակման պահին։']
].forEach(([e, a]) => children.push(...bullet(e, a)));

children.push(h2('5.7 Reminders', '5.7 Հիշեցումներ'));
[
  ['In-app dashboard: a panel listing upcoming hearings and deadlines across all cases, ordered by date.',
   'Ծրագրի վահանակ՝ բոլոր գործերի առաջիկա նիստերի և ժամկետների ցանկ՝ ըստ ամսաթվի։'],
  ['Email reminders: automatic email a configurable number of days before each hearing/deadline to the responsible lawyer(s).',
   'Էլ. փոստի հիշեցումներ՝ ինքնաշխատ նամակ յուրաքանչյուր նիստից/ժամկետից կարգավորելի օրեր առաջ՝ պատասխանատու իրավաբան(ներ)ին։'],
  ['Google Calendar integration: hearings and deadlines pushed to a shared firm calendar so they appear on phones automatically; updates/cancellations sync.',
   'Google Calendar ինտեգրում՝ նիստերն ու ժամկետները փոխանցվում են ընկերության ընդհանուր օրացույց, որպեսզի ինքնաշխատ հայտնվեն հեռախոսներում, փոփոխությունները/չեղարկումները համաժամանակացվում են։']
].forEach(([e, a]) => children.push(...bullet(e, a)));

children.push(h2('5.8 Search & Backend Visibility', '5.8 Որոնում և հետնամասի տեսանելիություն'));
[
  ['Global search by client name, internal number, official court number, sender, case type, stage, and location.',
   'Համընդհանուր որոնում՝ ըստ վստահորդի անվան, ներքին համարի, պաշտոնական դատական համարի, ուղարկողի, գործի տեսակի, փուլի և տեղի։'],
  ['Every new field is stored in the worksheet/database and readable directly from the backend, not only the UI.',
   'Յուրաքանչյուր նոր դաշտ պահվում է աղյուսակում/տվյալների բազայում և ընթեռնելի է ուղղակիորեն հետնամասից, ոչ միայն ինտերֆեյսից։']
].forEach(([e, a]) => children.push(...bullet(e, a)));

children.push(h2('5.9 Users & Access', '5.9 Օգտատերեր և մուտք'));
[
  ['The firm has chosen equal access: all staff can perform all actions.',
   'Ընկերությունն ընտրել է հավասար մուտք՝ բոլոր աշխատակիցները կարող են կատարել բոլոր գործողությունները։'],
  ['Therefore accountability is preserved through the audit log (who/what/when) for create, edit, migrate, archive and delete, and through the existing typed confirmation before any case deletion.',
   'Հետևաբար հաշվետվողականությունն ապահովվում է աուդիտի մատյանով (ով/ինչ/երբ)՝ ստեղծման, խմբագրման, տեղափոխման, արխիվացման և ջնջման համար, ինչպես նաև գործը ջնջելուց առաջ առկա հաստատման մեխանիզմով։'],
  ['RECOMMENDATION (optional): keep the option to later restrict deletion/migration to senior staff, so the design should not hard-code equal access in a way that blocks future roles.',
   'ԱՌԱՋԱՐԿ (ընտրովի)՝ պահել հնարավորությունը հետագայում ջնջումը/տեղափոխումը սահմանափակել ավագ աշխատակիցներով, ուստի դիզայնը չպետք է կոշտ ամրագրի հավասար մուտքն այնպես, որ արգելափակի ապագա դերերը։']
].forEach(([e, a]) => children.push(...bullet(e, a)));

// ---------- 6. DATA MODEL ----------
children.push(new Paragraph({ children: [new PageBreak()] }));
children.push(h1('6. Data Model Changes', '6. Տվյալների մոդելի փոփոխություններ'));
children.push(h2('6.1 New fields on the Cases table', '6.1 Նոր դաշտեր Cases աղյուսակում'));
children.push(dictTable([
  ['caseType', 'Case type (bankruptcy / civil / criminal / administrative / advisory / other).', 'Գործի տեսակ (սնանկություն / քաղաքացիական / քրեական / վարչական / խորհրդատվ. / այլ)։'],
  ['stage', 'Current stage: preparation / process / archive.', 'Ընթացիկ փուլ՝ նախապատրաստում / ընթացիկ / արխիվ։'],
  ['stageChangedAt', 'Timestamp of the last stage change.', 'Վերջին փուլափոխման ժամանակային դրոշմ։'],
  ['stageChangedBy', 'User who performed the last stage change.', 'Վերջին փուլափոխումը կատարած օգտատեր։'],
  ['officialCaseNumber', 'Official court-case number, added at trial.', 'Պաշտոնական դատական համար՝ ավելացվող դատաքննության ժամանակ։'],
  ['court', 'Court name and instance.', 'Դատարանի անվանում և ատյան։'],
  ['judge', 'Assigned judge.', 'Նշանակված դատավոր։'],
  ['caseStatus', 'Current procedural status / final outcome.', 'Ընթացիկ դատավարական կարգավիճակ / վերջնական արդյունք։'],
  ['archivedAt', 'Archival date/time (timestamp).', 'Արխիվացման ամսաթիվ/ժամ (ժամանակային դրոշմ)։'],
  ['storageNumber', 'Storage / box number.', 'Պահեստի / արկղի համար։'],
  ['storageRoom', 'Room or site.', 'Սենյակ կամ տարածք։'],
  ['storageShelf', 'Shelf / rack.', 'Դարակ։'],
  ['storageBox', 'Box.', 'Արկղ։'],
  ['locationHints', 'Free-text location hints.', 'Ազատ տեքստով տեղի հուշումներ։']
]));
children.push(h2('6.2 New tables', '6.2 Նոր աղյուսակներ'));
children.push(dictTable([
  ['IncomingMail', 'Mail register: id, dateReceived, sender, caseId, internalNumber, loggedBy, note.', 'Փոստի ռեգիստր՝ id, ստացման ամսաթիվ, ուղարկող, caseId, ներքին համար, գրանցող, նշում։'],
  ['Hearings', 'Hearing schedule: id, caseId, dateTime, note, reminderDaysBefore.', 'Նիստերի ժամանակացույց՝ id, caseId, ամսաթիվ/ժամ, նշում, հիշեցում քանի օր առաջ։'],
  ['Deadlines', 'Procedural deadlines: id, caseId, dueDate, description, reminderDaysBefore, done.', 'Դատավարական ժամկետներ՝ id, caseId, վերջնաժամկետ, նկարագրություն, հիշեցում, կատարված։'],
  ['CaseTypes', 'Editable list of case types (so the firm can manage them).', 'Գործերի տեսակների խմբագրելի ցանկ (ընկերության կառավարման համար)։']
]));
children.push(...bp({
  en: 'Existing tables (Cases, Documents, Debts, BankCertificates, CaseFiles, AuditLog) keep their current columns; new columns are appended, never reordered or removed, to protect the existing migration logic.',
  hy: 'Առկա աղյուսակները (Cases, Documents, Debts, BankCertificates, CaseFiles, AuditLog) պահում են իրենց ընթացիկ սյունակները, նոր սյունակներն ավելացվում են վերջում՝ երբեք չվերադասավորվելով կամ չհեռացվելով՝ առկա միգրացիայի տրամաբանությունը պաշտպանելու համար։'
}));

// ---------- 7. NON-FUNCTIONAL ----------
children.push(h1('7. Non-Functional Requirements', '7. Ոչ-ֆունկցիոնալ պահանջներ'));
[
  ['Security & confidentiality: the system stores sensitive client and legal data. Access must be restricted to authorised firm accounts; the web app must not be publicly accessible without authentication. Data stays within the firm’s Google Workspace / chosen infrastructure.',
   'Անվտանգություն և գաղտնիություն. համակարգը պահում է զգայուն վստահորդական և իրավական տվյալներ։ Մուտքը պետք է սահմանափակվի ընկերության լիազորված հաշիվներով, վեբ-հավելվածը չպետք է հանրությանը հասանելի լինի առանց նույնականացման։ Տվյալները մնում են ընկերության Google Workspace / ընտրված ենթակառուցվածքում։'],
  ['Performance: search and screen loads respond within ~2 seconds at the expected scale (under 1,000 cases).',
   'Արագագործություն. որոնումը և էկրանների բեռնումն արձագանքում են մոտ 2 վայրկյանում՝ սպասվող ծավալի դեպքում (1000-ից պակաս գործ)։'],
  ['Backups: automatic, scheduled backups of all worksheets/database and a documented restore procedure.',
   'Կրկնօրինակումներ. բոլոր աղյուսակների/բազայի ինքնաշխատ, ժամանակացույցով կրկնօրինակումներ և վերականգնման փաստաթղթավորված ընթացակարգ։'],
  ['Reliability: concurrent edits must not corrupt data (the current system already uses locking; this must be preserved).',
   'Հուսալիություն. զուգահեռ խմբագրումները չպետք է վնասեն տվյալները (ընթացիկ համակարգն արդեն օգտագործում է կողպում, սա պետք է պահպանվի)։'],
  ['Maintainability: documented code, configuration values kept out of hard-coding where the firm needs to change them (case types, reminder days).',
   'Սպասարկելիություն. փաստաթղթավորված կոդ, ընկերության կողմից փոփոխման ենթակա արժեքները (գործերի տեսակներ, հիշեցման օրեր) կոշտ ամրագրված չլինեն։'],
  ['Browser support: current desktop browsers (Chrome, Edge); usable on a tablet for archive/mail work.',
   'Բրաուզերների աջակցություն. ընթացիկ desktop բրաուզերներ (Chrome, Edge), օգտագործելի պլանշետի վրա՝ արխիվ/փոստ աշխատանքի համար։']
].forEach(([e, a]) => children.push(...bullet(e, a)));

// ---------- 8. MIGRATION ----------
children.push(h1('8. Data Migration', '8. Տվյալների միգրացիա'));
[
  ['Migrate all existing records 1:1 with no loss; verify counts before/after.',
   'Միգրացնել առկա բոլոր գրառումները 1:1՝ առանց կորստի, ստուգել քանակները նախքան և հետո։'],
  ['Set all existing cases to type = Bankruptcy and stage = Preparation (unless the firm marks specific ones otherwise during a review step).',
   'Սահմանել առկա բոլոր գործերը՝ տեսակ = Սնանկություն և փուլ = Նախապատրաստում (եթե ընկերությունը ստուգման քայլում որոշները այլ կերպ չնշի)։'],
  ['Preserve every existing internal number and Drive folder; no renumbering, no folder re-creation.',
   'Պահպանել առկա յուրաքանչյուր ներքին համար և Drive թղթապանակ, առանց վերահամարակալման, առանց թղթապանակների վերստեղծման։'],
  ['Provide a one-time assisted step for the firm to bulk-assign archive locations to already-archived cases.',
   'Տրամադրել միանվագ օժանդակ քայլ՝ ընկերության կողմից արդեն արխիվացված գործերին զանգվածաբար արխիվային տեղ նշանակելու համար։'],
  ['Keep a backup of the original data before migration (the current system already auto-backs up sheets before schema changes).',
   'Միգրացիայից առաջ պահել սկզբնական տվյալների կրկնօրինակ (ընթացիկ համակարգն արդեն ինքնաշխատ կրկնօրինակում է աղյուսակները սխեմայի փոփոխությունից առաջ)։']
].forEach(([e, a]) => children.push(...bullet(e, a)));

// ---------- 9. ACCEPTANCE ----------
children.push(h1('9. Acceptance Criteria', '9. Ընդունման չափանիշներ'));
children.push(...bp({
  en: 'The work is accepted only when ALL of the following are demonstrated on real migrated data:',
  hy: 'Աշխատանքն ընդունվում է միայն այն դեպքում, երբ հետևյալ ԲՈԼՈՐԸ ցուցադրվեն իրական միգրացված տվյալների վրա.'
}));
[
  ['Every pre-existing feature still works (Section 4 checklist passes 100%).',
   'Նախկինում առկա յուրաքանչյուր գործառույթ դեռ աշխատում է (բաժին 4-ի ստուգաթերթը անցնում է 100%)։'],
  ['Record counts match before and after migration; no data lost.',
   'Գրառումների քանակները համընկնում են միգրացիայից առաջ և հետո, տվյալներ չեն կորել։'],
  ['A case can be created, migrated Preparation→Process→Archive, and each migration is logged with user + timestamp.',
   'Գործը կարելի է ստեղծել, տեղափոխել Նախապատրաստում→Ընթացիկ→Արխիվ, և յուրաքանչյուր տեղափոխում գրանցվում է օգտատիրոջ + ժամանակային դրոշմով։'],
  ['A junior can find any case and its physical location from the mail screen in under ~10 seconds, by name or any number.',
   'Կրտսերը կարող է փոստի էկրանից ~10 վայրկյանից պակասում գտնել ցանկացած գործ և դրա ֆիզիկական տեղը՝ ըստ անվան կամ ցանկացած համարի։'],
  ['An archived file can be located by storage room/shelf/box from both the UI and the backend sheet.',
   'Արխիվային ֆայլը կարելի է գտնել ըստ սենյակ/դարակ/արկղի՝ թե՛ ինտերֆեյսից, թե՛ հետնամասի աղյուսակից։'],
  ['Hearings/deadlines appear in the dashboard, trigger email reminders, and sync to Google Calendar.',
   'Նիստերը/ժամկետները հայտնվում են վահանակում, գործարկում են էլ. փոստի հիշեցումներ և համաժամանակացվում Google Calendar-ի հետ։'],
  ['All new data is visible and searchable directly in the backend worksheets/database.',
   'Բոլոր նոր տվյալները տեսանելի և որոնելի են ուղղակիորեն հետնամասի աղյուսակներում/բազայում։'],
  ['Source code, deployment steps, backup/restore procedure and a short Armenian user guide are delivered.',
   'Մատուցվում են սկզբնական կոդը, տեղադրման քայլերը, կրկնօրինակման/վերականգնման ընթացակարգը և կարճ հայերեն օգտատիրոջ ուղեցույց։']
].forEach(([e, a]) => children.push(...num(e, a)));

// ---------- 10. DELIVERABLES ----------
children.push(h1('10. Deliverables', '10. Մատուցվելիքներ'));
[
  ['Working expanded system deployed to the firm’s environment.', 'Աշխատող ընդլայնված համակարգ՝ տեղադրված ընկերության միջավայրում։'],
  ['Full source code and configuration.', 'Ամբողջական սկզբնական կոդ և կարգավորումներ։'],
  ['Migration scripts and a migration report (counts before/after).', 'Միգրացիայի սկրիպտներ և միգրացիայի հաշվետվություն (քանակներ նախքան/հետո)։'],
  ['Deployment + backup/restore documentation.', 'Տեղադրման + կրկնօրինակման/վերականգնման փաստաթղթեր։'],
  ['Short Armenian user guide and a hand-over/training session.', 'Կարճ հայերեն օգտատիրոջ ուղեցույց և հանձնման/ուսուցման հանդիպում։'],
  ['Warranty/support period for defect fixes (term to be set in the contract).', 'Երաշխիքային/աջակցության ժամանակահատված թերությունների ուղղման համար (ժամկետը՝ պայմանագրով)։']
].forEach(([e, a]) => children.push(...bullet(e, a)));

// ---------- 11. ASSUMPTIONS ----------
children.push(h1('11. Assumptions & Items to Confirm with Bidder', '11. Ենթադրություններ և բիդերի հետ ճշտելու կետեր'));
[
  ['Scale assumed under 1,000 active+archived cases; if higher, the bidder must propose a database-backed design.',
   'Ենթադրվում է 1000-ից պակաս ակտիվ+արխիվ գործ, ավելիի դեպքում բիդերը պետք է առաջարկի տվյալների բազայի վրա հիմնված դիզայն։'],
  ['Email sending uses the firm’s Google Workspace account; calendar is a shared firm Google Calendar.',
   'Էլ. փոստն ուղարկվում է ընկերության Google Workspace հաշվից, օրացույցը ընկերության ընդհանուր Google Calendar է։'],
  ['Exact archive room/shelf/box labelling will be defined jointly during a short discovery step.',
   'Արխիվի սենյակ/դարակ/արկղ պիտակավորման ճշգրիտ սխեման կսահմանվի համատեղ՝ կարճ discovery քայլում։'],
  ['Final list of case types and reminder lead-times to be confirmed before development.',
   'Գործերի տեսակների վերջնական ցանկը և հիշեցումների ժամկետները ճշտվում են մինչև մշակումը։'],
  ['Whether to keep equal access or introduce roles later is the firm’s decision; design must not block future roles.',
   'Հավասար մուտք պահելը թե հետագայում դերեր ներմուծելը ընկերության որոշումն է, դիզայնը չպետք է արգելափակի ապագա դերերը։']
].forEach(([e, a]) => children.push(...bullet(e, a)));

// ---------- FOOTER NOTE ----------
children.push(new Paragraph({ spacing: { before: 360 },
  border: { top: { style: BorderStyle.SINGLE, size: 6, color: GOLD, space: 6 } },
  children: [new TextRun({ text: 'End of specification.  /  Առաջադրանքի ավարտ։', italics: true, color: HY })] }));

// ---------- DOCUMENT ----------
const doc = new Document({
  styles: {
    default: { document: { run: { font: 'Arial', size: 21 } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 30, bold: true, font: 'Arial', color: NAVY },
        paragraph: { spacing: { before: 280, after: 140 }, outlineLevel: 0 } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 24, bold: true, font: 'Arial', color: NAVY },
        paragraph: { spacing: { before: 200, after: 80 }, outlineLevel: 1 } }
    ]
  },
  numbering: { config: [
    { reference: 'b', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 540, hanging: 280 } } } }] },
    { reference: 'n', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 540, hanging: 280 } } } }] }
  ] },
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
    headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT,
      border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC', space: 4 } },
      children: [new TextRun({ text: 'IKV — Technical Specification / Տեխնիկական առաջադրանք', size: 16, color: HY })] })] }) },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: 'Page ', size: 16, color: HY }),
        new TextRun({ children: [PageNumber.CURRENT], size: 16, color: HY }),
        new TextRun({ text: '   •   Confidential / Գաղտնի', size: 16, color: HY })] })] }) },
    children
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync(path.join(__dirname, 'IKV_Technical_Specification.docx'), buffer);
  console.log('Written IKV_Technical_Specification.docx');
});
