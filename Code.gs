// ================================================================
// ԱՅ ՔԻ ՎԻ — Bankruptcy Case Management — Apps Script Backend
// ================================================================
// SETUP:
// 1. Go to script.google.com → New project → paste this code
// 2. Deploy as Web App (Execute as: Me, access restricted to the firm)
// 3. Copy the Web App URL into ikv_bankruptcy.html CONFIG.SCRIPT_URL
// ================================================================

const SPREADSHEET_ID = '1EdjhuXFFzNjxtWVPrcxueKJ5frS4qkxZnb7QqttQBxs'
const ROOT_FOLDER_NAME = 'IKV Bankruptcy Cases'
const MAX_UPLOAD_BYTES = 8 * 1024 * 1024
const REGISTER_TEMPLATE_DOC_ID = '1CfXZsHPy25vPk5i-KcxFvoggtefjYexQhrvhtptUtVE'
const PUBLIC_ACCESS_ENABLED = true
const DEFAULT_REMINDER_DAYS = 3
const DEFAULT_CASE_TYPES = [
  ['bankruptcy', 'Սնանկություն', 'YES', '10'],
  ['civil', 'Քաղաքացիական դատավարություն', 'YES', '20'],
  ['criminal', 'Քրեական պաշտպանություն', 'YES', '30'],
  ['administrative', 'Վարչական', 'YES', '40'],
  ['advisory', 'Խորհրդատվություն / Պայմանագրեր', 'YES', '50'],
  ['other', 'Այլ', 'YES', '60']
]

const REPRESENTATIVES = {
  hovhannes: {
    key: 'hovhannes',
    label: 'Հովհաննես Հարությունյան',
    fullNameFrom: 'Հովհաննես Գրիգորի Հարությունյանից',
    identityType: 'նույնականացման քարտ',
    identityNumber: '009987561',
    identityIssueDate: '20.02.2018',
    identityIssuer: '011',
    licenseNumber: '2332',
    notificationAddress: 'ք. Երևան, Այգեստան 7-րդ փողոց, 1-ին շենք, բն. 30',
    phone: '099-648-655',
    signatureName: 'Հովհաննես Հարությունյան'
  },
  karine: {
    key: 'karine',
    label: 'Կարինե Ավետիսյան',
    fullNameFrom: 'Կարինե Նաիրիի Ավետիսյանից',
    identityType: 'անձնագիր',
    identityNumber: 'AR0610106',
    identityIssueDate: '08.01.2018',
    identityIssuer: '001',
    licenseNumber: '',
    notificationAddress: 'ք. Երևան, Այգեստան 7-րդ փողոց, 1-ին շենք, բն. 30',
    phone: '041-648-656',
    signatureName: 'Կարինե Ավետիսյան'
  }
}
const REPRESENTATIVE = REPRESENTATIVES.hovhannes

/**
 * Run this function once from the Apps Script editor.
 * It opens Google's authorization dialog for Docs and Drive access.
 */
function authorizeDocumentGeneration() {
  const document = DocumentApp.create('IKV authorization test')
  const file = DriveApp.getFileById(document.getId())
  document.getBody().setText('Authorization completed successfully.')
  document.saveAndClose()
  file.setTrashed(true)
  return 'Google Docs and Drive authorization completed.'
}

const HEADERS = {
  Cases: ['id','applicantName','firstName','lastName','middleName','birthDate','passport','passportDate','passportBy','regAddr','notifAddr','psn','isMarried','spouseName','filingDate','employed','salaryAbove','banks','debts','spouseBanks','spouseDebts','lawyerApproved','lawyerApprovedBy','lawyerApprovedAt','createdAt','createdBy','spouseFirstName','spouseLastName','spouseMiddleName','spouseBirthDate','spousePassport','spousePassportDate','spousePassportBy','spousePsn','spouseRegAddr','spouseNotifAddr','assessmentJson','driveFolderId','driveFolderUrl','additionalIdentity','spouseAdditionalIdentity','noPassport','idCard','idCardDate','idCardBy','spouseNoPassport','spouseIdCard','spouseIdCardDate','spouseIdCardBy','noIdCard','spouseNoIdCard','internalNumber','caseType','stage','stageChangedAt','stageChangedBy','officialCaseNumber','court','courtInstance','judge','caseStatus','finalOutcome','archivedAt','storageNumber','storageRoom','storageShelf','storageBox','locationHints','responsibleEmails','updatedAt','updatedBy','clientEntityType','companyRegistrationNumber','companyTaxId'],
  Documents: ['id','caseId','typeId','subject','status','issueDate','expiryDate','appliedAt','updatedAt'],
  Debts: ['id','caseId','subject','creditor','contractNumber','contractDate','currency','principal','interest','penalty','totalAmount','dueDate','claimBasis','collateral','enforcementInfo','notes','createdAt','updatedAt'],
  BankCertificates: ['id','caseId','subject','bank','status','result','accountInfo','balance','currency','issueDate','expiryDate','appliedAt','notes','updatedAt'],
  CaseFiles: ['id','caseId','documentId','typeId','subject','fileName','mimeType','fileId','fileUrl','folderId','uploadedAt'],
  AuditLog: ['id','timestamp','user','action','entity','entityId','detail'],
  IncomingMail: ['id','dateReceived','sender','caseId','internalNumber','loggedBy','note'],
  Hearings: ['id','caseId','dateTime','note','reminderDaysBefore','calendarEventId','responsibleEmails','updatedAt','reminderSentAt'],
  Deadlines: ['id','caseId','dueDate','description','reminderDaysBefore','done','calendarEventId','responsibleEmails','updatedAt','reminderSentAt'],
  CaseTypes: ['key','label','enabled','sortOrder','updatedAt'],
  Settings: ['key','value','description','updatedAt']
}

const DOCUMENT_FOLDERS = {
  passport: '01 Personal Documents',
  psn: '01 Personal Documents',
  marriage: '01 Personal Documents',
  notarial_poa: '02 Legal Authorizations',
  proc_poa: '02 Legal Authorizations',
  credit_report: '03 Debts',
  loan_agr: '03 Debts',
  debt_stmt: '03 Debts',
  debt_account_statement: '03 Debts',
  debt_loan_agreement: '03 Debts',
  depository: '04 Government Certificates',
  cadastre: '04 Government Certificates',
  tax: '04 Government Certificates',
  tax_personal_account: '04 Government Certificates',
  social: '04 Government Certificates',
  territorial: '04 Government Certificates',
  minfin: '04 Government Certificates',
  mineco: '04 Government Certificates',
  mia: '04 Government Certificates',
  traffic: '04 Government Certificates',
  enforcement_service: '04 Government Certificates',
  rescue_service: '04 Government Certificates',
  civil_status: '04 Government Certificates',
  state_reg: '04 Government Certificates',
  banks_resp: '05 Bank Certificates',
  court_docs: '06 Court Documents',
  petition: '06 Court Documents',
  cred_service: '06 Court Documents',
  other_document: '08 Other Documents'
}

const CASE_FOLDER_NAMES = [
  '01 Personal Documents',
  '02 Legal Authorizations',
  '03 Debts',
  '04 Government Certificates',
  '05 Bank Certificates',
  '06 Court Documents',
  '07 Generated Documents',
  '08 Other Documents'
]

const STATE_REQUEST_TEMPLATES = {
  minfin: {
    typeId: 'minfin',
    title: ['ՀՀ ՖԻՆԱՆՍՆԵՐԻ ՆԱԽԱՐԱՐՈՒԹՅՈՒՆ'],
    heading: 'ԴԻՄՈՒՄ',
    paragraphs: [
      'Ղեկավարվելով «Տեղեկատվության ազատության մասին» ՀՀ օրենքի 6-րդ և 9-րդ, «Վարչարարության հիմունքների և վարչական վարույթի մասին» ՀՀ օրենքի 30-րդ, 36-րդ, 39-րդ, 40-րդ և «Փաստաբանության մասին» ՀՀ օրենքի 18-րդ հոդվածներով՝ խնդրում եմ տրամադրել տեղեկատվություն, այն մասին, թե արդյոք {{client.fullNameGenitive}} {{client.requestBodyIdentity}} անվամբ սույն հարցումը ներկայացնելու պահին և դրան նախորդող հինգ տարվա ընթացքում հանդիսացել է ՀՀ ֆինանսների նախարարության վարկառու, ինչպես նաև գրավատու պարտապան և/կամ ստանձնել է քաղաքացիաիրավական հարաբերություններից բխող այլ պարտավորություններ:'
    ]
  },
  rescue_service: {
    typeId: 'rescue_service',
    title: ['ՀՀ ՆԵՐՔԻՆ ԳՈՐԾԵՐԻ ՆԱԽԱՐԱՐՈՒԹՅԱՆ', 'ՓՐԿԱՐԱՐ ԾԱՌԱՅՈՒԹՅՈՒՆ'],
    heading: 'ՀԱՐՑՈՒՄ',
    paragraphs: [
      'Հարգելի գործընկերներ,',
      'Ղեկավարվելով «Տեղեկատվության ազատության մասին» ՀՀ օրենքի 6-րդ և 9-րդ, «Վարչարարության հիմունքների և վարչական վարույթի մասին» ՀՀ օրենքի 30-րդ, 36-րդ, 39-րդ, 40-րդ և «Փաստաբանության մասին» ՀՀ օրենքի 18-րդ հոդվածներով՝ խնդրում եմ տրամադրել տեղեկատվություն, այն մասին, թե ՀՀ ներքին գործերի նախարարության փրկարար ծառայության տվյալների շտեմարանում սույն հարցումը ներկայացնելու պահի դրությամբ և դրան նախորդող հինգ տարվա ընթացքում արդյոք առկա են {{client.fullNameGenitive}} {{client.requestBodyIdentity}} անվամբ գրանցված օբյեկտներ։'
    ]
  },
  depository: {
    typeId: 'depository',
    title: ['«ՀԱՅԱՍՏԱՆԻ ԿԵՆՏՐՈՆԱԿԱՆ ԴԵՊՈԶԻՏԱՐԻԱ»', 'ԲԲ ԸՆԿԵՐՈՒԹՅՈՒՆ'],
    heading: 'ՀԱՐՑՈՒՄ',
    paragraphs: [
      'Հարգելի գործընկերներ,',
      'Ղեկավարվելով «Տեղեկատվության ազատության մասին» ՀՀ օրենքի 6-րդ և 9-րդ, «Վարչարարության հիմունքների և վարչական վարույթի մասին» ՀՀ օրենքի 30-րդ, 36-րդ, 39-րդ, 40-րդ և «Փաստաբանության մասին» ՀՀ օրենքի 18-րդ հոդվածներով՝ խնդրում եմ տրամադրել տեղեկատվություն, այն մասին, թե «Հայաստանի կենտրոնական դեպոզիտարիա» ԲԲԸ-ի տվյալների շտեմարանում սույն հարցումը ներկայացնելու պահի դրությամբ և դրան նախորդող հինգ տարվա ընթացքում արդյոք գրանցված են {{client.fullNameGenitive}} {{client.requestBodyIdentity}} գույքի (արժեթղթերի), այլ իրավունքների առկայության/բացակայության վերաբերյալ տեղեկություններ, այդ թվում՝ ընկերության անվանումը, արժեթղթերի քանակը, անվանական արժեքը, ինչպես նաև դրանց նկատմամբ որևէ սահմանափակում կիրառված է, թե ոչ:'
    ],
    attachments: ['Անձը հաստատող փաստաթղթի սկանը', 'Ինձ տրված լիազորագրի սկանը']
  },
  mineco: {
    typeId: 'mineco',
    title: ['ՀՀ ԷԿՈՆՈՄԻԿԱՅԻ ՆԱԽԱՐԱՐՈՒԹՅՈՒՆ'],
    heading: 'ՀԱՐՑՈՒՄ',
    paragraphs: [
      'Հարգելի գործընկերներ,',
      'Ղեկավարվելով «Տեղեկատվության ազատության մասին» ՀՀ օրենքի 6-րդ և 9-րդ, «Վարչարարության հիմունքների և վարչական վարույթի մասին» ՀՀ օրենքի 30-րդ, 36-րդ, 39-րդ, 40-րդ և «Փաստաբանության մասին» ՀՀ օրենքի 18-րդ հոդվածներով՝ խնդրում եմ տրամադրել տեղեկատվություն, այն մասին, թե ՀՀ էկոնոմիկայի նախարարության տվյալների շտեմարանում սույն հարցումը ներկայացնելու պահի դրությամբ և դրան նախորդող հինգ տարվա ընթացքում արդյոք առկա են {{client.fullNameGenitive}} {{client.requestBodyIdentity}} անվամբ գույքի, այլ իրավունքների առկայության/բացակայության վերաբերյալ տեղեկություններ:'
    ]
  },
  cadastre: {
    typeId: 'cadastre',
    title: ['ՀՀ ԿԱԴԱՍՏՐԻ ԿՈՄԻՏԵ'],
    heading: 'ՀԱՐՑՈՒՄ',
    paragraphs: [
      'Հարգելի գործընկերներ,',
      'Ղեկավարվելով «Տեղեկատվության ազատության մասին» ՀՀ օրենքի 6-րդ և 9-րդ, «Վարչարարության հիմունքների և վարչական վարույթի մասին» ՀՀ օրենքի 30-րդ, 36-րդ, 39-րդ, 40-րդ և «Փաստաբանության մասին» ՀՀ օրենքի 18-րդ հոդվածներով՝ խնդրում եմ տրամադրել տեղեկատվություն, այն մասին, թե ՀՀ կադաստրի կոմիտեի տվյալների շտեմարանում սույն հարցումը ներկայացնելու պահի դրությամբ և դրան նախորդող հինգ տարվա ընթացքում արդյոք գրանցված են {{client.fullNameGenitive}} {{client.requestBodyIdentity}} անվամբ սեփականության, ինչպես նաև այլ գույքային իրավունքներ՝ կցելով դրանց պետական գրանցման համար հիմք հանդիսացած փաստաթղթերի վավերացված պատճենները:'
    ],
    attachments: ['Անձը հաստատող փաստաթղթի սկանը', 'Ինձ տրված լիազորագրի սկանը']
  },
  tax: {
    typeId: 'tax',
    title: ['ՀՀ ՊԵՏԱԿԱՆ ԵԿԱՄՈՒՏՆԵՐԻ ԿՈՄԻՏԵ'],
    heading: 'ՀԱՐՑՈՒՄ',
    paragraphs: [
      'Հարգելի գործընկերներ,',
      'Ղեկավարվելով «Տեղեկատվության ազատության մասին» ՀՀ օրենքի 6-րդ և 9-րդ, «Վարչարարության հիմունքների և վարչական վարույթի մասին» ՀՀ օրենքի 30-րդ, 36-րդ, 39-րդ, 40-րդ և «Փաստաբանության մասին» ՀՀ օրենքի 18-րդ հոդվածներով՝ խնդրում եմ տրամադրել տեղեկատվություն, այն մասին, թե ՀՀ պետական եկամուտների կոմիտեում մաքսային հսկողության ներքո գտնվող գույքի վերաբերյալ մաքսային հայտարարագրերի լրացման ավտոմատացված համակարգի տվյալների շտեմարանում սույն հարցումը ներկայացնելու պահի դրությամբ և դրան նախորդող հինգ տարվա ընթացքում արդյոք առկա են {{client.fullNameGenitive}} {{client.requestBodyIdentity}} անվամբ տվյալներ:'
    ]
  },
  tax_personal_account: {
    typeId: 'tax_personal_account',
    title: ['ՀՀ ՊԵՏԱԿԱՆ ԵԿԱՄՈՒՏՆԵՐԻ ԿՈՄԻՏԵ'],
    heading: 'ՀԱՐՑՈՒՄ',
    paragraphs: [
      'Հարգելի գործընկերներ,',
      'Ղեկավարվելով «Տեղեկատվության ազատության մասին» ՀՀ օրենքի 6-րդ և 9-րդ, «Վարչարարության հիմունքների և վարչական վարույթի մասին» ՀՀ օրենքի 30-րդ, 36-րդ, 39-րդ, 40-րդ և «Փաստաբանության մասին» ՀՀ օրենքի 18-րդ հոդվածներով՝ խնդրում եմ տրամադրել {{client.fullNameGenitive}} {{client.requestBodyIdentity}} անհատական հաշվի քաղվածքը և առկա հարկային պարտավորությունների վերաբերյալ տեղեկատվությունը:'
    ]
  },
  territorial: {
    typeId: 'territorial',
    title: ['ՀՀ ՏԱՐԱԾՔԱՅԻՆ ԿԱՌԱՎԱՐՄԱՆ ԵՎ', 'ԵՆԹԱԿԱՌՈՒՑՎԱԾՔՆԵՐԻ ՆԱԽԱՐԱՐՈՒԹՅՈՒՆ'],
    heading: 'ՀԱՐՑՈՒՄ',
    paragraphs: [
      'Հարգելի գործընկերներ,',
      'Ղեկավարվելով «Տեղեկատվության ազատության մասին» ՀՀ օրենքի 6-րդ և 9-րդ, «Վարչարարության հիմունքների և վարչական վարույթի մասին» ՀՀ օրենքի 30-րդ, 36-րդ, 39-րդ, 40-րդ և «Փաստաբանության մասին» ՀՀ օրենքի 18-րդ հոդվածներով՝ խնդրում եմ տրամադրել տեղեկատվություն, այն մասին, թե ՀՀ տարածքային կառավարման և ենթակառուցվածքների նախարարության տվյալների շտեմարանում սույն հարցումը ներկայացնելու պահի դրությամբ և դրան նախորդող հինգ տարվա ընթացքում արդյոք առկա են {{client.fullNameGenitive}} {{client.requestBodyIdentity}} անվամբ սեփականության, ինչպես նաև այլ գույքային իրավունքներ՝ կցելով դրանց պետական գրանցման համար հիմք հանդիսացած փաստաթղթերի պատճենները:'
    ]
  },
  social: {
    typeId: 'social',
    title: ['ՀՀ ՄԻԱՍՆԱԿԱՆ ՍՈՑԻԱԼԱԿԱՆ ԾԱՌԱՅՈՒԹՅՈՒՆ'],
    heading: 'ՀԱՐՑՈՒՄ',
    paragraphs: [
      'Հարգելի գործընկերներ,',
      'Ղեկավարվելով «Տեղեկատվության ազատության մասին» ՀՀ օրենքի 6-րդ և 9-րդ, «Վարչարարության հիմունքների և վարչական վարույթի մասին» ՀՀ օրենքի 30-րդ, 36-րդ, 39-րդ, 40-րդ և «Փաստաբանության մասին» ՀՀ օրենքի 18-րդ հոդվածներով՝ խնդրում եմ տրամադրել տեղեկատվություն, այն մասին, թե ՀՀ աշխատանքի և սոցիալական հարցերի նախարարության միասնական սոցիալական ծառայության պետական կենսաթոշակային համակարգի տվյալների շտեմարանում սույն հարցումը ներկայացնելու պահի դրությամբ և դրան նախորդող հինգ տարվա ընթացքում արդյոք առկա են {{client.fullNameGenitive}} {{client.requestBodyIdentity}} անվամբ գույքի, այլ իրավունքների առկայության/բացակայության վերաբերյալ տեղեկություններ:'
    ]
  },
  traffic: {
    typeId: 'traffic',
    title: ['ՀՀ ՆԵՐՔԻՆ ԳՈՐԾԵՐԻ ՆԱԽԱՐԱՐՈՒԹՅԱՆ', 'ՈՍՏԻԿԱՆՈՒԹՅԱՆ «ՃԱՆԱՊԱՐՀԱՅԻՆ', 'ՈՍՏԻԿԱՆՈՒԹՅՈՒՆ» ԾԱՌԱՅՈՒԹՅՈՒՆ'],
    heading: 'ՀԱՐՑՈՒՄ',
    paragraphs: [
      'Հարգելի գործընկերներ,',
      'Ղեկավարվելով «Տեղեկատվության ազատության մասին» ՀՀ օրենքի 6-րդ և 9-րդ, «Վարչարարության հիմունքների և վարչական վարույթի մասին» ՀՀ օրենքի 30-րդ, 36-րդ, 39-րդ, 40-րդ և «Փաստաբանության մասին» ՀՀ օրենքի 18-րդ հոդվածներով՝ խնդրում եմ տրամադրել տեղեկատվություն, այն մասին, թե ՀՀ ոստիկանության «Ճանապարհային ոստիկանություն» ծառայության տվյալների շտեմարանում սույն հարցումը ներկայացնելու պահի դրությամբ և դրան նախորդող հինգ տարվա ընթացքում արդյոք գրանցված են {{client.fullNameGenitive}} {{client.requestBodyIdentity}} անվամբ սեփականության, ինչպես նաև այլ գույքային իրավունքներ՝ կցելով դրանց պետական գրանցման համար հիմք հանդիսացած փաստաթղթերի պատճենները:'
    ]
  },
  enforcement_service: {
    typeId: 'enforcement_service',
    title: ['ՀՀ ՀԱՐԿԱԴԻՐ ԿԱՏԱՐՈՒՄՆ', 'ԱՊԱՀՈՎՈՂ ԾԱՌԱՅՈՒԹՅՈՒՆ'],
    heading: 'ՀԱՐՑՈՒՄ',
    paragraphs: [
      'Հարգելի գործընկերներ,',
      'Ղեկավարվելով «Տեղեկատվության ազատության մասին» ՀՀ օրենքի 6-րդ և 9-րդ, «Վարչարարության հիմունքների և վարչական վարույթի մասին» ՀՀ օրենքի 30-րդ, 36-րդ, 39-րդ, 40-րդ և «Փաստաբանության մասին» ՀՀ օրենքի 18-րդ հոդվածներով՝ խնդրում եմ թույլ տալ ծանոթանալ {{client.fullNameGenitive}} {{client.requestBodyIdentity}} վերաբերյալ հարուցված կատարողական վարույթի նյութերին։ Ինչպես նաև խնդրում եմ տրամադրել՝',
      'Ընթացիկ կատարողական վարույթների ընթացքում ինչ գույք, ակտիվներ կամ գույքային իրավունքներ են հայտնաբերվել և արգելադրվել (իրացվել, բռնագանձվել):',
      'Սույն հարցմանը պատասխանելու օրվա դրությամբ պարտապանը ինչ գույք կամ գույքային իրավունքներ ունի, որոնց վրա կիրառված են արգելանքներ և սահմանափակումներ:',
      'Խնդրում եմ տրամադրել գույքի հայտնաբերման ուղղությամբ ձեռնարկված քայլերի (հարցումներ, հայտարարագրեր, ամփոփաթերթեր) և դրանց մասով ստացված պատասխանների, կատարողական վարույթը ավարտելու, կասեցնելու մասին որոշումների պատշաճ վավերացված պատճենները:'
    ]
  },
  state_reg: {
    typeId: 'state_reg',
    title: ['ՀՀ ԱՆ ԻՐԱՎԱԲԱՆԱԿԱՆ ԱՆՁԱՆՑ', 'ՊԵՏԱԿԱՆ ՌԵԳԻՍՏՐԻ ԳՈՐԾԱԿԱԼՈՒԹՅՈՒՆ'],
    heading: 'ՀԱՐՑՈՒՄ',
    paragraphs: [
      'Հարգելի գործընկերներ,',
      'Ղեկավարվելով «Տեղեկատվության ազատության մասին» ՀՀ օրենքի 6-րդ և 9-րդ, «Վարչարարության հիմունքների և վարչական վարույթի մասին» ՀՀ օրենքի 30-րդ, 36-րդ, 39-րդ, 40-րդ և «Փաստաբանության մասին» ՀՀ օրենքի 18-րդ հոդվածներով՝ խնդրում եմ տրամադրել տեղեկատվություն, այն մասին, թե ՀՀ ԱՆ իրավաբանական անձանց պետական ռեգիստրի գործակալության տվյալների շտեմարանում սույն հարցումը ներկայացնելու պահի դրությամբ {{client.fullNameGenitive}} {{client.requestBodyIdentity}} անվամբ՝',
      'Արդյո՞ք առկա են գույքի, այլ իրավունքների առկայության/բացակայության վերաբերյալ տեղեկություններ:',
      'Արդյո՞ք հաշվառված է որպես անհատ ձեռնարկատեր:',
      'Խնդրում եմ սույն հարցման պատասխանը ուղարկել բնօրինակ տարբերակով՝ {{representative.notificationAddress}} հասցեով։'
    ]
  }
}

const POA_TEMPLATES = {
  proc_poa: {
    typeId: 'proc_poa',
    title: '/Լ Ի Ա Զ Ո Ր Ա Գ Ի Ր/',
    paragraphs: [
      'Ես՝ {{client.fullNameWithSuffix}} {{client.identityWithPsnAndAddress}} լիազորում եմ {{poa.representatives}} ինձ հետ կապված ցանկացած հարցերով լինել իմ ներկայացուցիչը, իմ փոխարեն հանդես գալ, ներկայացնել և պաշտպանել իմ շահերը Հայաստանի Հանրապետության բոլոր պետական և ոչ պետական մարմիններում, հիմնարկ ձեռնարկություններում, կազմակերպություններում, իրավաբանական անձանց կամ այդպիսի կարգավիճակից օգտվող ցանկացած այլ միավորներում:',
      'Լիազորում եմ հանդես գալ ՀՀ կադաստրի կոմիտեում, ՀՀ Արդարադատության նախարարությունում, ՀՀ ՔԿԱԳ սպասարկման կենտրոններում, ՀՀ Ոստիկանության «Ճանապարհային ոստիկանություն» ծառայությունում, ՀՀ տարածքային կառավարման և ենթակառուցվածքների նախարարությունում, ՀՀ պետական եկամուտների կոմիտեում, «Կենտրոնական դեպոզիտարիա» ԲԲ ընկերության հետ հարաբերություններում, ՀՀ աշխատանքի և սոցիալական հարցերի նախարարության, ՀՀ իրավաբանական անձանց պետական ռեգիստրի գործակալությունում, ՀՀ էկոնոմիկայի նախարարությունում, ՀՀ ֆինանսների նախարարությունում, ՀՀ Ներքին գործերի նախարարությունում, ՀՀ բարձր տեխնոլոգիական արդյունաբերության նախարարությունում, և հանձնել դիմումներ, բողոքներ, հարցումներ, ստանալ իմ անվամբ ցանկացած գույքի նկատմամբ գրանցված իրավունքների մասին առկայության(բացակայության) վերաբերյալ տեղեկություններ, ինչպես նաև ցանկացած այլ տեսակի տեղեկություններ:',
      'Լիազորում եմ իմ անունից հանդես գալ բոլոր ատյանների ու տեսակի դատարաններում, ՀՀ Արբիտրաժային դատարաններում՝ կատարելով իմ իրավունքների ու օրինական շահերի պաշտպանությանն ուղղված գործողություններ, այդ թվում նաև՝ դատական նիստի ժամանակի և վայրի, ինչպես նաև ՀՀ քաղաքացիական դատավարության օրենսգրքով նախատեսված դեպքերում՝ առանձին դատավարական գործողություններ կատարելու մասին ծանուցումներ ստանալը, հանձնելու դիմումներ ու այլ փաստաթղթեր, ստանալու ինձ առնչվող փաստաթղթեր ու տեղեկություններ՝ օժտելով ՀՀ վարչական դատավարության օրենսգրքի 22 հոդվածով, ՀՀ քաղաքացիական դատավարության օրենսգրքի 56 հոդվածի 1-ին մասի 1-12 կետերով սահմանված բոլոր իրավունքներով՝ 1) հայցադիմումը ստորագրելու. 2) արբիտրաժային համաձայնություն կնքելու և վեճն արբիտրաժ հանձնելու վերաբերյալ համաձայնություն տալու. 3) հայցապահանջներից ամբողջովին կամ մասնակիորեն հրաժարվելու. 4) հայցապահանջներն ամբողջովին կամ մասնակիորեն ընդունելու. 5) հայցի առարկան և հիմքը կամ դրանցից յուրաքանչյուրը փոխելու. 6) հաշտության համաձայնություն կնքելու. 7) հաշտարարության վերաբերյալ համաձայնություն կնքելու. 8) արտոնագրված հաշտարարի մասնակցությամբ հաշտարարական գործընթացին մասնակցելու. 9) լիազորություններն այլ անձի փոխանցելու (վերալիազորում կատարելու). 10) դատական ծանուցումները և դատավարական փաստաթղթերն ստանալու. 11) դատական ակտը բողոքարկելու. 12) կատարողական թերթ տալու վերաբերյալ դիմում ներկայացնելու, 14) բողոքարկելու դատական ակտը և դիմելու իրավասու անձանց` վճռաբեկ բողոք բերելու խնդրանքով:',
      'Լիազորում եմ լինել իմ ներկայացուցիչը «ԱՔՌԱ Քրեդիտ Ռեփորթինգ» ՓԲԸ-ում կատարել ցանկացած անհրաժեշտ գործողություն՝ կապված իմ վարկային պատմության ճշգրտման, տեղեկությունների վիճարկման և այլ հարցերով, ստանալ ցանկացած վարկային զեկույց։ Ինչպես նաև լիազորում եմ իմ անունից «ՀԱՅՓՈՍՏ» ՓԲԸ-ից ստանալու ինձ հասցեագրված ցանկացած առաքանի և նամակ:',
      'Լիազորում եմ լինել իմ ներկայացուցիչը ՀՀ Հարկադիր կատարումն ապահովող ծառայությունում, իմ փոխարեն մասնակցելու ինձ կամ ինձ լիազորած ցանկացած անձի հետ կապված ցանկացած կատարողական վարույթին, դատական ծանուցումները և դատավարական փաստաթղթերն ստանալ, կատարողական թերթ տալու վերաբերյալ դիմում ներկայացնել, ներկայացնել կատարողական թերթ և հետ վերցնել, բռնագանձված գույքն ու (կամ) դրամն ստանալ և հարկադիր կատարողի գործողությունների դեմ բողոքարկել: Լիազորում եմ նաև ֆինանսական համակարգի հաշտարարի հետ ցանկացած տեսակի իրավահարաբերություններում լինել իմ ներկայացուցիչը։',
      'Լիազորագիրը տրված է երեք տարի ժամկետով, վերալիազորման իրավունքով:'
    ]
  },
  notarial_poa: {
    typeId: 'notarial_poa',
    title: '/Լ Ի Ա Զ Ո Ր Ա Գ Ի Ր/',
    paragraphs: [
      'Ես՝ {{client.fullNameWithSuffix}} {{client.identityWithPsnAndAddress}} լիազորում եմ {{poa.representatives}} ինձ հետ կապված ցանկացած հարցերով լինել իմ ներկայացուցիչը, իմ անունից հանդես գալ ՀՀ պետական և ոչ պետական մարմիններում և կազմակերպություններում, ՀՀ Կենտրոնական Բանկում, ՀՀ Կենտրոնական Բանկի կողմից վերահսկվող բոլոր առևտրային բանկերում և վարկային կազմակերպություններում։ Լիազորում եմ հանդես գալ Պետական եկամուտների կոմիտեում, ստանալ հարկային գաղտնիք կազմող ցանկացած տեղեկատվություն, անհատական հաշվի քաղվածքներ, կատարել վճարումներ, ներկայացնել դիմումներ: Լիազորում եմ ստանալ և/կամ ներկայացնել փաստաթղթեր, պայմանագրեր, վարկային պայմանագրեր, այդ թվում՝ երաշխավորության պայմանագրեր, այդ պայմանագրերի լուծման, փոփոխման, լրացման համաձայնագրեր, ծանուցումներ, պայմանագրի լուծման առաջարկներ, դիմումներ, համաձայնություններ, տեղեկանքներ, թույլտվություններ, հայտարարություններ, տրամադրել տեղեկություններ, կատարել վճարումներ, ստանալ տեղեկություններ, տեղեկանքներ, ցանկացած բնույթի ինֆորմացիա՝ հաշվի և հաշվով կատարված գործառնությունների վերաբերյալ: Լիազորում եմ իմ անունից բացել և/կամ փակել բանկային հաշիվներ, իրականացնել տվյալների փոփոխություն, KYC փաստաթղթերի ստորագրություն, ստանալու «Բանկային ավանդների ներգրավման մասին» Հայաստանի Հանրապետության օրենքի 6-րդ հոդվածով նախատեսված հաշվի քաղվածք, վարկի քաղվածք կամ «Սպառողական կրեդիտավորման մասին» Հայաստանի Հանրապետության օրենքի 17-րդ հոդվածի 2-րդ մասով նախատեսված տեղեկատվություն պարունակող փաստաթուղթ, իրացնելու «Սպառողական կրեդիտավորման մասին» Հայաստանի Հանրապետության օրենքի 17-րդ հոդվածով ինձ վերապահված բոլոր իրավունքները, այդ թվում նաև լիազորում եմ հրաժարվելու «Սպառողական կրեդիտավորման մասին» Հայաստանի Հանրապետության օրենքի 17-րդ հոդվածի 1-ին և 2-րդ մասերով սահմանված տեղեկատվությունն էլեկտրոնային կապի միջոցով ստանալու իրավունքից՝ պարտադիր ներկայացման տեղեկատվությունը իր կողմից մատնանշված փոստային կամ կապի այլ միջոցներով ստանալու պայմանով, միակողմանի փոփոխելու հաղորդակցման միջոցը, «Սպառողական կրեդիտավորման մասին» Հայաստանի Հանրապետության օրենքի 17-րդ հոդվածի 1-ին և 2-րդ մասերով սահմանված տեղեկատվությունն ստանալու կրեդիտավորողի տարածքում՝ առձեռն:',
      'Լիազորում եմ ստանալ ինձ՝ որպես բանկի հաճախորդին սպասարկելու կապակցությամբ տվյալ բանկին հայտնի դարձած իմ հաշիվների վերաբերյալ տեղեկությունները, իմ հանձնարարությամբ կամ հօգուտ ինձ կատարված գործառնությունների վերաբերյալ տեղեկությունները, ինչպես նաև իմ առևտրային գաղտնիքը, գործունեության ցանկացած ծրագրի կամ մշակման, գյուտի, արդյունաբերական դիզայնի մասին տեղեկությունները և իմ վերաբերյալ ցանկացած այլ տեղեկություն, որը ես մտադիր եմ եղել գաղտնի պահել, և բանկը տեղյակ է կամ կարող էր տեղյակ լինել այդ մտադրության վերաբերյալ: Լիազորում եմ լինել իմ ներկայացուցիչը ԱՔՌԱ վարկային բյուրոյում ստանալ ցանկացած վարկային զեկույց։ Լիազորում եմ նաև իմ անունից հանդես գալ ՀՀ կենտրոնական դեպոզիտարիայում, կատարել հարցումներ, հանձնել և ստանալ ցանկացած փաստաթուղթ։ Ինչպես նաև լիազորում եմ իմ անունից «ՀԱՅՓՈՍՏ» ՓԲԸ-ից ստանալու ինձ հասցեագրված ցանկացած առաքանի և նամակ:',
      'Լիազորում եմ ինձ ներկայացնելու Հայաստանի Հանրապետության հարկադիր կատարման ծառայությունում իմ՝ որպես պահանջատեր և/կամ պարտապան և/կամ երրորդ անձ մասնակցությամբ հարուցված կամ հարուցվելիք բոլոր կատարողական վարույթներով։ Ներկայացուցչին վերապահում եմ կատարողական վարույթի հետ կապված օրենքով թույլատրելի բոլոր գործողությունների կատարման լիազորություն, այդ թվում՝ ծանոթանալ կատարողական վարույթի նյութերին, ստանալ դրանց պատճենները, կատարել քաղվածքներ, լուսապատճեններ և լուսանկարներ, մասնակցել բոլոր կատարողական գործողություններին, ներկայացնել դիմումներ, միջնորդություններ, ապացույցներ, բացատրություններ և դիրքորոշումներ, հայտնել բացարկներ, ստանալ կատարողական վարույթի բոլոր փաստաթղթերը, ծանուցումները և ակտերը։ Լիազորում եմ բողոքարկել հարկադիր կատարողի որոշումները, գործողությունները և անգործությունը, ինչպես նաև ներկայացնել համապատասխան դիմումներ և բողոքներ, լիազորում եմ վիճարկումն ու բողոքարկումն իրականացնել նաև դատական կարգով։ Լիազորում եմ ներկայացնել կատարման ենթակա ակտի հարկադիր կատարման դիմում։ Ներկայացնել միջնորդություն պահանջատիրոջ իրավահաջորդությամբ փոխարինման վերաբերյալ (պահանջի զիջման հիմքով)։ Լիազորում եմ ներկայացնել միջնորդություններ հարկադիր կատարողին՝ կատարման հետաձգման, տարաժամկետման, կատարման եղանակի կամ կարգի սահմանման կամ փոփոխման վերաբերյալ, ինչպես նաև ներկայացնել համապատասխան դիմումներ, հայցեր կամ հայցադիմումներ դատարան։ Լիազորում եմ իմ անունից հրաժարվել պահանջից ամբողջությամբ կամ մասնակիորեն, ստանալ պահանջատիրոջը հասանելիք գույքը, ներառյալ դրամական միջոցները, կնքել բռնագանձման հերթականությունը փոփոխելու մասին համաձայնություն։ Կնքել հաշտության համաձայնություն, իրականացնել օրենքով նախատեսված կամ չարգելված այլ բոլոր գործողությունները, որոնք առնչվում են կատարողական վարույթին, իրացնել բոլոր իրավունքներն ու լիազորությունները, որոնք բխում են կատարողական վարույթում այդ պահին իմ ունեցած իրավական կարգավիճակից: Լիազորում եմ կատարել օրենքով չարգելված բոլոր այլ գործողությունները, որոնք անհրաժեշտ են իմ իրավունքների և օրինական շահերի պաշտպանության համար կատարողական վարույթում:',
      'Լիազորագիրը տրված է երեք տարի ժամկետով, վերալիազորման իրավունքով:'
    ]
  }
}

const POA_REPRESENTATIVES_TEXT =
  'փաստաբան Հովհաննես Գրիգորի Հարությունյանին /նույնականացման քարտ 009987561, տրված 20.02.2018թ. 011-ի կողմից, անձնագիր՝ AR0598051, տրված՝ 21.02.2018, 011-ի կողմից, հաշվառված ք. Երևան, Արհեստավորների 2-րդ փողոց տուն 8 հասցեում, փաստաբանական գործունեության արտոնագիր N2332/, և/կամ Կարինե Նաիրիի Ավետիսյանին /անձնագիր AR0610106, տրված 08.01.2018, 001-ի կողմից, նույնականացման քարտ՝ 013922554, տրված՝ 27.01.2022թ., 011-ի կողմից, հաշվառված Գեղարքունիքի մարզ, գ. Կարճաղբյուր 15 փ. 5-րդ տուն հասցեում/, և/կամ Գրիգոր Նիկոլի Հարությունյանին /անձնագիր՝ AO0421258, տրված՝ 05.12.2025թ., 067-ի կողմից, նույնականացման քարտ՝ 011598737, տրված՝ 07.12.2020թ., 011-ի կողմից, հաշվառված ք. Վանաձոր, Երևանյան խճղ., 147, 2շ., 5բնկ. հասցեում/'

const ALLOWED_SHEETS = Object.freeze(Object.keys(HEADERS))
let spreadsheetCache = null

function getSpreadsheet() {
  if (!spreadsheetCache) spreadsheetCache = SpreadsheetApp.openById(SPREADSHEET_ID)
  return spreadsheetCache
}

function getOrCreateSheet(name) {
  if (!ALLOWED_SHEETS.includes(name)) {
    throw new Error('Unknown sheet: ' + name)
  }

  const ss = getSpreadsheet()
  let sheet = ss.getSheetByName(name)
  if (!sheet) {
    sheet = ss.insertSheet(name)
  }
  // Auto-create headers if row 1 is empty
  const firstRow = sheet.getRange(1, 1, 1, 1).getValue()
  if (!firstRow && HEADERS[name]) {
    sheet.getRange(1, 1, 1, HEADERS[name].length).setValues([HEADERS[name]])
    sheet.setFrozenRows(1)
    formatHeaderRow(sheet, HEADERS[name].length)
  } else {
    migrateSheetSchema(sheet, name)
  }
  if (name === 'CaseTypes') ensureDefaultCaseTypes(sheet)
  if (name === 'Settings') ensureDefaultSettings(sheet)
  return sheet
}

function ensureDefaultCaseTypes(sheet) {
  if (sheet.getLastRow() > 1) return
  const now = new Date().toISOString()
  sheet.getRange(2, 1, DEFAULT_CASE_TYPES.length, HEADERS.CaseTypes.length)
    .setValues(DEFAULT_CASE_TYPES.map(function(row) { return row.concat(now) }))
}

function ensureDefaultSettings(sheet) {
  if (sheet.getLastRow() > 1) return
  const now = new Date().toISOString()
  sheet.getRange(2, 1, 3, HEADERS.Settings.length).setValues([
    ['REMINDER_DAYS', String(DEFAULT_REMINDER_DAYS), 'Default email reminder lead time', now],
    ['SHARED_CALENDAR_ID', '', 'Shared Google Calendar ID', now],
    ['BACKUP_FOLDER_ID', '', 'Drive folder ID for scheduled spreadsheet backups', now]
  ])
}

function migrateSheetSchema(sheet, sheetName) {
  const expected = HEADERS[sheetName]
  const lastColumn = Math.max(sheet.getLastColumn(), 1)
  const rawHeaders = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0]
  const currentHeaders = rawHeaders.map(normalizeHeader)
  const expectedNormalized = expected.map(normalizeHeader)

  const alreadyAligned = expectedNormalized.every((header, index) =>
    currentHeaders[index] === header
  )
  if (alreadyAligned) {
    if (lastColumn < expected.length) {
      sheet.getRange(1, lastColumn + 1, 1, expected.length - lastColumn)
        .setValues([expected.slice(lastColumn)])
      formatHeaderRow(sheet, expected.length)
    }
    applySheetTextFormats(sheet)
    return
  }

  backupSheetBeforeMigration(sheet, sheetName)

  const sourceIndex = {}
  currentHeaders.forEach((header, index) => {
    if (header && sourceIndex[header] === undefined) sourceIndex[header] = index
  })

  const lastRow = Math.max(sheet.getLastRow(), 1)
  const sourceData = sheet.getRange(1, 1, lastRow, lastColumn).getValues()
  const migrated = [expected]
  for (let rowIndex = 1; rowIndex < sourceData.length; rowIndex++) {
    migrated.push(expectedNormalized.map((header, expectedIndex) => {
      const columnIndex = sourceIndex[header]
      if (columnIndex !== undefined) return sourceData[rowIndex][columnIndex]

      // Older app versions wrote rows in schema order even when headers were
      // renamed manually. Preserve that data by falling back to its position.
      return expectedIndex < sourceData[rowIndex].length
        ? sourceData[rowIndex][expectedIndex]
        : ''
    }))
  }

  sheet.clearContents()
  sheet.getRange(1, 1, migrated.length, expected.length).setValues(migrated)
  sheet.setFrozenRows(1)
  formatHeaderRow(sheet, expected.length)
  applySheetTextFormats(sheet)
}

function normalizeHeader(value) {
  return String(value || '')
    .replace(/^\uFEFF/, '')
    .trim()
    .toLowerCase()
}

function backupSheetBeforeMigration(sheet, sheetName) {
  const ss = sheet.getParent()
  const timezone = ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone()
  const stamp = Utilities.formatDate(new Date(), timezone, 'yyyyMMdd-HHmmss')
  let backupName = (sheetName + '_backup_' + stamp).slice(0, 90)
  let suffix = 1
  while (ss.getSheetByName(backupName)) {
    backupName = (sheetName + '_backup_' + stamp + '_' + suffix++).slice(0, 90)
  }
  sheet.copyTo(ss).setName(backupName).hideSheet()
}

function formatHeaderRow(sheet, width) {
  sheet.getRange(1, 1, 1, width)
    .setBackground('#1a1a2e')
    .setFontColor('#c9a84c')
    .setFontWeight('bold')
  applySheetTextFormats(sheet)
}

function applySheetTextFormats(sheet) {
  if (!sheet || sheet.getName() !== 'Cases') return
  ;['passport','spousePassport','idCard','spouseIdCard','passportBy','spousePassportBy','idCardBy','spouseIdCardBy','psn','spousePsn'].forEach(function(header) {
    const index = HEADERS.Cases.indexOf(header)
    if (index !== -1) sheet.getRange(1, index + 1, Math.max(sheet.getMaxRows(), 1), 1).setNumberFormat('@')
  })
}

function repairAllSheetSchemas() {
  Object.keys(HEADERS).forEach(name => getOrCreateSheet(name))
  return migrateExistingCases()
}

function doGet(e) {
  const action = e && e.parameter ? e.parameter.action : ''
  try {
    const user = assertAuthorizedUser()
    let result
    if (action === 'getCases') result = getRows('Cases')
    else if (action === 'getDashboardData') result = getDashboardData()
    else if (action === 'getCaseBundle') result = getCaseBundle(e.parameter.caseId)
    else if (action === 'getDocs') result = e.parameter.caseId ? getRows('Documents', 'caseId', e.parameter.caseId) : getRows('Documents')
    else if (action === 'getDebts') result = e.parameter.caseId ? getRows('Debts', 'caseId', e.parameter.caseId) : getRows('Debts')
    else if (action === 'getBankCertificates') result = e.parameter.caseId ? getRows('BankCertificates', 'caseId', e.parameter.caseId) : getRows('BankCertificates')
    else if (action === 'getFiles') result = e.parameter.caseId ? getRows('CaseFiles', 'caseId', e.parameter.caseId) : getRows('CaseFiles')
    else if (action === 'getHearings') result = e.parameter.caseId ? getRows('Hearings', 'caseId', e.parameter.caseId) : getRows('Hearings')
    else if (action === 'getDeadlines') result = e.parameter.caseId ? getRows('Deadlines', 'caseId', e.parameter.caseId) : getRows('Deadlines')
    else if (action === 'getIncomingMail') result = e.parameter.caseId ? getRows('IncomingMail', 'caseId', e.parameter.caseId) : getRows('IncomingMail')
    else if (action === 'getCaseTypes') result = getCaseTypes(e.parameter.includeDisabled === 'true')
    else if (action === 'getSettings') result = getSettings()
    else if (action === 'searchCases') result = searchCases(e.parameter.q)
    else if (action === 'getAudit') result = getRows('AuditLog').slice(-100).reverse()
    else if (action === 'repairIdentityTextCodes') result = repairIdentityTextCodes()
    else if (action === 'ping') result = { status: 'ok', time: new Date().toISOString(), user: user }
    else result = { error: 'Unknown action: ' + action }
    return jsonResponse(result)
  } catch(err) {
    return jsonResponse({ error: err.message })
  }
}

function doPost(e) {
  try {
    assertAuthorizedUser()
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('Missing request body')
    }

    const data = JSON.parse(e.postData.contents)
    const action = data.action
    let result
    if (action === 'saveCase') result = saveCaseWithFolders(validateRow('Cases', data.row))
    else if (action === 'saveDoc') result = upsertRow('Documents', validateRow('Documents', data.row))
    else if (action === 'saveDebt') result = upsertRow('Debts', validateRow('Debts', data.row))
    else if (action === 'saveBankCertificate') result = upsertRow('BankCertificates', validateRow('BankCertificates', data.row))
    else if (action === 'saveHearing') result = saveScheduledItem('Hearings', data.row)
    else if (action === 'deleteHearing') result = deleteScheduledItem('Hearings', data.id)
    else if (action === 'saveDeadline') result = saveScheduledItem('Deadlines', data.row)
    else if (action === 'deleteDeadline') result = deleteScheduledItem('Deadlines', data.id)
    else if (action === 'logIncomingMail') result = logIncomingMail(data)
    else if (action === 'saveCaseType') result = saveCaseType(data)
    else if (action === 'saveSetting') result = saveSetting(data)
    else if (action === 'migrateCaseStage') result = migrateCaseStage(data)
    else if (action === 'uploadDocument') result = uploadDocument(data)
    else if (action === 'generateRegisterRequest') result = generateRegisterRequest(data)
    else if (action === 'generateStateRequest') result = generateStateRequest(data)
    else if (action === 'generatePoa') result = generatePoa(data)
    else if (action === 'deleteCase') result = deleteCase(data)
    else if (action === 'importIngaApplicantVahagnSpouseCase') result = importIngaApplicantVahagnSpouseCase()
    else if (action === 'repairBankCertificateExpiryDates') result = repairBankCertificateExpiryDates()
    else if (action === 'repairIdentityTextCodes') result = repairIdentityTextCodes()
    else if (action === 'migrateExistingCases') result = migrateExistingCases()
    else if (action === 'runReminders') result = sendUpcomingReminders()
    else if (action === 'runBackup') result = backupSpreadsheet()
    else if (action === 'installMaintenanceTriggers') result = installMaintenanceTriggers()
    else if (action === 'audit') result = appendRow('AuditLog', validateRow('AuditLog', data.row))
    else result = { error: 'Unknown action: ' + action }
    return jsonResponse(result)
  } catch(err) {
    return jsonResponse({ error: err.message })
  }
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON)
}

function assertAuthorizedUser() {
  if (PUBLIC_ACCESS_ENABLED) return currentUserEmail()
  const props = PropertiesService.getScriptProperties()
  const email = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase()
  if (props.getProperty('IKV_ALLOW_UNAUTHENTICATED') === 'true') {
    return email || 'explicitly-allowed-unauthenticated-user'
  }
  if (!email) {
    throw new Error('Authentication required. Deploy the web app for firm accounts only.')
  }
  const allowedEmails = String(props.getProperty('IKV_ALLOWED_EMAILS') || '')
    .split(',').map(function(value) { return value.trim().toLowerCase() }).filter(Boolean)
  const allowedDomain = String(props.getProperty('IKV_ALLOWED_DOMAIN') || '').trim().toLowerCase()
  if (allowedEmails.length && allowedEmails.indexOf(email) === -1) {
    throw new Error('This account is not authorised for IKV.')
  }
  if (allowedDomain && !email.endsWith('@' + allowedDomain)) {
    throw new Error('This account is outside the authorised IKV domain.')
  }
  return email
}

function currentUserEmail() {
  return String(Session.getActiveUser().getEmail() || 'web-app-user').trim()
}

function rowToObject(sheetName, row) {
  return Object.fromEntries(HEADERS[sheetName].map(function(header, index) {
    return [header, row[index] || '']
  }))
}

function objectToRow(sheetName, object) {
  return HEADERS[sheetName].map(function(header) {
    return safeCellValue(object[header])
  })
}

function getSettings() {
  return Object.fromEntries(getRows('Settings').map(function(item) {
    return [item.key, item.value]
  }))
}

function getSetting(key, fallback) {
  const row = getRows('Settings').find(function(item) { return item.key === key })
  return row && row.value !== '' ? row.value : fallback
}

function saveSetting(data) {
  const key = String(data.key || '').trim().toUpperCase()
  if (['REMINDER_DAYS','SHARED_CALENDAR_ID','BACKUP_FOLDER_ID'].indexOf(key) === -1) {
    throw new Error('Unknown setting')
  }
  const existing = getRows('Settings').find(function(item) { return item.key === key }) || {}
  const row = objectToRow('Settings', {
    key: key,
    value: String(data.value || '').trim(),
    description: existing.description || String(data.description || ''),
    updatedAt: new Date().toISOString()
  })
  const result = upsertRow('Settings', row)
  auditAction('UPDATE', 'Setting', key, { value: key === 'SHARED_CALENDAR_ID' ? '[configured]' : data.value })
  return result
}

function getCaseTypes(includeDisabled) {
  return getRows('CaseTypes')
    .filter(function(item) { return includeDisabled || item.enabled !== 'NO' })
    .sort(function(a, b) { return Number(a.sortOrder || 0) - Number(b.sortOrder || 0) })
}

function saveCaseType(data) {
  const key = String(data.key || '').trim().toLowerCase().replace(/[^a-z0-9_-]/g, '')
  const label = String(data.label || '').trim()
  if (!key || !label) throw new Error('Case type key and label are required')
  const row = objectToRow('CaseTypes', {
    key: key,
    label: label,
    enabled: data.enabled === false || data.enabled === 'NO' ? 'NO' : 'YES',
    sortOrder: String(Number(data.sortOrder || 999)),
    updatedAt: new Date().toISOString()
  })
  const result = upsertRow('CaseTypes', row)
  auditAction('UPDATE', 'CaseType', key, { label: label, enabled: row[2] })
  return result
}

function auditAction(action, entity, entityId, detail) {
  return appendRow('AuditLog', validateRow('AuditLog', [
    Utilities.getUuid(), new Date().toISOString(), currentUserEmail(),
    action, entity, entityId, JSON.stringify(detail || {})
  ]))
}

function searchCases(query) {
  const needle = String(query || '').trim().toLocaleLowerCase()
  if (!needle) return []
  const mailByCase = {}
  getRows('IncomingMail').forEach(function(item) {
    if (!mailByCase[item.caseId]) mailByCase[item.caseId] = []
    mailByCase[item.caseId].push(item.sender)
  })
  return getRows('Cases').filter(function(item) {
    return [
      item.applicantName, item.spouseName, item.internalNumber, item.id,
      item.officialCaseNumber, item.caseType, item.stage, item.storageNumber,
      item.storageRoom, item.storageShelf, item.storageBox, item.locationHints,
      (mailByCase[item.id] || []).join(' ')
    ].some(function(value) {
      return String(value || '').toLocaleLowerCase().indexOf(needle) !== -1
    })
  }).slice(0, 50)
}

function getRows(sheetName, filterCol, filterVal) {
  const sheet = getOrCreateSheet(sheetName)
  const data = sheet.getDataRange().getDisplayValues()
  if (data.length < 2) return []
  const headers = data[0]
  let rows = data.slice(1).map(row =>
    Object.fromEntries(headers.map((h, i) => [h, row[i] || '']))
  ).filter(row => row[headers[0]] !== '') // skip empty rows

  if (filterCol) {
    if (!headers.includes(filterCol)) throw new Error('Unknown filter column: ' + filterCol)
    rows = rows.filter(r => r[filterCol] === filterVal)
  }
  return rows
}

function getDashboardData() {
  return {
    cases: getRows('Cases'),
    documents: getRows('Documents'),
    bankCertificates: getRows('BankCertificates'),
    hearings: getRows('Hearings'),
    deadlines: getRows('Deadlines'),
    caseTypes: getCaseTypes(false)
  }
}

function getCaseBundle(caseId) {
  if (!caseId) throw new Error('Missing case ID')
  const caseRow = getRows('Cases').find(row => row.id === caseId)
  if (!caseRow) throw new Error('Case not found')
  const folder = resolveCaseFolder(caseRow)
  return {
    case: caseRow,
    driveFolderId: folder.getId(),
    driveFolderUrl: folder.getUrl(),
    documents: getRows('Documents', 'caseId', caseId),
    debts: getRows('Debts', 'caseId', caseId),
    bankCertificates: getRows('BankCertificates', 'caseId', caseId),
    files: getRows('CaseFiles', 'caseId', caseId),
    hearings: getRows('Hearings', 'caseId', caseId),
    deadlines: getRows('Deadlines', 'caseId', caseId),
    incomingMail: getRows('IncomingMail', 'caseId', caseId)
  }
}

function upsertRow(sheetName, row) {
  const lock = LockService.getScriptLock()
  lock.waitLock(30000)
  try {
    const sheet = getOrCreateSheet(sheetName)
    const data = sheet.getDataRange().getDisplayValues()
    const id = row[0]

    // Find existing row by ID
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        sheet.getRange(i + 1, 1, 1, row.length).setValues([row])
        return { status: 'updated', id }
      }
    }
    sheet.appendRow(row)
    return { status: 'created', id }
  } finally {
    lock.releaseLock()
  }
}

function appendRow(sheetName, row) {
  const lock = LockService.getScriptLock()
  lock.waitLock(30000)
  try {
    const sheet = getOrCreateSheet(sheetName)
    sheet.appendRow(row)
    return { status: 'ok' }
  } finally {
    lock.releaseLock()
  }
}

function validateRow(sheetName, row) {
  if (!Array.isArray(row)) throw new Error('Row must be an array')
  if (row.length !== HEADERS[sheetName].length) {
    throw new Error(sheetName + ' row has an invalid number of columns')
  }

  const normalized = row.map(safeCellValue)
  if (sheetName === 'Cases') normalizeCaseTextCodes(normalized)
  if (!normalized[0]) throw new Error(sheetName + ' row is missing an ID')
  return normalized
}

function normalizeCaseTextCodes(row) {
  ;['passport','spousePassport','idCard','spouseIdCard','passportBy','spousePassportBy','idCardBy','spouseIdCardBy','psn','spousePsn'].forEach(function(header) {
    const index = HEADERS.Cases.indexOf(header)
    if (index === -1 || row[index] === '') return
    const next = /By$/.test(header) ? normalizeIssuerCode(row[index]) : normalizeDocumentNumber(row[index])
    row[index] = "'" + next
  })
}

function normalizeIssuerCode(value) {
  const raw = String(value || '').replace(/^'/, '').trim()
  return /^\d{1,2}$/.test(raw) ? raw.padStart(3, '0') : raw
}

function normalizeDocumentNumber(value) {
  return String(value || '').replace(/^'/, '').trim()
}

function repairIdentityTextCodes() {
  const sheet = getOrCreateSheet('Cases')
  applySheetTextFormats(sheet)
  const headers = HEADERS.Cases
  const codeHeaders = ['passportBy','spousePassportBy','idCardBy','spouseIdCardBy']
  const textHeaders = ['passport','spousePassport','idCard','spouseIdCard','psn','spousePsn']
  const columns = codeHeaders.concat(textHeaders).map(function(header) {
    return { header: header, index: headers.indexOf(header) }
  }).filter(function(item) {
    return item.index !== -1
  })
  const lastRow = sheet.getLastRow()
  if (lastRow < 2) return { status: 'ok', updatedCells: 0 }

  const range = sheet.getRange(2, 1, lastRow - 1, headers.length)
  const values = range.getDisplayValues()
  let updatedCells = 0
  values.forEach(function(row) {
    columns.forEach(function(item) {
      const raw = String(row[item.index] || '').replace(/^'/, '').trim()
      const next = codeHeaders.indexOf(item.header) !== -1
        ? normalizeIssuerCode(raw)
        : normalizeDocumentNumber(raw)
      if (next && raw !== next) {
        row[item.index] = "'" + next
        updatedCells++
      }
    })
  })
  range.setValues(values)
  applySheetTextFormats(sheet)
  return { status: 'ok', updatedCells: updatedCells }
}

function safeCellValue(value) {
  if (value === undefined || value === null) return ''
  const text = String(value).trim()
  return /^[=+\-@]/.test(text) ? "'" + text : text
}

function saveCaseWithFolders(row) {
  const idIndex = HEADERS.Cases.indexOf('id')
  const existing = getRows('Cases').find(function(item) { return item.id === row[idIndex] })
  const user = currentUserEmail()
  const now = new Date().toISOString()
  setCaseDefault(row, 'internalNumber', existing ? existing.internalNumber : nextInternalNumber())
  setCaseDefault(row, 'caseType', existing ? existing.caseType : 'bankruptcy')
  setCaseDefault(row, 'stage', existing ? existing.stage : 'preparation')
  setCaseDefault(row, 'stageChangedAt', existing ? existing.stageChangedAt : now)
  setCaseDefault(row, 'stageChangedBy', existing ? existing.stageChangedBy : user)
  setCaseDefault(row, 'createdAt', existing ? existing.createdAt : now)
  setCaseDefault(row, 'createdBy', existing ? existing.createdBy : user)
  row[HEADERS.Cases.indexOf('updatedAt')] = now
  row[HEADERS.Cases.indexOf('updatedBy')] = user
  validateCaseLifecycle(row)
  const folderIdIndex = HEADERS.Cases.indexOf('driveFolderId')
  const folderUrlIndex = HEADERS.Cases.indexOf('driveFolderUrl')
  let folder = null
  if (row[folderIdIndex]) {
    try {
      folder = DriveApp.getFolderById(row[folderIdIndex])
      if (folder.isTrashed()) folder = null
    } catch (err) {}
  }
  if (!folder) {
    folder = createCaseFolderTree(row[0], row[1])
    row[folderIdIndex] = folder.getId()
    row[folderUrlIndex] = folder.getUrl()
  } else {
    ensureCaseFolderTree(folder)
  }
  const saved = upsertRow('Cases', row)
  saved.driveFolderId = row[folderIdIndex]
  saved.driveFolderUrl = row[folderUrlIndex]
  saved.internalNumber = row[HEADERS.Cases.indexOf('internalNumber')]
  saved.caseType = row[HEADERS.Cases.indexOf('caseType')]
  saved.stage = row[HEADERS.Cases.indexOf('stage')]
  auditAction(existing ? 'UPDATE' : 'CREATE', 'Case', row[idIndex], {
    applicantName: row[HEADERS.Cases.indexOf('applicantName')],
    internalNumber: saved.internalNumber,
    caseType: saved.caseType,
    stage: saved.stage
  })
  return saved
}

function validateCaseLifecycle(row) {
  const stage = row[HEADERS.Cases.indexOf('stage')]
  function required(header, message) {
    if (!row[HEADERS.Cases.indexOf(header)]) throw new Error(message)
  }
  if (stage === 'process') {
    required('officialCaseNumber', 'Official court case number is required')
    required('court', 'Court is required')
    required('courtInstance', 'Court instance is required')
    required('judge', 'Judge is required')
  }
  if (stage === 'archive') {
    required('storageRoom', 'Archive room is required')
    required('storageShelf', 'Archive shelf is required')
    required('storageBox', 'Archive box is required')
    setCaseDefault(row, 'archivedAt', new Date().toISOString())
  }
}

function setCaseDefault(row, header, value) {
  const index = HEADERS.Cases.indexOf(header)
  if (index !== -1 && !row[index]) row[index] = value || ''
}

function nextInternalNumber() {
  const lock = LockService.getScriptLock()
  lock.waitLock(30000)
  try {
    const props = PropertiesService.getScriptProperties()
    const stored = Number(props.getProperty('IKV_CASE_SEQUENCE') || 0)
    const existingMax = getRows('Cases').reduce(function(max, item) {
      const match = String(item.internalNumber || '').match(/(\d+)$/)
      return Math.max(max, match ? Number(match[1]) : 0)
    }, 0)
    const next = Math.max(stored, existingMax) + 1
    props.setProperty('IKV_CASE_SEQUENCE', String(next))
    return String(next)
  } finally {
    lock.releaseLock()
  }
}

function createCaseFolderTree(caseId, applicantName) {
  const root = getOrCreateRootFolder()
  const folderName = sanitizeDriveName(applicantName + ' - ' + caseId)
  const existing = root.getFoldersByName(folderName)
  const caseFolder = existing.hasNext() ? existing.next() : root.createFolder(folderName)
  ensureCaseFolderTree(caseFolder)
  return caseFolder
}

function ensureCaseFolderTree(caseFolder) {
  CASE_FOLDER_NAMES.forEach(name => getOrCreateChildFolder(caseFolder, name))
  return caseFolder
}

function getOrCreateRootFolder() {
  const props = PropertiesService.getScriptProperties()
  const savedId = props.getProperty('IKV_ROOT_FOLDER_ID')
  if (savedId) {
    try {
      const saved = DriveApp.getFolderById(savedId)
      if (!saved.isTrashed()) return saved
    } catch (err) {}
  }
  const folders = DriveApp.getFoldersByName(ROOT_FOLDER_NAME)
  const root = folders.hasNext() ? folders.next() : DriveApp.createFolder(ROOT_FOLDER_NAME)
  props.setProperty('IKV_ROOT_FOLDER_ID', root.getId())
  return root
}

function getOrCreateChildFolder(parent, name) {
  const folders = parent.getFoldersByName(name)
  while (folders.hasNext()) {
    const folder = folders.next()
    if (!folder.isTrashed()) return folder
  }
  return parent.createFolder(name)
}

function resolveCaseFolder(caseRow) {
  let folder = null
  if (caseRow.driveFolderId) {
    try {
      folder = DriveApp.getFolderById(caseRow.driveFolderId)
      if (folder.isTrashed()) folder = null
    } catch (err) {}
  }
  if (folder) return ensureCaseFolderTree(folder)

  folder = createCaseFolderTree(caseRow.id, caseRow.applicantName)
  const sheet = getOrCreateSheet('Cases')
  const data = sheet.getDataRange().getDisplayValues()
  const folderIdColumn = HEADERS.Cases.indexOf('driveFolderId') + 1
  const folderUrlColumn = HEADERS.Cases.indexOf('driveFolderUrl') + 1
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === caseRow.id) {
      sheet.getRange(i + 1, folderIdColumn).setValue(folder.getId())
      sheet.getRange(i + 1, folderUrlColumn).setValue(folder.getUrl())
      break
    }
  }
  caseRow.driveFolderId = folder.getId()
  caseRow.driveFolderUrl = folder.getUrl()
  return folder
}

function migrateCaseStage(data) {
  const caseId = String(data.caseId || '').trim()
  const targetStage = String(data.targetStage || '').trim()
  const stages = ['preparation','process','archive']
  if (!caseId || stages.indexOf(targetStage) === -1) throw new Error('Invalid stage migration')

  const lock = LockService.getScriptLock()
  lock.waitLock(30000)
  try {
    const sheet = getOrCreateSheet('Cases')
    const values = sheet.getDataRange().getDisplayValues()
    const headers = values[0]
    const rowIndex = values.findIndex(function(row, index) { return index > 0 && row[0] === caseId })
    if (rowIndex === -1) throw new Error('Case not found')
    const item = Object.fromEntries(headers.map(function(header, index) {
      return [header, values[rowIndex][index] || '']
    }))
    const currentIndex = stages.indexOf(item.stage || 'preparation')
    const targetIndex = stages.indexOf(targetStage)
    if (Math.abs(targetIndex - currentIndex) !== 1) {
      throw new Error('Cases can move only one stage at a time')
    }
    if (targetStage === 'process') {
      if (!String(data.officialCaseNumber || item.officialCaseNumber || '').trim()) {
        throw new Error('Official court case number is required')
      }
      if (!String(data.court || item.court || '').trim()) throw new Error('Court is required')
      if (!String(data.courtInstance || item.courtInstance || '').trim()) throw new Error('Court instance is required')
      if (!String(data.judge || item.judge || '').trim()) throw new Error('Judge is required')
    }
    if (targetStage === 'archive') {
      if (!String(data.storageRoom || '').trim() ||
          !String(data.storageShelf || '').trim() ||
          !String(data.storageBox || '').trim()) {
        throw new Error('Archive room, shelf and box are required')
      }
    }

    const now = new Date().toISOString()
    const user = currentUserEmail()
    item.stage = targetStage
    item.stageChangedAt = now
    item.stageChangedBy = user
    item.updatedAt = now
    item.updatedBy = user
    ;['officialCaseNumber','court','courtInstance','judge','caseStatus','finalOutcome',
      'storageNumber','storageRoom','storageShelf','storageBox','locationHints'].forEach(function(key) {
      if (data[key] !== undefined) item[key] = String(data[key] || '').trim()
    })
    if (targetStage === 'archive') item.archivedAt = now
    if (targetStage !== 'archive' && currentIndex === 2) item.archivedAt = ''

    const nextRow = objectToRow('Cases', item)
    sheet.getRange(rowIndex + 1, 1, 1, nextRow.length).setValues([nextRow])
    const detail = {
      from: stages[currentIndex],
      to: targetStage,
      internalNumber: item.internalNumber,
      officialCaseNumber: item.officialCaseNumber,
      archiveLocation: archiveLocation(item)
    }
    getOrCreateSheet('AuditLog').appendRow(validateRow('AuditLog', [
      Utilities.getUuid(), now, user,
      targetIndex > currentIndex ? 'MIGRATE_FORWARD' : 'MIGRATE_BACKWARD',
      'Case', caseId, JSON.stringify(detail)
    ]))
    return item
  } finally {
    lock.releaseLock()
  }
}

function archiveLocation(item) {
  return [item.storageRoom, item.storageShelf, item.storageBox]
    .filter(Boolean).join(' / ')
}

function logIncomingMail(data) {
  const caseId = String(data.caseId || '').trim()
  const caseRow = getRows('Cases').find(function(item) { return item.id === caseId })
  if (!caseRow) throw new Error('Case not found')
  const item = {
    id: String(data.id || Utilities.getUuid()),
    dateReceived: String(data.dateReceived || new Date().toISOString()),
    sender: String(data.sender || '').trim(),
    caseId: caseId,
    internalNumber: caseRow.internalNumber,
    loggedBy: currentUserEmail(),
    note: String(data.note || '').trim()
  }
  appendRow('IncomingMail', validateRow('IncomingMail', objectToRow('IncomingMail', item)))
  auditAction('MAIL_RECEIVED', 'Case', caseId, {
    mailId: item.id, sender: item.sender, dateReceived: item.dateReceived
  })
  return item
}

function saveScheduledItem(sheetName, row) {
  const item = rowToObject(sheetName, validateRow(sheetName, row))
  const caseRow = getRows('Cases').find(function(value) { return value.id === item.caseId })
  if (!caseRow) throw new Error('Case not found')
  if (sheetName === 'Hearings' && !item.dateTime) throw new Error('Hearing date/time is required')
  if (sheetName === 'Deadlines' && (!item.dueDate || !item.description)) {
    throw new Error('Deadline date and description are required')
  }
  if (!item.reminderDaysBefore) item.reminderDaysBefore = getSetting('REMINDER_DAYS', DEFAULT_REMINDER_DAYS)
  if (!item.responsibleEmails) item.responsibleEmails = caseRow.responsibleEmails
  item.updatedAt = new Date().toISOString()
  item.calendarEventId = syncCalendarItem(sheetName, item, caseRow)
  const result = upsertRow(sheetName, validateRow(sheetName, objectToRow(sheetName, item)))
  auditAction('UPDATE', sheetName.slice(0, -1), item.id, {
    caseId: item.caseId,
    date: sheetName === 'Hearings' ? item.dateTime : item.dueDate
  })
  result.item = item
  return result
}

function deleteScheduledItem(sheetName, id) {
  const item = getRows(sheetName).find(function(value) { return value.id === id })
  if (!item) throw new Error('Item not found')
  deleteCalendarEvent(item.calendarEventId)
  const deleted = deleteRowsByValue(sheetName, 'id', id)
  auditAction('DELETE', sheetName.slice(0, -1), id, { caseId: item.caseId })
  return { status: 'deleted', id: id, deletedRows: deleted }
}

function syncCalendarItem(sheetName, item, caseRow) {
  const calendarId = String(getSetting('SHARED_CALENDAR_ID', '') || '').trim()
  if (!calendarId) return item.calendarEventId || ''
  const calendar = CalendarApp.getCalendarById(calendarId)
  if (!calendar) throw new Error('Shared calendar was not found')
  const isHearing = sheetName === 'Hearings'
  const start = new Date(isHearing ? item.dateTime : item.dueDate + 'T09:00:00')
  if (isNaN(start.getTime())) throw new Error('Invalid calendar date')
  const end = new Date(start.getTime() + (isHearing ? 60 : 30) * 60000)
  const title = isHearing
    ? 'Դատական նիստ · ' + caseRow.internalNumber + ' · ' + caseRow.applicantName
    : 'Ժամկետ · ' + caseRow.internalNumber + ' · ' + item.description
  const description = [
    caseRow.applicantName,
    'Ներքին համար՝ ' + caseRow.internalNumber,
    caseRow.officialCaseNumber ? 'Դատական համար՝ ' + caseRow.officialCaseNumber : '',
    isHearing ? item.note : item.description
  ].filter(Boolean).join('\n')
  let event = null
  if (item.calendarEventId) {
    try { event = calendar.getEventById(item.calendarEventId) } catch (err) {}
  }
  if (event) {
    event.setTitle(title).setTime(start, end).setDescription(description)
  } else {
    event = calendar.createEvent(title, start, end, { description: description })
  }
  return event.getId()
}

function deleteCalendarEvent(eventId) {
  if (!eventId) return
  const calendarId = String(getSetting('SHARED_CALENDAR_ID', '') || '').trim()
  if (!calendarId) return
  try {
    const event = CalendarApp.getCalendarById(calendarId).getEventById(eventId)
    if (event) event.deleteEvent()
  } catch (err) {}
}

function sendUpcomingReminders() {
  const now = new Date()
  const casesById = Object.fromEntries(getRows('Cases').map(function(item) { return [item.id, item] }))
  let sent = 0
  ;['Hearings','Deadlines'].forEach(function(sheetName) {
    const sheet = getOrCreateSheet(sheetName)
    const items = getRows(sheetName)
    items.forEach(function(item) {
      if (sheetName === 'Deadlines' && item.done === 'YES') return
      const target = new Date(sheetName === 'Hearings' ? item.dateTime : item.dueDate + 'T09:00:00')
      const days = Number(item.reminderDaysBefore || getSetting('REMINDER_DAYS', DEFAULT_REMINDER_DAYS))
      const hoursUntil = (target.getTime() - now.getTime()) / 3600000
      if (hoursUntil < 0 || hoursUntil > days * 24) return
      if (item.reminderSentAt) return
      const caseRow = casesById[item.caseId]
      if (!caseRow) return
      const recipients = String(item.responsibleEmails || caseRow.responsibleEmails || '')
        .split(',').map(function(value) { return value.trim() }).filter(Boolean)
      if (!recipients.length) return
      const subject = sheetName === 'Hearings'
        ? 'IKV հիշեցում․ դատական նիստ ' + caseRow.internalNumber
        : 'IKV հիշեցում․ դատավարական ժամկետ ' + caseRow.internalNumber
      const body = [
        'Գործ՝ ' + caseRow.applicantName,
        'Ներքին համար՝ ' + caseRow.internalNumber,
        'Ամսաթիվ՝ ' + (sheetName === 'Hearings' ? item.dateTime : item.dueDate),
        sheetName === 'Hearings' ? item.note : item.description
      ].filter(Boolean).join('\n')
      MailApp.sendEmail(recipients.join(','), subject, body)
      const rowNumber = items.findIndex(function(value) { return value.id === item.id }) + 2
      sheet.getRange(rowNumber, HEADERS[sheetName].indexOf('reminderSentAt') + 1)
        .setValue(new Date().toISOString())
      sent++
    })
  })
  return { status: 'ok', sent: sent }
}

function backupSpreadsheet() {
  const source = DriveApp.getFileById(SPREADSHEET_ID)
  const folderId = String(getSetting('BACKUP_FOLDER_ID', '') || '').trim()
  const folder = folderId ? DriveApp.getFolderById(folderId) : getOrCreateRootFolder()
  const timezone = getSpreadsheet().getSpreadsheetTimeZone() || Session.getScriptTimeZone()
  const stamp = Utilities.formatDate(new Date(), timezone, 'yyyy-MM-dd_HHmmss')
  const copy = source.makeCopy('IKV_Backup_' + stamp, folder)
  auditAction('BACKUP', 'Spreadsheet', SPREADSHEET_ID, { backupFileId: copy.getId() })
  return { status: 'ok', fileId: copy.getId(), fileUrl: copy.getUrl() }
}

function installMaintenanceTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (['sendUpcomingReminders','backupSpreadsheet'].indexOf(trigger.getHandlerFunction()) !== -1) {
      ScriptApp.deleteTrigger(trigger)
    }
  })
  ScriptApp.newTrigger('sendUpcomingReminders').timeBased().everyHours(1).create()
  ScriptApp.newTrigger('backupSpreadsheet').timeBased().everyDays(1).atHour(2).create()
  return 'Hourly reminders and daily backups are installed.'
}

function migrateExistingCases() {
  Object.keys(HEADERS).forEach(function(name) { getOrCreateSheet(name) })
  const sheet = getOrCreateSheet('Cases')
  const values = sheet.getDataRange().getDisplayValues()
  if (values.length < 2) return { status: 'ok', before: 0, after: 0, updated: 0 }
  backupSheetBeforeMigration(sheet, 'Cases_full_migration')
  const headers = values[0]
  const internalIndex = headers.indexOf('internalNumber')
  const typeIndex = headers.indexOf('caseType')
  const stageIndex = headers.indexOf('stage')
  const changedAtIndex = headers.indexOf('stageChangedAt')
  const changedByIndex = headers.indexOf('stageChangedBy')
  let updated = 0
  for (let i = 1; i < values.length; i++) {
    if (!values[i][0]) continue
    if (!values[i][internalIndex]) { values[i][internalIndex] = nextInternalNumber(); updated++ }
    if (!values[i][typeIndex]) { values[i][typeIndex] = 'bankruptcy'; updated++ }
    if (!values[i][stageIndex]) { values[i][stageIndex] = 'preparation'; updated++ }
    if (!values[i][changedAtIndex]) values[i][changedAtIndex] = new Date().toISOString()
    if (!values[i][changedByIndex]) values[i][changedByIndex] = currentUserEmail()
  }
  sheet.getRange(1, 1, values.length, headers.length).setValues(values)
  const count = values.filter(function(row, index) { return index > 0 && row[0] }).length
  auditAction('MIGRATION', 'Cases', 'all', { before: count, after: count, updatedCells: updated })
  return { status: 'ok', before: count, after: count, updated: updated }
}

function deleteCase(data) {
  const caseId = String(data.caseId || '')
  const confirmationName = String(data.confirmationName || '').trim()
  if (!caseId) throw new Error('Missing case ID')

  const caseRow = getRows('Cases').find(row => row.id === caseId)
  if (!caseRow) throw new Error('Case not found')
  if (!confirmationName || confirmationName !== String(caseRow.applicantName || '').trim()) {
    throw new Error('Client name confirmation does not match')
  }

  const lock = LockService.getScriptLock()
  lock.waitLock(30000)
  try {
    getRows('Hearings', 'caseId', caseId).forEach(function(item) {
      deleteCalendarEvent(item.calendarEventId)
    })
    getRows('Deadlines', 'caseId', caseId).forEach(function(item) {
      deleteCalendarEvent(item.calendarEventId)
    })
    let folderTrashed = false
    if (caseRow.driveFolderId) {
      try {
        const folder = DriveApp.getFolderById(caseRow.driveFolderId)
        if (!folder.isTrashed()) folder.setTrashed(true)
        folderTrashed = true
      } catch (err) {}
    }

    const deletedRows = {
      Documents: deleteRowsByValue('Documents', 'caseId', caseId),
      Debts: deleteRowsByValue('Debts', 'caseId', caseId),
      BankCertificates: deleteRowsByValue('BankCertificates', 'caseId', caseId),
      CaseFiles: deleteRowsByValue('CaseFiles', 'caseId', caseId),
      IncomingMail: deleteRowsByValue('IncomingMail', 'caseId', caseId),
      Hearings: deleteRowsByValue('Hearings', 'caseId', caseId),
      Deadlines: deleteRowsByValue('Deadlines', 'caseId', caseId),
      Cases: deleteRowsByValue('Cases', 'id', caseId)
    }
    const user = Session.getActiveUser().getEmail() || 'web-app-user'
    getOrCreateSheet('AuditLog').appendRow(validateRow('AuditLog', [
      Utilities.getUuid(), new Date().toISOString(), user, 'DELETE',
      'Case', caseId, JSON.stringify({
        applicantName: caseRow.applicantName,
        folderTrashed: folderTrashed,
        deletedRows: deletedRows
      })
    ]))
    return { status: 'deleted', caseId: caseId, folderTrashed: folderTrashed }
  } finally {
    lock.releaseLock()
  }
}

function deleteRowsByValue(sheetName, columnName, value) {
  const sheet = getOrCreateSheet(sheetName)
  const values = sheet.getDataRange().getDisplayValues()
  if (!values.length) return 0
  const columnIndex = values[0].indexOf(columnName)
  if (columnIndex === -1) throw new Error('Unknown column: ' + columnName)
  let deleted = 0
  for (let rowIndex = values.length - 1; rowIndex >= 1; rowIndex--) {
    if (values[rowIndex][columnIndex] === value) {
      sheet.deleteRow(rowIndex + 1)
      deleted++
    }
  }
  return deleted
}

/**
 * One-time, idempotent import for the second Inga/Vahagn case.
 * Run from the Apps Script editor after the original legacy import.
 */
function importIngaApplicantVahagnSpouseCase() {
  const source = getRows('Cases').find(function(row) {
    return row.applicantName === 'Վահագն Գարսեվանի Աբրահամյան' &&
      row.spouseName === 'Ինգա Արմենի Ամարյան'
  })
  if (!source) throw new Error('Source Vahagn/Inga case was not found')

  const targetId = 'legacy-inga-vahagn-reversed'
  const now = new Date().toISOString()
  const target = Object.assign({}, source, {
    id: targetId,
    applicantName: source.spouseName,
    firstName: source.spouseFirstName,
    lastName: source.spouseLastName,
    middleName: source.spouseMiddleName,
    birthDate: source.spouseBirthDate,
    passport: source.spousePassport,
    passportDate: source.spousePassportDate,
    passportBy: source.spousePassportBy,
    noPassport: source.spouseNoPassport,
    idCard: source.spouseIdCard,
    idCardDate: source.spouseIdCardDate,
    idCardBy: source.spouseIdCardBy,
    noIdCard: source.spouseNoIdCard,
    additionalIdentity: source.spouseAdditionalIdentity,
    regAddr: source.spouseRegAddr,
    notifAddr: source.spouseNotifAddr || source.spouseRegAddr,
    psn: source.spousePsn,
    spouseName: source.applicantName,
    spouseFirstName: source.firstName,
    spouseLastName: source.lastName,
    spouseMiddleName: source.middleName,
    spouseBirthDate: source.birthDate,
    spousePassport: source.passport,
    spousePassportDate: source.passportDate,
    spousePassportBy: source.passportBy,
    spouseNoPassport: source.noPassport,
    spouseIdCard: source.idCard,
    spouseIdCardDate: source.idCardDate,
    spouseIdCardBy: source.idCardBy,
    spouseNoIdCard: source.noIdCard,
    spouseAdditionalIdentity: source.additionalIdentity,
    spousePsn: source.psn,
    spouseRegAddr: source.regAddr,
    spouseNotifAddr: source.notifAddr || source.regAddr,
    createdAt: now,
    createdBy: 'legacy-reversed-import',
    assessmentJson: '',
    driveFolderId: '',
    driveFolderUrl: ''
  })

  const caseRow = HEADERS.Cases.map(function(header) {
    return target[header] || ''
  })
  saveCaseWithFolders(validateRow('Cases', caseRow))

  const copied = { documents: 0, bankCertificates: 0 }
  getRows('Documents', 'caseId', source.id).forEach(function(row) {
    const copy = Object.assign({}, row, {
      id: legacyCopyId('document', targetId, row.id),
      caseId: targetId,
      updatedAt: now
    })
    upsertRow('Documents', validateRow('Documents', HEADERS.Documents.map(function(header) {
      return copy[header] || ''
    })))
    copied.documents++
  })

  getRows('BankCertificates', 'caseId', source.id).forEach(function(row) {
    const copy = Object.assign({}, row, {
      id: legacyCopyId('bank', targetId, row.id),
      caseId: targetId,
      updatedAt: now
    })
    upsertRow('BankCertificates', validateRow('BankCertificates', HEADERS.BankCertificates.map(function(header) {
      return copy[header] || ''
    })))
    copied.bankCertificates++
  })

  return {
    status: 'ok',
    caseId: targetId,
    applicantName: target.applicantName,
    spouseName: target.spouseName,
    documents: copied.documents,
    bankCertificates: copied.bankCertificates
  }
}

function legacyCopyId(kind, caseId, sourceId) {
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    [kind, caseId, sourceId].join('|'),
    Utilities.Charset.UTF_8
  )
  return 'legacy-' + digest.slice(0, 12).map(function(byte) {
    return ('0' + ((byte + 256) % 256).toString(16)).slice(-2)
  }).join('')
}

function repairBankCertificateExpiryDates() {
  const sheet = getOrCreateSheet('BankCertificates')
  const values = sheet.getDataRange().getDisplayValues()
  if (values.length < 2) return { status: 'ok', repaired: 0 }

  const headers = values[0]
  const statusIndex = headers.indexOf('status')
  const issueDateIndex = headers.indexOf('issueDate')
  const expiryDateIndex = headers.indexOf('expiryDate')
  if ([statusIndex, issueDateIndex, expiryDateIndex].some(function(index) { return index === -1 })) {
    throw new Error('BankCertificates date columns were not found')
  }

  let repaired = 0
  for (let rowIndex = 1; rowIndex < values.length; rowIndex++) {
    const status = values[rowIndex][statusIndex]
    const issueDate = values[rowIndex][issueDateIndex]
    const expiryDate = values[rowIndex][expiryDateIndex]
    if (status === 'received' && issueDate && !expiryDate) {
      sheet.getRange(rowIndex + 1, expiryDateIndex + 1)
        .setValue(addDaysToIsoDate(issueDate, 90))
      repaired++
    }
  }
  return { status: 'ok', repaired: repaired }
}

function addDaysToIsoDate(dateText, days) {
  const parts = String(dateText || '').split('-').map(Number)
  if (parts.length !== 3 || parts.some(function(value) { return !value })) {
    throw new Error('Invalid date: ' + dateText)
  }
  const date = new Date(parts[0], parts[1] - 1, parts[2])
  date.setDate(date.getDate() + days)
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd')
}

function uploadDocument(data) {
  const caseId = String(data.caseId || '')
  const typeId = String(data.typeId || '')
  const subject = String(data.subject || '')
  const fileName = sanitizeDriveName(data.fileName || 'document')
  const mimeType = String(data.mimeType || 'application/octet-stream')
  const base64 = String(data.base64 || '')
  if (!caseId || !typeId || !subject || !base64) throw new Error('Missing upload data')

  const bytes = Utilities.base64Decode(base64)
  if (bytes.length > MAX_UPLOAD_BYTES) throw new Error('File exceeds the 8 MB upload limit')

  const caseRow = getRows('Cases').find(row => row.id === caseId)
  if (!caseRow) throw new Error('Case not found')
  const caseFolder = resolveCaseFolder(caseRow)
  const categoryFolder = getOrCreateChildFolder(caseFolder, DOCUMENT_FOLDERS[typeId] || '08 Other Documents')
  let subjectFolder = categoryFolder
  const folderPath = Array.isArray(data.folderPath) && data.folderPath.length
    ? data.folderPath
    : [data.folderName || subject]
  folderPath.forEach(function(folderName) {
    subjectFolder = getOrCreateChildFolder(subjectFolder, sanitizeDriveName(folderName))
  })
  const file = subjectFolder.createFile(Utilities.newBlob(bytes, mimeType, fileName))

  const documentId = String(data.documentId || '')
  const fileRow = [
    Utilities.getUuid(), caseId, documentId, typeId, subject, fileName, mimeType,
    file.getId(), file.getUrl(), subjectFolder.getId(), new Date().toISOString()
  ]
  appendRow('CaseFiles', validateRow('CaseFiles', fileRow))
  return {
    status: 'uploaded',
    fileId: file.getId(),
    fileUrl: file.getUrl(),
    fileName: fileName,
    folderId: subjectFolder.getId()
  }
}

function generateStateRequest(data) {
  const caseId = String(data.caseId || '')
  const typeId = String(data.typeId || '')
  const subject = String(data.subject || '')
  if (!caseId || !typeId || !subject) throw new Error('Missing case, request type, or subject')

  const template = STATE_REQUEST_TEMPLATES[typeId]
  if (!template) throw new Error('No request template is configured for: ' + typeId)

  const caseRow = getRows('Cases').find(row => row.id === caseId)
  if (!caseRow) throw new Error('Case not found')

  const person = getCaseRequestPerson(caseRow, subject)
  const representative = getRequestRepresentative(data.representative || data.representativeKey || '')
  validateRequestPerson(person)

  const caseFolder = resolveCaseFolder(caseRow)
  const generatedFolder = getOrCreateChildFolder(caseFolder, '07 Generated Documents')
  const requestFolder = getOrCreateChildFolder(generatedFolder, 'State Body Requests')
  const templateFolder = getOrCreateChildFolder(requestFolder, sanitizeDriveName(template.title.join(' ')))
  const subjectFolder = getOrCreateChildFolder(templateFolder, sanitizeDriveName(subject))
  const outputName = sanitizeDriveName(
    'Հարցում - ' + template.title.join(' ') + ' - ' + person.fullName + ' - ' +
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd')
  )

  const document = createStateRequestDocument(outputName, subjectFolder, template, person, representative)
  const file = DriveApp.getFileById(document.getId())
  const fileRow = [
    Utilities.getUuid(), caseId, '', typeId, subject, outputName,
    file.getMimeType(), file.getId(), file.getUrl(), subjectFolder.getId(), new Date().toISOString()
  ]
  appendRow('CaseFiles', validateRow('CaseFiles', fileRow))
  return {
    status: 'generated',
    fileId: file.getId(),
    fileUrl: file.getUrl(),
    fileName: outputName,
    folderId: subjectFolder.getId()
  }
}

function generatePoa(data) {
  const caseId = String(data.caseId || '')
  const typeId = String(data.typeId || '')
  const subject = String(data.subject || '')
  if (!caseId || !typeId || !subject) throw new Error('Missing case, POA type, or subject')

  const template = POA_TEMPLATES[typeId]
  if (!template) throw new Error('No POA template is configured for: ' + typeId)

  const caseRow = getRows('Cases').find(row => row.id === caseId)
  if (!caseRow) throw new Error('Case not found')

  const person = getCaseRequestPerson(caseRow, subject)
  validateRequestPerson(person)

  const caseFolder = resolveCaseFolder(caseRow)
  const generatedFolder = getOrCreateChildFolder(caseFolder, '07 Generated Documents')
  const poaFolder = getOrCreateChildFolder(generatedFolder, 'Powers of Attorney')
  const subjectFolder = getOrCreateChildFolder(poaFolder, sanitizeDriveName(subject))
  const outputName = sanitizeDriveName(
    (typeId === 'notarial_poa' ? 'Նոտարական լիազորագիր - ' : 'Դատավարական լիազորագիր - ') +
    person.fullName + ' - ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd')
  )

  const document = createPoaDocument(outputName, subjectFolder, template, person)
  const file = DriveApp.getFileById(document.getId())
  const fileRow = [
    Utilities.getUuid(), caseId, '', typeId, subject, outputName,
    file.getMimeType(), file.getId(), file.getUrl(), subjectFolder.getId(), new Date().toISOString()
  ]
  appendRow('CaseFiles', validateRow('CaseFiles', fileRow))
  return {
    status: 'generated',
    fileId: file.getId(),
    fileUrl: file.getUrl(),
    fileName: outputName,
    folderId: subjectFolder.getId()
  }
}

function createPoaDocument(outputName, folder, template, person) {
  const document = DocumentApp.create(outputName)
  const file = DriveApp.getFileById(document.getId())
  file.moveTo(folder)
  const body = document.getBody()
  body.clear()

  const center = DocumentApp.HorizontalAlignment.CENTER
  const justified = DocumentApp.HorizontalAlignment.JUSTIFY
  addPoaParagraph(body, template.title, center, true)
  addPoaParagraph(body, '')
  addPoaParagraph(body, formatPoaDateText(new Date()), center)
  template.paragraphs.forEach(text => {
    addPoaParagraph(body, replacePoaPlaceholders(text, person), justified)
  })
  addPoaParagraph(body, '')
  addPoaParagraph(body, '_________________________________________________________________________', center)
  addPoaParagraph(body, '(Լիազորողի, անուն, ազգանուն, ստորագրություն)', center)

  document.saveAndClose()
  return document
}

function addPoaParagraph(body, text, alignment, bold) {
  const paragraph = body.appendParagraph(text || '')
  if (alignment) paragraph.setAlignment(alignment)
  const editable = paragraph.editAsText()
  editable.setFontFamily('Merriweather')
  editable.setFontSize(12)
  editable.setBold(!!bold)
  return paragraph
}

function replacePoaPlaceholders(text, person) {
  return String(text || '')
    .split('{{client.fullNameWithSuffix}}').join(person.fullName + 'ս')
    .split('{{client.identityWithPsnAndAddress}}').join(person.identityWithPsnAndAddress)
    .split('{{poa.representatives}}').join(POA_REPRESENTATIVES_TEXT)
}

function formatPoaDateText(date) {
  const years = {
    2024: 'երկու հազար քսանչորս',
    2025: 'երկու հազար քսանհինգ',
    2026: 'երկու հազար քսանվեց',
    2027: 'երկու հազար քսանյոթ',
    2028: 'երկու հազար քսանութ'
  }
  const months = [
    'հունվարի','փետրվարի','մարտի','ապրիլի','մայիսի','հունիսի',
    'հուլիսի','օգոստոսի','սեպտեմբերի','հոկտեմբերի','նոյեմբերի','դեկտեմբերի'
  ]
  const days = {
    1:'մեկ',2:'երկուս',3:'երեք',4:'չորս',5:'հինգ',6:'վեց',7:'յոթ',8:'ութ',9:'ինը',10:'տաս',
    11:'տասնմեկ',12:'տասներկուս',13:'տասներեք',14:'տասնչորս',15:'տասնհինգ',16:'տասնվեց',
    17:'տասնյոթ',18:'տասնութ',19:'տասնինը',20:'քսան',21:'քսանմեկ',22:'քսաներկուս',
    23:'քսաներեք',24:'քսանչորս',25:'քսանհինգ',26:'քսանվեց',27:'քսանյոթ',
    28:'քսանութ',29:'քսանինը',30:'երեսուն',31:'երեսունմեկ'
  }
  return (years[date.getFullYear()] || String(date.getFullYear())) + ' թվականի ' +
    months[date.getMonth()] + ' ' + (days[date.getDate()] || date.getDate())
}

function getCaseRequestPerson(caseRow, subject) {
  const isApplicant = subject === caseRow.applicantName
  const person = isApplicant ? {
    firstName: caseRow.firstName,
    middleName: caseRow.middleName,
    lastName: caseRow.lastName,
    identityNumber: caseRow.passport,
    identityIssueDate: caseRow.passportDate,
    identityIssuer: caseRow.passportBy,
    noPassport: caseRow.noPassport,
    idCard: caseRow.idCard,
    idCardDate: caseRow.idCardDate,
    idCardBy: caseRow.idCardBy,
    noIdCard: caseRow.noIdCard,
    registrationAddress: caseRow.regAddr,
    psn: caseRow.psn,
    additionalIdentity: caseRow.additionalIdentity
  } : {
    firstName: caseRow.spouseFirstName,
    middleName: caseRow.spouseMiddleName,
    lastName: caseRow.spouseLastName,
    identityNumber: caseRow.spousePassport,
    identityIssueDate: caseRow.spousePassportDate,
    identityIssuer: caseRow.spousePassportBy,
    noPassport: caseRow.spouseNoPassport,
    idCard: caseRow.spouseIdCard,
    idCardDate: caseRow.spouseIdCardDate,
    idCardBy: caseRow.spouseIdCardBy,
    noIdCard: caseRow.spouseNoIdCard,
    registrationAddress: caseRow.spouseRegAddr,
    psn: caseRow.spousePsn,
    additionalIdentity: caseRow.spouseAdditionalIdentity
  }
  person.fullName = [person.firstName, person.middleName, person.lastName].filter(Boolean).join(' ') || subject
  person.fullNameGenitive = person.fullName + '-ի'
  person.identityNumber = stripSheetTextPrefix(person.identityNumber)
  person.idCard = stripSheetTextPrefix(person.idCard)
  person.identityIssuer = stripSheetTextPrefix(person.identityIssuer)
  person.idCardBy = stripSheetTextPrefix(person.idCardBy)
  person.psn = stripSheetTextPrefix(person.psn)
  person.identityIssueDate = formatRequestDate(person.identityIssueDate)
  person.idCardDate = formatRequestDate(person.idCardDate)
  const identityParts = getPersonIdentityParts(person)
  person.identityInline = '/' + identityParts.join(', ') + '/'
  person.identityWithPsnAndAddress = '/' + identityParts.concat([
    'ՀԾՀ՝ ' + person.psn,
    'հաշվառված ' + person.registrationAddress + ' հասցեում'
  ]).join(', ') + '/'
  person.requestBodyIdentity = '/' + identityParts.concat([
    'ՀԾՀ՝ ' + person.psn,
    'հաշվառման հասցե՝ ' + person.registrationAddress
  ]).join(', ') + '/'
  return person
}

function validateRequestPerson(person) {
  const identityParts = getPersonIdentityParts(person)
  const required = [
    ['full name', person.fullName],
    ['passport or identification card', identityParts.length ? 'YES' : ''],
    ['registration address', person.registrationAddress],
    ['public services number', person.psn]
  ]
  const missing = required.filter(item => !item[1]).map(item => item[0])
  if (missing.length) throw new Error('Missing client data: ' + missing.join(', '))
}

function getPersonIdentityParts(person) {
  const parts = []
  const hasPassport = String(person.noPassport || '').toUpperCase() !== 'YES' &&
    (person.identityNumber || person.identityIssueDate || person.identityIssuer)
  if (hasPassport) {
    parts.push('անձնագիր՝ ' + (person.identityNumber || '') +
      (person.identityIssueDate ? ', տրված՝ ' + person.identityIssueDate + 'թ.' : '') +
      (person.identityIssuer ? ' ' + person.identityIssuer + '-ի կողմից' : ''))
  }
  const hasIdCard = String(person.noIdCard || '').toUpperCase() !== 'YES' &&
    (person.idCard || person.idCardDate || person.idCardBy)
  if (hasIdCard) {
    parts.push('նույնականացման քարտ՝ ' + (person.idCard || '') +
      (person.idCardDate ? ', տրված՝ ' + person.idCardDate + 'թ.' : '') +
      (person.idCardBy ? ' ' + person.idCardBy + '-ի կողմից' : ''))
  }
  return parts.filter(part => !/չկա/i.test(part))
}

function stripSheetTextPrefix(value) {
  return String(value || '').replace(/^'/, '')
}

function getRequestRepresentative(input) {
  if (!input) return REPRESENTATIVE
  if (typeof input === 'string') return REPRESENTATIVES[input] || REPRESENTATIVE
  const key = String(input.key || '')
  const base = REPRESENTATIVES[key] || REPRESENTATIVE
  const merged = Object.assign({}, base, input)
  merged.identityType = merged.identityType || 'նույնականացման քարտ'
  merged.fullNameFrom = merged.fullNameFrom || merged.signatureName || base.fullNameFrom
  merged.signatureName = merged.signatureName || merged.fullNameFrom || base.signatureName
  merged.notificationAddress = merged.notificationAddress || base.notificationAddress
  merged.phone = merged.phone || base.phone
  return merged
}

function createStateRequestDocument(outputName, folder, template, person, representative) {
  const document = DocumentApp.create(outputName)
  const file = DriveApp.getFileById(document.getId())
  file.moveTo(folder)
  const body = document.getBody()
  body.clear()

  const right = DocumentApp.HorizontalAlignment.RIGHT
  const center = DocumentApp.HorizontalAlignment.CENTER
  const justified = DocumentApp.HorizontalAlignment.JUSTIFY

  template.title.forEach(line => addRequestParagraph(body, line, right, true))
  addRequestParagraph(body, '', right)
  addRequestParagraph(body, person.fullName, right, true)
  addRequestParagraph(body, person.identityInline, right)
  addRequestParagraph(body, 'հաշվառված՝ ' + person.registrationAddress + ' հասցեում', right)
  addRequestParagraph(body, '', right)
  addRequestParagraph(body, 'ներկայացուցիչ՝ ' + representative.fullNameFrom, right)
  addRequestParagraph(body, '(' + representative.identityType + '՝ ' + representative.identityNumber +
    ', տրված ' + formatRequestDate(representative.identityIssueDate) + 'թ. ' +
    representative.identityIssuer + '-ի կողմից', right)
  if (representative.licenseNumber) {
    addRequestParagraph(body, 'փաստաբանական գործունեության արտոնագիր թիվ ' + representative.licenseNumber + ')', right)
  } else {
    addRequestParagraph(body, ')', right)
  }
  addRequestParagraph(body, 'ԾԱՆՈՒՑՄԱՆ ՀԱՍՑԵ՝ ' + representative.notificationAddress, right, true)
  addRequestParagraph(body, 'հեռ. ' + representative.phone, right)
  addRequestParagraph(body, '')
  addRequestParagraph(body, template.heading || 'ՀԱՐՑՈՒՄ', center, true)
  addRequestParagraph(body, '')

  template.paragraphs.forEach(text => {
    addRequestParagraph(body, replaceRequestPlaceholders(text, person, representative), justified)
  })
  addRequestParagraph(body, '')
  addRequestParagraph(body, 'Կից ներկայացնում եմ')
  ;(template.attachments || defaultRequestAttachments(representative)).forEach(item => {
    const listItem = body.appendListItem(item)
    formatRequestText(listItem.editAsText())
  })
  addRequestParagraph(body, '')
  addRequestParagraph(body, 'Հարգանքներով՝')
  addRequestParagraph(body, '')
  addRequestParagraph(body, 'ներկայացուցիչ՝ ' + representative.signatureName + '____________')

  document.saveAndClose()
  return document
}

function defaultRequestAttachments(representative) {
  const attachments = ['Անձը հաստատող փաստաթղթի սկանը', 'Ինձ տրված լիազորագրի սկանը']
  if (representative.licenseNumber) attachments.push('Փաստաբանական գործունեության արտոնագրի սկանը')
  return attachments
}

function addRequestParagraph(body, text, alignment, bold, italic) {
  const paragraph = body.appendParagraph(text || '')
  if (alignment) paragraph.setAlignment(alignment)
  const editable = paragraph.editAsText()
  formatRequestText(editable)
  if (bold && text) editable.setBold(true)
  if (italic && text) editable.setItalic(true)
  return paragraph
}

function formatRequestText(text) {
  text.setFontFamily('Merriweather')
  text.setFontSize(12)
  text.setBold(false)
  text.setItalic(false)
}

function replaceRequestPlaceholders(text, person, representative) {
  const values = {
    '{{client.fullName}}': person.fullName,
    '{{client.fullNameGenitive}}': person.fullNameGenitive,
    '{{client.identityInline}}': person.identityInline,
    '{{client.requestBodyIdentity}}': person.requestBodyIdentity,
    '{{client.identityNumber}}': person.identityNumber,
    '{{client.identityIssueDate}}': person.identityIssueDate,
    '{{client.identityIssuer}}': person.identityIssuer,
    '{{client.registrationAddress}}': person.registrationAddress,
    '{{client.psn}}': person.psn,
    '{{representative.fullNameFrom}}': representative.fullNameFrom,
    '{{representative.notificationAddress}}': representative.notificationAddress,
    '{{representative.phone}}': representative.phone,
    '{{representative.signatureName}}': representative.signatureName
  }
  let result = String(text || '')
  Object.keys(values).forEach(key => {
    result = result.split(key).join(String(values[key] || ''))
  })
  return result
}

function generateRegisterRequest(data) {
  const caseId = String(data.caseId || '')
  const subject = String(data.subject || '')
  if (!caseId || !subject) throw new Error('Missing case or request subject')

  const caseRow = getRows('Cases').find(row => row.id === caseId)
  if (!caseRow) throw new Error('Case not found')

  const isApplicant = subject === caseRow.applicantName
  const person = isApplicant ? {
    firstName: caseRow.firstName,
    middleName: caseRow.middleName,
    lastName: caseRow.lastName,
    identityNumber: caseRow.passport,
    identityIssueDate: caseRow.passportDate,
    identityIssuer: caseRow.passportBy,
    registrationAddress: caseRow.regAddr,
    psn: caseRow.psn,
    additionalIdentity: caseRow.additionalIdentity
  } : {
    firstName: caseRow.spouseFirstName,
    middleName: caseRow.spouseMiddleName,
    lastName: caseRow.spouseLastName,
    identityNumber: caseRow.spousePassport,
    identityIssueDate: caseRow.spousePassportDate,
    identityIssuer: caseRow.spousePassportBy,
    registrationAddress: caseRow.spouseRegAddr,
    psn: caseRow.spousePsn,
    additionalIdentity: caseRow.spouseAdditionalIdentity
  }

  const fullName = [person.firstName, person.middleName, person.lastName].filter(Boolean).join(' ') || subject
  const representative = getRequestRepresentative(data.representative || data.representativeKey || '')
  const required = [
    ['full name', fullName],
    ['identity document number', person.identityNumber],
    ['identity document issue date', person.identityIssueDate],
    ['identity document issuer', person.identityIssuer],
    ['registration address', person.registrationAddress],
    ['public services number', person.psn]
  ]
  const missing = required.filter(item => !item[1]).map(item => item[0])
  if (missing.length) throw new Error('Missing client data: ' + missing.join(', '))

  const caseFolder = resolveCaseFolder(caseRow)
  const generatedFolder = getOrCreateChildFolder(caseFolder, '07 Generated Documents')
  const requestFolder = getOrCreateChildFolder(generatedFolder, 'State Register Requests')
  const subjectFolder = getOrCreateChildFolder(requestFolder, sanitizeDriveName(subject))
  const outputName = sanitizeDriveName(
    'Հարցում - Պետական ռեգիստր - ' + fullName + ' - ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd')
  )
  let file
  let document
  try {
    file = DriveApp.getFileById(REGISTER_TEMPLATE_DOC_ID).makeCopy(outputName, subjectFolder)
    document = DocumentApp.openById(file.getId())
  } catch (err) {
    document = createRegisterRequestFallback(outputName, subjectFolder)
    file = DriveApp.getFileById(document.getId())
  }
  const body = document.getBody()

  const replacements = {
    '{{client.fullNameGenitive}}': fullName + '-ի',
    '{{client.identityNumber}}': person.identityNumber,
    '{{client.identityIssueDate}}': formatRequestDate(person.identityIssueDate),
    '{{client.identityIssuer}}': person.identityIssuer,
    '{{client.registrationAddress}}': person.registrationAddress,
    '{{client.psn}}': person.psn,
    '{{representative.fullNameFrom}}': representative.fullNameFrom,
    '{{representative.identityNumber}}': representative.identityNumber,
    '{{representative.identityIssueDate}}': formatRequestDate(representative.identityIssueDate),
    '{{representative.identityIssuer}}': representative.identityIssuer,
    '{{representative.licenseNumber}}': representative.licenseNumber,
    '{{representative.notificationAddress}}': representative.notificationAddress,
    '{{representative.phone}}': representative.phone,
    '{{representative.signatureName}}': representative.signatureName
  }
  Object.keys(replacements).forEach(placeholder => {
    body.replaceText(escapeRegex(placeholder), String(replacements[placeholder] || ''))
  })
  normalizeRegisterRequestFormatting(body)
  document.saveAndClose()

  const fileRow = [
    Utilities.getUuid(), caseId, '', 'state_reg', subject, outputName,
    file.getMimeType(), file.getId(), file.getUrl(), subjectFolder.getId(), new Date().toISOString()
  ]
  appendRow('CaseFiles', validateRow('CaseFiles', fileRow))
  return {
    status: 'generated',
    fileId: file.getId(),
    fileUrl: file.getUrl(),
    fileName: outputName,
    folderId: subjectFolder.getId()
  }
}

function createRegisterRequestFallback(outputName, folder) {
  const document = DocumentApp.create(outputName)
  const file = DriveApp.getFileById(document.getId())
  file.moveTo(folder)
  const body = document.getBody()
  body.clear()

  const right = DocumentApp.HorizontalAlignment.RIGHT
  const center = DocumentApp.HorizontalAlignment.CENTER
  const justified = DocumentApp.HorizontalAlignment.JUSTIFY
  const add = function(text, alignment, bold, italic) {
    const paragraph = body.appendParagraph(text)
    if (alignment) paragraph.setAlignment(alignment)
    const editable = paragraph.editAsText()
    editable.setFontFamily('Merriweather').setFontSize(12)
    if (bold) editable.setBold(true)
    if (italic) editable.setItalic(true)
    return paragraph
  }

  add('ՀՀ ԱՆ ԻՐԱՎԱԲԱՆԱԿԱՆ ԱՆՁԱՆՑ', right, true)
  add('ՊԵՏԱԿԱՆ ՌԵԳԻՍՏՐԻ ԳՈՐԾԱԿԱԼՈՒԹՅՈՒՆ', right, true)
  add('', right)
  add('{{client.fullNameGenitive}}', right, true)
  add('/նույնականացման քարտ՝ {{client.identityNumber}}, տրված՝ {{client.identityIssueDate}}թ. {{client.identityIssuer}}-ի կողմից', right)
  add('Հաշվառման հասցե՝ {{client.registrationAddress}}/', right)
  add('', right)
  add('ներկայացուցիչ՝ {{representative.fullNameFrom}}', right)
  add('(նույնականացման քարտ {{representative.identityNumber}}, տրված {{representative.identityIssueDate}}թ. {{representative.identityIssuer}}-ի կողմից', right)
  add('փաստաբանական գործունեության արտոնագիր թիվ {{representative.licenseNumber}})', right)
  add('ԾԱՆՈՒՑՄԱՆ ՀԱՍՑԵ՝ {{representative.notificationAddress}}', right, true)
  add('հեռ. {{representative.phone}}', right)
  add('')
  add('ՀԱՐՑՈՒՄ', center, true)
  add('')
  add('Հարգելի գործընկերներ,')
  add('')
  add(
    'Ղեկավարվելով «Տեղեկատվության ազատության մասին» ՀՀ օրենքի 6-րդ և 9-րդ «Վարչարարության հիմունքների և վարչական վարույթի մասին» ՀՀ օրենքի 30-րդ, 36-րդ, 39-րդ, 40-րդ և «Փաստաբանության մասին» ՀՀ օրենքի 18-րդ հոդվածներով՝ խնդրում եմ տրամադրել տեղեկատվություն, այն մասին, թե ՀՀ ԱՆ իրավաբանական անձանց պետական ռեգիստրի գործակալության տվյալների շտեմարանում սույն հարցումը ներկայացնելու պահի դրությամբ {{client.fullNameGenitive}} /նույնականացման քարտ՝ {{client.identityNumber}}, տրված {{client.identityIssueDate}}թ. {{client.identityIssuer}}-ի կողմից, ՀԾՀ՝ {{client.psn}}/ վերաբերյալ',
    justified
  )
  add('')
  add('- Արդյոք առկան են գույքի, այլ իրավունքների առկայության/բացակայության վերաբերյալ տեղեկություններ:', null, false, true)
  add('')
  add('- Արդյոք հաշվառված է որպես անհատ ձեռնարկատեր:', null, false, true)
  add('')
  add('Կից ներկայացնում եմ')
  add('- Անձը հաստատող փաստաթղթի սկանը')
  add('- Ինձ տրված լիազորագրի սկանը')
  add('- Փաստաբանական գործունեության արտոնագրի սկանը')
  add('')
  add('Հարգանքներով՝')
  add('')
  add('ներկայացուցիչ՝ {{representative.signatureName}}____________')
  return document
}

function normalizeRegisterRequestFormatting(body) {
  const text = body.editAsText()
  text.setBold(false)
  text.setItalic(false)
  text.setFontFamily('Merriweather')
  text.setFontSize(12)

  setTextBold(body, 'ՀՀ ԱՆ ԻՐԱՎԱԲԱՆԱԿԱՆ ԱՆՁԱՆՑ')
  setTextBold(body, 'ՊԵՏԱԿԱՆ ՌԵԳԻՍՏՐԻ ԳՈՐԾԱԿԱԼՈՒԹՅՈՒՆ')
  setTextBold(body, 'ԾԱՆՈՒՑՄԱՆ ՀԱՍՑԵ՝')
  setTextBold(body, 'ՀԱՐՑՈՒՄ')
  setTextItalic(body, '- Արդյոք առկան են գույքի, այլ իրավունքների առկայության/բացակայության վերաբերյալ տեղեկություններ:')
  setTextItalic(body, '- Արդյոք հաշվառված է որպես անհատ ձեռնարկատեր:')
}

function setTextBold(body, value) {
  const found = body.findText(escapeRegex(value))
  if (!found) return
  const element = found.getElement().asText()
  element.setBold(found.getStartOffset(), found.getEndOffsetInclusive(), true)
}

function setTextItalic(body, value) {
  const found = body.findText(escapeRegex(value))
  if (!found) return
  const element = found.getElement().asText()
  element.setItalic(found.getStartOffset(), found.getEndOffsetInclusive(), true)
}

function formatRequestDate(value) {
  const match = String(value || '').match(/^(\d{4})-(\d{2})-(\d{2})$/)
  return match ? match[3] + '.' + match[2] + '.' + match[1] : String(value || '')
}

function escapeRegex(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

function sanitizeDriveName(value) {
  return String(value || '')
    .replace(/[\\/:*?"<>|#%{}[\]]/g, '-')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 180) || 'Untitled'
}

