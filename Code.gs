// ================================================================
// ԱՅ ՔԻ ՎԻ — Bankruptcy Case Management — Apps Script Backend
// ================================================================
// SETUP:
// 1. Go to script.google.com → New project → paste this code
// 2. Deploy as Web App (Execute as: Me, Who has access: Anyone)
// 3. Copy the Web App URL into ikv_bankruptcy.html CONFIG.SCRIPT_URL
// ================================================================

const SPREADSHEET_ID = '1EdjhuXFFzNjxtWVPrcxueKJ5frS4qkxZnb7QqttQBxs'
const ROOT_FOLDER_NAME = 'IKV Bankruptcy Cases'
const MAX_UPLOAD_BYTES = 8 * 1024 * 1024
const REGISTER_TEMPLATE_DOC_ID = '1CfXZsHPy25vPk5i-KcxFvoggtefjYexQhrvhtptUtVE'

const REPRESENTATIVE = {
  fullNameFrom: 'Հովհաննես Գրիգորի Հարությունյանից',
  identityNumber: '009987561',
  identityIssueDate: '20.02.2018',
  identityIssuer: '011',
  licenseNumber: '2332',
  notificationAddress: 'ք. Երևան, Այգեստան 7-րդ փողոց, 1-ին շենք, բն. 30',
  phone: '099-648-655',
  signatureName: 'Հովհաննես Հարությունյան'
}

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
  Cases: ['id','applicantName','firstName','lastName','middleName','birthDate','passport','passportDate','passportBy','regAddr','notifAddr','psn','isMarried','spouseName','filingDate','employed','salaryAbove','banks','debts','spouseBanks','spouseDebts','lawyerApproved','lawyerApprovedBy','lawyerApprovedAt','createdAt','createdBy','spouseFirstName','spouseLastName','spouseMiddleName','spouseBirthDate','spousePassport','spousePassportDate','spousePassportBy','spousePsn','spouseRegAddr','spouseNotifAddr','assessmentJson','driveFolderId','driveFolderUrl'],
  Documents: ['id','caseId','typeId','subject','status','issueDate','expiryDate','appliedAt','updatedAt'],
  Debts: ['id','caseId','subject','creditor','contractNumber','contractDate','currency','principal','interest','penalty','totalAmount','dueDate','claimBasis','collateral','enforcementInfo','notes','createdAt','updatedAt'],
  BankCertificates: ['id','caseId','subject','bank','status','result','accountInfo','balance','currency','issueDate','expiryDate','appliedAt','notes','updatedAt'],
  CaseFiles: ['id','caseId','documentId','typeId','subject','fileName','mimeType','fileId','fileUrl','folderId','uploadedAt'],
  AuditLog: ['id','timestamp','user','action','entity','entityId','detail']
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
  return sheet
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
}

function repairAllSheetSchemas() {
  Object.keys(HEADERS).forEach(name => getOrCreateSheet(name))
  return 'All sheet schemas are ready'
}

function doGet(e) {
  const action = e && e.parameter ? e.parameter.action : ''
  try {
    let result
    if (action === 'getCases') result = getRows('Cases')
    else if (action === 'getDashboardData') result = getDashboardData()
    else if (action === 'getCaseBundle') result = getCaseBundle(e.parameter.caseId)
    else if (action === 'getDocs') result = getRows('Documents', 'caseId', e.parameter.caseId)
    else if (action === 'getDebts') result = getRows('Debts', 'caseId', e.parameter.caseId)
    else if (action === 'getBankCertificates') result = getRows('BankCertificates', 'caseId', e.parameter.caseId)
    else if (action === 'getFiles') result = getRows('CaseFiles', 'caseId', e.parameter.caseId)
    else if (action === 'getAudit') result = getRows('AuditLog').slice(-100).reverse()
    else if (action === 'ping') result = { status: 'ok', time: new Date().toISOString() }
    else result = { error: 'Unknown action: ' + action }
    return jsonResponse(result)
  } catch(err) {
    return jsonResponse({ error: err.message })
  }
}

function doPost(e) {
  try {
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
    else if (action === 'uploadDocument') result = uploadDocument(data)
    else if (action === 'generateRegisterRequest') result = generateRegisterRequest(data)
    else if (action === 'deleteCase') result = deleteCase(data)
    else if (action === 'importIngaApplicantVahagnSpouseCase') result = importIngaApplicantVahagnSpouseCase()
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
    bankCertificates: getRows('BankCertificates')
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
    files: getRows('CaseFiles', 'caseId', caseId)
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
  if (!normalized[0]) throw new Error(sheetName + ' row is missing an ID')
  return normalized
}

function safeCellValue(value) {
  if (value === undefined || value === null) return ''
  const text = String(value).trim()
  return /^[=+\-@]/.test(text) ? "'" + text : text
}

function saveCaseWithFolders(row) {
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
  return saved
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
    psn: caseRow.psn
  } : {
    firstName: caseRow.spouseFirstName,
    middleName: caseRow.spouseMiddleName,
    lastName: caseRow.spouseLastName,
    identityNumber: caseRow.spousePassport,
    identityIssueDate: caseRow.spousePassportDate,
    identityIssuer: caseRow.spousePassportBy,
    registrationAddress: caseRow.spouseRegAddr,
    psn: caseRow.spousePsn
  }

  const fullName = [person.firstName, person.middleName, person.lastName].filter(Boolean).join(' ') || subject
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
    '{{representative.fullNameFrom}}': REPRESENTATIVE.fullNameFrom,
    '{{representative.identityNumber}}': REPRESENTATIVE.identityNumber,
    '{{representative.identityIssueDate}}': REPRESENTATIVE.identityIssueDate,
    '{{representative.identityIssuer}}': REPRESENTATIVE.identityIssuer,
    '{{representative.licenseNumber}}': REPRESENTATIVE.licenseNumber,
    '{{representative.notificationAddress}}': REPRESENTATIVE.notificationAddress,
    '{{representative.phone}}': REPRESENTATIVE.phone,
    '{{representative.signatureName}}': REPRESENTATIVE.signatureName
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
