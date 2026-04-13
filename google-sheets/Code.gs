const ADMIN_ACCESS_CODE = 'CHANGE_ME_BEFORE_PRODUCTION';
const ADMIN_EDITOR_EMAILS = [];

const SHEETS = {
  cases: 'Cases',
  documents: 'CaseDocuments',
  districts: 'Districts',
  campuses: 'Campuses',
  evaluators: 'Evaluators',
  calendars: 'DistrictCalendars',
  dashboard: 'Dashboard',
};

const CASE_TYPES = {
  initial: 'Initial',
  reevaluation: 'Re-evaluation',
};

const STATUSES = {
  referralReceived: 'Referral Received',
  inProgress: 'In Progress',
  complete: 'Complete',
};

const CASE_HEADERS = [
  'CaseID',
  'CaseType',
  'StudentName',
  'StudentID',
  'DOB',
  'Campus',
  'District',
  'LeadEvaluator',
  'Status',
  'ReferralReceivedDate',
  'ResponseDueDate',
  'ProjectedConsentDate',
  'ActualConsentDate',
  'ProjectedFIIEDueDate',
  'ActualFIIEDate',
  'ProjectedARDDueDate',
  'ActualARDDate',
  'ReevalDueDate',
  'Service_SchoolPsychologist',
  'Service_OccupationalTherapist',
  'Service_PhysicalTherapist',
  'Service_CounselingEvaluation',
  'Service_FBA',
  'Service_SpeechPathologist',
  'Service_VI',
  'Service_DHH',
  'Service_LanguageDominanceBilingual',
  'ServiceNotes',
  'ManualPrimaryDeadline',
  'ManualOverrideReason',
  'PrimaryDeadline',
  'CreatedAt',
  'UpdatedAt',
];

const DOCUMENT_HEADERS = ['DocumentID', 'CaseID', 'DocumentLabel', 'DocumentPath', 'AddedAt'];
const DISTRICT_HEADERS = ['District', 'ResponseSchoolDays', 'FIIESchoolDays', 'ARDCalendarDays', 'Active'];
const CAMPUS_HEADERS = ['Campus', 'District', 'Active'];
const EVALUATOR_HEADERS = ['LeadEvaluator', 'Email', 'Active'];
const CALENDAR_BASE_HEADERS = ['Date', 'Weekday'];

const SERVICE_FIELDS = [
  'Service_SchoolPsychologist',
  'Service_OccupationalTherapist',
  'Service_PhysicalTherapist',
  'Service_CounselingEvaluation',
  'Service_FBA',
  'Service_SpeechPathologist',
  'Service_VI',
  'Service_DHH',
  'Service_LanguageDominanceBilingual',
];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('SPED Status Reports')
    .addItem('Open App', 'showSpedSidebar')
    .addItem('Install / Repair Workbook', 'installSpedStatusReports')
    .addSeparator()
    .addItem('Refresh Dashboard', 'refreshDashboard')
    .addItem('Open Dashboard', 'openDashboard')
    .addSeparator()
    .addItem('Open Districts (Admin)', 'openDistrictsSheet')
    .addItem('Open Campuses (Admin)', 'openCampusesSheet')
    .addItem('Open Evaluators (Admin)', 'openEvaluatorsSheet')
    .addItem('Open Calendars (Admin)', 'openCalendarsSheet')
    .addItem('Sync Calendar Grid', 'syncDistrictCalendarGrid')
    .addItem('Reapply Admin Sheet Protection', 'reapplyAdminSheetProtection')
    .addToUi();
}

function installSpedStatusReports() {
  ensureWorkbookScaffold_({
    seedReferenceData: true,
    syncCalendar: true,
    showSheets: true,
    applyProtection: true,
  });
  refreshDashboard();
  SpreadsheetApp.getUi().alert(
    'SPED Status Reports is ready. Sheets are visible, and direct edits are restricted to configured admin accounts. Update the sample district, campus, evaluator, calendar, and admin settings before production use.'
  );
}

function syncDistrictCalendarGrid() {
  ensureWorkbookScaffold_();
  syncDistrictCalendarSheet_();
  SpreadsheetApp.getUi().alert(
    'District calendar grid synced. Weekends default to No, weekdays default to Yes, and existing Yes/No edits were preserved where possible.'
  );
}

function reapplyAdminSheetProtection() {
  applyAllSheetProtections_();
  SpreadsheetApp.getUi().alert(
    'Admin sheet protection was reapplied. Direct sheet edits are limited to the configured admin accounts.'
  );
}

function showSpedSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('SPED Status Reports');
  SpreadsheetApp.getUi().showSidebar(html);
}

function getAppBootstrap() {
  ensureWorkbookReady_();

  return {
    caseTypes: [CASE_TYPES.initial, CASE_TYPES.reevaluation],
    services: SERVICE_FIELDS,
    districts: getActiveColumnValues_(SHEETS.districts, 'District'),
    campuses: getActiveRows_(SHEETS.campuses).map((row) => ({
      campus: row.Campus,
      district: row.District,
    })),
    evaluators: getActiveRows_(SHEETS.evaluators).map((row) => ({
      evaluator: row.LeadEvaluator,
      email: row.Email,
    })),
  };
}

function ensureWorkbookReady_() {
  Object.values(SHEETS).forEach((sheetName) => {
    if (!SpreadsheetApp.getActive().getSheetByName(sheetName)) {
      throw new Error(
        'Workbook setup is incomplete. Run "SPED Status Reports -> Install / Repair Workbook" and try again.'
      );
    }
  });
}

function previewTimeline(input) {
  return normalizeTimelineForUi_(buildProjectedDates_(input));
}

function searchCases(studentId, caseType) {
  const normalizedStudentId = normalizeStudentId_(studentId);
  if (!normalizedStudentId) {
    return [];
  }

  const rows = getTableRows_(SHEETS.cases, CASE_HEADERS);
  return rows
    .filter((row) => {
      if (normalizeStudentId_(row.StudentID) !== normalizedStudentId) {
        return false;
      }
      if (caseType && row.CaseType !== caseType) {
        return false;
      }
      return true;
    })
    .map((row) => ({
      caseId: row.CaseID,
      caseType: row.CaseType,
      status: row.Status,
      studentName: row.StudentName,
      studentId: row.StudentID,
      primaryDeadline: formatDate_(row.PrimaryDeadline),
    }));
}

function getCaseDetails(caseId) {
  const record = findCaseRecord_(caseId);
  if (!record) {
    throw new Error(`Case not found: ${caseId}`);
  }
  const row = record.row;

  row.documentsText = getCaseDocumentsText_(caseId);
  row.timelinePreview = normalizeTimelineForUi_(
    buildProjectedDates_({
      caseType: row.CaseType,
      district: row.District,
      referralReceivedDate: row.ReferralReceivedDate,
      actualConsentDate: row.ActualConsentDate,
      actualFIIEDate: row.ActualFIIEDate,
      reevalDueDate: row.ReevalDueDate,
    })
  );

  return normalizeCaseForUi_(row);
}

function saveNewCase(payload) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    validatePayload_(payload, false);

    if (duplicateOpenCaseExists_(payload.studentId, payload.caseType)) {
      throw new Error(`An open ${payload.caseType} case already exists for this student.`);
    }

    const timeline = buildProjectedDates_(payload);
    const caseId = generateCaseId_(payload.caseType);
    const now = new Date();
    const row = buildStoredCaseRow_(caseId, payload, timeline, now, now);

    appendRow_(SHEETS.cases, row, CASE_HEADERS);
    replaceCaseDocuments_(caseId, payload.documentLinks || '');
    refreshDashboard_(false);

    return {
      ok: true,
      caseId,
      timeline: normalizeTimelineForUi_(timeline),
    };
  } finally {
    lock.releaseLock();
  }
}

function updateExistingCase(payload) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    validatePayload_(payload, true);

    const existingRecord = findCaseRecord_(payload.caseId);
    if (!existingRecord) {
      throw new Error(`Case not found: ${payload.caseId}`);
    }
    const existing = existingRecord.row;

    if (
      duplicateOpenCaseExists_(payload.studentId, existing.CaseType, payload.caseId)
    ) {
      throw new Error(`A different open ${existing.CaseType} case already exists for this student.`);
    }

    const merged = Object.assign({}, existing, {
      StudentName: payload.studentName,
      StudentID: normalizeStudentId_(payload.studentId),
      DOB: parseDate_(payload.dob),
      Campus: payload.campus,
      District: payload.district,
      LeadEvaluator: payload.leadEvaluator,
      ReferralReceivedDate: parseDate_(payload.referralReceivedDate),
      ReevalDueDate: parseDate_(payload.reevalDueDate),
      ActualConsentDate: parseDate_(payload.actualConsentDate),
      ActualFIIEDate: parseDate_(payload.actualFIIEDate),
      ActualARDDate: parseDate_(payload.actualARDDate),
      ServiceNotes: payload.serviceNotes || '',
      ManualPrimaryDeadline: parseDate_(payload.overridePrimaryDeadline),
      ManualOverrideReason: payload.overrideReason || '',
      UpdatedAt: new Date(),
    });

    SERVICE_FIELDS.forEach((field) => {
      merged[field] = payload.services && payload.services[field] ? 1 : 0;
    });

    const timeline = buildProjectedDates_({
      caseType: existing.CaseType,
      district: merged.District,
      referralReceivedDate: merged.ReferralReceivedDate,
      actualConsentDate: merged.ActualConsentDate,
      actualFIIEDate: merged.ActualFIIEDate,
      reevalDueDate: merged.ReevalDueDate,
    });

    merged.ResponseDueDate = timeline.responseDueDate;
    merged.ProjectedConsentDate = timeline.projectedConsentDate;
    merged.ProjectedFIIEDueDate = timeline.projectedFiiEDueDate;
    merged.ProjectedARDDueDate = timeline.projectedArdDueDate;
    merged.Status = determineStatus_(merged.ActualConsentDate, merged.ActualARDDate);
    merged.PrimaryDeadline = determinePrimaryDeadline_(
      existing.CaseType,
      merged.ResponseDueDate,
      merged.ProjectedFIIEDueDate,
      merged.ProjectedARDDueDate,
      merged.ActualConsentDate,
      merged.ActualFIIEDate,
      merged.ActualARDDate,
      merged.ReevalDueDate,
      merged.ManualPrimaryDeadline
    );

    writeCaseRowByIndex_(existingRecord.rowIndex, merged);
    replaceCaseDocuments_(payload.caseId, payload.documentLinks || '');
    refreshDashboard_(false);

    return {
      ok: true,
      caseId: payload.caseId,
      timeline: normalizeTimelineForUi_(timeline),
    };
  } finally {
    lock.releaseLock();
  }
}

function validateAdminCode(adminCode) {
  return String(adminCode || '') === String(ADMIN_ACCESS_CODE);
}

function openDashboard() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEETS.dashboard);
  if (sheet) {
    SpreadsheetApp.getActive().setActiveSheet(sheet);
  }
}

function openDistrictsSheet(adminCode) {
  openAdminSheet_(SHEETS.districts, adminCode);
}

function openCampusesSheet(adminCode) {
  openAdminSheet_(SHEETS.campuses, adminCode);
}

function openEvaluatorsSheet(adminCode) {
  openAdminSheet_(SHEETS.evaluators, adminCode);
}

function openCalendarsSheet(adminCode) {
  openAdminSheet_(SHEETS.calendars, adminCode);
}

function showAllSheets(adminCode) {
  if (!validateAdminCode(adminCode)) {
    throw new Error('Admin access denied.');
  }
  showAllSheets_();
  return true;
}

function applyAdminSheetProtection(adminCode) {
  if (!validateAdminCode(adminCode)) {
    throw new Error('Admin access denied.');
  }
  applyAllSheetProtections_();
  return true;
}

function refreshDashboard() {
  ensureWorkbookScaffold_();
  refreshDashboard_(true);
}

function refreshDashboard_(applyLayout) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEETS.dashboard);
  const cases = getTableRows_(SHEETS.cases, CASE_HEADERS);

  const headers = [
    'Case ID',
    'Case Type',
    'Student Name',
    'Student ID',
    'Campus',
    'District',
    'Lead Evaluator',
    'Status',
    'Primary Deadline',
    'Response Due',
    'Projected FIIE Due',
    'Projected ARD Due',
    'Re-eval Due',
    'Updated At',
  ];

  if (applyLayout || sheet.getRange(4, 1).getValue() !== headers[0]) {
    sheet.clear();
    sheet.getRange(1, 1).setValue('SPED Status Reports Dashboard').setFontWeight('bold').setFontSize(16);
    sheet
      .getRange(2, 1)
      .setValue('Use built-in filters on row 4 to filter by district, campus, evaluator, case type, or status.');
    sheet.getRange(4, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  } else if (sheet.getLastRow() > 4) {
    sheet.getRange(5, 1, sheet.getLastRow() - 4, headers.length).clearContent().setBackground('#ffffff');
  }

  const output = cases.map((row) => [
    row.CaseID,
    row.CaseType,
    row.StudentName,
    row.StudentID,
    row.Campus,
    row.District,
    row.LeadEvaluator,
    row.Status,
    row.PrimaryDeadline || '',
    row.ResponseDueDate || '',
    row.ProjectedFIIEDueDate || '',
    row.ProjectedARDDueDate || '',
    row.ReevalDueDate || '',
    row.UpdatedAt || '',
  ]);

  if (output.length) {
    sheet.getRange(5, 1, output.length, headers.length).setValues(output);
    applyDashboardFormatting_(sheet, 5, output.length);
  }

  if (sheet.getFilter()) {
    sheet.getFilter().remove();
  }
  sheet.getRange(4, 1, Math.max(output.length + 1, 2), headers.length).createFilter();
  if (applyLayout) {
    sheet.autoResizeColumns(1, headers.length);
  }
}

function ensureWorkbookScaffold_(options) {
  const settings = Object.assign({
    seedReferenceData: false,
    syncCalendar: false,
    showSheets: false,
    applyProtection: false,
  }, options || {});

  ensureSheet_(SHEETS.cases, CASE_HEADERS);
  ensureSheet_(SHEETS.documents, DOCUMENT_HEADERS);
  ensureSheet_(SHEETS.districts, DISTRICT_HEADERS);
  ensureSheet_(SHEETS.campuses, CAMPUS_HEADERS);
  ensureSheet_(SHEETS.evaluators, EVALUATOR_HEADERS);
  ensureSheet_(SHEETS.calendars, CALENDAR_BASE_HEADERS);
  ensureSheet_(SHEETS.dashboard, ['SPED Status Reports Dashboard']);

  if (settings.seedReferenceData) {
    seedReferenceData_();
  }

  if (settings.syncCalendar) {
    syncDistrictCalendarSheet_();
  }

  if (settings.showSheets) {
    showAllSheets_();
  }

  if (settings.applyProtection) {
    applyAllSheetProtections_();
  }
}

function ensureSheet_(sheetName, headers) {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  if (headers && headers.length) {
    const existingHeaders = sheet.getLastColumn() > 0 ? sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), headers.length)).getValues()[0] : [];
    const needsHeaders = headers.some((header, index) => existingHeaders[index] !== header);
    if (needsHeaders) {
      sheet.clear();
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
  }

  return sheet;
}

function applyProtection_(sheet) {
  const allowedEditors = getAdminEditorEmails_();
  const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  let protection = protections.length ? protections[0] : null;

  if (!protection) {
    protection = sheet.protect();
  }

  protection.setDescription(`SPED backend protection for ${sheet.getName()}`);
  protection.setWarningOnly(false);

  const editors = protection.getEditors();
  if (editors.length) {
    protection.removeEditors(editors);
  }

  if (allowedEditors.length) {
    protection.addEditors(allowedEditors);
  }

  if (protection.canDomainEdit()) {
    protection.setDomainEdit(false);
  }

  return protection;
}

function applyAllSheetProtections_() {
  Object.values(SHEETS).forEach((name) => {
    const sheet = SpreadsheetApp.getActive().getSheetByName(name);
    if (sheet) {
      applyProtection_(sheet);
    }
  });
}

function showAllSheets_() {
  Object.values(SHEETS).forEach((name) => {
    const sheet = SpreadsheetApp.getActive().getSheetByName(name);
    if (sheet) {
      sheet.showSheet();
    }
  });
}

function getAdminEditorEmails_() {
  const emails = [];
  ADMIN_EDITOR_EMAILS.forEach((email) => {
    if (email) {
      emails.push(String(email).trim());
    }
  });

  return [...new Set(emails.filter(Boolean))];
}

function syncDistrictCalendarSheet_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEETS.calendars);
  const districtNames = getActiveColumnValues_(SHEETS.districts, 'District');
  const headers = CALENDAR_BASE_HEADERS.concat(districtNames);
  const existingData = getCalendarExistingState_(sheet);
  const dateRows = buildSchoolYearDates_();

  const values = dateRows.map((dateValue) => {
    const key = isoDateKey_(dateValue);
    const existingRow = existingData[key] || {};

    return [
      dateValue,
      Utilities.formatDate(dateValue, Session.getScriptTimeZone(), 'EEEE'),
      ...districtNames.map((districtName) => {
        if (existingRow[districtName] !== undefined && existingRow[districtName] !== '') {
          return existingRow[districtName];
        }
        return isWeekend_(dateValue) ? 'No' : 'Yes';
      }),
    ];
  });

  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sheet.setFrozenRows(1);

  if (values.length) {
    sheet.getRange(2, 1, values.length, headers.length).setValues(values);
    sheet.getRange(2, 1, values.length, 1).setNumberFormat('mm/dd/yyyy');
  }

  sheet.autoResizeColumns(1, headers.length);
}

function getCalendarExistingState_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  const state = {};

  if (lastRow < 2 || lastColumn < 1) {
    return state;
  }

  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const values = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();

  values.forEach((row) => {
    const dateValue = parseDate_(row[0]);
    if (!dateValue) {
      return;
    }

    const key = isoDateKey_(dateValue);
    state[key] = {};

    for (let columnIndex = CALENDAR_BASE_HEADERS.length; columnIndex < headers.length; columnIndex += 1) {
      state[key][headers[columnIndex]] = row[columnIndex];
    }
  });

  return state;
}

function buildSchoolYearDates_() {
  const start = new Date(2026, 0, 1);
  const end = new Date(2027, 5, 1);
  const dates = [];
  let cursor = new Date(start);

  while (cursor <= end) {
    dates.push(new Date(cursor));
    cursor.setDate(cursor.getDate() + 1);
  }

  return dates;
}

function seedReferenceData_() {
  maybeAppendSeedRow_(SHEETS.districts, DISTRICT_HEADERS, {
    District: 'Sample ISD',
    ResponseSchoolDays: 15,
    FIIESchoolDays: 45,
    ARDCalendarDays: 30,
    Active: 'Yes',
  });

  maybeAppendSeedRow_(SHEETS.campuses, CAMPUS_HEADERS, {
    Campus: 'Sample Elementary',
    District: 'Sample ISD',
    Active: 'Yes',
  });

  maybeAppendSeedRow_(SHEETS.evaluators, EVALUATOR_HEADERS, {
    LeadEvaluator: 'Sample Evaluator',
    Email: 'sample@example.org',
    Active: 'Yes',
  });
}

function maybeAppendSeedRow_(sheetName, headers, objectRow) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (sheet.getLastRow() > 1) {
    return;
  }
  appendRow_(sheetName, objectRow, headers);
}

function appendRow_(sheetName, objectRow, headers) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  const values = headers.map((header) => (objectRow[header] === undefined ? '' : objectRow[header]));
  sheet.appendRow(values);
}

function writeCaseRowByIndex_(rowIndex, objectRow) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEETS.cases);
  const rowValues = CASE_HEADERS.map((header) => (objectRow[header] === undefined ? '' : objectRow[header]));
  sheet.getRange(rowIndex, 1, 1, CASE_HEADERS.length).setValues([rowValues]);
}

function getTableRows_(sheetName, headers) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }

  const values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  return values.map((row) => toObject_(headers, row));
}

function getActiveRows_(sheetName) {
  const headers = sheetName === SHEETS.campuses ? CAMPUS_HEADERS : sheetName === SHEETS.evaluators ? EVALUATOR_HEADERS : DISTRICT_HEADERS;
  return getTableRows_(sheetName, headers).filter((row) => String(row.Active).toLowerCase() === 'yes');
}

function getActiveColumnValues_(sheetName, columnName) {
  return getActiveRows_(sheetName).map((row) => row[columnName]);
}

function openAdminSheet_(sheetName, adminCode) {
  let resolvedCode = adminCode;

  if (!resolvedCode) {
    resolvedCode = SpreadsheetApp.getUi().prompt('Enter the admin access code.').getResponseText();
  }

  if (!validateAdminCode(resolvedCode)) {
    SpreadsheetApp.getUi().alert('Admin access denied.');
    return;
  }

  const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  sheet.showSheet();
  SpreadsheetApp.getActive().setActiveSheet(sheet);
}

function buildStoredCaseRow_(caseId, payload, timeline, createdAt, updatedAt) {
  const actualConsentDate = parseDate_(payload.actualConsentDate);
  const actualFiiEDate = parseDate_(payload.actualFIIEDate);
  const actualArdDate = parseDate_(payload.actualARDDate);
  const reevalDueDate = parseDate_(payload.reevalDueDate);
  const manualPrimaryDeadline = parseDate_(payload.overridePrimaryDeadline);

  const row = {
    CaseID: caseId,
    CaseType: payload.caseType,
    StudentName: payload.studentName,
    StudentID: normalizeStudentId_(payload.studentId),
    DOB: parseDate_(payload.dob),
    Campus: payload.campus,
    District: payload.district,
    LeadEvaluator: payload.leadEvaluator,
    Status: determineStatus_(actualConsentDate, actualArdDate),
    ReferralReceivedDate: parseDate_(payload.referralReceivedDate),
    ResponseDueDate: timeline.responseDueDate,
    ProjectedConsentDate: timeline.projectedConsentDate,
    ActualConsentDate: actualConsentDate,
    ProjectedFIIEDueDate: timeline.projectedFiiEDueDate,
    ActualFIIEDate: actualFiiEDate,
    ProjectedARDDueDate: timeline.projectedArdDueDate,
    ActualARDDate: actualArdDate,
    ReevalDueDate: reevalDueDate,
    ServiceNotes: payload.serviceNotes || '',
    ManualPrimaryDeadline: manualPrimaryDeadline,
    ManualOverrideReason: payload.overrideReason || '',
    PrimaryDeadline: determinePrimaryDeadline_(
      payload.caseType,
      timeline.responseDueDate,
      timeline.projectedFiiEDueDate,
      timeline.projectedArdDueDate,
      actualConsentDate,
      actualFiiEDate,
      actualArdDate,
      reevalDueDate,
      manualPrimaryDeadline
    ),
    CreatedAt: createdAt,
    UpdatedAt: updatedAt,
  };

  SERVICE_FIELDS.forEach((field) => {
    row[field] = payload.services && payload.services[field] ? 1 : 0;
  });

  return row;
}

function validatePayload_(payload, requireCaseId) {
  if (requireCaseId && !payload.caseId) {
    throw new Error('Case ID is required.');
  }

  if (!payload.caseType) {
    throw new Error('Case Type is required.');
  }

  if (!payload.studentName) {
    throw new Error('Student Name is required.');
  }

  if (!payload.studentId) {
    throw new Error('Student ID is required.');
  }

  if (!payload.district) {
    throw new Error('District is required.');
  }

  if (payload.caseType === CASE_TYPES.initial && !parseDate_(payload.referralReceivedDate)) {
    throw new Error('Referral Received Date is required for Initial cases.');
  }

  if (payload.caseType === CASE_TYPES.reevaluation && !parseDate_(payload.reevalDueDate)) {
    throw new Error('Re-evaluation Due Date is required for Re-evaluation cases.');
  }

  if (payload.overridePrimaryDeadline || payload.overrideReason) {
    if (!validateAdminCode(payload.adminCode)) {
      throw new Error('Admin code is required to apply a manual due-date override.');
    }
    if (!payload.overridePrimaryDeadline || !payload.overrideReason) {
      throw new Error('Both an override due date and an override reason are required.');
    }
  }
}

function buildProjectedDates_(input) {
  const caseType = input.caseType;
  const district = input.district;
  const referralDate = parseDate_(input.referralReceivedDate);
  const actualConsentDate = parseDate_(input.actualConsentDate);
  const actualFiiEDate = parseDate_(input.actualFIIEDate);
  const reevalDueDate = parseDate_(input.reevalDueDate);
  const calendarLookup = district ? getDistrictCalendarLookup_(district) : {};
  const districtConfig = getDistrictConfig_(district);

  const timeline = {
    responseDueDate: '',
    projectedConsentDate: '',
    projectedFiiEDueDate: '',
    projectedArdDueDate: '',
  };

  if (caseType === CASE_TYPES.initial && referralDate) {
    const responseDays = districtConfig.ResponseSchoolDays;
    const fiieDays = districtConfig.FIIESchoolDays;
    const ardDays = districtConfig.ARDCalendarDays;

    const responseDueDate = addInstructionalDays_(referralDate, responseDays, district, calendarLookup);
    const consentAnchor = actualConsentDate || responseDueDate;
    const fiiEDueDate = calculateInitialFiiEDueDate_(consentAnchor, district, fiieDays, calendarLookup);
    const ardAnchor = actualFiiEDate || fiiEDueDate;

    timeline.responseDueDate = responseDueDate;
    timeline.projectedConsentDate = responseDueDate;
    timeline.projectedFiiEDueDate = fiiEDueDate;
    timeline.projectedArdDueDate = addCalendarDays_(ardAnchor, ardDays);
  }

  if (caseType === CASE_TYPES.reevaluation && reevalDueDate) {
    const ardDays = districtConfig.ARDCalendarDays;
    timeline.projectedFiiEDueDate = reevalDueDate;
    timeline.projectedArdDueDate = addCalendarDays_(actualFiiEDate || reevalDueDate, ardDays);
  }

  return timeline;
}

function getDistrictConfig_(districtName) {
  const defaults = {
    ResponseSchoolDays: 15,
    FIIESchoolDays: 45,
    ARDCalendarDays: 30,
  };

  if (!districtName) {
    return defaults;
  }

  const rows = getTableRows_(SHEETS.districts, DISTRICT_HEADERS);
  const row = rows.find((item) => String(item.District).trim() === String(districtName).trim());
  if (!row) {
    return defaults;
  }

  return {
    ResponseSchoolDays: Number(row.ResponseSchoolDays) || defaults.ResponseSchoolDays,
    FIIESchoolDays: Number(row.FIIESchoolDays) || defaults.FIIESchoolDays,
    ARDCalendarDays: Number(row.ARDCalendarDays) || defaults.ARDCalendarDays,
  };
}

function addInstructionalDays_(startDate, daysToAdd, districtName, calendarLookup) {
  let cursor = new Date(startDate);
  let counted = 0;
  const lookup = calendarLookup || getDistrictCalendarLookup_(districtName);

  while (counted < Number(daysToAdd)) {
    cursor = addCalendarDays_(cursor, 1);
    if (isInstructionalDay_(cursor, districtName, lookup)) {
      counted += 1;
    }
  }

  return normalizeDateForStorage_(cursor);
}

function isInstructionalDay_(dateValue, districtName, calendarLookup) {
  if (isWeekend_(dateValue)) {
    return false;
  }

  const lookup = calendarLookup || getDistrictCalendarLookup_(districtName);
  const value = lookup[isoDateKey_(dateValue)];

  if (value === 'No') {
    return false;
  }
  if (value === 'Yes') {
    return true;
  }

  return true;
}

function getDistrictCalendarLookup_(districtName) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEETS.calendars);
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  const lookup = {};

  if (lastRow < 2 || lastColumn < 3) {
    return lookup;
  }

  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const districtColumn = headers.indexOf(districtName);
  if (districtColumn === -1) {
    return lookup;
  }

  const dateValues = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const districtValues = sheet.getRange(2, districtColumn + 1, lastRow - 1, 1).getValues();
  dateValues.forEach((row, index) => {
    const dateValue = parseDate_(row[0]);
    if (dateValue) {
      lookup[isoDateKey_(dateValue)] = String(districtValues[index][0] || '').trim();
    }
  });

  return lookup;
}

function addCalendarDays_(dateValue, dayCount) {
  const output = new Date(dateValue);
  output.setDate(output.getDate() + Number(dayCount));
  return normalizeDateForStorage_(output);
}

function calculateInitialFiiEDueDate_(consentDate, districtName, fiieDays, calendarLookup) {
  const lookup = calendarLookup || getDistrictCalendarLookup_(districtName);
  const lastInstructionalDay = getLastInstructionalDayOfSchoolYear_(consentDate, districtName, lookup);
  const june30DueDate = getJune30ForSchoolYear_(consentDate);

  if (lastInstructionalDay) {
    const remainingInstructionalDays = countInstructionalDaysBetween_(consentDate, lastInstructionalDay, districtName, lookup);
    if (remainingInstructionalDays >= 35 && remainingInstructionalDays < 45) {
      return june30DueDate;
    }
  }

  return addInstructionalDays_(consentDate, fiieDays, districtName, lookup);
}

function countInstructionalDaysBetween_(startDate, endDate, districtName, calendarLookup) {
  let cursor = normalizeDateForStorage_(startDate);
  const normalizedEndDate = normalizeDateForStorage_(endDate);
  let counted = 0;

  while (cursor < normalizedEndDate) {
    cursor = addCalendarDays_(cursor, 1);
    if (cursor <= normalizedEndDate && isInstructionalDay_(cursor, districtName, calendarLookup)) {
      counted += 1;
    }
  }

  return counted;
}

function getLastInstructionalDayOfSchoolYear_(referenceDate, districtName, calendarLookup) {
  let cursor = getSchoolYearInstructionEndDate_(referenceDate);
  const normalizedReferenceDate = normalizeDateForStorage_(referenceDate);

  while (cursor >= normalizedReferenceDate) {
    if (isInstructionalDay_(cursor, districtName, calendarLookup)) {
      return cursor;
    }
    cursor = addCalendarDays_(cursor, -1);
  }

  return '';
}

function getSchoolYearInstructionEndDate_(referenceDate) {
  const normalizedDate = normalizeDateForStorage_(referenceDate);
  const endYear = normalizedDate.getMonth() >= 6 ? normalizedDate.getFullYear() + 1 : normalizedDate.getFullYear();
  return createLocalDate_(endYear, 6, 1);
}

function getJune30ForSchoolYear_(referenceDate) {
  const normalizedDate = normalizeDateForStorage_(referenceDate);
  const dueYear = normalizedDate.getMonth() >= 6 ? normalizedDate.getFullYear() + 1 : normalizedDate.getFullYear();
  return createLocalDate_(dueYear, 6, 30);
}

function determineStatus_(actualConsentDate, actualArdDate) {
  if (actualArdDate) {
    return STATUSES.complete;
  }
  if (actualConsentDate) {
    return STATUSES.inProgress;
  }
  return STATUSES.referralReceived;
}

function determinePrimaryDeadline_(
  caseType,
  responseDueDate,
  projectedFiiEDueDate,
  projectedArdDueDate,
  actualConsentDate,
  actualFiiEDate,
  actualArdDate,
  reevalDueDate,
  manualPrimaryDeadline
) {
  if (manualPrimaryDeadline) {
    return manualPrimaryDeadline;
  }

  if (caseType === CASE_TYPES.initial) {
    if (!actualConsentDate) {
      return responseDueDate;
    }
    if (!actualFiiEDate) {
      return projectedFiiEDueDate;
    }
    if (!actualArdDate) {
      return projectedArdDueDate;
    }
    return actualArdDate;
  }

  if (!actualFiiEDate) {
    return reevalDueDate;
  }
  if (!actualArdDate) {
    return projectedArdDueDate;
  }
  return actualArdDate;
}

function duplicateOpenCaseExists_(studentId, caseType, ignoreCaseId) {
  const normalizedStudentId = normalizeStudentId_(studentId);
  const rows = getTableRows_(SHEETS.cases, CASE_HEADERS);
  return rows.some((row) => {
    if (normalizeStudentId_(row.StudentID) !== normalizedStudentId) {
      return false;
    }
    if (row.CaseType !== caseType) {
      return false;
    }
    if (row.Status === STATUSES.complete) {
      return false;
    }
    if (ignoreCaseId && row.CaseID === ignoreCaseId) {
      return false;
    }
    return true;
  });
}

function findCaseRecord_(caseId) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEETS.cases);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return null;
  }

  const headerMap = getHeaderMap_(CASE_HEADERS);
  const values = sheet.getRange(2, 1, lastRow - 1, CASE_HEADERS.length).getValues();

  for (let index = 0; index < values.length; index += 1) {
    if (String(values[index][headerMap.CaseID]).trim() === String(caseId).trim()) {
      return {
        row: toObject_(CASE_HEADERS, values[index]),
        rowIndex: index + 2,
      };
    }
  }

  return null;
}

function replaceCaseDocuments_(caseId, rawDocumentText) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEETS.documents);
  const lastRow = sheet.getLastRow();
  const values = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, DOCUMENT_HEADERS.length).getValues() : [];
  const kept = values.filter((row) => String(row[1]).trim() !== String(caseId).trim());
  const parsed = parseDocumentLines_(rawDocumentText, caseId);

  sheet.clearContents();
  sheet.getRange(1, 1, 1, DOCUMENT_HEADERS.length).setValues([DOCUMENT_HEADERS]).setFontWeight('bold');

  const finalRows = kept.concat(parsed);
  if (finalRows.length) {
    sheet.getRange(2, 1, finalRows.length, DOCUMENT_HEADERS.length).setValues(finalRows);
  }
}

function parseDocumentLines_(rawDocumentText, caseId) {
  const lines = String(rawDocumentText || '')
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .split('\n')
    .map((line) => line.trim())
    .filter(Boolean);

  return lines.map((line, index) => {
    const parts = line.split('|');
    const label = parts.length > 1 ? parts[0].trim() : `Document ${index + 1}`;
    const path = parts.length > 1 ? line.slice(line.indexOf('|') + 1).trim() : line;

    return [
      `${caseId}-DOC${String(index + 1).padStart(2, '0')}`,
      caseId,
      label,
      path,
      new Date(),
    ];
  });
}

function getCaseDocumentsText_(caseId) {
  const rows = getTableRows_(SHEETS.documents, DOCUMENT_HEADERS);
  return rows
    .filter((row) => row.CaseID === caseId)
    .map((row) => `${row.DocumentLabel}|${row.DocumentPath}`)
    .join('\n');
}

function generateCaseId_(caseType) {
  const prefix = caseType === CASE_TYPES.reevaluation ? 'REE' : 'INI';
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');
  return `${prefix}-${stamp}`;
}

function normalizeTimelineForUi_(timeline) {
  return {
    responseDueDate: formatDate_(timeline.responseDueDate),
    projectedConsentDate: formatDate_(timeline.projectedConsentDate),
    projectedFiiEDueDate: formatDate_(timeline.projectedFiiEDueDate),
    projectedArdDueDate: formatDate_(timeline.projectedArdDueDate),
  };
}

function normalizeCaseForUi_(row) {
  const output = Object.assign({}, row, {
    DOB: formatDate_(row.DOB),
    ReferralReceivedDate: formatDate_(row.ReferralReceivedDate),
    ResponseDueDate: formatDate_(row.ResponseDueDate),
    ProjectedConsentDate: formatDate_(row.ProjectedConsentDate),
    ActualConsentDate: formatDate_(row.ActualConsentDate),
    ProjectedFIIEDueDate: formatDate_(row.ProjectedFIIEDueDate),
    ActualFIIEDate: formatDate_(row.ActualFIIEDate),
    ProjectedARDDueDate: formatDate_(row.ProjectedARDDueDate),
    ActualARDDate: formatDate_(row.ActualARDDate),
    ReevalDueDate: formatDate_(row.ReevalDueDate),
    ManualPrimaryDeadline: formatDate_(row.ManualPrimaryDeadline),
    PrimaryDeadline: formatDate_(row.PrimaryDeadline),
    CreatedAt: formatDateTime_(row.CreatedAt),
    UpdatedAt: formatDateTime_(row.UpdatedAt),
  });

  output.services = {};
  SERVICE_FIELDS.forEach((field) => {
    output.services[field] = Number(row[field]) === 1;
  });

  return output;
}

function parseDate_(value) {
  if (!value) {
    return '';
  }
  if (Object.prototype.toString.call(value) === '[object Date]' && !Number.isNaN(value.getTime())) {
    return normalizeDateForStorage_(value);
  }
  if (typeof value === 'string') {
    const isoMatch = value.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (isoMatch) {
      return createLocalDate_(Number(isoMatch[1]), Number(isoMatch[2]), Number(isoMatch[3]));
    }

    const slashMatch = value.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (slashMatch) {
      return createLocalDate_(Number(slashMatch[3]), Number(slashMatch[1]), Number(slashMatch[2]));
    }
  }
  const parsed = new Date(value);
  return Number.isNaN(parsed.getTime()) ? '' : normalizeDateForStorage_(parsed);
}

function formatDate_(value) {
  const dateValue = parseDate_(value);
  if (!dateValue) {
    return '';
  }
  return Utilities.formatDate(dateValue, Session.getScriptTimeZone(), 'MM/dd/yyyy');
}

function formatDateTime_(value) {
  const dateValue = parseDate_(value);
  if (!dateValue) {
    return '';
  }
  return Utilities.formatDate(dateValue, Session.getScriptTimeZone(), 'MM/dd/yyyy HH:mm:ss');
}

function sameDay_(left, right) {
  return (
    left.getFullYear() === right.getFullYear() &&
    left.getMonth() === right.getMonth() &&
    left.getDate() === right.getDate()
  );
}

function isWeekend_(dateValue) {
  const day = dateValue.getDay();
  return day === 0 || day === 6;
}

function isoDateKey_(dateValue) {
  return Utilities.formatDate(dateValue, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function createLocalDate_(year, month, day) {
  return new Date(year, month - 1, day, 12, 0, 0, 0);
}

function normalizeDateForStorage_(dateValue) {
  return new Date(dateValue.getFullYear(), dateValue.getMonth(), dateValue.getDate(), 12, 0, 0, 0);
}

function normalizeStudentId_(value) {
  const text = String(value === undefined || value === null ? '' : value).trim();
  if (!text) {
    return '';
  }

  if (/^\d+(\.0+)?$/.test(text)) {
    return text.replace(/\.0+$/, '');
  }

  return text.toUpperCase();
}

function toObject_(headers, values) {
  return headers.reduce((acc, header, index) => {
    acc[header] = values[index];
    return acc;
  }, {});
}

function getHeaderMap_(headers) {
  return headers.reduce((acc, header, index) => {
    acc[header] = index;
    return acc;
  }, {});
}

function applyDashboardFormatting_(sheet, startRow, rowCount) {
  const range = sheet.getRange(startRow, 1, rowCount, 14);
  const values = range.getValues();
  const backgrounds = values.map((row) => {
    const status = row[7];
    const deadline = parseDate_(row[8]);

    if (status === STATUSES.complete) {
      return new Array(14).fill('#c6efce');
    }
    if (deadline) {
      const today = new Date();
      today.setHours(0, 0, 0, 0);
      const compare = new Date(deadline);
      compare.setHours(0, 0, 0, 0);
      const diffDays = Math.floor((compare - today) / 86400000);

      if (diffDays < 0) {
        return new Array(14).fill('#ffc7ce');
      }
      if (diffDays <= 7) {
        return new Array(14).fill('#ffeb9c');
      }
    }

    return new Array(14).fill('#ffffff');
  });

  range.setBackgrounds(backgrounds);
}
