const ADMIN_ACCESS_CODE = 'CHANGE_ME_BEFORE_PRODUCTION';

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
    .addToUi();
}

function installSpedStatusReports() {
  ensureWorkbookScaffold_();
  seedReferenceData_();
  syncDistrictCalendarSheet_();
  hideBackendSheets_();
  refreshDashboard();
  SpreadsheetApp.getUi().alert(
    'SPED Status Reports is ready. Update the sample district, campus, evaluator, calendar, and admin settings before production use.'
  );
}

function syncDistrictCalendarGrid() {
  ensureWorkbookScaffold_();
  syncDistrictCalendarSheet_();
  SpreadsheetApp.getUi().alert(
    'District calendar grid synced. Weekends default to No, weekdays default to Yes, and existing Yes/No edits were preserved where possible.'
  );
}

function showSpedSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('SPED Status Reports');
  SpreadsheetApp.getUi().showSidebar(html);
}

function getAppBootstrap() {
  ensureWorkbookScaffold_();

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

function previewTimeline(input) {
  return normalizeTimelineForUi_(buildProjectedDates_(input));
}

function searchCases(studentId, caseType) {
  const rows = getTableRows_(SHEETS.cases, CASE_HEADERS);
  return rows
    .filter((row) => {
      if (studentId && String(row.StudentID).trim() !== String(studentId).trim()) {
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
  const row = findCaseRow_(caseId);
  if (!row) {
    throw new Error(`Case not found: ${caseId}`);
  }

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
    ensureWorkbookScaffold_();
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
    refreshDashboard();
    SpreadsheetApp.flush();

    return {
      ok: true,
      caseId,
      timeline: normalizeTimelineForUi_(timeline),
    };
  } finally {
    SpreadsheetApp.flush();
    lock.releaseLock();
  }
}

function updateExistingCase(payload) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    ensureWorkbookScaffold_();
    validatePayload_(payload, true);

    const existing = findCaseRow_(payload.caseId);
    if (!existing) {
      throw new Error(`Case not found: ${payload.caseId}`);
    }

    if (
      duplicateOpenCaseExists_(payload.studentId, existing.CaseType, payload.caseId)
    ) {
      throw new Error(`A different open ${existing.CaseType} case already exists for this student.`);
    }

    const merged = Object.assign({}, existing, {
      StudentName: payload.studentName,
      StudentID: payload.studentId,
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

    writeCaseRow_(payload.caseId, merged);
    replaceCaseDocuments_(payload.caseId, payload.documentLinks || '');
    refreshDashboard();
    SpreadsheetApp.flush();

    return {
      ok: true,
      caseId: payload.caseId,
      timeline: normalizeTimelineForUi_(timeline),
    };
  } finally {
    SpreadsheetApp.flush();
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

function showBackendSheets(adminCode) {
  if (!validateAdminCode(adminCode)) {
    throw new Error('Admin access denied.');
  }
  Object.values(SHEETS).forEach((name) => {
    const sheet = SpreadsheetApp.getActive().getSheetByName(name);
    if (sheet) {
      sheet.showSheet();
    }
  });
  return true;
}

function hideBackendSheets(adminCode) {
  if (!validateAdminCode(adminCode)) {
    throw new Error('Admin access denied.');
  }
  hideBackendSheets_();
  return true;
}

function refreshDashboard() {
  ensureWorkbookScaffold_();

  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEETS.dashboard);
  const cases = getTableRows_(SHEETS.cases, CASE_HEADERS);

  sheet.clear();
  sheet.getRange(1, 1).setValue('SPED Status Reports Dashboard').setFontWeight('bold').setFontSize(16);
  sheet
    .getRange(2, 1)
    .setValue('Use built-in filters on row 4 to filter by district, campus, evaluator, case type, or status.');

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

  sheet.getRange(4, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');

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
  sheet.autoResizeColumns(1, headers.length);
}

function ensureWorkbookScaffold_() {
  ensureSheet_(SHEETS.cases, CASE_HEADERS);
  ensureSheet_(SHEETS.documents, DOCUMENT_HEADERS);
  ensureSheet_(SHEETS.districts, DISTRICT_HEADERS);
  ensureSheet_(SHEETS.campuses, CAMPUS_HEADERS);
  ensureSheet_(SHEETS.evaluators, EVALUATOR_HEADERS);
  ensureSheet_(SHEETS.calendars, CALENDAR_BASE_HEADERS);
  ensureSheet_(SHEETS.dashboard, ['SPED Status Reports Dashboard']);
  syncDistrictCalendarSheet_();
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

  if (sheetName !== SHEETS.dashboard) {
    applyProtection_(sheet);
  }

  return sheet;
}

function applyProtection_(sheet) {
  const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  if (protections.length) {
    return protections[0];
  }

  const protection = sheet.protect();
  protection.setDescription(`SPED backend protection for ${sheet.getName()}`);
  return protection;
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
  const today = new Date();
  const year = today.getMonth() >= 6 ? today.getFullYear() : today.getFullYear() - 1;
  const start = new Date(year, 6, 1);
  const end = new Date(year + 1, 5, 30);
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

function writeCaseRow_(caseId, objectRow) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEETS.cases);
  const values = sheet.getDataRange().getValues();
  const headerMap = getHeaderMap_(CASE_HEADERS);

  for (let rowIndex = 1; rowIndex < values.length; rowIndex += 1) {
    if (String(values[rowIndex][headerMap.CaseID]).trim() === String(caseId).trim()) {
      const rowValues = CASE_HEADERS.map((header) => (objectRow[header] === undefined ? '' : objectRow[header]));
      sheet.getRange(rowIndex + 1, 1, 1, CASE_HEADERS.length).setValues([rowValues]);
      return;
    }
  }

  throw new Error(`Case not found: ${caseId}`);
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

function hideBackendSheets_() {
  const dashboard = SpreadsheetApp.getActive().getSheetByName(SHEETS.dashboard);
  Object.entries(SHEETS).forEach(([, name]) => {
    const sheet = SpreadsheetApp.getActive().getSheetByName(name);
    if (!sheet) {
      return;
    }
    if (name === SHEETS.dashboard) {
      sheet.showSheet();
    } else {
      sheet.hideSheet();
    }
  });
  if (dashboard) {
    SpreadsheetApp.getActive().setActiveSheet(dashboard);
  }
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
    StudentID: payload.studentId,
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

  const timeline = {
    responseDueDate: '',
    projectedConsentDate: '',
    projectedFiiEDueDate: '',
    projectedArdDueDate: '',
  };

  if (caseType === CASE_TYPES.initial && referralDate) {
    const responseDays = getDistrictRule_(district, 'ResponseSchoolDays', 15);
    const fiieDays = getDistrictRule_(district, 'FIIESchoolDays', 45);
    const ardDays = getDistrictRule_(district, 'ARDCalendarDays', 30);

    const responseDueDate = addInstructionalDays_(referralDate, responseDays, district);
    const consentAnchor = actualConsentDate || responseDueDate;
    const fiiEDueDate = addInstructionalDays_(consentAnchor, fiieDays, district);
    const ardAnchor = actualFiiEDate || fiiEDueDate;

    timeline.responseDueDate = responseDueDate;
    timeline.projectedConsentDate = responseDueDate;
    timeline.projectedFiiEDueDate = fiiEDueDate;
    timeline.projectedArdDueDate = addCalendarDays_(ardAnchor, ardDays);
  }

  if (caseType === CASE_TYPES.reevaluation && reevalDueDate) {
    const ardDays = getDistrictRule_(district, 'ARDCalendarDays', 30);
    timeline.projectedFiiEDueDate = reevalDueDate;
    timeline.projectedArdDueDate = addCalendarDays_(actualFiiEDate || reevalDueDate, ardDays);
  }

  return timeline;
}

function getDistrictRule_(districtName, ruleColumn, fallbackValue) {
  const rows = getTableRows_(SHEETS.districts, DISTRICT_HEADERS);
  const row = rows.find((item) => String(item.District).trim() === String(districtName).trim());
  if (!row || row[ruleColumn] === '') {
    return fallbackValue;
  }
  return Number(row[ruleColumn]) || fallbackValue;
}

function addInstructionalDays_(startDate, daysToAdd, districtName) {
  let cursor = new Date(startDate);
  let counted = 0;
  const calendarLookup = getDistrictCalendarLookup_(districtName);

  while (counted < Number(daysToAdd)) {
    cursor = addCalendarDays_(cursor, 1);
    if (isInstructionalDay_(cursor, districtName, calendarLookup)) {
      counted += 1;
    }
  }

  return cursor;
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

  const values = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();
  values.forEach((row) => {
    const dateValue = parseDate_(row[0]);
    if (dateValue) {
      lookup[isoDateKey_(dateValue)] = String(row[districtColumn] || '').trim();
    }
  });

  return lookup;
}

function addCalendarDays_(dateValue, dayCount) {
  const output = new Date(dateValue);
  output.setDate(output.getDate() + Number(dayCount));
  return output;
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
  const rows = getTableRows_(SHEETS.cases, CASE_HEADERS);
  return rows.some((row) => {
    if (String(row.StudentID).trim() !== String(studentId).trim()) {
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

function findCaseRow_(caseId) {
  const rows = getTableRows_(SHEETS.cases, CASE_HEADERS);
  return rows.find((row) => String(row.CaseID).trim() === String(caseId).trim()) || null;
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
    return value;
  }
  const parsed = new Date(value);
  return Number.isNaN(parsed.getTime()) ? '' : parsed;
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
