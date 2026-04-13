const ADMIN_ACCESS_CODE = 'CHANGE_ME_BEFORE_PRODUCTION';
const ADMIN_EDITOR_EMAILS = [];
const UPLOADS_FOLDER_ID = '';
const DISTRICT_DASHBOARD_PREFIX = 'Dashboard - ';
const DOCUMENT_PROPERTIES = {
  pendingCaseId: 'SPED_PENDING_CASE_ID',
};

const SHEETS = {
  cases: 'Cases',
  archive: 'ArchiveCases',
  documents: 'CaseDocuments',
  tests: 'DueDateTests',
  districts: 'Districts',
  campuses: 'Campuses',
  evaluators: 'Evaluators',
  calendars: 'DistrictCalendars',
  settings: 'Settings',
  audit: 'AuditLog',
  dashboard: 'Dashboard',
  summaryCaseType: 'SummaryByCaseType',
  summaryEvaluator: 'SummaryByEvaluator',
  summaryDistrictCaseType: 'SummaryByDistrictCaseType',
};

const CASE_TYPES = {
  initial: 'Initial',
  reevaluation: 'Re-evaluation',
};

const STATUSES = {
  referralReceived: 'Referral Received',
  responseSent: 'Response Sent',
  consentReceived: 'Consent Received',
  evaluationInProgress: 'Evaluation in Progress',
  evaluationComplete: 'Evaluation Complete',
  ardScheduled: 'ARD Scheduled',
  completed: 'Completed',
};

const CASE_HEADERS = [
  'CaseID',
  'CaseType',
  'StudentName',
  'StudentID',
  'GradeLevel',
  'DOB',
  'Campus',
  'District',
  'LeadEvaluator',
  'Status',
  'ReferralReceivedDate',
  'ResponseDueDate',
  'ResponseSentDate',
  'ProjectedConsentDate',
  'ActualConsentDate',
  'ProjectedFIIEDueDate',
  'ActualFIIEDate',
  'EvaluationStartedDate',
  'ARDScheduledDate',
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
  'VarianceExplanation',
  'ManualPrimaryDeadline',
  'ManualOverrideReason',
  'PrimaryDeadline',
  'CreatedAt',
  'UpdatedAt',
];

const DOCUMENT_HEADERS = ['DocumentID', 'CaseID', 'DocumentLabel', 'DocumentPath', 'AddedAt'];
const ARCHIVE_HEADERS = CASE_HEADERS.concat(['ArchivedAt']);
const TEST_HEADERS = [
  'ScenarioName',
  'CaseType',
  'District',
  'ReferralReceivedDate',
  'ActualConsentDate',
  'ActualFIIEDate',
  'ReevalDueDate',
  'ExpectedResponseDueDate',
  'ExpectedProjectedConsentDate',
  'ExpectedProjectedFIIEDueDate',
  'ExpectedProjectedARDDueDate',
  'ActualResponseDueDate',
  'ActualProjectedConsentDate',
  'ActualProjectedFIIEDueDate',
  'ActualProjectedARDDueDate',
  'Result',
  'Notes',
];
const DISTRICT_HEADERS = ['District', 'ResponseSchoolDays', 'FIIESchoolDays', 'ARDCalendarDays', 'Active'];
const CAMPUS_HEADERS = ['Campus', 'District', 'Active'];
const EVALUATOR_HEADERS = ['LeadEvaluator', 'Email', 'Active'];
const CALENDAR_BASE_HEADERS = ['Date', 'Weekday'];
const SETTINGS_HEADERS = ['SettingKey', 'SettingValue', 'Description'];
const AUDIT_HEADERS = ['AuditID', 'CaseID', 'Action', 'FieldName', 'OldValue', 'NewValue', 'ChangedAt'];

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

const SERVICE_LABELS = {
  Service_SchoolPsychologist: 'School Psychologist',
  Service_OccupationalTherapist: 'Occupational Therapist',
  Service_PhysicalTherapist: 'Physical Therapist',
  Service_CounselingEvaluation: 'Counseling Evaluation',
  Service_FBA: 'FBA',
  Service_SpeechPathologist: 'Speech Pathologist',
  Service_VI: 'VI',
  Service_DHH: 'DHH',
  Service_LanguageDominanceBilingual: 'Language Dominance / Bilingual',
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('SPED Status Reports')
    .addItem('Open App', 'showSpedSidebar')
    .addItem('Install / Repair Workbook', 'installSpedStatusReports')
    .addSeparator()
    .addItem('Refresh Dashboard', 'refreshDashboard')
    .addItem('Open Dashboard', 'openDashboard')
    .addItem('Open Selected Case', 'openSelectedCaseFromActiveRow')
    .addSeparator()
    .addItem('Open Districts (Admin)', 'openDistrictsSheet')
    .addItem('Open Campuses (Admin)', 'openCampusesSheet')
    .addItem('Open Evaluators (Admin)', 'openEvaluatorsSheet')
    .addItem('Open Calendars (Admin)', 'openCalendarsSheet')
    .addItem('Open Due Date Tests (Admin)', 'openDueDateTestsSheet')
    .addItem('Open Settings (Admin)', 'openSettingsSheet')
    .addItem('Open Audit Log (Admin)', 'openAuditSheet')
    .addItem('Open Archive (Admin)', 'openArchiveSheet')
    .addItem('Open District Case Type Summary', 'openDistrictCaseTypeSummarySheet')
    .addItem('Open Case Type Summary', 'openCaseTypeSummarySheet')
    .addItem('Open Evaluator Summary', 'openEvaluatorSummarySheet')
    .addItem('Archive Completed Cases (Admin)', 'archiveCompletedCases')
    .addItem('Restore Selected Archived Case (Admin)', 'restoreSelectedArchivedCase')
    .addItem('Refresh Due Date Tests', 'refreshDueDateTests')
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
    pendingCaseId: consumePendingCaseId_(),
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

function getLateDateWarnings(payload) {
  return getLateDateWarnings_(payload, buildProjectedDates_(payload));
}

function getQuickCaseList(filterName, evaluatorName) {
  ensureWorkbookReady_();

  const activeCases = getTableRows_(SHEETS.cases, CASE_HEADERS)
    .filter((row) => row.Status !== STATUSES.completed);
  const today = normalizeDateForStorage_(new Date());
  const normalizedEvaluator = normalizeComparisonText_(evaluatorName);
  let filtered = activeCases;

  if (filterName === 'overdue') {
    filtered = activeCases.filter((row) => {
      const deadline = parseDate_(row.PrimaryDeadline);
      return deadline && deadline.getTime() < today.getTime();
    });
  } else if (filterName === 'dueThisWeek') {
    filtered = activeCases.filter((row) => {
      const deadline = parseDate_(row.PrimaryDeadline);
      if (!deadline) {
        return false;
      }
      const diffDays = Math.floor((deadline - today) / 86400000);
      return diffDays >= 0 && diffDays <= 7;
    });
  } else if (filterName === 'myCases') {
    if (!normalizedEvaluator) {
      throw new Error('Choose a lead evaluator before loading evaluator cases.');
    }
    filtered = activeCases.filter((row) => normalizeComparisonText_(row.LeadEvaluator) === normalizedEvaluator);
  }

  filtered.sort((left, right) => {
    const leftDeadline = parseDate_(left.PrimaryDeadline);
    const rightDeadline = parseDate_(right.PrimaryDeadline);
    if (!leftDeadline && !rightDeadline) {
      return String(left.StudentName || '').localeCompare(String(right.StudentName || ''));
    }
    if (!leftDeadline) {
      return 1;
    }
    if (!rightDeadline) {
      return -1;
    }
    return leftDeadline.getTime() - rightDeadline.getTime();
  });

  return filtered.slice(0, 100).map((row) => ({
    caseId: row.CaseID,
    caseType: row.CaseType,
    studentName: row.StudentName,
    studentId: row.StudentID,
    leadEvaluator: row.LeadEvaluator,
    status: row.Status,
    primaryDeadline: formatDate_(row.PrimaryDeadline),
  }));
}

function getSelectedCaseDetails() {
  const selection = getSelectedCaseReference_();
  if (!selection || selection.location !== 'active') {
    throw new Error('Select a case row from Cases or a dashboard first.');
  }
  return getCaseDetails(selection.caseId);
}

function searchCases(studentId, caseType) {
  const normalizedStudentId = normalizeStudentId_(studentId);
  if (!normalizedStudentId) {
    return [];
  }

  const rows = getTableRows_(SHEETS.cases, CASE_HEADERS);
  return rows
    .filter((row) => {
      if (!String(row.CaseID || '').trim()) {
        return false;
      }
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
  if (!String(caseId || '').trim()) {
    throw new Error('A valid Case ID is required to load an existing case.');
  }

  const record = findCaseRecord_(caseId);
  if (!record) {
    throw new Error(`Case not found: ${caseId}`);
  }
  const row = record.row;

  row.documents = normalizeDocumentsForUi_(getCaseDocuments_(caseId));
  row.documentsText = buildCaseDocumentsText_(row.documents);
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

  return toHtmlSafeObject_(normalizeCaseForUi_(row));
}

function saveNewCase(payload) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    validatePayload_(payload, false);

    const duplicate = findOpenDuplicateCase_(payload, payload.caseType);
    if (duplicate) {
      throw new Error(buildDuplicateCaseMessage_(duplicate, payload.caseType));
    }

    const timeline = buildProjectedDates_(payload);
    validateVarianceNotes_(payload, timeline);
    const caseId = generateCaseId_(payload.caseType);
    const now = new Date();
    const row = buildStoredCaseRow_(caseId, payload, timeline, now, now);

    appendRow_(SHEETS.cases, row, CASE_HEADERS);
    replaceCaseDocuments_(caseId, payload.documentLinks || '');
    logCaseCreation_(row);
    logDocumentsUpdate_(caseId, '', payload.documentLinks || '');
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

    const duplicate = findOpenDuplicateCase_(payload, existing.CaseType, payload.caseId);
    if (duplicate) {
      throw new Error(buildDuplicateCaseMessage_(duplicate, existing.CaseType));
    }

    const merged = Object.assign({}, existing, {
      StudentName: payload.studentName,
      StudentID: normalizeStudentId_(payload.studentId),
      GradeLevel: String(payload.gradeLevel || '').trim(),
      DOB: parseDate_(payload.dob),
      Campus: payload.campus,
      District: payload.district,
      LeadEvaluator: payload.leadEvaluator,
      ReferralReceivedDate: parseDate_(payload.referralReceivedDate),
      ReevalDueDate: parseDate_(payload.reevalDueDate),
      ResponseSentDate: parseDate_(payload.responseSentDate),
      ActualConsentDate: parseDate_(payload.actualConsentDate),
      ActualFIIEDate: parseDate_(payload.actualFIIEDate),
      EvaluationStartedDate: parseDate_(payload.evaluationStartedDate),
      ARDScheduledDate: parseDate_(payload.ardScheduledDate),
      ActualARDDate: parseDate_(payload.actualARDDate),
      ServiceNotes: payload.serviceNotes || '',
      VarianceExplanation: payload.varianceExplanation || '',
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
    validateVarianceNotes_(merged, timeline);

    merged.ResponseDueDate = timeline.responseDueDate;
    merged.ProjectedConsentDate = timeline.projectedConsentDate;
    merged.ProjectedFIIEDueDate = timeline.projectedFiiEDueDate;
    merged.ProjectedARDDueDate = timeline.projectedArdDueDate;
    merged.Status = determineStatus_(
      merged.ResponseSentDate,
      merged.ActualConsentDate,
      merged.EvaluationStartedDate,
      merged.ActualFIIEDate,
      merged.ARDScheduledDate,
      merged.ActualARDDate
    );
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
    const previousDocumentsText = getCaseDocumentsText_(payload.caseId);
    replaceCaseDocuments_(payload.caseId, payload.documentLinks || '');
    logCaseUpdate_(existing, merged);
    logDocumentsUpdate_(payload.caseId, previousDocumentsText, payload.documentLinks || '');
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

function openSelectedCaseFromActiveRow() {
  ensureWorkbookReady_();
  const selection = getSelectedCaseReference_();
  if (!selection || selection.location !== 'active') {
    throw new Error('Select a case row from Cases or a dashboard first.');
  }
  setPendingCaseId_(selection.caseId);
  showSpedSidebar();
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

function openDueDateTestsSheet(adminCode) {
  openAdminSheet_(SHEETS.tests, adminCode);
}

function openSettingsSheet(adminCode) {
  openAdminSheet_(SHEETS.settings, adminCode);
}

function openAuditSheet(adminCode) {
  openAdminSheet_(SHEETS.audit, adminCode);
}

function openArchiveSheet(adminCode) {
  openAdminSheet_(SHEETS.archive, adminCode);
}

function openCaseTypeSummarySheet() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEETS.summaryCaseType);
  if (sheet) {
    SpreadsheetApp.getActive().setActiveSheet(sheet);
  }
}

function openDistrictCaseTypeSummarySheet() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEETS.summaryDistrictCaseType);
  if (sheet) {
    SpreadsheetApp.getActive().setActiveSheet(sheet);
  }
}

function openEvaluatorSummarySheet() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEETS.summaryEvaluator);
  if (sheet) {
    SpreadsheetApp.getActive().setActiveSheet(sheet);
  }
}

function archiveCompletedCases(adminCode) {
  let resolvedCode = adminCode;
  if (!resolvedCode) {
    resolvedCode = SpreadsheetApp.getUi().prompt('Enter the admin access code to archive completed cases.').getResponseText();
  }

  if (!validateAdminCode(resolvedCode)) {
    throw new Error('Admin access denied.');
  }

  ensureWorkbookScaffold_();

  const completedCases = getTableRows_(SHEETS.cases, CASE_HEADERS).filter((row) => row.Status === STATUSES.completed);
  if (!completedCases.length) {
    SpreadsheetApp.getUi().alert('No completed cases are ready to archive.');
    return 0;
  }

  const archiveSheet = SpreadsheetApp.getActive().getSheetByName(SHEETS.archive);
  const archivedAt = new Date();
  const archiveValues = completedCases.map((row) => ARCHIVE_HEADERS.map((header) => (
    header === 'ArchivedAt' ? archivedAt : (row[header] === undefined ? '' : row[header])
  )));

  archiveSheet.getRange(archiveSheet.getLastRow() + 1, 1, archiveValues.length, ARCHIVE_HEADERS.length).setValues(archiveValues);

  const remainingCases = getTableRows_(SHEETS.cases, CASE_HEADERS).filter((row) => row.Status !== STATUSES.completed);
  rewriteSheetRows_(SHEETS.cases, CASE_HEADERS, remainingCases);

  appendAuditRows_(completedCases.map((row) => buildAuditRow_(row.CaseID, 'Archive', 'Status', row.Status, 'Archived')));
  refreshDashboard_(true);
  SpreadsheetApp.getUi().alert(`${completedCases.length} completed case(s) were archived.`);
  return completedCases.length;
}

function restoreSelectedArchivedCase(adminCode) {
  let resolvedCode = adminCode;
  if (!resolvedCode) {
    resolvedCode = SpreadsheetApp.getUi().prompt('Enter the admin access code to restore an archived case.').getResponseText();
  }

  if (!validateAdminCode(resolvedCode)) {
    throw new Error('Admin access denied.');
  }

  ensureWorkbookScaffold_();

  const selection = getSelectedCaseReference_();
  if (!selection || selection.location !== 'archive') {
    throw new Error('Select a row from ArchiveCases first.');
  }

  const restoredCaseId = restoreArchivedCaseById_(selection.caseId);
  SpreadsheetApp.getUi().alert(`${restoredCaseId} was restored to the active Cases sheet.`);
  return restoredCaseId;
}

function refreshDueDateTests() {
  ensureWorkbookScaffold_();
  const results = refreshDueDateTests_();
  SpreadsheetApp.getUi().alert(`Due date tests refreshed. ${results.passCount} passed, ${results.failCount} failed, ${results.checkCount} need expected dates.`);
  return results;
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
  const cases = getTableRows_(SHEETS.cases, CASE_HEADERS);
  const documentsByCase = getCaseDocumentsMap_();

  renderDashboardSheet_(
    ensureSheet_(SHEETS.dashboard, ['SPED Status Reports Dashboard']),
    'SPED Status Reports Dashboard',
    'Use built-in filters on row 4 to filter by district, campus, evaluator, case type, or status.',
    cases,
    documentsByCase,
    applyLayout
  );

  getActiveColumnValues_(SHEETS.districts, 'District').forEach((districtName) => {
    const districtSheetName = getDistrictDashboardSheetName_(districtName);
    const existingSheet = ss.getSheetByName(districtSheetName);
    const districtSheet = existingSheet || ss.insertSheet(districtSheetName);
    renderDashboardSheet_(
      districtSheet,
      `${districtName} SPED Status Reports`,
      `District dashboard for ${districtName}.`,
      cases.filter((row) => String(row.District).trim() === String(districtName).trim()),
      documentsByCase,
      applyLayout
    );
    if ((applyLayout || !existingSheet) && getAdminEditorEmails_().length) {
      applyProtection_(districtSheet);
    }
  });

  renderSummarySheet_(
    ss.getSheetByName(SHEETS.summaryCaseType),
    'Summary By Case Type',
    'Active case summary split by Initial and Re-evaluation.',
    getCaseTypeSummaryRows_(cases),
    applyLayout
  );

  renderSummarySheet_(
    ss.getSheetByName(SHEETS.summaryEvaluator),
    'Summary By Evaluator',
    'Active case summary split by evaluator.',
    getEvaluatorSummaryRows_(cases),
    applyLayout
  );

  renderSummarySheet_(
    ss.getSheetByName(SHEETS.summaryDistrictCaseType),
    'Summary By District And Case Type',
    'Active case summary split by district and then by Initial / Re-evaluation.',
    getDistrictCaseTypeSummaryRows_(cases),
    applyLayout
  );
}

function renderDashboardSheet_(sheet, title, subtitle, cases, documentsByCase, applyLayout) {
  const headers = getDashboardHeaders_();
  const records = cases.map((row) => buildDashboardRecord_(row, documentsByCase[row.CaseID] || []));
  const output = records.map((record) => record.values);
  const summary = getDashboardSummary_(cases);
  const headerRow = 7;
  const dataStartRow = 8;

  if (applyLayout || sheet.getRange(headerRow, 1).getValue() !== headers[0]) {
    sheet.clear();
    sheet.getRange(1, 1).setValue(title).setFontWeight('bold').setFontSize(16);
    sheet.getRange(2, 1).setValue(subtitle);
    sheet.setFrozenRows(headerRow);
  } else if (sheet.getLastRow() >= dataStartRow) {
    sheet.getRange(dataStartRow, 1, sheet.getLastRow() - dataStartRow + 1, headers.length).clearContent().setBackground('#ffffff');
  }

  renderDashboardSummary_(sheet, summary);
  sheet.getRange(headerRow, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');

  if (output.length) {
    sheet.getRange(dataStartRow, 1, output.length, headers.length).setValues(output);
    applyDashboardFormatting_(sheet, dataStartRow, records);
    sheet.getRange(dataStartRow, 4, output.length, 2).setNumberFormat('@');
    sheet.getRange(dataStartRow, headers.length - 1, output.length, 2).setWrap(true);
  }

  if (sheet.getFilter()) {
    sheet.getFilter().remove();
  }
  sheet.getRange(headerRow, 1, Math.max(output.length + 1, 2), headers.length).createFilter();
  if (applyLayout) {
    sheet.autoResizeColumns(1, headers.length);
  }
}

function renderDashboardSummary_(sheet, summary) {
  sheet.getRange(4, 1, 2, 10).breakApart();
  const cards = [
    ['Total Active', summary.totalActive],
    ['Overdue', summary.overdue],
    ['Due This Week', summary.dueThisWeek],
    ['Due This Month', summary.dueThisMonth],
    ['Completed', summary.completed],
  ];
  const backgrounds = ['#d9ead3', '#ff9999', '#fff2cc', '#f4cccc', '#d0e0e3'];

  cards.forEach((card, index) => {
    const column = index * 2 + 1;
    sheet.getRange(4, column, 1, 2).merge().setValue(card[0]).setFontWeight('bold').setBackground(backgrounds[index]);
    sheet
      .getRange(5, column, 1, 2)
      .merge()
      .setValue(card[1])
      .setFontSize(16)
      .setFontWeight('bold')
      .setBackground('#ffffff')
      .setNumberFormat('0');
  });
}

function getDashboardSummary_(cases) {
  const today = normalizeDateForStorage_(new Date());
  let overdue = 0;
  let dueThisWeek = 0;
  let dueThisMonth = 0;
  let completed = 0;

  cases.forEach((row) => {
    if (row.Status === STATUSES.completed) {
      completed += 1;
      return;
    }

    const deadline = parseDate_(row.PrimaryDeadline);
    if (!deadline) {
      return;
    }

    const diffDays = Math.floor((deadline - today) / 86400000);
    if (diffDays < 0) {
      overdue += 1;
    }
    if (diffDays >= 0 && diffDays <= 7) {
      dueThisWeek += 1;
    }
    if (diffDays >= 0 && diffDays <= 30) {
      dueThisMonth += 1;
    }
  });

  return {
    totalActive: cases.length,
    overdue,
    dueThisWeek,
    dueThisMonth,
    completed,
  };
}

function renderSummarySheet_(sheet, title, subtitle, rows, applyLayout) {
  const headers = ['Breakdown', 'Total Active', 'Overdue', 'Due This Week', 'Due This Month', 'Completed'];
  const headerRow = 4;
  const dataStartRow = 5;

  if (applyLayout || sheet.getRange(headerRow, 1).getValue() !== headers[0]) {
    sheet.clear();
    sheet.getRange(1, 1).setValue(title).setFontWeight('bold').setFontSize(16);
    sheet.getRange(2, 1).setValue(subtitle);
    sheet.setFrozenRows(headerRow);
  } else if (sheet.getLastRow() >= dataStartRow) {
    sheet.getRange(dataStartRow, 1, sheet.getLastRow() - dataStartRow + 1, headers.length).clearContent().setBackground('#ffffff');
  }

  sheet.getRange(headerRow, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');

  if (rows.length) {
    const values = rows.map((row) => [
      row.label,
      row.totalActive,
      row.overdue,
      row.dueThisWeek,
      row.dueThisMonth,
      row.completed,
    ]);
    sheet.getRange(dataStartRow, 1, values.length, headers.length).setValues(values);
    sheet.getRange(dataStartRow, 2, values.length, headers.length - 1).setNumberFormat('0');
    applySummarySheetFormatting_(sheet, dataStartRow, rows.length, headers.length);
  }

  if (sheet.getFilter()) {
    sheet.getFilter().remove();
  }
  sheet.getRange(headerRow, 1, Math.max(rows.length + 1, 2), headers.length).createFilter();
  if (applyLayout) {
    sheet.autoResizeColumns(1, headers.length);
  }
}

function getCaseTypeSummaryRows_(cases) {
  return [
    buildSummaryRow_('Initial', cases.filter((row) => row.CaseType === CASE_TYPES.initial)),
    buildSummaryRow_('Re-evaluation', cases.filter((row) => row.CaseType === CASE_TYPES.reevaluation)),
  ];
}

function getEvaluatorSummaryRows_(cases) {
  const evaluatorNames = [...new Set(cases.map((row) => String(row.LeadEvaluator || '').trim()).filter(Boolean))].sort();
  return evaluatorNames.map((evaluatorName) => (
    buildSummaryRow_(evaluatorName, cases.filter((row) => String(row.LeadEvaluator || '').trim() === evaluatorName))
  ));
}

function getDistrictCaseTypeSummaryRows_(cases) {
  const districtNames = [...new Set(cases.map((row) => String(row.District || '').trim()).filter(Boolean))].sort();
  const rows = [];

  districtNames.forEach((districtName) => {
    const districtCases = cases.filter((row) => String(row.District || '').trim() === districtName);
    rows.push(buildSummaryRow_(`${districtName} | Initial`, districtCases.filter((row) => row.CaseType === CASE_TYPES.initial)));
    rows.push(buildSummaryRow_(`${districtName} | Re-evaluation`, districtCases.filter((row) => row.CaseType === CASE_TYPES.reevaluation)));
  });

  return rows;
}

function buildSummaryRow_(label, cases) {
  const summary = getDashboardSummary_(cases);
  return {
    label,
    totalActive: summary.totalActive,
    overdue: summary.overdue,
    dueThisWeek: summary.dueThisWeek,
    dueThisMonth: summary.dueThisMonth,
    completed: summary.completed,
  };
}

function applySummarySheetFormatting_(sheet, startRow, rowCount, columnCount) {
  const range = sheet.getRange(startRow, 1, rowCount, columnCount);
  const values = range.getValues();
  const backgrounds = values.map((row) => {
    const output = new Array(columnCount).fill('#ffffff');
    if (Number(row[2]) > 0) {
      output[2] = getSettingValue_('RedDeadlineColor', '#ff9999');
    }
    if (Number(row[3]) > 0) {
      output[3] = '#fff2cc';
    }
    if (Number(row[4]) > 0) {
      output[4] = getSettingValue_('PinkDeadlineColor', '#f4cccc');
    }
    if (Number(row[5]) > 0) {
      output[5] = '#d9ead3';
    }
    return output;
  });
  range.setBackgrounds(backgrounds);
}

function getDashboardHeaders_() {
  return [
    'Case ID',
    'Case Type',
    'Student Name',
    'Student ID',
    'Grade Level',
    'Campus',
    'District',
    'Eval Lead',
    'Referral Date',
    'Response Due Date',
    'Consent Date',
    'Evaluation Due Date',
    'Evaluation Date',
    'ARD Due Date',
    'ARD Date',
    'Status',
    ...SERVICE_FIELDS.map((field) => SERVICE_LABELS[field] || field),
    'Uploads',
    'Notes',
  ];
}

function buildDashboardRecord_(row, documents) {
  const evaluationDueDate = getEvaluationDueDate_(row, {
    projectedFiiEDueDate: row.ProjectedFIIEDueDate,
  });
  return {
    values: [
      row.CaseID,
      row.CaseType,
      row.StudentName,
      row.StudentID,
      row.GradeLevel || '',
      row.Campus,
      row.District,
      row.LeadEvaluator,
      row.ReferralReceivedDate || '',
      row.ResponseDueDate || '',
      row.ActualConsentDate || '',
      evaluationDueDate || '',
      row.ActualFIIEDate || '',
      row.ProjectedARDDueDate || '',
      row.ActualARDDate || '',
      row.Status,
      ...SERVICE_FIELDS.map((field) => (Number(row[field]) === 1 ? 'Yes' : 'No')),
      documents.map((item) => `${item.DocumentLabel}|${item.DocumentPath}`).join('\n'),
      buildCombinedNotes_(row.ServiceNotes, row.VarianceExplanation),
    ],
    status: row.Status,
    responseDueDate: row.ResponseDueDate,
    responseCompletedDate: row.ResponseSentDate,
    consentDueDate: row.ProjectedConsentDate,
    consentCompletedDate: row.ActualConsentDate,
    evaluationDueDate,
    evaluationCompletedDate: row.ActualFIIEDate,
    ardDueDate: row.ProjectedARDDueDate,
    ardCompletedDate: row.ActualARDDate,
  };
}

function buildCombinedNotes_(serviceNotes, varianceExplanation) {
  const parts = [];
  if (String(serviceNotes || '').trim()) {
    parts.push(`Service Notes: ${String(serviceNotes).trim()}`);
  }
  if (String(varianceExplanation || '').trim()) {
    parts.push(`Variance: ${String(varianceExplanation).trim()}`);
  }
  return parts.join('\n');
}

function getDistrictDashboardSheetName_(districtName) {
  const cleanedName = String(districtName || '')
    .replace(/[\[\]\\/?*:]/g, '-')
    .trim();
  return `${DISTRICT_DASHBOARD_PREFIX}${cleanedName}`.slice(0, 99) || `${DISTRICT_DASHBOARD_PREFIX}District`;
}

function ensureWorkbookScaffold_(options) {
  const settings = Object.assign({
    seedReferenceData: false,
    syncCalendar: false,
    showSheets: false,
    applyProtection: false,
  }, options || {});

  ensureSheet_(SHEETS.cases, CASE_HEADERS);
  ensureSheet_(SHEETS.archive, ARCHIVE_HEADERS);
  ensureSheet_(SHEETS.documents, DOCUMENT_HEADERS);
  ensureSheet_(SHEETS.tests, TEST_HEADERS);
  ensureSheet_(SHEETS.districts, DISTRICT_HEADERS);
  ensureSheet_(SHEETS.campuses, CAMPUS_HEADERS);
  ensureSheet_(SHEETS.evaluators, EVALUATOR_HEADERS);
  ensureSheet_(SHEETS.calendars, CALENDAR_BASE_HEADERS);
  ensureSheet_(SHEETS.settings, SETTINGS_HEADERS);
  ensureSheet_(SHEETS.audit, AUDIT_HEADERS);
  ensureSheet_(SHEETS.dashboard, ['SPED Status Reports Dashboard']);
  ensureSheet_(SHEETS.summaryCaseType, ['Summary By Case Type']);
  ensureSheet_(SHEETS.summaryEvaluator, ['Summary By Evaluator']);
  ensureSheet_(SHEETS.summaryDistrictCaseType, ['Summary By District And Case Type']);

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
    const existingHeaders = sheet.getLastColumn() > 0 ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0] : [];
    const needsHeaders = headers.some((header, index) => existingHeaders[index] !== header);
    if (needsHeaders) {
      const preservedRows = getPreservedSheetRows_(sheet, existingHeaders, headers);
      sheet.clear();
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
      if (preservedRows.length) {
        sheet.getRange(2, 1, preservedRows.length, headers.length).setValues(preservedRows);
      }
      sheet.setFrozenRows(1);
    }
  }

  applySheetColumnFormats_(sheetName, sheet);
  return sheet;
}

function getPreservedSheetRows_(sheet, existingHeaders, targetHeaders) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2 || !existingHeaders.length) {
    return [];
  }

  const values = sheet.getRange(2, 1, lastRow - 1, existingHeaders.length).getValues();
  const headerMap = getHeaderMap_(existingHeaders);

  return values.map((row) => targetHeaders.map((header) => {
    if (headerMap[header] !== undefined) {
      return row[headerMap[header]];
    }
    return '';
  }));
}

function applySheetColumnFormats_(sheetName, sheet) {
  if (
    sheetName !== SHEETS.cases &&
    sheetName !== SHEETS.archive &&
    sheetName !== SHEETS.tests
  ) {
    return;
  }

  if (sheet.getMaxRows() < 2) {
    return;
  }

  if (sheetName === SHEETS.tests) {
    const headerMap = getHeaderMap_(TEST_HEADERS);
    const dateColumns = [
      'ReferralReceivedDate',
      'ActualConsentDate',
      'ActualFIIEDate',
      'ReevalDueDate',
      'ExpectedResponseDueDate',
      'ExpectedProjectedConsentDate',
      'ExpectedProjectedFIIEDueDate',
      'ExpectedProjectedARDDueDate',
      'ActualResponseDueDate',
      'ActualProjectedConsentDate',
      'ActualProjectedFIIEDueDate',
      'ActualProjectedARDDueDate',
    ];
    dateColumns.forEach((header) => {
      sheet.getRange(2, headerMap[header] + 1, Math.max(sheet.getMaxRows() - 1, 1), 1).setNumberFormat('mm/dd/yyyy');
    });
    return;
  }

  const headerMap = getHeaderMap_(sheetName === SHEETS.archive ? ARCHIVE_HEADERS : CASE_HEADERS);
  sheet.getRange(2, headerMap.StudentID + 1, Math.max(sheet.getMaxRows() - 1, 1), 1).setNumberFormat('@');
  sheet.getRange(2, headerMap.GradeLevel + 1, Math.max(sheet.getMaxRows() - 1, 1), 1).setNumberFormat('@');
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
  getManagedSheetNames_().forEach((name) => {
    const sheet = SpreadsheetApp.getActive().getSheetByName(name);
    if (sheet) {
      applyProtection_(sheet);
    }
  });
}

function showAllSheets_() {
  getManagedSheetNames_().forEach((name) => {
    const sheet = SpreadsheetApp.getActive().getSheetByName(name);
    if (sheet) {
      sheet.showSheet();
    }
  });
}

function getManagedSheetNames_() {
  const districtDashboardNames = getActiveColumnValues_(SHEETS.districts, 'District').map((districtName) => (
    getDistrictDashboardSheetName_(districtName)
  ));
  return [...new Set(Object.values(SHEETS).concat(districtDashboardNames))];
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

function getSettingValue_(settingKey, fallbackValue) {
  const rows = getTableRows_(SHEETS.settings, SETTINGS_HEADERS);
  const match = rows.find((row) => String(row.SettingKey).trim() === String(settingKey).trim());
  if (!match || match.SettingValue === '') {
    return fallbackValue;
  }
  return match.SettingValue;
}

function getNumericSetting_(settingKey, fallbackValue) {
  const value = Number(getSettingValue_(settingKey, fallbackValue));
  return Number.isFinite(value) ? value : fallbackValue;
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
  seedSettings_();
  seedDueDateTests_();

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

function seedDueDateTests_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEETS.tests);
  if (sheet.getLastRow() > 1) {
    return;
  }

  const activeDistricts = getActiveColumnValues_(SHEETS.districts, 'District');
  const districtName = activeDistricts[0] || 'Sample ISD';
  const seedRows = [
    {
      ScenarioName: 'Initial referral baseline',
      CaseType: CASE_TYPES.initial,
      District: districtName,
      ReferralReceivedDate: createLocalDate_(2026, 1, 12),
      ActualConsentDate: '',
      ActualFIIEDate: '',
      ReevalDueDate: '',
      ExpectedResponseDueDate: '',
      ExpectedProjectedConsentDate: '',
      ExpectedProjectedFIIEDueDate: '',
      ExpectedProjectedARDDueDate: '',
      Notes: 'Fill expected dates after you verify your district calendar.',
    },
    {
      ScenarioName: 'Re-evaluation baseline',
      CaseType: CASE_TYPES.reevaluation,
      District: districtName,
      ReferralReceivedDate: '',
      ActualConsentDate: '',
      ActualFIIEDate: '',
      ReevalDueDate: createLocalDate_(2026, 10, 15),
      ExpectedResponseDueDate: '',
      ExpectedProjectedConsentDate: '',
      ExpectedProjectedFIIEDueDate: '',
      ExpectedProjectedARDDueDate: '',
      Notes: 'Use this row to verify ARD projection from a re-evaluation due date.',
    },
  ];

  const values = seedRows.map((row) => TEST_HEADERS.map((header) => (row[header] === undefined ? '' : row[header])));
  sheet.getRange(2, 1, values.length, TEST_HEADERS.length).setValues(values);
}

function seedSettings_() {
  ensureSettingRow_('UpcomingWarningDays', '30', 'How many days before an incomplete deadline should turn pink.');
  ensureSettingRow_('PinkDeadlineColor', '#f4cccc', 'Color used when a deadline is within the warning window.');
  ensureSettingRow_('RedDeadlineColor', '#ff9999', 'Color used when a due date is missed.');
  ensureSettingRow_('UploadsFolderIdOverride', '', 'Optional Google Drive folder ID used for uploaded documents.');
}

function ensureSettingRow_(settingKey, settingValue, description) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEETS.settings);
  const rows = getTableRows_(SHEETS.settings, SETTINGS_HEADERS);
  if (rows.some((row) => String(row.SettingKey).trim() === String(settingKey).trim())) {
    return;
  }

  appendRow_(SHEETS.settings, {
    SettingKey: settingKey,
    SettingValue: settingValue,
    Description: description,
  }, SETTINGS_HEADERS);
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

function rewriteSheetRows_(sheetName, headers, rows) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sheet.setFrozenRows(1);

  if (rows.length) {
    const values = rows.map((row) => headers.map((header) => (row[header] === undefined ? '' : row[header])));
    sheet.getRange(2, 1, values.length, headers.length).setValues(values);
  }

  applySheetColumnFormats_(sheetName, sheet);
}

function refreshDueDateTests_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEETS.tests);
  const rows = getTableRows_(SHEETS.tests, TEST_HEADERS);
  const actualHeaders = [
    'ActualResponseDueDate',
    'ActualProjectedConsentDate',
    'ActualProjectedFIIEDueDate',
    'ActualProjectedARDDueDate',
  ];
  const headerMap = getHeaderMap_(TEST_HEADERS);
  const actualOutput = [];
  const resultValues = [];
  const resultBackgrounds = [];
  let passCount = 0;
  let failCount = 0;
  let checkCount = 0;

  rows.forEach((row, index) => {
    const timeline = buildProjectedDates_({
      caseType: row.CaseType,
      district: row.District,
      referralReceivedDate: row.ReferralReceivedDate,
      actualConsentDate: row.ActualConsentDate,
      actualFIIEDate: row.ActualFIIEDate,
      reevalDueDate: row.ReevalDueDate,
    });

    const actualValues = [
      timeline.responseDueDate || '',
      timeline.projectedConsentDate || '',
      timeline.projectedFiiEDueDate || '',
      timeline.projectedArdDueDate || '',
    ];

    actualOutput.push(actualValues);

    const comparisons = [
      ['ExpectedResponseDueDate', timeline.responseDueDate],
      ['ExpectedProjectedConsentDate', timeline.projectedConsentDate],
      ['ExpectedProjectedFIIEDueDate', timeline.projectedFiiEDueDate],
      ['ExpectedProjectedARDDueDate', timeline.projectedArdDueDate],
    ];

    let hasExpected = false;
    let mismatch = false;
    comparisons.forEach(([expectedHeader, actualValue]) => {
      const expectedValue = parseDate_(row[expectedHeader]);
      if (!expectedValue) {
        return;
      }
      hasExpected = true;
      const actualDate = parseDate_(actualValue);
      if (!actualDate || !sameDay_(expectedValue, actualDate)) {
        mismatch = true;
      }
    });

    let result = 'CHECK';
    let background = '#fff2cc';

    if (hasExpected && !mismatch) {
      result = 'PASS';
      background = '#d9ead3';
      passCount += 1;
    } else if (hasExpected && mismatch) {
      result = 'FAIL';
      background = '#f4cccc';
      failCount += 1;
    } else {
      checkCount += 1;
    }

    resultValues.push([result]);
    resultBackgrounds.push([background]);
  });

  if (rows.length) {
    sheet.getRange(2, headerMap[actualHeaders[0]] + 1, rows.length, actualHeaders.length).setValues(actualOutput);
    sheet.getRange(2, headerMap.Result + 1, rows.length, 1)
      .setValues(resultValues)
      .setBackgrounds(resultBackgrounds)
      .setHorizontalAlignment('center')
      .setFontWeight('bold');
  }

  return {
    passCount,
    failCount,
    checkCount,
  };
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

function setPendingCaseId_(caseId) {
  PropertiesService.getDocumentProperties().setProperty(DOCUMENT_PROPERTIES.pendingCaseId, String(caseId || '').trim());
}

function consumePendingCaseId_() {
  const properties = PropertiesService.getDocumentProperties();
  const caseId = String(properties.getProperty(DOCUMENT_PROPERTIES.pendingCaseId) || '').trim();
  if (caseId) {
    properties.deleteProperty(DOCUMENT_PROPERTIES.pendingCaseId);
  }
  return caseId;
}

function buildStoredCaseRow_(caseId, payload, timeline, createdAt, updatedAt) {
  const responseSentDate = parseDate_(payload.responseSentDate);
  const actualConsentDate = parseDate_(payload.actualConsentDate);
  const actualFiiEDate = parseDate_(payload.actualFIIEDate);
  const evaluationStartedDate = parseDate_(payload.evaluationStartedDate);
  const ardScheduledDate = parseDate_(payload.ardScheduledDate);
  const actualArdDate = parseDate_(payload.actualARDDate);
  const reevalDueDate = parseDate_(payload.reevalDueDate);
  const manualPrimaryDeadline = parseDate_(payload.overridePrimaryDeadline);

  const row = {
    CaseID: caseId,
    CaseType: payload.caseType,
    StudentName: payload.studentName,
    StudentID: normalizeStudentId_(payload.studentId),
    GradeLevel: String(payload.gradeLevel || '').trim(),
    DOB: parseDate_(payload.dob),
    Campus: payload.campus,
    District: payload.district,
    LeadEvaluator: payload.leadEvaluator,
    Status: determineStatus_(
      responseSentDate,
      actualConsentDate,
      evaluationStartedDate,
      actualFiiEDate,
      ardScheduledDate,
      actualArdDate
    ),
    ReferralReceivedDate: parseDate_(payload.referralReceivedDate),
    ResponseDueDate: timeline.responseDueDate,
    ResponseSentDate: responseSentDate,
    ProjectedConsentDate: timeline.projectedConsentDate,
    ActualConsentDate: actualConsentDate,
    ProjectedFIIEDueDate: timeline.projectedFiiEDueDate,
    ActualFIIEDate: actualFiiEDate,
    EvaluationStartedDate: evaluationStartedDate,
    ARDScheduledDate: ardScheduledDate,
    ProjectedARDDueDate: timeline.projectedArdDueDate,
    ActualARDDate: actualArdDate,
    ReevalDueDate: reevalDueDate,
    ServiceNotes: payload.serviceNotes || '',
    VarianceExplanation: payload.varianceExplanation || '',
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

  validateDateOrder_(payload);

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

function validateVarianceNotes_(payload, timeline) {
  const notes = String(payload.varianceExplanation || '').trim();
  const lateDates = getLateDateWarnings_(payload, timeline);

  if (lateDates.length && !notes) {
    throw new Error(`Notes are required when actual dates are later than due dates: ${lateDates.join(', ')}.`);
  }
}

function validateDateOrder_(payload) {
  const referralReceivedDate = parseDate_(payload.referralReceivedDate);
  const responseSentDate = parseDate_(payload.responseSentDate);
  const actualConsentDate = parseDate_(payload.actualConsentDate);
  const evaluationStartedDate = parseDate_(payload.evaluationStartedDate);
  const actualFiiEDate = parseDate_(payload.actualFIIEDate);
  const ardScheduledDate = parseDate_(payload.ardScheduledDate);
  const actualArdDate = parseDate_(payload.actualARDDate);

  validateDateSequence_('Response Sent Date', responseSentDate, 'Referral Received Date', referralReceivedDate);
  validateDateSequence_('Consent Date', actualConsentDate, 'Referral Received Date', referralReceivedDate);
  validateDateSequence_('Consent Date', actualConsentDate, 'Response Sent Date', responseSentDate);
  validateDateSequence_('Evaluation Started Date', evaluationStartedDate, 'Consent Date', actualConsentDate);
  validateDateSequence_('Evaluation Date', actualFiiEDate, 'Consent Date', actualConsentDate);
  validateDateSequence_('Evaluation Date', actualFiiEDate, 'Evaluation Started Date', evaluationStartedDate);
  validateDateSequence_('ARD Scheduled Date', ardScheduledDate, 'Evaluation Date', actualFiiEDate);
  validateDateSequence_('ARD Date', actualArdDate, 'Evaluation Date', actualFiiEDate);
  validateDateSequence_('ARD Date', actualArdDate, 'ARD Scheduled Date', ardScheduledDate);

  if (payload.caseType === CASE_TYPES.initial && !referralReceivedDate) {
    return;
  }
}

function validateDateSequence_(laterLabel, laterDate, earlierLabel, earlierDate) {
  if (!laterDate || !earlierDate) {
    return;
  }
  if (laterDate.getTime() < earlierDate.getTime()) {
    throw new Error(`${laterLabel} cannot be earlier than ${earlierLabel}.`);
  }
}

function getLateDateWarnings_(payload, timeline) {
  const lateDates = [];
  const responseSentDate = parseDate_(payload.responseSentDate);
  const actualConsentDate = parseDate_(payload.actualConsentDate);
  const actualFiiEDate = parseDate_(payload.actualFIIEDate);
  const actualArdDate = parseDate_(payload.actualARDDate);
  const responseDueDate = parseDate_(timeline.responseDueDate);
  const consentDueDate = parseDate_(timeline.projectedConsentDate);
  const evaluationDueDate = parseDate_(getEvaluationDueDate_(payload, timeline));
  const ardDueDate = parseDate_(timeline.projectedArdDueDate);

  if (isAfterDate_(responseSentDate, responseDueDate)) {
    lateDates.push('Response Sent Date');
  }
  if (isAfterDate_(actualConsentDate, consentDueDate)) {
    lateDates.push('Consent Date');
  }
  if (isAfterDate_(actualFiiEDate, evaluationDueDate)) {
    lateDates.push('Evaluation Date');
  }
  if (isAfterDate_(actualArdDate, ardDueDate)) {
    lateDates.push('ARD Date');
  }

  return lateDates;
}

function restoreArchivedCaseById_(caseId) {
  const archivedRecord = findArchivedCaseRecord_(caseId);
  if (!archivedRecord) {
    throw new Error(`Archived case not found: ${caseId}`);
  }

  const duplicate = findOpenDuplicateCase_(archivedRecord.row, archivedRecord.row.CaseType);
  if (duplicate) {
    throw new Error(buildDuplicateCaseMessage_(duplicate, archivedRecord.row.CaseType));
  }

  appendRow_(SHEETS.cases, archivedRecord.row, CASE_HEADERS);

  const remainingArchivedRows = getTableRows_(SHEETS.archive, ARCHIVE_HEADERS).filter((row) => row.CaseID !== caseId);
  rewriteSheetRows_(SHEETS.archive, ARCHIVE_HEADERS, remainingArchivedRows);

  appendAuditRows_([
    buildAuditRow_(caseId, 'Restore', 'Status', 'Archived', archivedRecord.row.Status),
  ]);

  refreshDashboard_(true);
  return caseId;
}

function removeCaseDocument(caseId, documentId, deleteFromDrive) {
  ensureWorkbookReady_();

  if (!String(caseId || '').trim() || !String(documentId || '').trim()) {
    throw new Error('A case and document selection are required.');
  }

  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEETS.documents);
  const rows = getTableRows_(SHEETS.documents, DOCUMENT_HEADERS);
  const target = rows.find((row) => row.CaseID === caseId && row.DocumentID === documentId);

  if (!target) {
    throw new Error('That document link no longer exists. Refresh the case and try again.');
  }

  if (deleteFromDrive) {
    deleteDriveFileIfPossible_(target.DocumentPath);
  }

  const remaining = rows.filter((row) => !(row.CaseID === caseId && row.DocumentID === documentId));
  rewriteSheetRows_(SHEETS.documents, DOCUMENT_HEADERS, remaining);

  const previousDocumentsText = buildCaseDocumentsText_(rows.filter((row) => row.CaseID === caseId));
  const nextDocumentsText = buildCaseDocumentsText_(remaining.filter((row) => row.CaseID === caseId));
  logDocumentsUpdate_(caseId, previousDocumentsText, nextDocumentsText);
  refreshDashboard_(false);

  return toHtmlSafeObject_({
    documents: normalizeDocumentsForUi_(remaining.filter((row) => row.CaseID === caseId)),
    documentsText: nextDocumentsText,
  });
}

function deleteDriveFileIfPossible_(documentPath) {
  const fileId = extractDriveFileId_(documentPath);
  if (!fileId) {
    return false;
  }

  try {
    DriveApp.getFileById(fileId).setTrashed(true);
    return true;
  } catch (error) {
    throw new Error('The Google Drive file could not be deleted with the current permissions.');
  }
}

function extractDriveFileId_(documentPath) {
  const text = String(documentPath || '').trim();
  if (!text) {
    return '';
  }

  const patterns = [
    /\/d\/([a-zA-Z0-9_-]{20,})/,
    /[?&]id=([a-zA-Z0-9_-]{20,})/,
    /^([a-zA-Z0-9_-]{20,})$/,
  ];

  for (let index = 0; index < patterns.length; index += 1) {
    const match = text.match(patterns[index]);
    if (match) {
      return match[1];
    }
  }

  return '';
}

function logCaseCreation_(row) {
  const auditRows = CASE_HEADERS
    .filter((header) => header !== 'CreatedAt' && header !== 'UpdatedAt')
    .map((header) => buildAuditRow_(row.CaseID, 'Create', header, '', row[header]));
  appendAuditRows_(auditRows);
}

function logCaseUpdate_(existingRow, updatedRow) {
  const auditRows = CASE_HEADERS
    .filter((header) => header !== 'CreatedAt' && header !== 'UpdatedAt')
    .filter((header) => auditStringify_(existingRow[header]) !== auditStringify_(updatedRow[header]))
    .map((header) => buildAuditRow_(updatedRow.CaseID, 'Update', header, existingRow[header], updatedRow[header]));
  appendAuditRows_(auditRows);
}

function logDocumentsUpdate_(caseId, previousDocuments, nextDocuments) {
  if (previousDocuments === nextDocuments) {
    return;
  }
  appendAuditRows_([
    buildAuditRow_(caseId, 'Documents', 'DocumentLinks', previousDocuments, nextDocuments),
  ]);
}

function buildAuditRow_(caseId, action, fieldName, oldValue, newValue) {
  return {
    AuditID: `AUD-${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss-SSS')}-${Math.floor(Math.random() * 1000)}`,
    CaseID: caseId,
    Action: action,
    FieldName: fieldName,
    OldValue: auditStringify_(oldValue),
    NewValue: auditStringify_(newValue),
    ChangedAt: new Date(),
  };
}

function appendAuditRows_(auditRows) {
  if (!auditRows.length) {
    return;
  }

  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEETS.audit);
  const values = auditRows.map((row) => AUDIT_HEADERS.map((header) => (row[header] === undefined ? '' : row[header])));
  sheet.getRange(sheet.getLastRow() + 1, 1, values.length, AUDIT_HEADERS.length).setValues(values);
}

function auditStringify_(value) {
  if (value === undefined || value === null || value === '') {
    return '';
  }
  if (Object.prototype.toString.call(value) === '[object Date]' && !Number.isNaN(value.getTime())) {
    return formatDateTime_(value) || formatDate_(value);
  }
  if (typeof value === 'string' && (/^\d{4}-\d{2}-\d{2}$/.test(value) || /^\d{1,2}\/\d{1,2}\/\d{4}$/.test(value))) {
    return formatDate_(value);
  }
  return String(value);
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

function determineStatus_(
  responseSentDate,
  actualConsentDate,
  evaluationStartedDate,
  actualFiiEDate,
  ardScheduledDate,
  actualArdDate
) {
  if (actualArdDate) {
    return STATUSES.completed;
  }
  if (ardScheduledDate) {
    return STATUSES.ardScheduled;
  }
  if (actualFiiEDate) {
    return STATUSES.evaluationComplete;
  }
  if (evaluationStartedDate) {
    return STATUSES.evaluationInProgress;
  }
  if (actualConsentDate) {
    return STATUSES.consentReceived;
  }
  if (responseSentDate) {
    return STATUSES.responseSent;
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

function getEvaluationDueDate_(row, timeline) {
  if ((row.caseType || row.CaseType) === CASE_TYPES.reevaluation) {
    return row.reevalDueDate || row.ReevalDueDate || '';
  }
  return row.projectedFiiEDueDate || row.ProjectedFIIEDueDate || (timeline ? timeline.projectedFiiEDueDate : '');
}

function findOpenDuplicateCase_(candidate, caseType, ignoreCaseId) {
  const candidateIdentity = buildCaseIdentity_(candidate);
  const rows = getTableRows_(SHEETS.cases, CASE_HEADERS);

  for (let index = 0; index < rows.length; index += 1) {
    const row = rows[index];
    if (row.CaseType !== caseType) {
      continue;
    }
    if (row.Status === STATUSES.completed) {
      continue;
    }
    if (ignoreCaseId && row.CaseID === ignoreCaseId) {
      continue;
    }

    const rowIdentity = buildCaseIdentity_(row);
    if (candidateIdentity.studentId && rowIdentity.studentId && candidateIdentity.studentId === rowIdentity.studentId) {
      return {
        caseId: row.CaseID,
        matchReason: 'Student ID',
        studentName: row.StudentName,
        studentId: row.StudentID,
      };
    }

    if (
      candidateIdentity.studentName &&
      candidateIdentity.dobKey &&
      candidateIdentity.campus &&
      candidateIdentity.studentName === rowIdentity.studentName &&
      candidateIdentity.dobKey === rowIdentity.dobKey &&
      candidateIdentity.campus === rowIdentity.campus
    ) {
      return {
        caseId: row.CaseID,
        matchReason: 'Student Name + DOB + Campus',
        studentName: row.StudentName,
        studentId: row.StudentID,
      };
    }
  }

  return null;
}

function buildDuplicateCaseMessage_(duplicate, caseType) {
  return `A possible duplicate open ${caseType} case already exists: ${duplicate.caseId} (${duplicate.studentName || 'Unknown Student'}${duplicate.studentId ? `, Student ID ${duplicate.studentId}` : ''}). Match found by ${duplicate.matchReason}.`;
}

function buildCaseIdentity_(source) {
  return {
    studentId: normalizeStudentId_(source.StudentID !== undefined ? source.StudentID : source.studentId),
    studentName: normalizeComparisonText_(source.StudentName !== undefined ? source.StudentName : source.studentName),
    dobKey: getDateKey_(source.DOB !== undefined ? source.DOB : source.dob),
    campus: normalizeComparisonText_(source.Campus !== undefined ? source.Campus : source.campus),
  };
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
    if (!String(values[index][headerMap.CaseID] || '').trim()) {
      continue;
    }
    if (String(values[index][headerMap.CaseID]).trim() === String(caseId).trim()) {
      return {
        row: toObject_(CASE_HEADERS, values[index]),
        rowIndex: index + 2,
      };
    }
  }

  return null;
}

function findArchivedCaseRecord_(caseId) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEETS.archive);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return null;
  }

  const headerMap = getHeaderMap_(ARCHIVE_HEADERS);
  const values = sheet.getRange(2, 1, lastRow - 1, ARCHIVE_HEADERS.length).getValues();

  for (let index = 0; index < values.length; index += 1) {
    if (!String(values[index][headerMap.CaseID] || '').trim()) {
      continue;
    }
    if (String(values[index][headerMap.CaseID]).trim() === String(caseId).trim()) {
      return {
        row: toObject_(ARCHIVE_HEADERS, values[index]),
        rowIndex: index + 2,
      };
    }
  }

  return null;
}

function getSelectedCaseReference_() {
  const spreadsheet = SpreadsheetApp.getActive();
  const range = spreadsheet.getActiveRange();
  if (!range) {
    return null;
  }

  const caseId = String(range.getSheet().getRange(range.getRow(), 1).getDisplayValue() || '').trim();
  if (!caseId) {
    return null;
  }

  if (findCaseRecord_(caseId)) {
    return {
      caseId,
      location: 'active',
    };
  }

  if (findArchivedCaseRecord_(caseId)) {
    return {
      caseId,
      location: 'archive',
    };
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

function uploadFilesToCase(caseId, files) {
  ensureWorkbookReady_();

  if (!caseId) {
    throw new Error('Save or load the case before uploading files.');
  }

  const uploads = Array.isArray(files) ? files : [];
  if (!uploads.length) {
    return toHtmlSafeObject_({
      documents: normalizeDocumentsForUi_(getCaseDocuments_(caseId)),
      documentsText: getCaseDocumentsText_(caseId),
    });
  }

  const folder = getCaseUploadFolder_(caseId);
  const previousDocumentsText = getCaseDocumentsText_(caseId);
  const existingDocuments = getCaseDocuments_(caseId).map((item) => ({
    label: item.DocumentLabel,
    path: item.DocumentPath,
  }));

  uploads.forEach((file) => {
    if (!file || !file.name || !file.content) {
      return;
    }
    const blob = Utilities.newBlob(
      Utilities.base64Decode(file.content),
      file.mimeType || MimeType.PLAIN_TEXT,
      file.name
    );
    const createdFile = folder.createFile(blob);
    existingDocuments.push({
      label: file.name,
      path: createdFile.getUrl(),
    });
  });

  const nextDocumentsText = buildCaseDocumentsText_(existingDocuments);
  replaceCaseDocuments_(caseId, nextDocumentsText);
  logDocumentsUpdate_(caseId, previousDocumentsText, nextDocumentsText);
  refreshDashboard_(false);

  return toHtmlSafeObject_({
    documents: normalizeDocumentsForUi_(getCaseDocuments_(caseId)),
    documentsText: getCaseDocumentsText_(caseId),
  });
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
  return buildCaseDocumentsText_(getCaseDocuments_(caseId));
}

function buildCaseDocumentsText_(documents) {
  return (documents || [])
    .map((row) => `${row.DocumentLabel || row.label}|${row.DocumentPath || row.path}`)
    .join('\n');
}

function getCaseDocuments_(caseId) {
  const rows = getTableRows_(SHEETS.documents, DOCUMENT_HEADERS);
  return rows.filter((row) => row.CaseID === caseId);
}

function normalizeDocumentsForUi_(documents) {
  return (documents || []).map((row) => ({
    DocumentID: row.DocumentID || '',
    CaseID: row.CaseID || '',
    DocumentLabel: row.DocumentLabel || '',
    DocumentPath: row.DocumentPath || '',
    AddedAt: formatDateTime_(row.AddedAt),
  }));
}

function getCaseDocumentsMap_() {
  const rows = getTableRows_(SHEETS.documents, DOCUMENT_HEADERS);
  return rows.reduce((acc, row) => {
    if (!acc[row.CaseID]) {
      acc[row.CaseID] = [];
    }
    acc[row.CaseID].push(row);
    return acc;
  }, {});
}

function getCaseUploadFolder_(caseId) {
  const parentFolder = getUploadsRootFolder_();
  const existingFolders = parentFolder.getFoldersByName(caseId);
  return existingFolders.hasNext() ? existingFolders.next() : parentFolder.createFolder(caseId);
}

function getUploadsRootFolder_() {
  const folderId = String(getSettingValue_('UploadsFolderIdOverride', UPLOADS_FOLDER_ID) || '').trim();
  if (folderId) {
    return DriveApp.getFolderById(folderId);
  }

  const spreadsheetFile = DriveApp.getFileById(SpreadsheetApp.getActive().getId());
  const parentFolders = spreadsheetFile.getParents();
  const folderName = 'SPED Status Report Uploads';

  if (parentFolders.hasNext()) {
    const parent = parentFolders.next();
    const existing = parent.getFoldersByName(folderName);
    return existing.hasNext() ? existing.next() : parent.createFolder(folderName);
  }

  const rootFolder = DriveApp.getRootFolder();
  const existing = rootFolder.getFoldersByName(folderName);
  return existing.hasNext() ? existing.next() : rootFolder.createFolder(folderName);
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
    ResponseSentDate: formatDate_(row.ResponseSentDate),
    ProjectedConsentDate: formatDate_(row.ProjectedConsentDate),
    ActualConsentDate: formatDate_(row.ActualConsentDate),
    ProjectedFIIEDueDate: formatDate_(row.ProjectedFIIEDueDate),
    ActualFIIEDate: formatDate_(row.ActualFIIEDate),
    EvaluationStartedDate: formatDate_(row.EvaluationStartedDate),
    ARDScheduledDate: formatDate_(row.ARDScheduledDate),
    ProjectedARDDueDate: formatDate_(row.ProjectedARDDueDate),
    ActualARDDate: formatDate_(row.ActualARDDate),
    ReevalDueDate: formatDate_(row.ReevalDueDate),
    ManualPrimaryDeadline: formatDate_(row.ManualPrimaryDeadline),
    PrimaryDeadline: formatDate_(row.PrimaryDeadline),
    CreatedAt: formatDateTime_(row.CreatedAt),
    UpdatedAt: formatDateTime_(row.UpdatedAt),
    ServiceNotes: row.ServiceNotes || '',
    VarianceExplanation: row.VarianceExplanation || '',
  });

  output.services = {};
  SERVICE_FIELDS.forEach((field) => {
    output.services[field] = Number(row[field]) === 1;
  });

  return output;
}

function toHtmlSafeObject_(value) {
  return JSON.parse(JSON.stringify(value, (key, currentValue) => {
    if (currentValue === undefined) {
      return '';
    }
    if (Object.prototype.toString.call(currentValue) === '[object Date]' && !Number.isNaN(currentValue.getTime())) {
      return currentValue.toISOString();
    }
    return currentValue;
  }));
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

function isAfterDate_(left, right) {
  const leftDate = parseDate_(left);
  const rightDate = parseDate_(right);
  if (!leftDate || !rightDate) {
    return false;
  }
  return leftDate.getTime() > rightDate.getTime();
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

  if (/^\d+\.0+$/.test(text)) {
    return text.replace(/\.0+$/, '');
  }

  return text;
}

function normalizeComparisonText_(value) {
  return String(value === undefined || value === null ? '' : value)
    .trim()
    .replace(/\s+/g, ' ')
    .toLowerCase();
}

function getDateKey_(value) {
  const dateValue = parseDate_(value);
  return dateValue ? isoDateKey_(dateValue) : '';
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

function applyDashboardFormatting_(sheet, startRow, records) {
  const headers = getDashboardHeaders_();
  const range = sheet.getRange(startRow, 1, records.length, headers.length);
  const backgrounds = records.map(() => new Array(headers.length).fill('#ffffff'));
  const dueDateColors = {
    lateColor: getSettingValue_('RedDeadlineColor', '#ff9999'),
    warningColor: getSettingValue_('PinkDeadlineColor', '#f4cccc'),
    warningDays: getNumericSetting_('UpcomingWarningDays', 30),
  };
  const statusColumnIndex = headers.indexOf('Status');
  const responseDueColumnIndex = headers.indexOf('Response Due Date');
  const consentDateColumnIndex = headers.indexOf('Consent Date');
  const evaluationDueColumnIndex = headers.indexOf('Evaluation Due Date');
  const evaluationDateColumnIndex = headers.indexOf('Evaluation Date');
  const ardDueColumnIndex = headers.indexOf('ARD Due Date');
  const ardDateColumnIndex = headers.indexOf('ARD Date');
  const statusColors = {
    [STATUSES.referralReceived]: '#d9eaf7',
    [STATUSES.responseSent]: '#f4cccc',
    [STATUSES.consentReceived]: '#fff2cc',
    [STATUSES.evaluationInProgress]: '#fce5cd',
    [STATUSES.evaluationComplete]: '#d9ead3',
    [STATUSES.ardScheduled]: '#d0e0e3',
    [STATUSES.completed]: '#b6d7a8',
  };

  records.forEach((record, rowIndex) => {
    backgrounds[rowIndex][statusColumnIndex] = statusColors[record.status] || '#ffffff';
    applyDueDateColor_(backgrounds[rowIndex], responseDueColumnIndex, record.responseDueDate, record.responseCompletedDate, -1, dueDateColors);
    applyDueDateColor_(backgrounds[rowIndex], consentDateColumnIndex, record.consentDueDate, record.consentCompletedDate, consentDateColumnIndex, dueDateColors);
    applyDueDateColor_(backgrounds[rowIndex], evaluationDueColumnIndex, record.evaluationDueDate, record.evaluationCompletedDate, evaluationDateColumnIndex, dueDateColors);
    applyDueDateColor_(backgrounds[rowIndex], ardDueColumnIndex, record.ardDueDate, record.ardCompletedDate, ardDateColumnIndex, dueDateColors);
  });

  range.setBackgrounds(backgrounds);
}

function applyDueDateColor_(rowBackgrounds, dueColumnIndex, dueDate, completedDate, completedColumnIndex, colorConfig) {
  if (!dueDate) {
    return;
  }

  const due = parseDate_(dueDate);
  if (!due) {
    return;
  }

  if (completedDate) {
    if (isAfterDate_(completedDate, dueDate)) {
      const targetColumnIndex = completedColumnIndex >= 0 ? completedColumnIndex : dueColumnIndex;
      rowBackgrounds[targetColumnIndex] = colorConfig.lateColor;
    }
    return;
  }

  const today = normalizeDateForStorage_(new Date());
  const diffDays = Math.floor((due - today) / 86400000);
  if (diffDays < 0) {
    rowBackgrounds[dueColumnIndex] = colorConfig.lateColor;
    return;
  }
  if (diffDays <= colorConfig.warningDays) {
    rowBackgrounds[dueColumnIndex] = colorConfig.warningColor;
  }
}
