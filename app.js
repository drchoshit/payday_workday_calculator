const STORAGE_KEY = "payroll-app-config-v1";
const EXPORT_APP_ID = "payroll-auto-settlement";
const EXPORT_VERSION = 1;
const WEEKDAYS = ["일", "월", "화", "수", "목", "금", "토"];

const HEADER_CANDIDATES = {
  date: ["영업일자", "근무일자", "일자", "날짜"],
  name: ["이름", "성명", "근무자"],
  category: ["카테고리", "직급", "역할", "구분"],
  start: ["등원시각", "출근시각", "출근시간", "등원시간", "시작시각", "시작시간"],
  end: ["하원시각", "퇴근시각", "퇴근시간", "하원시간", "종료시각", "종료시간"],
  stayHours: ["체류시간(시간)", "체류시간", "근무시간", "시간환산", "시간"],
  note: ["비고", "메모"],
};

const DEFAULT_SETTINGS = {
  useWeeklyHoliday: true,
  useHolidayPremium: true,
  useOvertimePremium: true,
  useNightPremium: true,
  taxRate: 3.3,
  defaultHourlyWage: 10320,
  autoFillMissing: true,
  defaultMissingHours: 8,
  defaultStartTime: "09:00",
  defaultNightStart: "22:30",
  nightEndTime: "06:00",
  weeklyThresholdHours: 15,
};

const SHIFT_IDS = ["1T", "2T", "3T"];
const WEEKDAY_ORDER = [1, 2, 3, 4, 5, 6, 0];
const MIN_AUTO_SCHEDULE_MONTH = "2026-05";

const DEFAULT_SHIFT_TEMPLATES = [
  { id: "1T", label: "1타임", start: "07:30", end: "15:30" },
  { id: "2T", label: "2타임", start: "14:30", end: "22:30" },
  { id: "3T", label: "3타임", start: "21:30", end: "02:00" },
];

const DEFAULT_CAREER_LEVELS = [
  { id: "starter", name: "Starter", min: 0, max: 180, wage: 10320 },
  { id: "junior-b", name: "Junior B", min: 181, max: 340, wage: 11400 },
  { id: "junior-a", name: "Junior A", min: 341, max: 550, wage: 12500 },
  { id: "senior-b", name: "Senior B", min: 551, max: 790, wage: 13800 },
  { id: "senior-a", name: "Senior A", min: 791, max: 990, wage: 15200 },
  { id: "expert", name: "Expert", min: 991, max: null, wage: 17000 },
];

const persisted = loadPersistedConfig();

const state = {
  settings: { ...DEFAULT_SETTINGS, ...(persisted.settings || {}) },
  holidays: sanitizeHolidayList(persisted.holidays || []),
  employeeSettings: persisted.employeeSettings || {},
  entries: normalizePersistedEntries(persisted.entries || []),
  selectedMonth: cleanText(persisted.selectedMonth || ""),
  selectedEmployee: cleanText(persisted.selectedEmployee || ""),
  scheduleMonth: cleanText(persisted.scheduleMonth || ""),
  scheduleTemplates: normalizeScheduleTemplates(persisted.scheduleTemplates || []),
  scheduleFixedRules: normalizeScheduleFixedRules(persisted.scheduleFixedRules || []),
  scheduleAssignments: normalizeScheduleAssignments(persisted.scheduleAssignments || {}),
  scheduleExcludedManagers: normalizeNameList(persisted.scheduleExcludedManagers || []),
  careerLevels: normalizeCareerLevels(persisted.careerLevels || []),
  careerMonth: cleanText(persisted.careerMonth || ""),
  managerProfiles: normalizeManagerProfiles(persisted.managerProfiles || {}),
  selectedScheduleSlotId: cleanText(persisted.selectedScheduleSlotId || ""),
};

applyDefaultWageMigration();

const wonFormatter = new Intl.NumberFormat("ko-KR");

const dom = {
  tabs: Array.from(document.querySelectorAll(".tab-btn")),
  tabPanels: Array.from(document.querySelectorAll(".tab-panel")),
  fileInput: document.getElementById("fileInput"),
  monthSelect: document.getElementById("monthSelect"),
  summaryMonthSelect: document.getElementById("summaryMonthSelect"),
  downloadSummaryExcelBtn: document.getElementById("downloadSummaryExcelBtn"),
  recalculateBtn: document.getElementById("recalculateBtn"),
  saveAllBtn: document.getElementById("saveAllBtn"),
  loadAllBtn: document.getElementById("loadAllBtn"),
  loadStateFileInput: document.getElementById("loadStateFileInput"),
  uploadStatus: document.getElementById("uploadStatus"),

  taxRateInput: document.getElementById("taxRateInput"),
  defaultHourlyWageInput: document.getElementById("defaultHourlyWageInput"),
  autoFillMissingInput: document.getElementById("autoFillMissingInput"),
  defaultMissingHoursInput: document.getElementById("defaultMissingHoursInput"),
  defaultStartTimeInput: document.getElementById("defaultStartTimeInput"),
  defaultNightStartInput: document.getElementById("defaultNightStartInput"),
  nightEndTimeInput: document.getElementById("nightEndTimeInput"),
  weeklyThresholdInput: document.getElementById("weeklyThresholdInput"),

  holidayDateInput: document.getElementById("holidayDateInput"),
  addHolidayBtn: document.getElementById("addHolidayBtn"),
  holidayList: document.getElementById("holidayList"),

  employeeSettingsBody: document.getElementById("employeeSettingsBody"),
  missingDetectedBody: document.getElementById("missingDetectedBody"),
  missingAppliedBody: document.getElementById("missingAppliedBody"),

  missingEmployeeSelect: document.getElementById("missingEmployeeSelect"),
  missingEmployeeName: document.getElementById("missingEmployeeName"),
  missingDate: document.getElementById("missingDate"),
  missingCategory: document.getElementById("missingCategory"),
  missingStart: document.getElementById("missingStart"),
  missingEnd: document.getElementById("missingEnd"),
  missingHours: document.getElementById("missingHours"),
  missingNote: document.getElementById("missingNote"),
  addMissingShiftBtn: document.getElementById("addMissingShiftBtn"),

  previewBody: document.getElementById("previewBody"),

  statementEmployeeSelect: document.getElementById("statementEmployeeSelect"),
  printStatementBtn: document.getElementById("printStatementBtn"),
  statementMeta: document.getElementById("statementMeta"),
  statementBody: document.getElementById("statementBody"),
  weeklySummaryBody: document.getElementById("weeklySummaryBody"),
  totalHoursCell: document.getElementById("totalHoursCell"),
  basePayCell: document.getElementById("basePayCell"),
  allowanceCell: document.getElementById("allowanceCell"),
  grossPayCell: document.getElementById("grossPayCell"),
  netPayCell: document.getElementById("netPayCell"),

  summaryBody: document.getElementById("summaryBody"),
  summaryMeta: document.getElementById("summaryMeta"),
  summaryTotalHours: document.getElementById("summaryTotalHours"),
  summaryTotalGross: document.getElementById("summaryTotalGross"),
  summaryTotalNet: document.getElementById("summaryTotalNet"),
  monthlyArchiveBody: document.getElementById("monthlyArchiveBody"),

  scheduleMonthInput: document.getElementById("scheduleMonthInput"),
  generateScheduleBtn: document.getElementById("generateScheduleBtn"),
  printScheduleBtn: document.getElementById("printScheduleBtn"),
  shiftTemplateBody: document.getElementById("shiftTemplateBody"),
  fixedRuleMode: document.getElementById("fixedRuleMode"),
  fixedRuleWeekday: document.getElementById("fixedRuleWeekday"),
  fixedRuleDate: document.getElementById("fixedRuleDate"),
  fixedRuleShift: document.getElementById("fixedRuleShift"),
  fixedRuleManager: document.getElementById("fixedRuleManager"),
  fixedRuleNote: document.getElementById("fixedRuleNote"),
  addFixedRuleBtn: document.getElementById("addFixedRuleBtn"),
  fixedRuleBody: document.getElementById("fixedRuleBody"),
  newManagerNameInput: document.getElementById("newManagerNameInput"),
  addManagerRowBtn: document.getElementById("addManagerRowBtn"),
  allAvailabilityOnBtn: document.getElementById("allAvailabilityOnBtn"),
  allAvailabilityOffBtn: document.getElementById("allAvailabilityOffBtn"),
  shift3AvailabilityOnBtn: document.getElementById("shift3AvailabilityOnBtn"),
  shift3AvailabilityOffBtn: document.getElementById("shift3AvailabilityOffBtn"),
  scheduleManagerBody: document.getElementById("scheduleManagerBody"),
  scheduleCalendarHead: document.getElementById("scheduleCalendarHead"),
  scheduleCalendarBody: document.getElementById("scheduleCalendarBody"),
  scheduleHeadcountCalendarHead: document.getElementById("scheduleHeadcountCalendarHead"),
  scheduleHeadcountCalendarBody: document.getElementById("scheduleHeadcountCalendarBody"),
  selectedSlotMeta: document.getElementById("selectedSlotMeta"),
  selectedSlotWorkers: document.getElementById("selectedSlotWorkers"),
  selectedSlotPrimaryManager: document.getElementById("selectedSlotPrimaryManager"),
  selectedSlotPrimaryStart: document.getElementById("selectedSlotPrimaryStart"),
  selectedSlotPrimaryEnd: document.getElementById("selectedSlotPrimaryEnd"),
  selectedSlotSecondaryManager: document.getElementById("selectedSlotSecondaryManager"),
  selectedSlotSecondaryStart: document.getElementById("selectedSlotSecondaryStart"),
  selectedSlotSecondaryEnd: document.getElementById("selectedSlotSecondaryEnd"),
  selectedSlotNote: document.getElementById("selectedSlotNote"),
  saveSelectedSlotBtn: document.getElementById("saveSelectedSlotBtn"),
  scheduleSummaryBody: document.getElementById("scheduleSummaryBody"),

  addCareerLevelBtn: document.getElementById("addCareerLevelBtn"),
  careerLevelBody: document.getElementById("careerLevelBody"),
  careerMonthSelect: document.getElementById("careerMonthSelect"),
  careerManagerBody: document.getElementById("careerManagerBody"),
};

init();

function init() {
  initializeAdvancedDefaults();
  hydrateSettingInputs();
  bindEvents();
  syncEmployeeSettings();
  syncCareerRatesToEmployeeSettings();
  renderHolidayList();
  renderMonthOptions();
  renderEmployeeSelectors();
  renderEmployeeSettingsTable();
  renderDetectedMissingTable();
  renderAppliedMissingTable();
  renderPreview();
  renderStatement();
  renderSummary();
  renderMonthlyArchive();
  renderScheduleAi();
  renderCareer();
}

function bindEvents() {
  dom.tabs.forEach((tab) => {
    tab.addEventListener("click", () => {
      switchTab(tab.dataset.tab || "setup");
    });
  });

  dom.fileInput.addEventListener("change", handleFileUpload);
  dom.monthSelect.addEventListener("change", () => {
    state.selectedMonth = dom.monthSelect.value;
    ensureSelectedEmployeeForMonth();
    syncMonthSelectors();
    persistConfig();
    renderEmployeeSelectors();
    renderDetectedMissingTable();
    renderAppliedMissingTable();
    renderStatement();
    renderSummary();
    renderMonthlyArchive();
  });
  dom.summaryMonthSelect.addEventListener("change", () => {
    state.selectedMonth = dom.summaryMonthSelect.value;
    ensureSelectedEmployeeForMonth();
    syncMonthSelectors();
    persistConfig();
    renderEmployeeSelectors();
    renderDetectedMissingTable();
    renderAppliedMissingTable();
    renderStatement();
    renderSummary();
    renderMonthlyArchive();
  });
  dom.recalculateBtn.addEventListener("click", () => {
    if (!state.entries.length) {
      setStatus("재계산할 데이터가 없습니다.");
      return;
    }
    rebuildEntriesWithCurrentDefaults();
    setStatus("현재 설정 기준으로 누락 데이터/기본값을 다시 계산했습니다.");
    renderEverything();
  });
  dom.saveAllBtn.addEventListener("click", exportAllDataFile);
  dom.loadAllBtn.addEventListener("click", () => dom.loadStateFileInput.click());
  dom.loadStateFileInput.addEventListener("change", handleLoadStateFile);
  dom.downloadSummaryExcelBtn.addEventListener("click", downloadSummaryExcel);

  bindSettingListeners();

  dom.addHolidayBtn.addEventListener("click", addHoliday);
  dom.holidayList.addEventListener("click", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLButtonElement)) return;
    if (!target.dataset.date) return;
    removeHoliday(target.dataset.date);
  });

  dom.employeeSettingsBody.addEventListener("input", handleEmployeeSettingChange);
  dom.employeeSettingsBody.addEventListener("click", handleEmployeeSettingAction);
  dom.addMissingShiftBtn.addEventListener("click", addMissingShift);
  dom.missingDetectedBody.addEventListener("input", handleDetectedMissingInput);
  dom.missingDetectedBody.addEventListener("click", handleDetectedMissingAction);
  dom.missingAppliedBody.addEventListener("input", handleAppliedMissingInput);
  dom.missingAppliedBody.addEventListener("click", handleAppliedMissingAction);
  dom.monthlyArchiveBody.addEventListener("click", handleMonthlyArchiveAction);

  dom.scheduleMonthInput.addEventListener("change", () => {
    state.scheduleMonth = cleanText(dom.scheduleMonthInput.value || "");
    if (!state.careerMonth) {
      state.careerMonth = state.scheduleMonth;
    }
    persistConfig();
    renderScheduleAi();
    renderCareer();
  });
  dom.generateScheduleBtn.addEventListener("click", generateAutoSchedule);
  dom.printScheduleBtn.addEventListener("click", printScheduleCalendar);
  dom.shiftTemplateBody.addEventListener("input", handleShiftTemplateInput);
  dom.shiftTemplateBody.addEventListener("change", handleShiftTemplateInput);
  dom.addFixedRuleBtn.addEventListener("click", addFixedRule);
  dom.fixedRuleMode.addEventListener("change", syncFixedRuleModeUi);
  dom.fixedRuleBody.addEventListener("click", handleFixedRuleAction);
  dom.addManagerRowBtn.addEventListener("click", addManagerRow);
  dom.allAvailabilityOnBtn.addEventListener("click", () => setAllAvailability(true));
  dom.allAvailabilityOffBtn.addEventListener("click", () => setAllAvailability(false));
  dom.shift3AvailabilityOnBtn.addEventListener("click", () => setShiftAvailabilityForAll("3T", true));
  dom.shift3AvailabilityOffBtn.addEventListener("click", () => setShiftAvailabilityForAll("3T", false));
  dom.scheduleManagerBody.addEventListener("input", handleScheduleManagerInput);
  dom.scheduleManagerBody.addEventListener("change", handleScheduleManagerInput);
  dom.scheduleManagerBody.addEventListener("click", handleScheduleManagerAction);
  dom.scheduleCalendarBody.addEventListener("click", handleScheduleCalendarClick);
  dom.scheduleHeadcountCalendarBody.addEventListener("click", handleScheduleCalendarClick);
  dom.scheduleHeadcountCalendarBody.addEventListener("change", handleScheduleHeadcountCalendarChange);
  dom.selectedSlotWorkers.addEventListener("change", () => {
    const workers = Math.min(2, Math.max(1, Math.round(toNumber(dom.selectedSlotWorkers.value, 1))));
    toggleSecondaryEditor(workers > 1);
  });
  dom.saveSelectedSlotBtn.addEventListener("click", saveSelectedScheduleSlot);

  dom.addCareerLevelBtn.addEventListener("click", addCareerLevel);
  dom.careerLevelBody.addEventListener("input", handleCareerLevelInput);
  dom.careerLevelBody.addEventListener("change", handleCareerLevelInput);
  dom.careerLevelBody.addEventListener("click", handleCareerLevelAction);
  dom.careerMonthSelect.addEventListener("change", () => {
    state.careerMonth = cleanText(dom.careerMonthSelect.value || "");
    persistConfig();
    renderCareer();
  });
  dom.careerManagerBody.addEventListener("input", handleCareerManagerInput);
  dom.careerManagerBody.addEventListener("change", handleCareerManagerInput);

  dom.statementEmployeeSelect.addEventListener("change", () => {
    state.selectedEmployee = dom.statementEmployeeSelect.value;
    persistConfig();
    renderStatement();
  });

  dom.printStatementBtn.addEventListener("click", () => window.print());

  dom.statementBody.addEventListener("change", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) return;
    if (target.classList.contains("hour-editor")) {
      const input = target;
      const entryId = input.dataset.id;
      if (!entryId) return;
      const value = parseDurationInput(input.value);
      if (!Number.isFinite(value) || value < 0) return;
      const entry = state.entries.find((item) => item.id === entryId);
      if (!entry) return;
      entry.adjustedHours = round2(value);
      entry.reportedHours = round2(value);
      entry.isAdjustedManually = true;
      persistConfig();
      renderStatement();
      renderSummary();
    }
  });

  dom.statementBody.addEventListener("click", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLButtonElement)) return;
    if (!target.classList.contains("delete-entry-btn")) return;
    const entryId = target.dataset.id;
    if (!entryId) return;
    const next = state.entries.filter((entry) => entry.id !== entryId);
    if (next.length === state.entries.length) return;
    state.entries = next;
    renderEverything();
    setStatus("선택한 근무기록을 삭제했습니다.");
  });
}

function bindSettingListeners() {
  const checkboxBindings = [["autoFillMissingInput", "autoFillMissing"]];
  checkboxBindings.forEach(([inputId, settingKey]) => {
    dom[inputId].addEventListener("change", () => {
      state.settings[settingKey] = dom[inputId].checked;
      persistConfig();
      renderStatement();
      renderSummary();
    });
  });

  const numberBindings = [
    ["taxRateInput", "taxRate"],
    ["defaultHourlyWageInput", "defaultHourlyWage"],
    ["defaultMissingHoursInput", "defaultMissingHours"],
    ["weeklyThresholdInput", "weeklyThresholdHours"],
  ];
  numberBindings.forEach(([inputId, settingKey]) => {
    dom[inputId].addEventListener("input", () => {
      const value = toNumber(dom[inputId].value, DEFAULT_SETTINGS[settingKey]);
      state.settings[settingKey] = value;
      persistConfig();
      renderStatement();
      renderSummary();
    });
  });

  const timeBindings = [
    ["defaultStartTimeInput", "defaultStartTime"],
    ["defaultNightStartInput", "defaultNightStart"],
    ["nightEndTimeInput", "nightEndTime"],
  ];
  timeBindings.forEach(([inputId, settingKey]) => {
    dom[inputId].addEventListener("change", () => {
      state.settings[settingKey] = dom[inputId].value || DEFAULT_SETTINGS[settingKey];
      persistConfig();
      renderStatement();
      renderSummary();
    });
  });
}

function hydrateSettingInputs() {
  dom.taxRateInput.value = String(state.settings.taxRate);
  dom.defaultHourlyWageInput.value = String(state.settings.defaultHourlyWage);
  dom.autoFillMissingInput.checked = Boolean(state.settings.autoFillMissing);
  dom.defaultMissingHoursInput.value = String(state.settings.defaultMissingHours);
  dom.defaultStartTimeInput.value = state.settings.defaultStartTime;
  dom.defaultNightStartInput.value = state.settings.defaultNightStart;
  dom.nightEndTimeInput.value = state.settings.nightEndTime;
  dom.weeklyThresholdInput.value = String(state.settings.weeklyThresholdHours);
}

function switchTab(tabName) {
  dom.tabs.forEach((tab) => {
    tab.classList.toggle("active", tab.dataset.tab === tabName);
  });
  dom.tabPanels.forEach((panel) => {
    panel.classList.toggle("active", panel.id === `tab-${tabName}`);
  });
}

async function handleFileUpload(event) {
  const input = event.target;
  if (!(input instanceof HTMLInputElement) || !input.files?.length) return;
  const file = input.files[0];
  setStatus(`"${file.name}" 분석 중입니다...`);
  try {
    const parsedEntries = await parseWorkEntries(file);
    const uploadedMonths = Array.from(
      new Set(parsedEntries.map((entry) => entry.workDate.slice(0, 7)))
    ).sort();
    const keepEntries = state.entries.filter((entry) => {
      const month = entry.workDate.slice(0, 7);
      if (!uploadedMonths.includes(month)) return true;
      return entry.source === "manual";
    });
    state.entries = keepEntries.concat(parsedEntries).sort(compareEntries);
    state.selectedMonth = uploadedMonths.slice(-1)[0] || getAvailableMonths().slice(-1)[0] || "";
    syncEmployeeSettings();
    state.selectedEmployee = getEmployeesInMonth(state.selectedMonth)[0] || "";
    renderEverything();
    setStatus(
      `${parsedEntries.length}건을 불러왔고, ${uploadedMonths.join(
        ", "
      )} 월 데이터는 갱신했습니다. 다른 월 데이터는 유지됩니다.`
    );
  } catch (error) {
    const message = error instanceof Error ? error.message : "파일을 읽는 중 오류가 발생했습니다.";
    setStatus(message, true);
  }
}

async function parseWorkEntries(file) {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
  const firstSheetName = workbook.SheetNames[0];
  if (!firstSheetName) {
    throw new Error("시트가 비어 있습니다.");
  }
  const worksheet = workbook.Sheets[firstSheetName];
  const rows = XLSX.utils.sheet_to_json(worksheet, { defval: "", raw: true });
  if (!rows.length) {
    throw new Error("엑셀에 데이터가 없습니다.");
  }

  const headerMap = buildHeaderMap(rows[0]);
  if (!headerMap.name || !headerMap.date) {
    throw new Error("헤더를 찾지 못했습니다. '영업일자', '이름' 컬럼이 필요합니다.");
  }

  const parsed = rows
    .map((row, index) => parseRowToEntry(row, headerMap, index))
    .filter(Boolean)
    .sort(compareEntries);

  return parsed;
}

function buildHeaderMap(firstRow) {
  const keys = Object.keys(firstRow || {});
  const map = {};
  Object.entries(HEADER_CANDIDATES).forEach(([field, candidates]) => {
    map[field] = findHeaderKey(keys, candidates);
  });
  return map;
}

function findHeaderKey(keys, candidates) {
  const normalizedKeys = keys.map((key) => ({ key, normalized: normalizeHeaderKey(key) }));
  for (const candidate of candidates) {
    const candidateNorm = normalizeHeaderKey(candidate);
    const match = normalizedKeys.find(
      (item) => item.normalized === candidateNorm || item.normalized.includes(candidateNorm)
    );
    if (match) return match.key;
  }
  return "";
}

function normalizeHeaderKey(value) {
  return String(value)
    .trim()
    .toLowerCase()
    .replace(/\s+/g, "")
    .replace(/[()\-_/]/g, "");
}

function parseRowToEntry(row, headerMap, index) {
  const name = cleanText(row[headerMap.name]);
  if (!name) return null;

  const workDateObj = parseExcelDate(row[headerMap.date]);
  if (!workDateObj) return null;
  const workDate = formatDateKey(workDateObj);

  const startRaw = headerMap.start ? row[headerMap.start] : "";
  const endRaw = headerMap.end ? row[headerMap.end] : "";
  const reportedRaw = headerMap.stayHours ? row[headerMap.stayHours] : "";

  const start = parseExcelDateTime(startRaw, workDateObj);
  const end = parseExcelDateTime(endRaw, workDateObj);
  const reportedHours = parseHoursValue(reportedRaw);

  return normalizeShift({
    id: `excel-${Date.now()}-${index}-${Math.random().toString(36).slice(2, 7)}`,
    source: "excel",
    name,
    category: cleanText(row[headerMap.category]),
    workDate,
    start,
    end,
    reportedHours,
    note: cleanText(row[headerMap.note]),
    isAdjustedManually: false,
  });
}

function normalizeShift(raw) {
  const defaultHours = Math.max(0, toNumber(state.settings.defaultMissingHours, 8));
  const autoFill = Boolean(state.settings.autoFillMissing);
  const fallbackDate = parseDateOnly(raw.workDate);

  let originalStart = raw.start ? new Date(raw.start) : null;
  let originalEnd = raw.end ? new Date(raw.end) : null;
  if (originalStart && originalEnd && originalEnd <= originalStart) {
    originalEnd = addDays(originalEnd, 1);
  }

  let start = originalStart ? new Date(originalStart) : null;
  let end = originalEnd ? new Date(originalEnd) : null;
  const missingStart = !start;
  const missingEnd = !end;

  if (autoFill && fallbackDate) {
    if (!start && !end) {
      start = combineDateAndTime(fallbackDate, state.settings.defaultStartTime);
      end = addHours(start, defaultHours);
    } else if (!start && end) {
      start = addHours(end, -defaultHours);
    } else if (start && !end) {
      end = addHours(start, defaultHours);
    }
  }

  if (start && end && end <= start) {
    end = addDays(end, 1);
  }

  const computedHours = start && end ? round2(diffHours(start, end)) : null;
  const reportedHours =
    Number.isFinite(raw.reportedHours) && raw.reportedHours >= 0 ? round2(raw.reportedHours) : null;

  let adjustedHours = reportedHours ?? computedHours;
  if (!Number.isFinite(adjustedHours)) adjustedHours = defaultHours;
  adjustedHours = Math.max(0, round2(adjustedHours));

  return {
    id: raw.id || `manual-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
    source: raw.source || "manual",
    name: cleanText(raw.name),
    category: cleanText(raw.category),
    workDate: raw.workDate,
    note: cleanText(raw.note),
    reportedHours,
    adjustedHours,
    isAdjustedManually: Boolean(raw.isAdjustedManually),
    missingStart,
    missingEnd,
    missingOriginDetected:
      typeof raw.missingOriginDetected === "boolean"
        ? raw.missingOriginDetected
        : raw.source === "excel" && (missingStart || missingEnd),
    missingFixApplied: Boolean(raw.missingFixApplied),
    originalStartIso: originalStart ? formatDateTimeKey(originalStart) : "",
    originalEndIso: originalEnd ? formatDateTimeKey(originalEnd) : "",
    startIso: start ? formatDateTimeKey(start) : "",
    endIso: end ? formatDateTimeKey(end) : "",
  };
}

function rebuildEntriesWithCurrentDefaults() {
  state.entries = state.entries
    .map((entry) => {
      const normalized = normalizeShift({
        id: entry.id,
        source: entry.source,
        name: entry.name,
        category: entry.category,
        workDate: entry.workDate,
        start: parseDateTimeKey(entry.originalStartIso),
        end: parseDateTimeKey(entry.originalEndIso),
        reportedHours: entry.reportedHours,
        note: entry.note,
        isAdjustedManually: entry.isAdjustedManually,
        missingOriginDetected: entry.missingOriginDetected,
        missingFixApplied: entry.missingFixApplied,
      });
      if (entry.isAdjustedManually) {
        normalized.adjustedHours = entry.adjustedHours;
        normalized.isAdjustedManually = true;
      }
      return normalized;
    })
    .sort(compareEntries);
}

function addHoliday() {
  const date = dom.holidayDateInput.value;
  if (!date) return;
  if (!state.holidays.includes(date)) {
    state.holidays.push(date);
    state.holidays = sanitizeHolidayList(state.holidays);
    persistConfig();
    renderHolidayList();
    renderStatement();
    renderSummary();
  }
  dom.holidayDateInput.value = "";
}

function removeHoliday(date) {
  state.holidays = state.holidays.filter((item) => item !== date);
  persistConfig();
  renderHolidayList();
  renderStatement();
  renderSummary();
}

function renderHolidayList() {
  if (!state.holidays.length) {
    dom.holidayList.innerHTML = '<span class="hint">등록된 공휴일이 없습니다.</span>';
    return;
  }
  dom.holidayList.innerHTML = state.holidays
    .map((date) => {
      const sunday = parseDateOnly(date)?.getDay() === 0;
      return `<span class="chip">${formatDateForDisplay(date)}${sunday ? " (일요일)" : ""}
        <button type="button" data-date="${date}">삭제</button>
      </span>`;
    })
    .join("");
}

function addMissingShift() {
  const selectedExisting = dom.missingEmployeeSelect.value;
  const customName = cleanText(dom.missingEmployeeName.value);
  const isNewEmployeeSelection = selectedExisting === "__new__";
  const name = isNewEmployeeSelection ? customName : customName || selectedExisting;
  const workDate = dom.missingDate.value;
  if (!name || !workDate) {
    setStatus("누락 근무 추가 시 근무자 이름과 근무일은 필수입니다.", true);
    return;
  }
  if (isNewEmployeeSelection && !customName) {
    setStatus("신규 근무자 추가를 선택한 경우 이름을 입력하세요.", true);
    return;
  }

  const dateObj = parseDateOnly(workDate);
  if (!dateObj) {
    setStatus("근무일 형식이 올바르지 않습니다.", true);
    return;
  }

  const start = dom.missingStart.value ? combineDateAndTime(dateObj, dom.missingStart.value) : null;
  const end = dom.missingEnd.value ? combineDateAndTime(dateObj, dom.missingEnd.value) : null;
  const reportedHoursRaw = dom.missingHours.value.trim() ? toNumber(dom.missingHours.value, NaN) : NaN;
  const reportedHours = Number.isFinite(reportedHoursRaw) ? reportedHoursRaw : null;

  const entry = normalizeShift({
    id: `manual-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
    source: "manual",
    name,
    category: cleanText(dom.missingCategory.value),
    workDate,
    start,
    end,
    reportedHours,
    note: cleanText(dom.missingNote.value) || "수동 누락 추가",
    isAdjustedManually: Number.isFinite(reportedHours),
  });

  if (Number.isFinite(reportedHours)) {
    entry.adjustedHours = round2(reportedHours);
  }

  state.entries.push(entry);
  state.entries.sort(compareEntries);
  syncEmployeeSettings();

  if (!state.selectedMonth) {
    state.selectedMonth = workDate.slice(0, 7);
  }
  if (!state.selectedEmployee) {
    state.selectedEmployee = name;
  }

  clearMissingShiftForm();
  renderEverything();
  setStatus("누락 근무를 추가했습니다.");
}

function clearMissingShiftForm() {
  dom.missingEmployeeName.value = "";
  dom.missingDate.value = "";
  dom.missingCategory.value = "";
  dom.missingStart.value = "";
  dom.missingEnd.value = "";
  dom.missingHours.value = "";
  dom.missingNote.value = "";
}

function handleEmployeeSettingChange(event) {
  const target = event.target;
  if (!(target instanceof HTMLInputElement)) return;
  const name = target.dataset.name;
  if (!name) return;
  const config = ensureEmployeeSetting(name);
  if (target.classList.contains("rate-input")) {
    config.hourlyRate = Math.max(0, toNumber(target.value, state.settings.defaultHourlyWage));
  }
  if (target.classList.contains("night-start-input")) {
    config.nightStart = target.value || state.settings.defaultNightStart;
  }
  if (target.classList.contains("employee-allowance-toggle")) {
    const key = target.dataset.key;
    if (key) {
      config[key] = target.checked;
    }
  }
  persistConfig();
  renderStatement();
  renderSummary();
}

function handleEmployeeSettingAction(event) {
  const target = event.target;
  if (!(target instanceof HTMLButtonElement)) return;
  if (!target.classList.contains("delete-employee-btn")) return;
  const name = target.dataset.name;
  if (!name) return;

  const beforeCount = state.entries.length;
  state.entries = state.entries.filter((entry) => entry.name !== name);
  delete state.employeeSettings[name];
  delete state.managerProfiles[name];
  state.scheduleExcludedManagers = state.scheduleExcludedManagers.filter((item) => item !== name);
  state.scheduleFixedRules = state.scheduleFixedRules.filter((rule) => rule.manager !== name);
  Object.keys(state.scheduleAssignments).forEach((month) => {
    state.scheduleAssignments[month] = state.scheduleAssignments[month].map((row) => {
      const next = { ...row };
      if (next.primaryManager === name) {
        next.primaryManager = "";
        next.primaryStart = "";
        next.primaryEnd = "";
      }
      if (next.secondaryManager === name) {
        next.secondaryManager = "";
        next.secondaryStart = "";
        next.secondaryEnd = "";
      }
      return next;
    });
  });

  if (beforeCount === state.entries.length) {
    setStatus("삭제할 근무자 데이터가 없습니다.", true);
    renderEverything();
    return;
  }

  setStatus(`${name} 근무자 데이터를 삭제했습니다.`);
  renderEverything();
}

function renderEverything() {
  const months = getAvailableMonths();
  if (!months.includes(state.selectedMonth)) {
    state.selectedMonth = months.slice(-1)[0] || "";
  }

  const employeesInMonth = getEmployeesInMonth(state.selectedMonth);
  if (!employeesInMonth.includes(state.selectedEmployee)) {
    state.selectedEmployee = employeesInMonth[0] || "";
  }

  syncEmployeeSettings();
  syncCareerRatesToEmployeeSettings();
  renderMonthOptions();
  renderEmployeeSelectors();
  renderEmployeeSettingsTable();
  renderDetectedMissingTable();
  renderAppliedMissingTable();
  renderPreview();
  renderStatement();
  renderSummary();
  renderMonthlyArchive();
  renderScheduleAi();
  renderCareer();
  persistConfig();
}

function renderMonthOptions() {
  const months = getAvailableMonths();
  if (!months.length) {
    state.selectedMonth = "";
    dom.monthSelect.innerHTML = '<option value="">월 선택</option>';
    dom.summaryMonthSelect.innerHTML = '<option value="">월 선택</option>';
    dom.monthSelect.value = "";
    dom.summaryMonthSelect.value = "";
    return;
  }
  const optionHtml = months
    .map((month) => `<option value="${month}">${formatMonthLabel(month)}</option>`)
    .join("");
  dom.monthSelect.innerHTML = optionHtml;
  dom.summaryMonthSelect.innerHTML = optionHtml;

  const fallbackMonth = months[months.length - 1];
  state.selectedMonth = months.includes(state.selectedMonth) ? state.selectedMonth : fallbackMonth;
  syncMonthSelectors();
}

function syncMonthSelectors() {
  if (dom.monthSelect.value !== state.selectedMonth) {
    dom.monthSelect.value = state.selectedMonth;
  }
  if (dom.summaryMonthSelect.value !== state.selectedMonth) {
    dom.summaryMonthSelect.value = state.selectedMonth;
  }
}

function ensureSelectedEmployeeForMonth() {
  const employees = getEmployeesInMonth(state.selectedMonth);
  if (!employees.includes(state.selectedEmployee)) {
    state.selectedEmployee = employees[0] || "";
  }
}

function renderEmployeeSelectors() {
  const employees = getEmployeesInMonth(state.selectedMonth);
  const placeholder = '<option value="">근무자 선택</option>';
  const newEmployeeOption = '<option value="__new__">+ 신규 근무자 추가</option>';
  const optionHtml = employees
    .map((name) => `<option value="${escapeHtml(name)}">${escapeHtml(name)}</option>`)
    .join("");

  dom.statementEmployeeSelect.innerHTML = placeholder + optionHtml;
  dom.missingEmployeeSelect.innerHTML = placeholder + newEmployeeOption + optionHtml;

  if (employees.includes(state.selectedEmployee)) {
    dom.statementEmployeeSelect.value = state.selectedEmployee;
  } else {
    state.selectedEmployee = employees[0] || "";
    dom.statementEmployeeSelect.value = state.selectedEmployee;
  }
}

function renderEmployeeSettingsTable() {
  const employees = getEmployees();
  if (!employees.length) {
    dom.employeeSettingsBody.innerHTML =
      '<tr><td colspan="9" class="empty">업로드 후 자동 생성됩니다.</td></tr>';
    return;
  }

  syncEmployeeSettings();
  dom.employeeSettingsBody.innerHTML = employees
    .map((name) => {
      const config = ensureEmployeeSetting(name);
      const category = getMostUsedCategory(name);
      return `<tr>
        <td>${escapeHtml(name)}</td>
        <td>${escapeHtml(category || "-")}</td>
        <td>
          <input class="rate-input" data-name="${escapeHtml(name)}" type="number" min="0" step="10" value="${Math.round(
            config.hourlyRate
          )}" />
        </td>
        <td>
          <input class="night-start-input" data-name="${escapeHtml(name)}" type="time" value="${config.nightStart}" />
        </td>
        <td>
          <label class="switch compact-switch">
            <input class="employee-allowance-toggle" data-key="useWeeklyHoliday" data-name="${escapeHtml(
              name
            )}" type="checkbox" ${config.useWeeklyHoliday ? "checked" : ""} />
          </label>
        </td>
        <td>
          <label class="switch compact-switch">
            <input class="employee-allowance-toggle" data-key="useHolidayPremium" data-name="${escapeHtml(
              name
            )}" type="checkbox" ${config.useHolidayPremium ? "checked" : ""} />
          </label>
        </td>
        <td>
          <label class="switch compact-switch">
            <input class="employee-allowance-toggle" data-key="useOvertimePremium" data-name="${escapeHtml(
              name
            )}" type="checkbox" ${config.useOvertimePremium ? "checked" : ""} />
          </label>
        </td>
        <td>
          <label class="switch compact-switch">
            <input class="employee-allowance-toggle" data-key="useNightPremium" data-name="${escapeHtml(
              name
            )}" type="checkbox" ${config.useNightPremium ? "checked" : ""} />
          </label>
        </td>
        <td>
          <button class="btn secondary small-btn delete-employee-btn" data-name="${escapeHtml(
            name
          )}" type="button">삭제</button>
        </td>
      </tr>`;
    })
    .join("");
}

function renderDetectedMissingTable() {
  const targets = getDetectedMissingEntries();
  if (!targets.length) {
    dom.missingDetectedBody.innerHTML =
      '<tr><td colspan="7" class="empty">누락 감지 항목이 없습니다.</td></tr>';
    return;
  }

  dom.missingDetectedBody.innerHTML = targets
    .map((entry) => {
      const statusParts = [];
      if (entry.missingStart) statusParts.push("출근 누락");
      if (entry.missingEnd) statusParts.push("퇴근 누락");
      if (!statusParts.length) statusParts.push("시간 보정");
      return `<tr>
        <td>${escapeHtml(entry.name)}</td>
        <td>${formatDateWithDay(entry.workDate)}</td>
        <td>${statusParts.join(" / ")}</td>
        <td>
          <input class="missing-fix-start" data-id="${entry.id}" type="time" value="${
            entry.originalStartIso ? formatTimeFromIso(entry.originalStartIso) : ""
          }" />
        </td>
        <td>
          <input class="missing-fix-end" data-id="${entry.id}" type="time" value="${
            entry.originalEndIso ? formatTimeFromIso(entry.originalEndIso) : ""
          }" />
        </td>
        <td>
          <input class="missing-fix-hours" data-id="${entry.id}" type="text" placeholder="예: 8시간 30분" value="${formatDurationText(
            entry.adjustedHours
          )}" />
        </td>
        <td>
          <button class="btn secondary small-btn apply-missing-fix-btn" data-id="${entry.id}" type="button">적용</button>
        </td>
      </tr>`;
    })
    .join("");
}

function getDetectedMissingEntries() {
  return state.entries
    .filter((entry) => entry.source === "excel" && (entry.missingStart || entry.missingEnd))
    .filter((entry) => !entry.missingFixApplied)
    .filter((entry) => !state.selectedMonth || entry.workDate.startsWith(state.selectedMonth))
    .sort(compareEntries);
}

function handleDetectedMissingInput(event) {
  const target = event.target;
  if (!(target instanceof HTMLInputElement)) return;
  if (!target.classList.contains("missing-fix-start") && !target.classList.contains("missing-fix-end")) return;

  const row = target.closest("tr");
  if (!row) return;
  const startInput = row.querySelector(".missing-fix-start");
  const endInput = row.querySelector(".missing-fix-end");
  const hoursInput = row.querySelector(".missing-fix-hours");
  if (!(startInput instanceof HTMLInputElement)) return;
  if (!(endInput instanceof HTMLInputElement)) return;
  if (!(hoursInput instanceof HTMLInputElement)) return;

  const entryId = target.dataset.id;
  if (!entryId) return;
  const entry = state.entries.find((item) => item.id === entryId);
  if (!entry) return;

  const autoHours = calculateDurationFromTimeInputs(entry.workDate, startInput.value, endInput.value);
  if (Number.isFinite(autoHours)) {
    hoursInput.value = formatDurationText(autoHours);
  }
}

function handleDetectedMissingAction(event) {
  const target = event.target;
  if (!(target instanceof HTMLButtonElement)) return;
  if (!target.classList.contains("apply-missing-fix-btn")) return;
  const entryId = target.dataset.id;
  if (!entryId) return;

  const entry = state.entries.find((item) => item.id === entryId);
  if (!entry) return;

  const startInput = dom.missingDetectedBody.querySelector(`.missing-fix-start[data-id="${entryId}"]`);
  const endInput = dom.missingDetectedBody.querySelector(`.missing-fix-end[data-id="${entryId}"]`);
  const hoursInput = dom.missingDetectedBody.querySelector(`.missing-fix-hours[data-id="${entryId}"]`);
  if (!(startInput instanceof HTMLInputElement)) return;
  if (!(endInput instanceof HTMLInputElement)) return;
  if (!(hoursInput instanceof HTMLInputElement)) return;

  const day = parseDateOnly(entry.workDate);
  if (!day) return;

  const startRaw = startInput.value.trim();
  const endRaw = endInput.value.trim();
  const hoursRaw = hoursInput.value.trim();

  let nextStart = startRaw ? combineDateAndTime(day, startRaw) : parseDateTimeKey(entry.originalStartIso);
  let nextEnd = endRaw ? combineDateAndTime(day, endRaw) : parseDateTimeKey(entry.originalEndIso);
  let durationHours = parseDurationInput(hoursRaw);

  if (!Number.isFinite(durationHours)) {
    const autoHours = calculateDurationFromTimeInputs(entry.workDate, startRaw, endRaw);
    if (Number.isFinite(autoHours)) {
      durationHours = autoHours;
    }
  }

  if (!nextStart && !nextEnd && !Number.isFinite(durationHours)) {
    setStatus("보정할 값을 하나 이상 입력하세요.", true);
    return;
  }

  if (nextStart && nextEnd && nextEnd <= nextStart) {
    nextEnd = addDays(nextEnd, 1);
  }

  const replaced = normalizeShift({
    id: entry.id,
    source: entry.source,
    name: entry.name,
    category: entry.category,
    workDate: entry.workDate,
    start: nextStart,
    end: nextEnd,
    reportedHours: Number.isFinite(durationHours) ? durationHours : entry.reportedHours,
    note: entry.note || "누락 보정",
    isAdjustedManually: Number.isFinite(durationHours),
  });

  if (Number.isFinite(durationHours)) {
    replaced.adjustedHours = round2(durationHours);
    replaced.reportedHours = round2(durationHours);
    replaced.isAdjustedManually = true;
  } else if (entry.isAdjustedManually) {
    replaced.adjustedHours = entry.adjustedHours;
    replaced.isAdjustedManually = true;
  }

  replaced.missingOriginDetected = true;
  replaced.missingFixApplied = true;

  const idx = state.entries.findIndex((item) => item.id === entryId);
  if (idx >= 0) {
    state.entries[idx] = replaced;
    state.entries.sort(compareEntries);
    setStatus(`${entry.name} ${formatDateForDisplay(entry.workDate)} 누락 기록을 보정했습니다.`);
    renderEverything();
  }
}

function getAppliedMissingEntries() {
  return state.entries
    .filter((entry) => entry.source === "excel" && entry.missingOriginDetected && entry.missingFixApplied)
    .filter((entry) => !state.selectedMonth || entry.workDate.startsWith(state.selectedMonth))
    .sort(compareEntries);
}

function renderAppliedMissingTable() {
  const rows = getAppliedMissingEntries();
  if (!rows.length) {
    dom.missingAppliedBody.innerHTML =
      '<tr><td colspan="6" class="empty">아직 적용 완료된 누락 보정 기록이 없습니다.</td></tr>';
    return;
  }

  dom.missingAppliedBody.innerHTML = rows
    .map((entry) => {
      const startValue = entry.startIso ? formatTimeFromIso(entry.startIso) : "";
      const endValue = entry.endIso ? formatTimeFromIso(entry.endIso) : "";
      return `<tr>
        <td>${escapeHtml(entry.name)}</td>
        <td>${formatDateWithDay(entry.workDate)}</td>
        <td><input class="applied-fix-start" data-id="${entry.id}" type="time" value="${startValue}" /></td>
        <td><input class="applied-fix-end" data-id="${entry.id}" type="time" value="${endValue}" /></td>
        <td>
          <input class="applied-fix-hours" data-id="${entry.id}" type="text" value="${formatDurationText(
            entry.adjustedHours
          )}" />
        </td>
        <td>
          <button class="btn secondary small-btn apply-applied-fix-btn" data-id="${entry.id}" type="button">수정 적용</button>
        </td>
      </tr>`;
    })
    .join("");
}

function handleAppliedMissingInput(event) {
  const target = event.target;
  if (!(target instanceof HTMLInputElement)) return;
  if (!target.classList.contains("applied-fix-start") && !target.classList.contains("applied-fix-end")) return;

  const row = target.closest("tr");
  if (!row) return;
  const startInput = row.querySelector(".applied-fix-start");
  const endInput = row.querySelector(".applied-fix-end");
  const hoursInput = row.querySelector(".applied-fix-hours");
  if (!(startInput instanceof HTMLInputElement)) return;
  if (!(endInput instanceof HTMLInputElement)) return;
  if (!(hoursInput instanceof HTMLInputElement)) return;

  const entryId = target.dataset.id;
  if (!entryId) return;
  const entry = state.entries.find((item) => item.id === entryId);
  if (!entry) return;

  const autoHours = calculateDurationFromTimeInputs(entry.workDate, startInput.value, endInput.value);
  if (Number.isFinite(autoHours)) {
    hoursInput.value = formatDurationText(autoHours);
  }
}

function handleAppliedMissingAction(event) {
  const target = event.target;
  if (!(target instanceof HTMLButtonElement)) return;
  if (!target.classList.contains("apply-applied-fix-btn")) return;
  const entryId = target.dataset.id;
  if (!entryId) return;
  const entry = state.entries.find((item) => item.id === entryId);
  if (!entry) return;

  const startInput = dom.missingAppliedBody.querySelector(`.applied-fix-start[data-id="${entryId}"]`);
  const endInput = dom.missingAppliedBody.querySelector(`.applied-fix-end[data-id="${entryId}"]`);
  const hoursInput = dom.missingAppliedBody.querySelector(`.applied-fix-hours[data-id="${entryId}"]`);
  if (!(startInput instanceof HTMLInputElement)) return;
  if (!(endInput instanceof HTMLInputElement)) return;
  if (!(hoursInput instanceof HTMLInputElement)) return;

  const day = parseDateOnly(entry.workDate);
  if (!day) return;

  const startRaw = startInput.value.trim();
  const endRaw = endInput.value.trim();
  const hoursRaw = hoursInput.value.trim();

  let nextStart = startRaw ? combineDateAndTime(day, startRaw) : null;
  let nextEnd = endRaw ? combineDateAndTime(day, endRaw) : null;
  let durationHours = parseDurationInput(hoursRaw);

  if (!Number.isFinite(durationHours)) {
    const autoHours = calculateDurationFromTimeInputs(entry.workDate, startRaw, endRaw);
    if (Number.isFinite(autoHours)) {
      durationHours = autoHours;
    }
  }

  if (!nextStart && !nextEnd && !Number.isFinite(durationHours)) {
    setStatus("수정할 값을 하나 이상 입력하세요.", true);
    return;
  }
  if (nextStart && nextEnd && nextEnd <= nextStart) {
    nextEnd = addDays(nextEnd, 1);
  }

  const replaced = normalizeShift({
    id: entry.id,
    source: entry.source,
    name: entry.name,
    category: entry.category,
    workDate: entry.workDate,
    start: nextStart,
    end: nextEnd,
    reportedHours: Number.isFinite(durationHours) ? durationHours : entry.reportedHours,
    note: entry.note || "누락 보정",
    isAdjustedManually: true,
    missingOriginDetected: true,
    missingFixApplied: true,
  });

  if (Number.isFinite(durationHours)) {
    replaced.adjustedHours = round2(durationHours);
    replaced.reportedHours = round2(durationHours);
  } else {
    replaced.adjustedHours = entry.adjustedHours;
    replaced.reportedHours = entry.reportedHours;
  }

  replaced.missingOriginDetected = true;
  replaced.missingFixApplied = true;

  const idx = state.entries.findIndex((item) => item.id === entryId);
  if (idx >= 0) {
    state.entries[idx] = replaced;
    state.entries.sort(compareEntries);
    setStatus(`${entry.name} ${formatDateForDisplay(entry.workDate)} 보정 기록을 수정했습니다.`);
    renderEverything();
  }
}

function renderPreview() {
  if (!state.entries.length) {
    dom.previewBody.innerHTML =
      '<tr><td colspan="7" class="empty">아직 업로드된 데이터가 없습니다.</td></tr>';
    return;
  }

  const rows = state.entries.slice(0, 20);
  dom.previewBody.innerHTML = rows
    .map((entry) => {
      return `<tr>
        <td>${formatDateForDisplay(entry.workDate)}</td>
        <td>${escapeHtml(entry.name)}</td>
        <td>${escapeHtml(entry.category || "-")}</td>
        <td>${entry.startIso ? formatTimeFromIso(entry.startIso) : "-"}</td>
        <td>${entry.endIso ? formatTimeFromIso(entry.endIso) : "-"}</td>
        <td>${formatDurationText(entry.adjustedHours)}</td>
        <td>${escapeHtml(entry.note || "-")}</td>
      </tr>`;
    })
    .join("");
}

function renderStatement() {
  if (!state.selectedMonth || !state.selectedEmployee) {
    dom.statementMeta.textContent = "근무자와 정산 월을 선택하세요.";
    dom.statementBody.innerHTML =
      '<tr><td colspan="12" class="empty">표시할 근무 내역이 없습니다.</td></tr>';
    dom.weeklySummaryBody.innerHTML = '<tr><td colspan="6" class="empty">데이터가 없습니다.</td></tr>';
    renderTotalsEmpty();
    return;
  }

  const payroll = computeEmployeeMonthPayroll(state.selectedEmployee, state.selectedMonth);
  const employeeConfig = ensureEmployeeSetting(state.selectedEmployee);

  dom.statementMeta.textContent = `${formatMonthLabel(state.selectedMonth)} · ${
    state.selectedEmployee
  } · 시급 ${formatWon(employeeConfig.hourlyRate)} · ${payroll.shifts.length}건`;

  if (!payroll.shifts.length) {
    dom.statementBody.innerHTML =
      '<tr><td colspan="12" class="empty">선택한 월의 근무기록이 없습니다.</td></tr>';
    dom.weeklySummaryBody.innerHTML = '<tr><td colspan="6" class="empty">데이터가 없습니다.</td></tr>';
    renderTotalsEmpty();
    return;
  }

  dom.statementBody.innerHTML = payroll.shifts
    .map((shift) => {
      return `<tr>
        <td>${shift.weekIndex}주차</td>
        <td>${formatDateWithDay(shift.entry.workDate)}</td>
        <td>${shift.entry.startIso ? formatTimeFromIso(shift.entry.startIso) : "-"}</td>
        <td>${shift.entry.endIso ? formatTimeFromIso(shift.entry.endIso) : "-"}</td>
        <td>${formatDurationText(shift.entry.adjustedHours)}</td>
        <td>
          <input class="hour-editor" data-id="${shift.entry.id}" type="text" value="${formatDurationText(
            shift.entry.adjustedHours
          )}" />
        </td>
        <td>${formatWon(shift.basePay)}</td>
        <td>${formatWon(shift.overtimePremium)}</td>
        <td>${formatWon(shift.nightPremium)}</td>
        <td>${formatWon(shift.holidayPremium)}</td>
        <td>${formatWon(shift.shiftPay)}</td>
        <td><button class="btn secondary delete-entry-btn" data-id="${shift.entry.id}" type="button">삭제</button></td>
      </tr>`;
    })
    .join("");

  if (!payroll.weeks.length) {
    dom.weeklySummaryBody.innerHTML = '<tr><td colspan="6" class="empty">데이터가 없습니다.</td></tr>';
  } else {
    dom.weeklySummaryBody.innerHTML = payroll.weeks
      .map((week) => {
        const premiumSum =
          week.overtimePremium + week.nightPremium + week.holidayPremium + week.weeklyHolidayPay;
        return `<tr>
          <td>${week.weekIndex}주차</td>
          <td>${formatDurationText(week.hours)}</td>
          <td>${week.workDays}</td>
          <td>${formatWon(week.basePay)}</td>
          <td>${formatWon(premiumSum)}</td>
          <td>${formatWon(week.totalPay)}</td>
        </tr>`;
      })
      .join("");
  }

  dom.totalHoursCell.textContent = formatDurationText(payroll.totals.hours);
  dom.basePayCell.textContent = formatWon(payroll.totals.basePay);
  dom.allowanceCell.textContent = formatWon(payroll.totals.allowances);
  dom.grossPayCell.textContent = formatWon(payroll.totals.grossPay);
  dom.netPayCell.textContent = formatWon(payroll.totals.netPay);
}

function renderTotalsEmpty() {
  dom.totalHoursCell.textContent = "0시간 0분";
  dom.basePayCell.textContent = "₩0";
  dom.allowanceCell.textContent = "₩0";
  dom.grossPayCell.textContent = "₩0";
  dom.netPayCell.textContent = "₩0";
}

function renderSummary() {
  if (!state.selectedMonth) {
    dom.summaryMeta.textContent = "요약할 월을 선택하세요.";
    dom.summaryBody.innerHTML = '<tr><td colspan="4" class="empty">정산할 데이터가 없습니다.</td></tr>';
    dom.summaryTotalHours.textContent = "0시간 0분";
    dom.summaryTotalGross.textContent = "₩0";
    dom.summaryTotalNet.textContent = "₩0";
    return;
  }

  const employees = getEmployeesInMonth(state.selectedMonth);
  dom.summaryMeta.textContent = `${formatMonthLabel(state.selectedMonth)} 기준 근무자별 급여 요약`;
  if (!employees.length) {
    dom.summaryBody.innerHTML = '<tr><td colspan="4" class="empty">정산할 데이터가 없습니다.</td></tr>';
    dom.summaryTotalHours.textContent = "0시간 0분";
    dom.summaryTotalGross.textContent = "₩0";
    dom.summaryTotalNet.textContent = "₩0";
    return;
  }

  const rows = buildSummaryRowsForMonth(state.selectedMonth);

  dom.summaryBody.innerHTML = rows
    .map(
      ({ name, payroll }) => `<tr>
      <td>${escapeHtml(name)}</td>
      <td>${formatDurationText(payroll.totals.hours)}</td>
      <td>${formatWon(payroll.totals.grossPay)}</td>
      <td>${formatWon(payroll.totals.netPay)}</td>
    </tr>`
    )
    .join("");

  const totals = rows.reduce(
    (acc, item) => {
      acc.hours += item.payroll.totals.hours;
      acc.gross += item.payroll.totals.grossPay;
      acc.net += item.payroll.totals.netPay;
      return acc;
    },
    { hours: 0, gross: 0, net: 0 }
  );

  dom.summaryTotalHours.textContent = formatDurationText(totals.hours);
  dom.summaryTotalGross.textContent = formatWon(totals.gross);
  dom.summaryTotalNet.textContent = formatWon(totals.net);
}

function buildSummaryRowsForMonth(month) {
  return getEmployeesInMonth(month)
    .map((name) => ({
      name,
      payroll: computeEmployeeMonthPayroll(name, month),
    }))
    .sort((a, b) => b.payroll.totals.grossPay - a.payroll.totals.grossPay);
}

function downloadSummaryExcel() {
  if (!state.selectedMonth) {
    setStatus("다운로드할 월을 먼저 선택하세요.", true);
    return;
  }

  const rows = buildSummaryRowsForMonth(state.selectedMonth);
  if (!rows.length) {
    setStatus("선택한 월에 다운로드할 급여 요약 데이터가 없습니다.", true);
    return;
  }

  const dataRows = rows.map((row, index) => ({
    번호: index + 1,
    이름: row.name,
    근무시간: formatDurationText(row.payroll.totals.hours),
    "근무시간(시간)": round2(row.payroll.totals.hours),
    "급여 총계(원)": roundCurrency(row.payroll.totals.grossPay),
    "세후 급여(원)": roundCurrency(row.payroll.totals.netPay),
  }));

  const totals = rows.reduce(
    (acc, row) => {
      acc.hours += row.payroll.totals.hours;
      acc.gross += row.payroll.totals.grossPay;
      acc.net += row.payroll.totals.netPay;
      return acc;
    },
    { hours: 0, gross: 0, net: 0 }
  );
  dataRows.push({
    번호: "",
    이름: "합계",
    근무시간: formatDurationText(totals.hours),
    "근무시간(시간)": round2(totals.hours),
    "급여 총계(원)": roundCurrency(totals.gross),
    "세후 급여(원)": roundCurrency(totals.net),
  });

  const worksheet = XLSX.utils.json_to_sheet(dataRows);
  worksheet["!cols"] = [
    { wch: 8 },
    { wch: 16 },
    { wch: 14 },
    { wch: 14 },
    { wch: 16 },
    { wch: 16 },
  ];

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "전체 급여 요약");
  XLSX.writeFile(workbook, `근무자별_급여요약_${state.selectedMonth}.xlsx`);
  setStatus(`${formatMonthLabel(state.selectedMonth)} 급여 요약 엑셀을 다운로드했습니다.`);
}

function renderMonthlyArchive() {
  const months = getAvailableMonths();
  if (!months.length) {
    dom.monthlyArchiveBody.innerHTML =
      '<tr><td colspan="6" class="empty">저장된 월별 데이터가 없습니다.</td></tr>';
    return;
  }

  const rows = months.map((month) => buildMonthSnapshot(month));
  dom.monthlyArchiveBody.innerHTML = rows
    .map((row) => {
      const selectedBadge = row.month === state.selectedMonth ? ' <span class="badge">현재 선택</span>' : "";
      return `<tr>
        <td>${formatMonthLabel(row.month)}${selectedBadge}</td>
        <td>${row.employeeCount}</td>
        <td>${formatDurationText(row.totalHours)}</td>
        <td>${formatWon(row.totalGross)}</td>
        <td>${formatWon(row.totalNet)}</td>
        <td>
          <button class="btn secondary small-btn view-month-btn" data-month="${row.month}" type="button">이 달 보기</button>
        </td>
      </tr>`;
    })
    .join("");
}

function buildMonthSnapshot(month) {
  const employees = getEmployeesInMonth(month);
  const totals = employees.reduce(
    (acc, name) => {
      const payroll = computeEmployeeMonthPayroll(name, month);
      acc.hours += payroll.totals.hours;
      acc.gross += payroll.totals.grossPay;
      acc.net += payroll.totals.netPay;
      return acc;
    },
    { hours: 0, gross: 0, net: 0 }
  );
  return {
    month,
    employeeCount: employees.length,
    totalHours: round2(totals.hours),
    totalGross: roundCurrency(totals.gross),
    totalNet: roundCurrency(totals.net),
  };
}

function handleMonthlyArchiveAction(event) {
  const target = event.target;
  if (!(target instanceof HTMLButtonElement)) return;
  if (!target.classList.contains("view-month-btn")) return;
  const month = target.dataset.month;
  if (!month) return;
  state.selectedMonth = month;
  ensureSelectedEmployeeForMonth();
  syncMonthSelectors();
  persistConfig();
  renderEmployeeSelectors();
  renderDetectedMissingTable();
  renderAppliedMissingTable();
  renderStatement();
  renderSummary();
  renderMonthlyArchive();
}

function initializeAdvancedDefaults() {
  if (!state.scheduleTemplates.length) {
    state.scheduleTemplates = cloneDefaultShiftTemplates();
  }
  if (!state.careerLevels.length) {
    state.careerLevels = cloneDefaultCareerLevels();
  }
  if (!state.scheduleMonth) {
    state.scheduleMonth = getDefaultScheduleMonth();
  }
  if (!state.careerMonth) {
    state.careerMonth = state.selectedMonth || state.scheduleMonth;
  }
  syncManagerProfiles();
}

function getDefaultScheduleMonth() {
  const base =
    state.scheduleMonth || state.selectedMonth || formatDateKey(new Date()).slice(0, 7);
  return base < MIN_AUTO_SCHEDULE_MONTH ? MIN_AUTO_SCHEDULE_MONTH : base;
}

function renderScheduleAi() {
  syncManagerProfiles();
  renderScheduleMonthInput();
  ensureScheduleRowsForMonth(state.scheduleMonth);
  ensureSelectedScheduleSlot(state.scheduleMonth);
  renderShiftTemplateTable();
  renderFixedRuleSelectors();
  renderFixedRuleTable();
  renderScheduleManagerTable();
  renderScheduleResultTable();
  renderScheduleHeadcountCalendarTable();
  renderSelectedSlotEditor();
  renderScheduleSummaryTable();
}

function printScheduleCalendar() {
  const activeTab = dom.tabs.find((tab) => tab.classList.contains("active"))?.dataset.tab || "setup";
  switchTab("schedule-ai");
  document.body.classList.add("print-schedule-mode");

  const cleanup = () => {
    document.body.classList.remove("print-schedule-mode");
    switchTab(activeTab);
    window.removeEventListener("afterprint", cleanup);
  };
  window.addEventListener("afterprint", cleanup);
  window.print();
  setTimeout(() => {
    if (document.body.classList.contains("print-schedule-mode")) {
      cleanup();
    }
  }, 1200);
}

function renderScheduleMonthInput() {
  if (!state.scheduleMonth) {
    state.scheduleMonth = getDefaultScheduleMonth();
  }
  if (dom.scheduleMonthInput.value !== state.scheduleMonth) {
    dom.scheduleMonthInput.value = state.scheduleMonth;
  }
}

function renderShiftTemplateTable() {
  if (!state.scheduleTemplates.length) {
    dom.shiftTemplateBody.innerHTML = '<tr><td colspan="3" class="empty">타임 설정이 없습니다.</td></tr>';
    return;
  }
  dom.shiftTemplateBody.innerHTML = state.scheduleTemplates
    .map(
      (template) => `<tr>
        <td>${escapeHtml(template.label || template.id)}</td>
        <td><input class="shift-template-start" data-shift-id="${template.id}" type="time" value="${template.start}" /></td>
        <td><input class="shift-template-end" data-shift-id="${template.id}" type="time" value="${template.end}" /></td>
      </tr>`
    )
    .join("");
}

function handleShiftTemplateInput(event) {
  const target = event.target;
  if (!(target instanceof HTMLInputElement)) return;
  if (
    !target.classList.contains("shift-template-start") &&
    !target.classList.contains("shift-template-end")
  ) {
    return;
  }
  const shiftId = target.dataset.shiftId;
  if (!shiftId) return;
  const template = state.scheduleTemplates.find((item) => item.id === shiftId);
  if (!template) return;

  const prevStart = template.start;
  const prevEnd = template.end;

  if (target.classList.contains("shift-template-start")) {
    template.start = normalizeTimeText(target.value, template.start);
  }
  if (target.classList.contains("shift-template-end")) {
    template.end = normalizeTimeText(target.value, template.end);
  }

  Object.values(state.scheduleAssignments).forEach((rows) => {
    rows.forEach((row) => {
      if (row.shiftId !== shiftId) return;
      if (row.shiftStart === prevStart) row.shiftStart = template.start;
      if (row.shiftEnd === prevEnd) row.shiftEnd = template.end;
      if (row.primaryStart === prevStart) row.primaryStart = template.start;
      if (row.primaryEnd === prevEnd) row.primaryEnd = template.end;
      if (row.secondaryStart === prevStart) row.secondaryStart = template.start;
      if (row.secondaryEnd === prevEnd) row.secondaryEnd = template.end;
    });
  });

  persistConfig();
  renderScheduleAi();
  renderCareer();
}

function renderFixedRuleSelectors() {
  dom.fixedRuleWeekday.innerHTML = WEEKDAY_ORDER.map(
    (day) => `<option value="${day}">${WEEKDAYS[day]}</option>`
  ).join("");

  dom.fixedRuleShift.innerHTML = state.scheduleTemplates
    .map((template) => `<option value="${template.id}">${escapeHtml(template.label || template.id)}</option>`)
    .join("");

  const managerNames = getManagerNames();
  dom.fixedRuleManager.innerHTML =
    '<option value="">근무자 선택</option>' +
    managerNames
      .map(
        (name) =>
          `<option value="${escapeHtml(name)}">${escapeHtml(formatManagerDisplayName(name))}</option>`
      )
      .join("");

  syncFixedRuleModeUi();
}

function syncFixedRuleModeUi() {
  const mode = dom.fixedRuleMode.value === "date" ? "date" : "weekly";
  dom.fixedRuleWeekday.disabled = mode !== "weekly";
  dom.fixedRuleDate.disabled = mode !== "date";
}

function addFixedRule() {
  const mode = dom.fixedRuleMode.value === "date" ? "date" : "weekly";
  const shiftId = cleanText(dom.fixedRuleShift.value);
  const manager = stripManagerTitle(cleanText(dom.fixedRuleManager.value));
  const note = cleanText(dom.fixedRuleNote.value);

  if (!shiftId || !manager) {
    setStatus("고정 배정 규칙 추가 시 타임과 근무자는 필수입니다.", true);
    return;
  }

  const rule = {
    id: `rule-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
    mode,
    shiftId,
    weekday: mode === "weekly" ? toNumber(dom.fixedRuleWeekday.value, 1) : null,
    date: mode === "date" ? cleanText(dom.fixedRuleDate.value) : "",
    manager,
    note,
  };

  if (mode === "date" && !/^\d{4}-\d{2}-\d{2}$/.test(rule.date)) {
    setStatus("특정 날짜 규칙은 날짜를 입력해야 합니다.", true);
    return;
  }

  state.scheduleFixedRules.push(rule);
  persistConfig();
  renderScheduleAi();

  dom.fixedRuleNote.value = "";
}

function handleFixedRuleAction(event) {
  const target = event.target;
  if (!(target instanceof HTMLButtonElement)) return;
  if (!target.classList.contains("delete-fixed-rule-btn")) return;
  const ruleId = target.dataset.id;
  if (!ruleId) return;
  state.scheduleFixedRules = state.scheduleFixedRules.filter((item) => item.id !== ruleId);
  persistConfig();
  renderScheduleAi();
}

function renderFixedRuleTable() {
  if (!state.scheduleFixedRules.length) {
    dom.fixedRuleBody.innerHTML =
      '<tr><td colspan="6" class="empty">등록된 고정 배정 규칙이 없습니다.</td></tr>';
    return;
  }

  dom.fixedRuleBody.innerHTML = state.scheduleFixedRules
    .map((rule) => {
      const label = rule.mode === "weekly" ? "매주" : "날짜";
      const dayOrDate =
        rule.mode === "weekly"
          ? WEEKDAYS[toNumber(rule.weekday, 0)] || "-"
          : formatDateForDisplay(rule.date || "");
      const shiftLabel = getShiftTemplate(rule.shiftId)?.label || rule.shiftId;
      return `<tr>
        <td>${label}</td>
        <td>${dayOrDate}</td>
        <td>${escapeHtml(shiftLabel)}</td>
        <td>${escapeHtml(formatManagerDisplayName(rule.manager))}</td>
        <td>${escapeHtml(rule.note || "-")}</td>
        <td><button class="btn secondary small-btn delete-fixed-rule-btn" data-id="${rule.id}" type="button">삭제</button></td>
      </tr>`;
    })
    .join("");
}

function renderScheduleManagerTable() {
  const managers = getManagerNames();
  if (!managers.length) {
    dom.scheduleManagerBody.innerHTML =
      '<tr><td colspan="13" class="empty">관리자 데이터가 없습니다.</td></tr>';
    return;
  }

  dom.scheduleManagerBody.innerHTML = managers
    .map((name) => {
      const profile = ensureManagerProfile(name);
      const employee = ensureEmployeeSetting(name);
      const weekdaysHtml = WEEKDAY_ORDER.map((weekday) =>
        buildAvailabilityCell(name, profile, weekday)
      ).join("");
      return `<tr>
        <td>${escapeHtml(formatManagerDisplayName(name))}</td>
        <td>${formatWon(employee.hourlyRate)}</td>
        <td>
          <input class="schedule-manager-desired" data-name="${escapeHtml(
            name
          )}" type="number" min="0" step="1" value="${toNumber(profile.desiredShifts, 0)}" />
        </td>
        <td>
          <input class="schedule-manager-priority" data-name="${escapeHtml(
            name
          )}" type="number" min="1" step="1" value="${toNumber(profile.priority, 5)}" />
        </td>
        <td>
          <input class="schedule-manager-point" data-name="${escapeHtml(
            name
          )}" type="number" min="1" max="3" step="0.1" value="${toNumber(profile.pointPerShift, 1)}" />
        </td>
        ${weekdaysHtml}
        <td>
          <button class="btn secondary small-btn delete-manager-row-btn" data-name="${escapeHtml(
            name
          )}" type="button">삭제</button>
        </td>
      </tr>`;
    })
    .join("");
}

function buildAvailabilityCell(name, profile, weekday) {
  const shiftHtml = SHIFT_IDS.map((shiftId) => {
    const checked = Boolean(profile.availability?.[weekday]?.[shiftId]);
    return `<label class="tiny-check">
      <input class="schedule-manager-availability" data-name="${escapeHtml(
        name
      )}" data-day="${weekday}" data-shift-id="${shiftId}" type="checkbox" ${
        checked ? "checked" : ""
      } />
      <span>${shiftId}</span>
    </label>`;
  }).join("");
  return `<td><div class="tiny-check-grid">${shiftHtml}</div></td>`;
}

function handleScheduleManagerInput(event) {
  const target = event.target;
  if (!(target instanceof HTMLInputElement)) return;
  const name = target.dataset.name;
  if (!name) return;
  const profile = ensureManagerProfile(name);

  if (target.classList.contains("schedule-manager-desired")) {
    profile.desiredShifts = Math.max(0, toNumber(target.value, 0));
  }
  if (target.classList.contains("schedule-manager-priority")) {
    profile.priority = Math.max(1, Math.round(toNumber(target.value, 1)));
  }
  if (target.classList.contains("schedule-manager-point")) {
    profile.pointPerShift = Math.min(3, Math.max(1, round2(toNumber(target.value, 1))));
    if (event.type === "change") {
      syncCareerRatesToEmployeeSettings();
      renderEmployeeSettingsTable();
      renderStatement();
      renderSummary();
      renderCareer();
      renderScheduleAi();
    }
  }
  if (target.classList.contains("schedule-manager-availability")) {
    const weekday = toNumber(target.dataset.day, 0);
    const shiftId = cleanText(target.dataset.shiftId);
    ensureManagerProfileAvailability(profile, weekday);
    if (shiftId) {
      profile.availability[weekday][shiftId] = target.checked;
    }
  }

  persistConfig();
}

function handleScheduleManagerAction(event) {
  const target = event.target;
  if (!(target instanceof HTMLButtonElement)) return;
  if (!target.classList.contains("delete-manager-row-btn")) return;
  const name = cleanText(target.dataset.name);
  if (!name) return;
  removeManagerFromSchedule(name);
}

function addManagerRow() {
  const name = stripManagerTitle(cleanText(dom.newManagerNameInput.value));
  if (!name) {
    setStatus("추가할 관리자 이름을 입력하세요.", true);
    return;
  }
  if (getManagerNames().includes(name)) {
    setStatus("이미 등록된 관리자입니다.", true);
    return;
  }
  state.scheduleExcludedManagers = state.scheduleExcludedManagers.filter((item) => item !== name);
  state.managerProfiles[name] = createDefaultManagerProfile();
  ensureEmployeeSetting(name);
  dom.newManagerNameInput.value = "";
  syncCareerRatesToEmployeeSettings();
  persistConfig();
  renderScheduleAi();
  renderCareer();
  renderEmployeeSettingsTable();
  setStatus(`${name} 관리자를 배정 조건에 추가했습니다.`);
}

function setAllAvailability(checked) {
  getManagerNames().forEach((name) => {
    const profile = ensureManagerProfile(name);
    for (let day = 0; day < 7; day += 1) {
      ensureManagerProfileAvailability(profile, day);
      SHIFT_IDS.forEach((shiftId) => {
        profile.availability[day][shiftId] = checked;
      });
    }
  });
  persistConfig();
  renderScheduleManagerTable();
}

function setShiftAvailabilityForAll(shiftId, checked) {
  if (!SHIFT_IDS.includes(shiftId)) return;
  getManagerNames().forEach((name) => {
    const profile = ensureManagerProfile(name);
    for (let day = 0; day < 7; day += 1) {
      ensureManagerProfileAvailability(profile, day);
      profile.availability[day][shiftId] = checked;
    }
  });
  persistConfig();
  renderScheduleManagerTable();
}

function removeManagerFromSchedule(name) {
  if (!state.scheduleExcludedManagers.includes(name)) {
    state.scheduleExcludedManagers.push(name);
    state.scheduleExcludedManagers = normalizeNameList(state.scheduleExcludedManagers);
  }
  delete state.managerProfiles[name];
  state.scheduleFixedRules = state.scheduleFixedRules.filter((rule) => rule.manager !== name);
  Object.keys(state.scheduleAssignments).forEach((month) => {
    state.scheduleAssignments[month] = state.scheduleAssignments[month].map((row) => {
      const next = { ...row };
      if (next.primaryManager === name) {
        next.primaryManager = "";
        next.primaryStart = "";
        next.primaryEnd = "";
      }
      if (next.secondaryManager === name) {
        next.secondaryManager = "";
        next.secondaryStart = "";
        next.secondaryEnd = "";
      }
      return normalizeScheduleRow(next);
    });
  });
  syncCareerRatesToEmployeeSettings();
  persistConfig();
  renderScheduleAi();
  renderCareer();
  renderEmployeeSettingsTable();
  renderStatement();
  renderSummary();
  setStatus(`${name} 관리자를 배정 조건에서 삭제했습니다.`);
}

function generateAutoSchedule() {
  if (!state.scheduleMonth) {
    state.scheduleMonth = getDefaultScheduleMonth();
  }
  if (state.scheduleMonth < MIN_AUTO_SCHEDULE_MONTH) {
    setStatus(
      `자동 배정은 ${MIN_AUTO_SCHEDULE_MONTH}부터 사용합니다. 배정 월을 ${MIN_AUTO_SCHEDULE_MONTH} 이후로 선택하세요.`,
      true
    );
    return;
  }

  const managers = getManagerNames();
  if (!managers.length) {
    setStatus("자동 배정을 할 관리자 데이터가 없습니다.", true);
    return;
  }

  const rows = buildAutoScheduleForMonth(state.scheduleMonth, managers);
  state.scheduleAssignments[state.scheduleMonth] = rows;
  syncCareerRatesToEmployeeSettings();
  persistConfig();
  renderScheduleAi();
  renderCareer();
  renderEmployeeSettingsTable();
  renderStatement();
  renderSummary();
  setStatus(`${formatMonthLabel(state.scheduleMonth)} 관리자 자동 배정을 생성했습니다.`);
}

function buildAutoScheduleForMonth(month, managers) {
  const dates = listDatesInMonth(month);
  const assignmentMap = new Map(managers.map((name) => [name, { shiftUnits: 0, hours: 0 }]));
  const existingRows = getScheduleRowsForMonth(month);
  const workerSlotMap = new Map(
    existingRows.map((row) => [buildScheduleSlotKey(row.date, row.shiftId), toNumber(row.workerSlots, 1)])
  );
  const rows = [];

  dates.forEach((dateKey) => {
    const date = parseDateOnly(dateKey);
    if (!date) return;
    const weekday = date.getDay();

    state.scheduleTemplates.forEach((template) => {
      const fixedRule = findFixedRule(dateKey, weekday, template.id);
      const workerSlots = Math.min(
        2,
        Math.max(1, Math.round(toNumber(workerSlotMap.get(buildScheduleSlotKey(dateKey, template.id)), 1)))
      );
      let manager = fixedRule?.manager || "";
      let secondaryManager = "";

      if (!manager) {
        manager = pickAutoManager(managers, weekday, template.id, assignmentMap, new Set());
      }
      if (workerSlots > 1) {
        const excluded = new Set(manager ? [manager] : []);
        secondaryManager = pickAutoManager(managers, weekday, template.id, assignmentMap, excluded);
      }

      const shiftHours = calculateTimeRangeHours(dateKey, template.start, template.end);
      if (manager) {
        const current = assignmentMap.get(manager) || { shiftUnits: 0, hours: 0 };
        current.shiftUnits += 1;
        current.hours += shiftHours;
        assignmentMap.set(manager, current);
      }
      if (secondaryManager) {
        const current = assignmentMap.get(secondaryManager) || { shiftUnits: 0, hours: 0 };
        current.shiftUnits += 1;
        current.hours += shiftHours;
        assignmentMap.set(secondaryManager, current);
      }

      rows.push(
        normalizeScheduleRow({
          id: `sch-${month}-${dateKey}-${template.id}-${Math.random().toString(36).slice(2, 7)}`,
          date: dateKey,
          weekday,
          shiftId: template.id,
          workerSlots,
          shiftStart: template.start,
          shiftEnd: template.end,
          primaryManager: manager,
          primaryStart: manager ? template.start : "",
          primaryEnd: manager ? template.end : "",
          secondaryManager,
          secondaryStart: secondaryManager ? template.start : "",
          secondaryEnd: secondaryManager ? template.end : "",
          note: fixedRule?.note || "",
          source: fixedRule ? "고정규칙" : manager ? "자동배정" : "미배정",
        })
      );
    });
  });

  return rows.sort((a, b) => {
    if (a.date < b.date) return -1;
    if (a.date > b.date) return 1;
    return a.shiftId.localeCompare(b.shiftId);
  });
}

function findFixedRule(dateKey, weekday, shiftId) {
  const exact = state.scheduleFixedRules.find(
    (rule) => rule.mode === "date" && rule.date === dateKey && rule.shiftId === shiftId
  );
  if (exact) return exact;
  return state.scheduleFixedRules.find(
    (rule) => rule.mode === "weekly" && toNumber(rule.weekday, -1) === weekday && rule.shiftId === shiftId
  );
}

function pickAutoManager(managers, weekday, shiftId, assignmentMap, excluded = new Set()) {
  const candidates = managers.filter(
    (name) => !excluded.has(name) && isManagerAvailable(name, weekday, shiftId)
  );
  if (!candidates.length) return "";

  candidates.sort((a, b) => {
    const profileA = ensureManagerProfile(a);
    const profileB = ensureManagerProfile(b);
    const assignedA = toNumber(assignmentMap.get(a)?.shiftUnits, 0);
    const assignedB = toNumber(assignmentMap.get(b)?.shiftUnits, 0);
    const remainA = toNumber(profileA.desiredShifts, 0) - assignedA;
    const remainB = toNumber(profileB.desiredShifts, 0) - assignedB;

    const unmetA = remainA > 0 ? 0 : 1;
    const unmetB = remainB > 0 ? 0 : 1;
    if (unmetA !== unmetB) return unmetA - unmetB;
    if (remainA !== remainB) return remainB - remainA;

    const priorityA = toNumber(profileA.priority, 999);
    const priorityB = toNumber(profileB.priority, 999);
    if (priorityA !== priorityB) return priorityA - priorityB;

    if (assignedA !== assignedB) return assignedA - assignedB;
    return a.localeCompare(b, "ko-KR");
  });

  return candidates[0] || "";
}

function isManagerAvailable(name, weekday, shiftId) {
  const profile = ensureManagerProfile(name);
  ensureManagerProfileAvailability(profile, weekday);
  return Boolean(profile.availability?.[weekday]?.[shiftId]);
}

function renderScheduleResultTable() {
  const month = state.scheduleMonth;
  const rows = getScheduleRowsForMonth(month);
  if (!rows.length) {
    dom.scheduleCalendarBody.innerHTML =
      '<tr><td colspan="8" class="empty">자동 배정 결과가 없습니다.</td></tr>';
    return;
  }

  renderCalendarWeekHeader(dom.scheduleCalendarHead, true);
  const rowsByDateShift = new Map(
    rows.map((row) => [buildScheduleSlotKey(row.date, row.shiftId), row])
  );
  const weeks = buildCalendarWeeks(month);

  dom.scheduleCalendarBody.innerHTML = weeks
    .map((week) => {
      return renderAssignmentWeekRows(week, rowsByDateShift);
    })
    .join("");
}

function renderScheduleHeadcountCalendarTable() {
  const month = state.scheduleMonth;
  const rows = getScheduleRowsForMonth(month);
  if (!rows.length) {
    dom.scheduleHeadcountCalendarBody.innerHTML =
      '<tr><td colspan="7" class="empty">인원 설정 데이터가 없습니다.</td></tr>';
    return;
  }

  renderCalendarWeekHeader(dom.scheduleHeadcountCalendarHead);
  const rowsByDateShift = new Map(rows.map((row) => [buildScheduleSlotKey(row.date, row.shiftId), row]));
  const weeks = buildCalendarWeeks(month);

  dom.scheduleHeadcountCalendarBody.innerHTML = weeks
    .map((week) => {
      const cells = week.map((dateKey) => renderHeadcountDayCell(dateKey, rowsByDateShift)).join("");
      return `<tr>${cells}</tr>`;
    })
    .join("");
}

function renderCalendarWeekHeader(target, withShiftColumn = false) {
  const shiftHead = withShiftColumn ? '<th class="assignment-time-head">타임</th>' : "";
  target.innerHTML = `<tr>
    ${shiftHead}
    <th>월</th>
    <th>화</th>
    <th>수</th>
    <th>목</th>
    <th>금</th>
    <th>토</th>
    <th>일</th>
  </tr>`;
}

function renderAssignmentWeekRows(week, rowsByDateShift) {
  const dateCells = week.map((dateKey) => renderAssignmentDateCell(dateKey)).join("");

  const slotRows = SHIFT_IDS.map((shiftId) => {
    const slotCells = week
      .map((dateKey) => renderAssignmentSlotCell(dateKey, shiftId, rowsByDateShift))
      .join("");
    return `<tr class="assignment-shift-row">
      <th class="assignment-shift-label">${shiftId}</th>
      ${slotCells}
    </tr>`;
  }).join("");

  return `<tr class="assignment-date-row">
    <th class="assignment-shift-spacer"></th>
    ${dateCells}
  </tr>${slotRows}`;
}

function renderAssignmentDateCell(dateKey) {
  if (!dateKey) {
    return '<td class="assignment-date-cell assignment-empty-cell"></td>';
  }
  const date = parseDateOnly(dateKey);
  const dayLabel = Number(cleanText(dateKey).slice(-2));
  const dayIndex = date ? date.getDay() : 1;
  const dayName = WEEKDAYS[dayIndex] || "";
  const dayClass = getWeekdayCardClass(dayIndex);
  return `<td class="assignment-date-cell ${dayClass}">
    <div class="assignment-date-head"><span>${dayLabel}일</span><span>${dayName}</span></div>
  </td>`;
}

function renderAssignmentSlotCell(dateKey, shiftId, rowsByDateShift) {
  if (!dateKey) {
    return '<td class="assignment-slot-cell assignment-empty-cell"></td>';
  }
  const date = parseDateOnly(dateKey);
  const dayIndex = date ? date.getDay() : 1;
  const dayClass = getWeekdayCardClass(dayIndex);
  const row = rowsByDateShift.get(buildScheduleSlotKey(dateKey, shiftId));
  if (!row) {
    return `<td class="assignment-slot-cell ${dayClass}"></td>`;
  }
  const activeClass = row.id === state.selectedScheduleSlotId ? " active" : "";
  const slotStyle = buildAssignmentSlotInlineStyle(row);
  const styleAttr = slotStyle ? ` style="${slotStyle}"` : "";
  return `<td class="assignment-slot-cell ${dayClass}">
    <button class="schedule-slot-btn assignment-slot-btn${activeClass}" data-id="${row.id}" type="button"${styleAttr}>
      <div class="assignment-slot-managers">${buildAssignmentSlotLabelHtml(row)}</div>
    </button>
  </td>`;
}

function renderHeadcountDayCell(dateKey, rowsByDateShift) {
  if (!dateKey) {
    return '<td><div class="calendar-day-card calendar-day-card-empty"></div></td>';
  }
  const date = parseDateOnly(dateKey);
  const dayLabel = Number(cleanText(dateKey).slice(-2));
  const dayIndex = date ? date.getDay() : 1;
  const dayName = WEEKDAYS[dayIndex] || "";
  const dayClass = getWeekdayCardClass(dayIndex);
  const slotsHtml = SHIFT_IDS.map((shiftId) => {
    const row = rowsByDateShift.get(buildScheduleSlotKey(dateKey, shiftId));
    if (!row) return "";
    const shiftClass = getShiftClass(shiftId);
    const activeClass = row.id === state.selectedScheduleSlotId ? " active" : "";
    const selectedWorkers = Math.min(2, Math.max(1, Math.round(toNumber(row.workerSlots, 1))));
    return `<div class="headcount-slot ${shiftClass}${activeClass}">
      <button class="headcount-slot-btn" data-id="${row.id}" type="button">${shiftId}</button>
      <select class="headcount-slot-workers" data-id="${row.id}">
        <option value="1" ${selectedWorkers === 1 ? "selected" : ""}>1명</option>
        <option value="2" ${selectedWorkers === 2 ? "selected" : ""}>2명</option>
      </select>
    </div>`;
  }).join("");

  return `<td>
    <div class="calendar-day-card ${dayClass}">
      <div class="calendar-day-head"><span>${dayLabel}일</span><span>${dayName}</span></div>
      <div class="calendar-day-slots">${slotsHtml}</div>
    </div>
  </td>`;
}

function getWeekdayCardClass(dayIndex) {
  if (dayIndex === 1) return "day-w1";
  if (dayIndex === 2) return "day-w2";
  if (dayIndex === 3) return "day-w3";
  if (dayIndex === 4) return "day-w4";
  if (dayIndex === 5) return "day-w5";
  if (dayIndex === 6) return "day-w6";
  return "day-w0";
}

function getShiftClass(shiftId) {
  if (shiftId === "1T") return "shift-1t";
  if (shiftId === "2T") return "shift-2t";
  if (shiftId === "3T") return "shift-3t";
  return "";
}

function buildAssignmentSlotLabelHtml(row) {
  const names = [row.primaryManager, row.secondaryManager].filter(Boolean);
  if (!names.length) {
    return '<span class="assignment-slot-empty">미배정</span>';
  }
  const label = names.map((name) => formatManagerDisplayName(name)).join(" · ");
  return `<span class="assignment-slot-name" title="${escapeHtml(label)}">${escapeHtml(label)}</span>`;
}

function buildAssignmentSlotInlineStyle(row) {
  const names = [row.primaryManager, row.secondaryManager].filter(Boolean);
  if (!names.length) return "";
  const primary = getManagerColorStyle(names[0]);
  const textColor = "#20375f";
  if (names.length === 1) {
    return `background:${primary.bg};border-color:${primary.border};color:${textColor};`;
  }
  const secondary = getManagerColorStyle(names[1]);
  return `background:linear-gradient(120deg, ${primary.bg} 0%, ${primary.bg} 48%, ${secondary.bg} 52%, ${secondary.bg} 100%);border-color:${primary.border};color:${textColor};`;
}

function getManagerColorStyle(name) {
  const text = stripManagerTitle(cleanText(name));
  let hash = 0;
  for (let i = 0; i < text.length; i += 1) {
    hash = (hash * 31 + text.charCodeAt(i)) % 360;
  }
  const hue = (hash + 360) % 360;
  return {
    bg: `hsl(${hue} 70% 92%)`,
    border: `hsl(${hue} 55% 70%)`,
    text: `hsl(${hue} 50% 28%)`,
  };
}

function buildManagerOptionsHtml() {
  const managers = getManagerNames().map(
    (name) =>
      `<option value="${escapeHtml(name)}">${escapeHtml(formatManagerDisplayName(name))}</option>`
  );
  return ['<option value="">미배정</option>', ...managers].join("");
}

function handleScheduleCalendarClick(event) {
  const target = event.target;
  if (!(target instanceof HTMLElement)) return;
  const button = target.closest(".schedule-slot-btn, .headcount-slot-btn");
  if (!(button instanceof HTMLButtonElement)) return;
  const rowId = button.dataset.id;
  if (!rowId) return;
  state.selectedScheduleSlotId = state.selectedScheduleSlotId === rowId ? "" : rowId;
  persistConfig();
  renderScheduleAi();
}

function handleScheduleHeadcountCalendarChange(event) {
  const target = event.target;
  if (!(target instanceof HTMLSelectElement)) return;
  if (!target.classList.contains("headcount-slot-workers")) return;
  const rowId = target.dataset.id;
  if (!rowId) return;
  const rows = getScheduleRowsForMonth(state.scheduleMonth);
  const row = rows.find((item) => item.id === rowId);
  if (!row) return;
  row.workerSlots = Math.min(2, Math.max(1, Math.round(toNumber(target.value, 1))));
  autoFillManagersForRow(row, rows);
  if (row.workerSlots === 1) {
    row.secondaryManager = "";
    row.secondaryStart = "";
    row.secondaryEnd = "";
  }
  if (row.source !== "고정규칙") row.source = "수동수정";
  state.scheduleAssignments[state.scheduleMonth] = rows.map((item) =>
    item.id === row.id ? normalizeScheduleRow(row) : item
  );
  state.selectedScheduleSlotId = row.id;
  syncManagerProfiles();
  syncCareerRatesToEmployeeSettings();
  persistConfig();
  renderScheduleAi();
  renderCareer();
}

function autoFillManagersForRow(row, rows) {
  const managers = getManagerNames();
  const weekday = parseDateOnly(row.date)?.getDay();
  if (!Number.isFinite(weekday)) return;
  const assignmentMap = new Map();
  rows.forEach((item) => {
    if (item.id === row.id) return;
    if (item.primaryManager) {
      const current = assignmentMap.get(item.primaryManager) || { shiftUnits: 0, hours: 0 };
      current.shiftUnits += 1;
      assignmentMap.set(item.primaryManager, current);
    }
    if (item.secondaryManager) {
      const current = assignmentMap.get(item.secondaryManager) || { shiftUnits: 0, hours: 0 };
      current.shiftUnits += 1;
      assignmentMap.set(item.secondaryManager, current);
    }
  });

  if (!row.primaryManager) {
    row.primaryManager = pickAutoManager(managers, weekday, row.shiftId, assignmentMap, new Set());
    if (row.primaryManager) {
      row.primaryStart = row.primaryStart || row.shiftStart;
      row.primaryEnd = row.primaryEnd || row.shiftEnd;
    }
  }
  if (row.workerSlots > 1 && !row.secondaryManager) {
    const excluded = new Set(row.primaryManager ? [row.primaryManager] : []);
    row.secondaryManager = pickAutoManager(managers, weekday, row.shiftId, assignmentMap, excluded);
    if (row.secondaryManager) {
      row.secondaryStart = row.secondaryStart || row.shiftStart;
      row.secondaryEnd = row.secondaryEnd || row.shiftEnd;
    }
  }
}

function renderSelectedSlotEditor() {
  const row = getSelectedScheduleSlot();
  if (!row) {
    dom.selectedSlotMeta.textContent = "캘린더에서 타임을 선택하세요.";
    dom.selectedSlotWorkers.value = "1";
    dom.selectedSlotPrimaryManager.innerHTML = buildManagerOptionsHtml();
    dom.selectedSlotPrimaryStart.value = "";
    dom.selectedSlotPrimaryEnd.value = "";
    dom.selectedSlotSecondaryManager.innerHTML = buildManagerOptionsHtml();
    dom.selectedSlotSecondaryStart.value = "";
    dom.selectedSlotSecondaryEnd.value = "";
    dom.selectedSlotNote.value = "";
    toggleSecondaryEditor(false);
    return;
  }

  dom.selectedSlotMeta.textContent = `${formatDateForDisplay(row.date)} · ${row.shiftId} · ${row.shiftStart}~${
    row.shiftEnd
  }`;
  const optionsHtml = buildManagerOptionsHtml();
  dom.selectedSlotWorkers.value = String(Math.min(2, Math.max(1, Math.round(toNumber(row.workerSlots, 1)))));
  dom.selectedSlotPrimaryManager.innerHTML = optionsHtml;
  dom.selectedSlotPrimaryManager.value = row.primaryManager || "";
  dom.selectedSlotPrimaryStart.value = row.primaryStart || "";
  dom.selectedSlotPrimaryEnd.value = row.primaryEnd || "";
  dom.selectedSlotSecondaryManager.innerHTML = optionsHtml;
  dom.selectedSlotSecondaryManager.value = row.secondaryManager || "";
  dom.selectedSlotSecondaryStart.value = row.secondaryStart || "";
  dom.selectedSlotSecondaryEnd.value = row.secondaryEnd || "";
  dom.selectedSlotNote.value = row.note || "";
  toggleSecondaryEditor(row.workerSlots > 1);
}

function saveSelectedScheduleSlot() {
  const row = getSelectedScheduleSlot();
  if (!row) {
    setStatus("먼저 캘린더에서 수정할 타임을 선택하세요.", true);
    return;
  }

  const workerSlots = Math.min(2, Math.max(1, Math.round(toNumber(dom.selectedSlotWorkers.value, 1))));
  row.workerSlots = workerSlots;
  row.primaryManager = cleanText(dom.selectedSlotPrimaryManager.value);
  row.primaryStart = normalizeTimeText(dom.selectedSlotPrimaryStart.value, "");
  row.primaryEnd = normalizeTimeText(dom.selectedSlotPrimaryEnd.value, "");
  row.secondaryManager = workerSlots > 1 ? cleanText(dom.selectedSlotSecondaryManager.value) : "";
  row.secondaryStart = workerSlots > 1 ? normalizeTimeText(dom.selectedSlotSecondaryStart.value, "") : "";
  row.secondaryEnd = workerSlots > 1 ? normalizeTimeText(dom.selectedSlotSecondaryEnd.value, "") : "";
  row.note = cleanText(dom.selectedSlotNote.value);

  if (row.secondaryManager && row.secondaryManager === row.primaryManager) {
    row.secondaryManager = "";
    row.secondaryStart = "";
    row.secondaryEnd = "";
  }
  if (row.workerSlots === 1) {
    row.secondaryManager = "";
    row.secondaryStart = "";
    row.secondaryEnd = "";
  }
  if (row.source !== "고정규칙") row.source = "수동수정";

  const rows = getScheduleRowsForMonth(state.scheduleMonth);
  state.scheduleAssignments[state.scheduleMonth] = rows.map((item) =>
    item.id === row.id ? normalizeScheduleRow(row) : item
  );

  syncManagerProfiles();
  syncCareerRatesToEmployeeSettings();
  persistConfig();
  renderScheduleAi();
  renderCareer();
  renderEmployeeSettingsTable();
  renderStatement();
  renderSummary();
}

function toggleSecondaryEditor(enabled) {
  dom.selectedSlotSecondaryManager.disabled = !enabled;
  dom.selectedSlotSecondaryStart.disabled = !enabled;
  dom.selectedSlotSecondaryEnd.disabled = !enabled;
}

function getSelectedScheduleSlot() {
  if (!state.selectedScheduleSlotId) return null;
  const rows = getScheduleRowsForMonth(state.scheduleMonth);
  return rows.find((item) => item.id === state.selectedScheduleSlotId) || null;
}

function renderScheduleSummaryTable() {
  const month = state.scheduleMonth;
  const summary = getScheduleMonthContribution(month);
  const rows = Array.from(summary.entries()).sort((a, b) => b[1].shiftUnits - a[1].shiftUnits);

  if (!rows.length) {
    dom.scheduleSummaryBody.innerHTML =
      '<tr><td colspan="3" class="empty">요약할 배정 데이터가 없습니다.</td></tr>';
    return;
  }

  dom.scheduleSummaryBody.innerHTML = rows
    .map(
      ([name, item]) => `<tr>
      <td>${escapeHtml(formatManagerDisplayName(name))}</td>
      <td>${round2(item.shiftUnits)}T</td>
      <td>${formatDurationText(item.hours)}</td>
    </tr>`
    )
    .join("");
}

function getScheduleRowsForMonth(month) {
  const key = cleanText(month || "");
  if (!key) return [];
  if (!Array.isArray(state.scheduleAssignments[key])) {
    state.scheduleAssignments[key] = [];
  }
  return state.scheduleAssignments[key];
}

function ensureScheduleRowsForMonth(month) {
  const key = cleanText(month || "");
  if (!/^\d{4}-\d{2}$/.test(key)) return;
  const rows = getScheduleRowsForMonth(key);
  if (rows.length) return;
  const generated = [];
  listDatesInMonth(key).forEach((dateKey) => {
    const date = parseDateOnly(dateKey);
    if (!date) return;
    const weekday = date.getDay();
    state.scheduleTemplates.forEach((template) => {
      generated.push(
        normalizeScheduleRow({
          id: `sch-init-${key}-${dateKey}-${template.id}`,
          date: dateKey,
          weekday,
          shiftId: template.id,
          workerSlots: 1,
          shiftStart: template.start,
          shiftEnd: template.end,
          primaryManager: "",
          primaryStart: "",
          primaryEnd: "",
          secondaryManager: "",
          secondaryStart: "",
          secondaryEnd: "",
          note: "",
          source: "미배정",
        })
      );
    });
  });
  state.scheduleAssignments[key] = generated;
  persistConfig();
}

function ensureSelectedScheduleSlot(month) {
  const rows = getScheduleRowsForMonth(month);
  if (!rows.length) {
    state.selectedScheduleSlotId = "";
    return;
  }
  if (!state.selectedScheduleSlotId) {
    return;
  }
  const exists = rows.some((row) => row.id === state.selectedScheduleSlotId);
  if (!exists) {
    state.selectedScheduleSlotId = "";
  }
}

function buildScheduleSlotKey(dateKey, shiftId) {
  return `${dateKey}__${shiftId}`;
}

function buildCalendarWeeks(month) {
  const dates = listDatesInMonth(month);
  if (!dates.length) return [];
  const firstDate = parseDateOnly(`${month}-01`);
  if (!firstDate) return [];
  const mondayStartIndex = (firstDate.getDay() + 6) % 7;
  const cells = new Array(mondayStartIndex).fill("");
  cells.push(...dates);
  while (cells.length % 7 !== 0) {
    cells.push("");
  }
  const weeks = [];
  for (let i = 0; i < cells.length; i += 7) {
    weeks.push(cells.slice(i, i + 7));
  }
  return weeks;
}

function getScheduleMonthContribution(month) {
  const result = new Map();
  const rows = getScheduleRowsForMonth(month);
  rows.forEach((row) => addScheduleRowContribution(result, row));
  return result;
}

function addScheduleRowContribution(map, row) {
  const shiftHours = calculateTimeRangeHours(row.date, row.shiftStart, row.shiftEnd);
  const primaryHours =
    row.primaryManager && row.primaryStart && row.primaryEnd
      ? calculateTimeRangeHours(row.date, row.primaryStart, row.primaryEnd)
      : row.primaryManager
        ? shiftHours
        : 0;
  const secondaryHours =
    row.secondaryManager && row.secondaryStart && row.secondaryEnd
      ? calculateTimeRangeHours(row.date, row.secondaryStart, row.secondaryEnd)
      : row.secondaryManager
        ? shiftHours
        : 0;

  if (row.primaryManager) {
    const info = map.get(row.primaryManager) || { shiftUnits: 0, hours: 0 };
    info.hours += primaryHours;
    info.shiftUnits += shiftHours > 0 ? primaryHours / shiftHours : 0;
    map.set(row.primaryManager, info);
  }
  if (row.secondaryManager) {
    const info = map.get(row.secondaryManager) || { shiftUnits: 0, hours: 0 };
    info.hours += secondaryHours;
    info.shiftUnits += shiftHours > 0 ? secondaryHours / shiftHours : 0;
    map.set(row.secondaryManager, info);
  }
}

function calculateTimeRangeHours(dateKey, startText, endText) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(cleanText(dateKey))) return 0;
  const hours = calculateDurationFromTimeInputs(dateKey, startText, endText);
  return Number.isFinite(hours) && hours > 0 ? round2(hours) : 0;
}

function listDatesInMonth(month) {
  const match = /^(\d{4})-(\d{2})$/.exec(month || "");
  if (!match) return [];
  const year = toNumber(match[1], 0);
  const monthIndex = toNumber(match[2], 1) - 1;
  const lastDate = new Date(year, monthIndex + 1, 0).getDate();
  const dates = [];
  for (let day = 1; day <= lastDate; day += 1) {
    dates.push(`${year}-${pad2(monthIndex + 1)}-${pad2(day)}`);
  }
  return dates;
}

function getShiftTemplate(shiftId) {
  return state.scheduleTemplates.find((item) => item.id === shiftId) || null;
}

function syncManagerProfiles() {
  const managers = getManagerNames();
  managers.forEach((name) => ensureManagerProfile(name));
}

function ensureManagerProfile(name) {
  const key = cleanText(name);
  if (!key) {
    return createDefaultManagerProfile();
  }
  if (!state.managerProfiles[key]) {
    state.managerProfiles[key] = createDefaultManagerProfile();
  }
  const profile = state.managerProfiles[key];
  profile.desiredShifts = Math.max(0, toNumber(profile.desiredShifts, 0));
  profile.priority = Math.max(1, Math.round(toNumber(profile.priority, 5)));
  profile.pointPerShift = Math.min(3, Math.max(1, round2(toNumber(profile.pointPerShift, 1))));
  profile.basePoints = Math.max(0, round2(toNumber(profile.basePoints, 0)));
  profile.baseHours = Math.max(0, round2(toNumber(profile.baseHours, 0)));
  if (!profile.availability || typeof profile.availability !== "object") {
    profile.availability = {};
  }
  for (let day = 0; day < 7; day += 1) {
    ensureManagerProfileAvailability(profile, day);
  }
  return profile;
}

function ensureManagerProfileAvailability(profile, day) {
  if (!profile.availability[day] || typeof profile.availability[day] !== "object") {
    profile.availability[day] = {};
  }
  SHIFT_IDS.forEach((shiftId) => {
    if (typeof profile.availability[day][shiftId] !== "boolean") {
      profile.availability[day][shiftId] = true;
    }
  });
}

function createDefaultManagerProfile() {
  const availability = {};
  for (let day = 0; day < 7; day += 1) {
    availability[day] = {};
    SHIFT_IDS.forEach((shiftId) => {
      availability[day][shiftId] = true;
    });
  }
  return {
    desiredShifts: 0,
    priority: 5,
    pointPerShift: 1,
    basePoints: 0,
    baseHours: 0,
    availability,
  };
}

function getManagerNames() {
  const excluded = new Set(state.scheduleExcludedManagers || []);
  const names = new Set();
  state.entries.forEach((entry) => {
    if (isManagerCategory(entry.category) && !excluded.has(entry.name)) {
      names.add(entry.name);
    }
  });
  Object.keys(state.managerProfiles || {}).forEach((name) => {
    if (!excluded.has(name)) names.add(name);
  });
  state.scheduleFixedRules.forEach((rule) => {
    if (rule.manager && !excluded.has(rule.manager)) names.add(rule.manager);
  });
  Object.values(state.scheduleAssignments || {}).forEach((rows) => {
    rows.forEach((row) => {
      if (row.primaryManager && !excluded.has(row.primaryManager)) names.add(row.primaryManager);
      if (row.secondaryManager && !excluded.has(row.secondaryManager)) names.add(row.secondaryManager);
    });
  });
  if (!names.size) {
    state.entries.forEach((entry) => {
      if (!excluded.has(entry.name)) names.add(entry.name);
    });
  }
  return Array.from(names).sort((a, b) => a.localeCompare(b, "ko-KR"));
}

function isManagerCategory(category) {
  return /관리/.test(cleanText(category));
}

function stripManagerTitle(name) {
  return cleanText(name).replace(/\s*쌤$/, "");
}

function formatManagerDisplayName(name) {
  const base = stripManagerTitle(name);
  return base ? `${base}쌤` : "";
}

function renderCareer() {
  syncManagerProfiles();
  renderCareerMonthOptions();
  renderCareerLevelTable();
  syncCareerRatesToEmployeeSettings();
  renderCareerManagerTable();
}

function renderCareerMonthOptions() {
  const monthSet = new Set([
    ...getAvailableMonths(),
    ...Object.keys(state.scheduleAssignments || {}),
    state.scheduleMonth,
    state.selectedMonth,
    state.careerMonth,
  ]);
  const months = Array.from(monthSet)
    .filter((month) => /^\d{4}-\d{2}$/.test(month))
    .sort();

  if (!months.length) {
    dom.careerMonthSelect.innerHTML = '<option value="">월 선택</option>';
    state.careerMonth = state.careerMonth || state.scheduleMonth || "";
    dom.careerMonthSelect.value = state.careerMonth;
    return;
  }

  const fallbackMonth = months[months.length - 1];
  state.careerMonth = months.includes(state.careerMonth) ? state.careerMonth : fallbackMonth;
  dom.careerMonthSelect.innerHTML = months
    .map((month) => `<option value="${month}">${formatMonthLabel(month)}</option>`)
    .join("");
  dom.careerMonthSelect.value = state.careerMonth;
}

function renderCareerLevelTable() {
  if (!state.careerLevels.length) {
    dom.careerLevelBody.innerHTML = '<tr><td colspan="5" class="empty">직급 기준이 없습니다.</td></tr>';
    return;
  }
  dom.careerLevelBody.innerHTML = state.careerLevels
    .map(
      (level) => `<tr>
      <td><input class="career-level-name" data-id="${level.id}" type="text" value="${escapeHtml(
        level.name
      )}" /></td>
      <td><input class="career-level-min" data-id="${level.id}" type="number" min="0" step="0.1" value="${toNumber(
        level.min,
        0
      )}" /></td>
      <td><input class="career-level-max" data-id="${level.id}" type="number" min="0" step="0.1" value="${
        Number.isFinite(level.max) ? toNumber(level.max, 0) : ""
      }" placeholder="비우면 무제한" /></td>
      <td><input class="career-level-wage" data-id="${level.id}" type="number" min="0" step="10" value="${toNumber(
        level.wage,
        0
      )}" /></td>
      <td><button class="btn secondary small-btn delete-career-level-btn" data-id="${
        level.id
      }" type="button">삭제</button></td>
    </tr>`
    )
    .join("");
}

function addCareerLevel() {
  state.careerLevels.push({
    id: `level-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
    name: "New Level",
    min: 0,
    max: null,
    wage: state.settings.defaultHourlyWage,
  });
  state.careerLevels = normalizeCareerLevels(state.careerLevels);
  syncCareerRatesToEmployeeSettings();
  persistConfig();
  renderCareer();
  renderScheduleAi();
  renderEmployeeSettingsTable();
  renderStatement();
  renderSummary();
}

function handleCareerLevelInput(event) {
  const target = event.target;
  if (!(target instanceof HTMLInputElement)) return;
  const levelId = target.dataset.id;
  if (!levelId) return;
  const level = state.careerLevels.find((item) => item.id === levelId);
  if (!level) return;

  if (target.classList.contains("career-level-name")) {
    level.name = cleanText(target.value) || "Unnamed";
  }
  if (target.classList.contains("career-level-min")) {
    level.min = Math.max(0, round2(toNumber(target.value, 0)));
  }
  if (target.classList.contains("career-level-max")) {
    const raw = cleanText(target.value);
    level.max = raw ? Math.max(0, round2(toNumber(raw, 0))) : null;
  }
  if (target.classList.contains("career-level-wage")) {
    level.wage = Math.max(0, roundCurrency(toNumber(target.value, state.settings.defaultHourlyWage)));
  }

  if (event.type !== "change") {
    persistConfig();
    return;
  }

  state.careerLevels = normalizeCareerLevels(state.careerLevels);
  syncCareerRatesToEmployeeSettings();
  persistConfig();
  renderCareer();
  renderScheduleAi();
  renderEmployeeSettingsTable();
  renderStatement();
  renderSummary();
}

function handleCareerLevelAction(event) {
  const target = event.target;
  if (!(target instanceof HTMLButtonElement)) return;
  if (!target.classList.contains("delete-career-level-btn")) return;
  const levelId = target.dataset.id;
  if (!levelId) return;
  state.careerLevels = state.careerLevels.filter((item) => item.id !== levelId);
  if (!state.careerLevels.length) {
    state.careerLevels = cloneDefaultCareerLevels();
  }
  state.careerLevels = normalizeCareerLevels(state.careerLevels);
  syncCareerRatesToEmployeeSettings();
  persistConfig();
  renderCareer();
  renderScheduleAi();
  renderEmployeeSettingsTable();
  renderStatement();
  renderSummary();
}

function renderCareerManagerTable() {
  const managers = getManagerNames();
  if (!managers.length) {
    dom.careerManagerBody.innerHTML = '<tr><td colspan="9" class="empty">관리자 데이터가 없습니다.</td></tr>';
    return;
  }

  dom.careerManagerBody.innerHTML = managers
    .map((name) => {
      const profile = ensureManagerProfile(name);
      const snapshot = computeManagerCareerSnapshot(name, state.careerMonth);
      return `<tr>
        <td>${escapeHtml(name)}</td>
        <td><input class="career-manager-base-points" data-name="${escapeHtml(
          name
        )}" type="number" min="0" step="0.1" value="${toNumber(profile.basePoints, 0)}" /></td>
        <td><input class="career-manager-base-hours" data-name="${escapeHtml(
          name
        )}" type="number" min="0" step="0.1" value="${toNumber(profile.baseHours, 0)}" /></td>
        <td><input class="career-manager-point" data-name="${escapeHtml(
          name
        )}" type="number" min="1" max="3" step="0.1" value="${toNumber(profile.pointPerShift, 1)}" /></td>
        <td>${round2(snapshot.assignedShiftUnits)}T</td>
        <td>${formatDurationText(snapshot.assignedHours)}</td>
        <td>${round2(snapshot.totalPoints)}P</td>
        <td>${escapeHtml(snapshot.levelName)}</td>
        <td>${formatWon(snapshot.appliedWage)}</td>
      </tr>`;
    })
    .join("");
}

function handleCareerManagerInput(event) {
  const target = event.target;
  if (!(target instanceof HTMLInputElement)) return;
  const name = target.dataset.name;
  if (!name) return;
  const profile = ensureManagerProfile(name);

  if (target.classList.contains("career-manager-base-points")) {
    profile.basePoints = Math.max(0, round2(toNumber(target.value, 0)));
  }
  if (target.classList.contains("career-manager-base-hours")) {
    profile.baseHours = Math.max(0, round2(toNumber(target.value, 0)));
  }
  if (target.classList.contains("career-manager-point")) {
    profile.pointPerShift = Math.min(3, Math.max(1, round2(toNumber(target.value, 1))));
  }

  if (event.type !== "change") {
    persistConfig();
    return;
  }

  syncCareerRatesToEmployeeSettings();
  persistConfig();
  renderCareer();
  renderScheduleAi();
  renderEmployeeSettingsTable();
  renderStatement();
  renderSummary();
}

function computeManagerCareerSnapshot(name, limitMonth) {
  const profile = ensureManagerProfile(name);
  const contribution = getManagerContributionUpToMonth(name, limitMonth);
  const assignedShiftUnits = round2(contribution.shiftUnits);
  const assignedHours = round2(contribution.hours);
  const totalPoints = round2(profile.basePoints + assignedShiftUnits * profile.pointPerShift);
  const level = resolveCareerLevel(totalPoints);
  const appliedWage = level ? level.wage : ensureEmployeeSetting(name).hourlyRate;
  return {
    name,
    assignedShiftUnits,
    assignedHours,
    totalPoints,
    levelName: level ? level.name : "-",
    appliedWage: Math.max(0, roundCurrency(appliedWage)),
  };
}

function getManagerContributionUpToMonth(name, limitMonth) {
  const result = { shiftUnits: 0, hours: 0 };
  const months = Object.keys(state.scheduleAssignments || {})
    .filter((month) => /^\d{4}-\d{2}$/.test(month))
    .sort();

  months.forEach((month) => {
    if (limitMonth && month > limitMonth) return;
    const rows = getScheduleRowsForMonth(month);
    rows.forEach((row) => {
      const temp = new Map();
      addScheduleRowContribution(temp, row);
      const info = temp.get(name);
      if (!info) return;
      result.shiftUnits += info.shiftUnits;
      result.hours += info.hours;
    });
  });

  return {
    shiftUnits: round2(result.shiftUnits),
    hours: round2(result.hours),
  };
}

function resolveCareerLevel(points) {
  const target = toNumber(points, 0);
  const levels = state.careerLevels.slice().sort((a, b) => toNumber(a.min, 0) - toNumber(b.min, 0));
  return (
    levels.find((level) => {
      const min = toNumber(level.min, 0);
      const max = Number.isFinite(level.max) ? toNumber(level.max, Infinity) : Infinity;
      return target >= min && target <= max;
    }) || null
  );
}

function syncCareerRatesToEmployeeSettings() {
  const managers = getManagerNames();
  if (!managers.length) return false;
  const refMonth = state.careerMonth || state.scheduleMonth || state.selectedMonth;
  let changed = false;
  managers.forEach((name) => {
    const snapshot = computeManagerCareerSnapshot(name, refMonth);
    const config = ensureEmployeeSetting(name);
    if (roundCurrency(config.hourlyRate) !== roundCurrency(snapshot.appliedWage)) {
      config.hourlyRate = roundCurrency(snapshot.appliedWage);
      changed = true;
    }
  });
  return changed;
}

function computeEmployeeMonthPayroll(name, month) {
  const config = ensureEmployeeSetting(name);
  const hourlyRate = Math.max(0, toNumber(config.hourlyRate, state.settings.defaultHourlyWage));

  const shifts = state.entries
    .filter((entry) => entry.name === name && entry.workDate.startsWith(month))
    .map((entry) => computeShiftPayroll(entry, hourlyRate, config))
    .sort((a, b) => compareEntries(a.entry, b.entry));

  const weekMap = new Map();
  shifts.forEach((shift) => {
    const existing = weekMap.get(shift.weekIndex);
    if (existing) {
      existing.hours += shift.adjustedHours;
      existing.basePay += shift.basePay;
      existing.overtimePremium += shift.overtimePremium;
      existing.nightPremium += shift.nightPremium;
      existing.holidayPremium += shift.holidayPremium;
      existing.totalPay += shift.shiftPay;
      existing.daySet.add(shift.entry.workDate);
    } else {
      weekMap.set(shift.weekIndex, {
        weekIndex: shift.weekIndex,
        hours: shift.adjustedHours,
        basePay: shift.basePay,
        overtimePremium: shift.overtimePremium,
        nightPremium: shift.nightPremium,
        holidayPremium: shift.holidayPremium,
        weeklyHolidayPay: 0,
        totalPay: shift.shiftPay,
        daySet: new Set([shift.entry.workDate]),
      });
    }
  });

  if (config.useWeeklyHoliday) {
    weekMap.forEach((week) => {
      const threshold = Math.max(0, toNumber(state.settings.weeklyThresholdHours, 15));
      if (week.hours >= threshold) {
        const weeklyHolidayHours = Math.min(8, Math.max(0, (week.hours / 40) * 8));
        week.weeklyHolidayPay = roundCurrency(weeklyHolidayHours * hourlyRate);
        week.totalPay += week.weeklyHolidayPay;
      }
    });
  }

  const weeks = Array.from(weekMap.values())
    .map((week) => ({
      ...week,
      hours: round2(week.hours),
      basePay: roundCurrency(week.basePay),
      overtimePremium: roundCurrency(week.overtimePremium),
      nightPremium: roundCurrency(week.nightPremium),
      holidayPremium: roundCurrency(week.holidayPremium),
      weeklyHolidayPay: roundCurrency(week.weeklyHolidayPay),
      totalPay: roundCurrency(week.totalPay),
      workDays: week.daySet.size,
    }))
    .sort((a, b) => a.weekIndex - b.weekIndex);

  const totals = shifts.reduce(
    (acc, shift) => {
      acc.hours += shift.adjustedHours;
      acc.basePay += shift.basePay;
      acc.overtimePremium += shift.overtimePremium;
      acc.nightPremium += shift.nightPremium;
      acc.holidayPremium += shift.holidayPremium;
      return acc;
    },
    { hours: 0, basePay: 0, overtimePremium: 0, nightPremium: 0, holidayPremium: 0 }
  );

  const weeklyHolidayPay = weeks.reduce((sum, week) => sum + week.weeklyHolidayPay, 0);
  const allowances = totals.overtimePremium + totals.nightPremium + totals.holidayPremium + weeklyHolidayPay;
  const grossPay = totals.basePay + allowances;
  const tax = grossPay * (Math.max(0, toNumber(state.settings.taxRate, 0)) / 100);
  const netPay = grossPay - tax;

  return {
    shifts,
    weeks,
    totals: {
      hours: round2(totals.hours),
      basePay: roundCurrency(totals.basePay),
      overtimePremium: roundCurrency(totals.overtimePremium),
      nightPremium: roundCurrency(totals.nightPremium),
      holidayPremium: roundCurrency(totals.holidayPremium),
      weeklyHolidayPay: roundCurrency(weeklyHolidayPay),
      allowances: roundCurrency(allowances),
      grossPay: roundCurrency(grossPay),
      netPay: roundCurrency(netPay),
    },
  };
}

function computeShiftPayroll(entry, hourlyRate, config) {
  const adjustedHours = Math.max(0, toNumber(entry.adjustedHours, 0));
  const start = parseDateTimeKey(entry.startIso);
  const end = parseDateTimeKey(entry.endIso);
  const clockHours = start && end ? Math.max(0, diffHours(start, end)) : 0;
  const ratio = clockHours > 0 ? adjustedHours / clockHours : 1;

  const overtimeHours = config.useOvertimePremium ? Math.max(0, adjustedHours - 8) : 0;
  const rawNightHours =
    config.useNightPremium && start && end
      ? calculateNightHours(
          start,
          end,
          config.nightStart || state.settings.defaultNightStart,
          state.settings.nightEndTime
        )
      : 0;
  const nightHours = Math.min(adjustedHours, rawNightHours * ratio);
  const isHolidayDate = isOfficialHolidayDate(entry.workDate);
  const holidayHours = config.useHolidayPremium && isHolidayDate ? adjustedHours : 0;

  const basePay = adjustedHours * hourlyRate;
  const overtimePremium = overtimeHours * hourlyRate * 0.5;
  const nightPremium = nightHours * hourlyRate * 0.5;
  const holidayPremium = holidayHours * hourlyRate * 0.5;
  const shiftPay = basePay + overtimePremium + nightPremium + holidayPremium;

  return {
    entry,
    weekIndex: getWeekOfMonth(entry.workDate),
    adjustedHours: round2(adjustedHours),
    clockHours: round2(clockHours),
    overtimeHours: round2(overtimeHours),
    nightHours: round2(nightHours),
    holidayHours: round2(holidayHours),
    isHolidayDate,
    basePay: roundCurrency(basePay),
    overtimePremium: roundCurrency(overtimePremium),
    nightPremium: roundCurrency(nightPremium),
    holidayPremium: roundCurrency(holidayPremium),
    shiftPay: roundCurrency(shiftPay),
  };
}

function calculateNightHours(start, end, nightStart, nightEnd) {
  const [startHour, startMinute] = parseTimeParts(nightStart, "22:30");
  const [endHour, endMinute] = parseTimeParts(nightEnd, "06:00");

  let cursor = startOfDay(addDays(start, -1));
  const last = startOfDay(end);
  let total = 0;

  while (cursor <= last) {
    const windowStart = new Date(cursor);
    windowStart.setHours(startHour, startMinute, 0, 0);
    const windowEnd = new Date(cursor);
    windowEnd.setHours(endHour, endMinute, 0, 0);
    if (windowEnd <= windowStart) {
      windowEnd.setDate(windowEnd.getDate() + 1);
    }
    total += overlapHours(start, end, windowStart, windowEnd);
    cursor.setDate(cursor.getDate() + 1);
  }

  return round2(total);
}

function overlapHours(aStart, aEnd, bStart, bEnd) {
  const start = Math.max(aStart.getTime(), bStart.getTime());
  const end = Math.min(aEnd.getTime(), bEnd.getTime());
  if (end <= start) return 0;
  return (end - start) / 3_600_000;
}

function getWeekOfMonth(dateKey) {
  const date = parseDateOnly(dateKey);
  if (!date) return 0;
  const first = new Date(date.getFullYear(), date.getMonth(), 1);
  return Math.floor((date.getDate() + first.getDay() - 1) / 7) + 1;
}

function isOfficialHolidayDate(dateKey) {
  const date = parseDateOnly(dateKey);
  if (!date) return false;
  if (date.getDay() === 0) return false;
  return state.holidays.includes(dateKey);
}

function getAvailableMonths() {
  const months = Array.from(new Set(state.entries.map((entry) => entry.workDate.slice(0, 7))));
  return months.sort();
}

function getEmployees() {
  const names = new Set(state.entries.map((entry) => entry.name));
  Object.keys(state.managerProfiles || {}).forEach((name) => names.add(name));
  return Array.from(names).sort((a, b) => a.localeCompare(b, "ko-KR"));
}

function getEmployeesInMonth(month) {
  if (!month) return [];
  return Array.from(
    new Set(state.entries.filter((entry) => entry.workDate.startsWith(month)).map((entry) => entry.name))
  ).sort((a, b) => a.localeCompare(b, "ko-KR"));
}

function getMostUsedCategory(name) {
  const counts = new Map();
  state.entries
    .filter((entry) => entry.name === name && entry.category)
    .forEach((entry) => {
      counts.set(entry.category, (counts.get(entry.category) || 0) + 1);
    });
  let top = "";
  let max = 0;
  counts.forEach((count, category) => {
    if (count > max) {
      max = count;
      top = category;
    }
  });
  return top;
}

function syncEmployeeSettings() {
  getEmployees().forEach((name) => {
    ensureEmployeeSetting(name);
  });
}

function ensureEmployeeSetting(name) {
  if (!state.employeeSettings[name]) {
    state.employeeSettings[name] = {
      hourlyRate: Math.max(0, toNumber(state.settings.defaultHourlyWage, 0)),
      nightStart: state.settings.defaultNightStart,
      useWeeklyHoliday: Boolean(state.settings.useWeeklyHoliday),
      useHolidayPremium: Boolean(state.settings.useHolidayPremium),
      useOvertimePremium: Boolean(state.settings.useOvertimePremium),
      useNightPremium: Boolean(state.settings.useNightPremium),
    };
  }
  if (!state.employeeSettings[name].nightStart) {
    state.employeeSettings[name].nightStart = state.settings.defaultNightStart;
  }
  if (!Number.isFinite(toNumber(state.employeeSettings[name].hourlyRate, NaN))) {
    state.employeeSettings[name].hourlyRate = state.settings.defaultHourlyWage;
  }
  if (typeof state.employeeSettings[name].useWeeklyHoliday !== "boolean") {
    state.employeeSettings[name].useWeeklyHoliday = Boolean(state.settings.useWeeklyHoliday);
  }
  if (typeof state.employeeSettings[name].useHolidayPremium !== "boolean") {
    state.employeeSettings[name].useHolidayPremium = Boolean(state.settings.useHolidayPremium);
  }
  if (typeof state.employeeSettings[name].useOvertimePremium !== "boolean") {
    state.employeeSettings[name].useOvertimePremium = Boolean(state.settings.useOvertimePremium);
  }
  if (typeof state.employeeSettings[name].useNightPremium !== "boolean") {
    state.employeeSettings[name].useNightPremium = Boolean(state.settings.useNightPremium);
  }
  return state.employeeSettings[name];
}

function parseExcelDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return startOfDay(value);
  }
  if (typeof value === "number" && Number.isFinite(value)) {
    return startOfDay(excelSerialToDate(value));
  }
  if (typeof value === "string") {
    const trimmed = value.trim();
    if (!trimmed) return null;
    const parsed = parseDateString(trimmed);
    return parsed ? startOfDay(parsed) : null;
  }
  return null;
}

function parseExcelDateTime(value, fallbackDate) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return new Date(value);
  }
  if (typeof value === "number" && Number.isFinite(value)) {
    const parsed = excelSerialToDate(value);
    if (parsed.getFullYear() < 1900 && fallbackDate) {
      const merged = new Date(fallbackDate);
      merged.setHours(parsed.getHours(), parsed.getMinutes(), parsed.getSeconds(), 0);
      return merged;
    }
    return parsed;
  }
  if (typeof value === "string") {
    const trimmed = value.trim();
    if (!trimmed) return null;
    const timeOnly = /^(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?$/.exec(trimmed);
    if (timeOnly) {
      if (!fallbackDate) return null;
      const hours = toNumber(timeOnly[1], 0);
      const minutes = toNumber(timeOnly[2], 0);
      const seconds = toNumber(timeOnly[3], 0);
      const merged = new Date(fallbackDate);
      merged.setHours(hours, minutes, seconds, 0);
      return merged;
    }
    const parsed = parseDateString(trimmed);
    return parsed;
  }
  return null;
}

function parseDateString(value) {
  const normalized = value
    .replace(/\./g, "-")
    .replace(/\//g, "-")
    .replace(/\s+/g, " ")
    .trim();

  const krMeridiem = /(\d{4}-\d{1,2}-\d{1,2})\s*(오전|오후)\s*(\d{1,2}):(\d{1,2})/.exec(normalized);
  if (krMeridiem) {
    const date = parseDateOnly(krMeridiem[1]);
    if (!date) return null;
    let hour = toNumber(krMeridiem[3], 0);
    const minute = toNumber(krMeridiem[4], 0);
    if (krMeridiem[2] === "오후" && hour < 12) hour += 12;
    if (krMeridiem[2] === "오전" && hour === 12) hour = 0;
    date.setHours(hour, minute, 0, 0);
    return date;
  }

  const dateTimeMatch =
    /^(\d{4})-(\d{1,2})-(\d{1,2})(?:\s+|T)(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?$/.exec(normalized);
  if (dateTimeMatch) {
    const y = toNumber(dateTimeMatch[1], 0);
    const m = toNumber(dateTimeMatch[2], 1) - 1;
    const d = toNumber(dateTimeMatch[3], 1);
    const hh = toNumber(dateTimeMatch[4], 0);
    const mm = toNumber(dateTimeMatch[5], 0);
    const ss = toNumber(dateTimeMatch[6], 0);
    return new Date(y, m, d, hh, mm, ss, 0);
  }

  const dateMatch = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(normalized);
  if (dateMatch) {
    return new Date(toNumber(dateMatch[1], 0), toNumber(dateMatch[2], 1) - 1, toNumber(dateMatch[3], 1), 0, 0, 0, 0);
  }

  const native = new Date(normalized);
  if (!Number.isNaN(native.getTime())) return native;
  return null;
}

function parseHoursValue(value) {
  if (typeof value === "number" && Number.isFinite(value)) return round2(value);
  if (typeof value === "string") {
    const trimmed = value.trim();
    if (!trimmed) return null;
    const hhmm = /(\d+)\s*[:시]\s*(\d{1,2})/.exec(trimmed);
    if (hhmm) {
      return round2(toNumber(hhmm[1], 0) + toNumber(hhmm[2], 0) / 60);
    }
    const numeric = trimmed.replace(/[^\d.-]/g, "");
    const number = toNumber(numeric, NaN);
    if (Number.isFinite(number)) return round2(number);
  }
  return null;
}

function excelSerialToDate(serial) {
  const wholeDays = Math.floor(serial - 25569);
  const milliseconds = wholeDays * 86400 * 1000;
  const wholeDate = new Date(milliseconds);
  const dayFraction = serial - Math.floor(serial);
  const totalSeconds = Math.round(dayFraction * 86400);
  const hours = Math.floor(totalSeconds / 3600);
  const minutes = Math.floor((totalSeconds % 3600) / 60);
  const seconds = totalSeconds % 60;

  return new Date(
    wholeDate.getUTCFullYear(),
    wholeDate.getUTCMonth(),
    wholeDate.getUTCDate(),
    hours,
    minutes,
    seconds,
    0
  );
}

function parseDateOnly(dateKey) {
  if (!dateKey) return null;
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(dateKey);
  if (!match) return null;
  return new Date(toNumber(match[1], 0), toNumber(match[2], 1) - 1, toNumber(match[3], 1), 0, 0, 0, 0);
}

function parseDateTimeKey(value) {
  if (!value) return null;
  const match = /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})$/.exec(value);
  if (!match) return null;
  return new Date(
    toNumber(match[1], 0),
    toNumber(match[2], 1) - 1,
    toNumber(match[3], 1),
    toNumber(match[4], 0),
    toNumber(match[5], 0),
    0,
    0
  );
}

function formatDateKey(date) {
  return `${date.getFullYear()}-${pad2(date.getMonth() + 1)}-${pad2(date.getDate())}`;
}

function formatDateTimeKey(date) {
  return `${formatDateKey(date)}T${pad2(date.getHours())}:${pad2(date.getMinutes())}`;
}

function formatDateForDisplay(dateKey) {
  const date = parseDateOnly(dateKey);
  if (!date) return dateKey;
  return `${date.getFullYear()}-${pad2(date.getMonth() + 1)}-${pad2(date.getDate())}`;
}

function formatDateWithDay(dateKey) {
  const date = parseDateOnly(dateKey);
  if (!date) return dateKey;
  return `${date.getMonth() + 1}/${date.getDate()}(${WEEKDAYS[date.getDay()]})`;
}

function formatMonthLabel(month) {
  const match = /^(\d{4})-(\d{2})$/.exec(month);
  if (!match) return month;
  return `${match[1]}년 ${toNumber(match[2], 0)}월`;
}

function formatTimeFromIso(iso) {
  const date = parseDateTimeKey(iso);
  if (!date) return "-";
  return `${pad2(date.getHours())}:${pad2(date.getMinutes())}`;
}

function formatHours(value) {
  const fixed = round2(toNumber(value, 0)).toFixed(2);
  return fixed.replace(/\.00$/, "").replace(/(\.\d)0$/, "$1");
}

function formatDurationText(value) {
  const totalMinutes = Math.max(0, Math.round(toNumber(value, 0) * 60));
  const hours = Math.floor(totalMinutes / 60);
  const minutes = totalMinutes % 60;
  return `${hours}시간 ${minutes}분`;
}

function parseDurationInput(value) {
  const text = String(value ?? "").trim();
  if (!text) return NaN;

  const direct = toNumber(text, NaN);
  if (Number.isFinite(direct)) return round2(direct);

  const hhmmMatch = /^(\d{1,3})\s*:\s*(\d{1,2})$/.exec(text);
  if (hhmmMatch) {
    const hours = toNumber(hhmmMatch[1], 0);
    const minutes = toNumber(hhmmMatch[2], 0);
    return round2(hours + minutes / 60);
  }

  const hourMinuteMatch = /(?:(\d+(?:\.\d+)?)\s*시간)?\s*(?:(\d{1,2})\s*분)?/.exec(text);
  if (hourMinuteMatch) {
    const hasHour = hourMinuteMatch[1] !== undefined;
    const hasMinute = hourMinuteMatch[2] !== undefined;
    if (hasHour || hasMinute) {
      const hours = toNumber(hourMinuteMatch[1], 0);
      const minutes = toNumber(hourMinuteMatch[2], 0);
      return round2(hours + minutes / 60);
    }
  }

  return NaN;
}

function calculateDurationFromTimeInputs(workDate, startText, endText) {
  const day = parseDateOnly(workDate);
  if (!day) return NaN;
  const start = String(startText || "").trim();
  const end = String(endText || "").trim();
  if (!start || !end) return NaN;
  if (!/^\d{2}:\d{2}$/.test(start) || !/^\d{2}:\d{2}$/.test(end)) return NaN;

  let startDt = combineDateAndTime(day, start);
  let endDt = combineDateAndTime(day, end);
  if (endDt <= startDt) {
    endDt = addDays(endDt, 1);
  }
  return round2(diffHours(startDt, endDt));
}

function formatWon(value) {
  return `₩${wonFormatter.format(Math.round(toNumber(value, 0)))}`;
}

function sanitizeHolidayList(list) {
  return Array.from(new Set(list.filter((item) => /^\d{4}-\d{2}-\d{2}$/.test(item)))).sort();
}

function combineDateAndTime(date, timeText) {
  const [hours, minutes] = parseTimeParts(timeText, "09:00");
  return new Date(date.getFullYear(), date.getMonth(), date.getDate(), hours, minutes, 0, 0);
}

function parseTimeParts(value, fallback) {
  const source = value && /^\d{2}:\d{2}$/.test(value) ? value : fallback;
  const [hh, mm] = source.split(":");
  return [toNumber(hh, 0), toNumber(mm, 0)];
}

function diffHours(start, end) {
  return (end.getTime() - start.getTime()) / 3_600_000;
}

function addDays(date, days) {
  const next = new Date(date);
  next.setDate(next.getDate() + days);
  return next;
}

function addHours(date, hours) {
  const next = new Date(date);
  next.setTime(next.getTime() + hours * 3_600_000);
  return next;
}

function startOfDay(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate(), 0, 0, 0, 0);
}

function toNumber(value, fallback) {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : fallback;
}

function round2(value) {
  return Math.round(toNumber(value, 0) * 100) / 100;
}

function roundCurrency(value) {
  return Math.round(toNumber(value, 0));
}

function compareEntries(a, b) {
  const aDate = `${a.workDate} ${a.startIso || "99:99"}`;
  const bDate = `${b.workDate} ${b.startIso || "99:99"}`;
  if (aDate < bDate) return -1;
  if (aDate > bDate) return 1;
  return a.name.localeCompare(b.name, "ko-KR");
}

function cleanText(value) {
  return String(value ?? "")
    .replace(/\s+/g, " ")
    .trim();
}

function normalizeNameList(list) {
  if (!Array.isArray(list)) return [];
  return Array.from(new Set(list.map((item) => cleanText(item)).filter(Boolean))).sort((a, b) =>
    a.localeCompare(b, "ko-KR")
  );
}

function normalizePersistedEntries(list) {
  if (!Array.isArray(list)) return [];
  return list
    .map((entry, index) => normalizePersistedEntry(entry, index))
    .filter(Boolean)
    .sort(compareEntries);
}

function normalizePersistedEntry(raw, index) {
  if (!raw || typeof raw !== "object") return null;
  const workDate = cleanText(raw.workDate);
  const name = cleanText(raw.name);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(workDate) || !name) return null;

  const reportedHoursNum = toNumber(raw.reportedHours, NaN);
  const adjustedHoursNum = toNumber(raw.adjustedHours, NaN);
  const baseHours = Number.isFinite(adjustedHoursNum)
    ? adjustedHoursNum
    : Number.isFinite(reportedHoursNum)
      ? reportedHoursNum
      : DEFAULT_SETTINGS.defaultMissingHours;

  return {
    id: cleanText(raw.id) || `restored-${Date.now()}-${index}`,
    source: raw.source === "manual" ? "manual" : "excel",
    name,
    category: cleanText(raw.category),
    workDate,
    note: cleanText(raw.note),
    reportedHours: Number.isFinite(reportedHoursNum) ? round2(reportedHoursNum) : null,
    adjustedHours: Math.max(0, round2(baseHours)),
    isAdjustedManually: Boolean(raw.isAdjustedManually),
    missingStart: Boolean(raw.missingStart),
    missingEnd: Boolean(raw.missingEnd),
    missingOriginDetected:
      typeof raw.missingOriginDetected === "boolean"
        ? raw.missingOriginDetected
        : raw.source === "excel" && (Boolean(raw.missingStart) || Boolean(raw.missingEnd)),
    missingFixApplied: Boolean(raw.missingFixApplied),
    originalStartIso: normalizePersistedIso(raw.originalStartIso),
    originalEndIso: normalizePersistedIso(raw.originalEndIso),
    startIso: normalizePersistedIso(raw.startIso),
    endIso: normalizePersistedIso(raw.endIso),
  };
}

function normalizePersistedIso(value) {
  const text = cleanText(value);
  return /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}$/.test(text) ? text : "";
}

function normalizeTimeText(value, fallback = "") {
  const text = cleanText(value);
  if (/^\d{2}:\d{2}$/.test(text)) return text;
  if (/^\d{1}:\d{2}$/.test(text)) {
    const [h, m] = text.split(":");
    return `${pad2(toNumber(h, 0))}:${pad2(toNumber(m, 0))}`;
  }
  return fallback && /^\d{2}:\d{2}$/.test(fallback) ? fallback : "";
}

function cloneDefaultShiftTemplates() {
  return DEFAULT_SHIFT_TEMPLATES.map((item) => ({ ...item }));
}

function cloneDefaultCareerLevels() {
  return DEFAULT_CAREER_LEVELS.map((item) => ({ ...item }));
}

function normalizeScheduleTemplates(list) {
  const source = Array.isArray(list) ? list : [];
  const byId = new Map(source.map((item) => [cleanText(item?.id), item]));
  return DEFAULT_SHIFT_TEMPLATES.map((defaults) => {
    const raw = byId.get(defaults.id) || {};
    return {
      id: defaults.id,
      label: cleanText(raw.label || defaults.label),
      start: normalizeTimeText(raw.start, defaults.start),
      end: normalizeTimeText(raw.end, defaults.end),
    };
  });
}

function normalizeScheduleFixedRules(list) {
  if (!Array.isArray(list)) return [];
  return list
    .map((raw, index) => {
      const mode = raw?.mode === "date" ? "date" : "weekly";
      const shiftId = SHIFT_IDS.includes(cleanText(raw?.shiftId)) ? cleanText(raw.shiftId) : "1T";
      const weekday = Math.max(0, Math.min(6, Math.round(toNumber(raw?.weekday, 1))));
      const date = cleanText(raw?.date);
      const manager = stripManagerTitle(cleanText(raw?.manager));
      if (!manager) return null;
      if (mode === "date" && !/^\d{4}-\d{2}-\d{2}$/.test(date)) return null;
      return {
        id: cleanText(raw?.id) || `rule-restored-${index}`,
        mode,
        shiftId,
        weekday: mode === "weekly" ? weekday : null,
        date: mode === "date" ? date : "",
        manager,
        note: cleanText(raw?.note),
      };
    })
    .filter(Boolean);
}

function normalizeScheduleAssignments(value) {
  if (!value || typeof value !== "object") return {};
  const next = {};
  Object.entries(value).forEach(([month, rows]) => {
    if (!/^\d{4}-\d{2}$/.test(cleanText(month))) return;
    if (!Array.isArray(rows)) {
      next[month] = [];
      return;
    }
    next[month] = rows
      .map((row, index) => normalizeScheduleRow({ ...row, id: row?.id || `sch-restored-${month}-${index}` }))
      .filter(Boolean)
      .sort((a, b) => {
        if (a.date < b.date) return -1;
        if (a.date > b.date) return 1;
        return a.shiftId.localeCompare(b.shiftId);
      });
  });
  return next;
}

function normalizeScheduleRow(raw) {
  if (!raw || typeof raw !== "object") return null;
  const date = cleanText(raw.date);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return null;
  const shiftId = SHIFT_IDS.includes(cleanText(raw.shiftId)) ? cleanText(raw.shiftId) : "1T";
  const template = DEFAULT_SHIFT_TEMPLATES.find((item) => item.id === shiftId) || DEFAULT_SHIFT_TEMPLATES[0];
  const weekday = parseDateOnly(date)?.getDay();
  const inferredWorkerSlots = cleanText(raw.secondaryManager) ? 2 : 1;
  return {
    id: cleanText(raw.id) || `sch-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
    date,
    weekday: Number.isFinite(toNumber(raw.weekday, NaN)) ? toNumber(raw.weekday, 0) : weekday,
    shiftId,
    workerSlots: Math.min(2, Math.max(1, Math.round(toNumber(raw.workerSlots, inferredWorkerSlots)))),
    shiftStart: normalizeTimeText(raw.shiftStart, template.start),
    shiftEnd: normalizeTimeText(raw.shiftEnd, template.end),
    primaryManager: stripManagerTitle(cleanText(raw.primaryManager)),
    primaryStart: normalizeTimeText(raw.primaryStart, ""),
    primaryEnd: normalizeTimeText(raw.primaryEnd, ""),
    secondaryManager: stripManagerTitle(cleanText(raw.secondaryManager)),
    secondaryStart: normalizeTimeText(raw.secondaryStart, ""),
    secondaryEnd: normalizeTimeText(raw.secondaryEnd, ""),
    note: cleanText(raw.note),
    source: cleanText(raw.source || "수동"),
  };
}

function normalizeManagerProfiles(value) {
  if (!value || typeof value !== "object") return {};
  const next = {};
  Object.entries(value).forEach(([name, raw]) => {
    const key = stripManagerTitle(cleanText(name));
    if (!key) return;
    const base = createDefaultManagerProfile();
    const profile = {
      desiredShifts: Math.max(0, toNumber(raw?.desiredShifts, base.desiredShifts)),
      priority: Math.max(1, Math.round(toNumber(raw?.priority, base.priority))),
      pointPerShift: Math.min(3, Math.max(1, round2(toNumber(raw?.pointPerShift, base.pointPerShift)))),
      basePoints: Math.max(0, round2(toNumber(raw?.basePoints, 0))),
      baseHours: Math.max(0, round2(toNumber(raw?.baseHours, 0))),
      availability: base.availability,
    };
    if (raw?.availability && typeof raw.availability === "object") {
      for (let day = 0; day < 7; day += 1) {
        if (!profile.availability[day]) profile.availability[day] = {};
        SHIFT_IDS.forEach((shiftId) => {
          const valueAtDay = raw.availability?.[day]?.[shiftId];
          profile.availability[day][shiftId] =
            typeof valueAtDay === "boolean" ? valueAtDay : profile.availability[day][shiftId];
        });
      }
    }
    next[key] = profile;
  });
  return next;
}

function normalizeCareerLevels(list) {
  const source = Array.isArray(list) ? list : [];
  const mapped = source
    .map((raw, index) => {
      const id = cleanText(raw?.id) || `career-${index}`;
      const name = cleanText(raw?.name) || `Level ${index + 1}`;
      const min = Math.max(0, round2(toNumber(raw?.min, 0)));
      const maxRaw = cleanText(raw?.max);
      const hasFiniteMax = maxRaw !== "" && Number.isFinite(toNumber(raw?.max, NaN));
      const max = hasFiniteMax ? Math.max(min, round2(toNumber(raw?.max, min))) : null;
      const wage = Math.max(0, roundCurrency(toNumber(raw?.wage, 0)));
      return { id, name, min, max, wage };
    })
    .filter((item) => item.wage > 0);

  const usable = mapped.length ? mapped : cloneDefaultCareerLevels();
  usable.sort((a, b) => {
    const minDiff = toNumber(a.min, 0) - toNumber(b.min, 0);
    if (minDiff !== 0) return minDiff;
    return a.name.localeCompare(b.name, "ko-KR");
  });
  return usable;
}

function pad2(number) {
  return String(number).padStart(2, "0");
}

function setStatus(message, isError = false) {
  dom.uploadStatus.textContent = message;
  dom.uploadStatus.classList.toggle("text-danger", isError);
}

function buildPersistPayload() {
  return {
    settings: state.settings,
    holidays: state.holidays,
    employeeSettings: state.employeeSettings,
    entries: state.entries,
    selectedMonth: state.selectedMonth,
    selectedEmployee: state.selectedEmployee,
    scheduleMonth: state.scheduleMonth,
    scheduleTemplates: state.scheduleTemplates,
    scheduleFixedRules: state.scheduleFixedRules,
    scheduleAssignments: state.scheduleAssignments,
    scheduleExcludedManagers: state.scheduleExcludedManagers,
    careerLevels: state.careerLevels,
    careerMonth: state.careerMonth,
    managerProfiles: state.managerProfiles,
    selectedScheduleSlotId: state.selectedScheduleSlotId,
  };
}

function buildOutputsSnapshot() {
  const months = getAvailableMonths();
  const monthSnapshots = months.map((month) => {
    const employees = getEmployeesInMonth(month);
    const employeeSnapshots = employees.map((name) => {
      const payroll = computeEmployeeMonthPayroll(name, month);
      return {
        name,
        totals: payroll.totals,
        weeks: payroll.weeks.map((week) => ({
          weekIndex: week.weekIndex,
          hours: week.hours,
          workDays: week.workDays,
          basePay: week.basePay,
          overtimePremium: week.overtimePremium,
          nightPremium: week.nightPremium,
          holidayPremium: week.holidayPremium,
          weeklyHolidayPay: week.weeklyHolidayPay,
          totalPay: week.totalPay,
        })),
        shifts: payroll.shifts.map((shift) => ({
          workDate: shift.entry.workDate,
          startIso: shift.entry.startIso,
          endIso: shift.entry.endIso,
          adjustedHours: shift.entry.adjustedHours,
          basePay: shift.basePay,
          overtimePremium: shift.overtimePremium,
          nightPremium: shift.nightPremium,
          holidayPremium: shift.holidayPremium,
          shiftPay: shift.shiftPay,
        })),
      };
    });

    const totals = employeeSnapshots.reduce(
      (acc, item) => {
        acc.hours += item.totals.hours;
        acc.basePay += item.totals.basePay;
        acc.allowances += item.totals.allowances;
        acc.gross += item.totals.grossPay;
        acc.net += item.totals.netPay;
        return acc;
      },
      { hours: 0, basePay: 0, allowances: 0, gross: 0, net: 0 }
    );

    return {
      month,
      monthLabel: formatMonthLabel(month),
      employeeCount: employees.length,
      totals: {
        hours: round2(totals.hours),
        basePay: roundCurrency(totals.basePay),
        allowances: roundCurrency(totals.allowances),
        grossPay: roundCurrency(totals.gross),
        netPay: roundCurrency(totals.net),
      },
      employees: employeeSnapshots,
    };
  });

  const scheduleMonths = Object.keys(state.scheduleAssignments || {})
    .filter((month) => /^\d{4}-\d{2}$/.test(month))
    .sort();
  const scheduleSnapshots = scheduleMonths.map((month) => {
    const rows = getScheduleRowsForMonth(month);
    const summary = Array.from(getScheduleMonthContribution(month).entries()).map(([name, info]) => ({
      name,
      shiftUnits: round2(info.shiftUnits),
      hours: round2(info.hours),
    }));
    return {
      month,
      monthLabel: formatMonthLabel(month),
      rows,
      summary,
    };
  });

  const careerReferenceMonth = state.careerMonth || state.scheduleMonth || state.selectedMonth;
  const careerSnapshot = getManagerNames().map((name) => computeManagerCareerSnapshot(name, careerReferenceMonth));

  return {
    generatedAt: new Date().toISOString(),
    monthSnapshots,
    scheduleSnapshots,
    career: {
      referenceMonth: careerReferenceMonth,
      levels: state.careerLevels,
      managerSnapshots: careerSnapshot,
    },
  };
}

function exportAllDataFile() {
  const payload = {
    app: EXPORT_APP_ID,
    version: EXPORT_VERSION,
    exportedAt: new Date().toISOString(),
    data: buildPersistPayload(),
    outputs: buildOutputsSnapshot(),
  };

  const json = JSON.stringify(payload, null, 2);
  const blob = new Blob([json], { type: "application/json;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  const stamp = buildExportTimeStamp(new Date());
  link.href = url;
  link.download = `payroll_backup_${stamp}.json`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
  setStatus("전체 데이터(입력/결과 포함) 저장 파일을 만들었습니다.");
}

async function handleLoadStateFile(event) {
  const input = event.target;
  if (!(input instanceof HTMLInputElement) || !input.files?.length) return;
  const file = input.files[0];
  try {
    const text = await file.text();
    const parsed = JSON.parse(text);
    const data = parsed && typeof parsed === "object" && parsed.data ? parsed.data : parsed;
    if (!data || typeof data !== "object") {
      throw new Error("저장 파일 형식이 올바르지 않습니다.");
    }
    restoreFromImportedData(data);
    renderEverything();
    setStatus(`"${file.name}" 파일을 불러왔습니다. 저장된 월별 데이터가 복원되었습니다.`);
  } catch (error) {
    const message =
      error instanceof Error ? error.message : "저장 파일을 불러오는 중 오류가 발생했습니다.";
    setStatus(message, true);
  } finally {
    input.value = "";
  }
}

function restoreFromImportedData(data) {
  state.settings = {
    ...DEFAULT_SETTINGS,
    ...((data.settings && typeof data.settings === "object") ? data.settings : {}),
  };
  state.holidays = sanitizeHolidayList(Array.isArray(data.holidays) ? data.holidays : []);
  state.employeeSettings =
    data.employeeSettings && typeof data.employeeSettings === "object" ? data.employeeSettings : {};
  state.entries = normalizePersistedEntries(Array.isArray(data.entries) ? data.entries : []);
  state.selectedMonth = cleanText(data.selectedMonth || "");
  state.selectedEmployee = cleanText(data.selectedEmployee || "");
  state.scheduleMonth = cleanText(data.scheduleMonth || "");
  state.scheduleTemplates = normalizeScheduleTemplates(data.scheduleTemplates || []);
  state.scheduleFixedRules = normalizeScheduleFixedRules(data.scheduleFixedRules || []);
  state.scheduleAssignments = normalizeScheduleAssignments(data.scheduleAssignments || {});
  state.scheduleExcludedManagers = normalizeNameList(data.scheduleExcludedManagers || []);
  state.careerLevels = normalizeCareerLevels(data.careerLevels || []);
  state.careerMonth = cleanText(data.careerMonth || "");
  state.managerProfiles = normalizeManagerProfiles(data.managerProfiles || {});
  state.selectedScheduleSlotId = cleanText(data.selectedScheduleSlotId || "");
  applyDefaultWageMigration();
  initializeAdvancedDefaults();
  syncEmployeeSettings();
  syncCareerRatesToEmployeeSettings();
  hydrateSettingInputs();
}

function buildExportTimeStamp(date) {
  return `${date.getFullYear()}${pad2(date.getMonth() + 1)}${pad2(date.getDate())}_${pad2(
    date.getHours()
  )}${pad2(date.getMinutes())}${pad2(date.getSeconds())}`;
}

function applyDefaultWageMigration() {
  let changed = false;
  if (toNumber(state.settings.defaultHourlyWage, NaN) === 11000) {
    state.settings.defaultHourlyWage = 10320;
    changed = true;
  }
  Object.values(state.employeeSettings).forEach((config) => {
    if (toNumber(config.hourlyRate, NaN) === 11000) {
      config.hourlyRate = 10320;
      changed = true;
    }
  });
  state.careerLevels.forEach((level) => {
    if (toNumber(level.wage, NaN) === 11000) {
      level.wage = 10320;
      changed = true;
    }
  });
  if (changed) {
    persistConfig();
  }
}

function persistConfig() {
  const payload = buildPersistPayload();
  localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
}

function loadPersistedConfig() {
  const raw = localStorage.getItem(STORAGE_KEY);
  if (!raw) return {};
  try {
    return JSON.parse(raw);
  } catch {
    return {};
  }
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&#39;");
}
