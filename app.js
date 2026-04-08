const STORAGE_KEY = "payroll-app-config-v1";
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
  defaultHourlyWage: 11000,
  autoFillMissing: true,
  defaultMissingHours: 8,
  defaultStartTime: "09:00",
  defaultNightStart: "22:30",
  nightEndTime: "06:00",
  weeklyThresholdHours: 15,
  weeklyMultiplier: 1,
};

const persisted = loadPersistedConfig();

const state = {
  settings: { ...DEFAULT_SETTINGS, ...(persisted.settings || {}) },
  holidays: sanitizeHolidayList(persisted.holidays || []),
  employeeSettings: persisted.employeeSettings || {},
  entries: [],
  selectedMonth: "",
  selectedEmployee: "",
};

const wonFormatter = new Intl.NumberFormat("ko-KR");

const dom = {
  tabs: Array.from(document.querySelectorAll(".tab-btn")),
  tabPanels: Array.from(document.querySelectorAll(".tab-panel")),
  fileInput: document.getElementById("fileInput"),
  monthSelect: document.getElementById("monthSelect"),
  recalculateBtn: document.getElementById("recalculateBtn"),
  uploadStatus: document.getElementById("uploadStatus"),

  useWeeklyHolidayInput: document.getElementById("useWeeklyHolidayInput"),
  useHolidayPremiumInput: document.getElementById("useHolidayPremiumInput"),
  useOvertimePremiumInput: document.getElementById("useOvertimePremiumInput"),
  useNightPremiumInput: document.getElementById("useNightPremiumInput"),
  taxRateInput: document.getElementById("taxRateInput"),
  defaultHourlyWageInput: document.getElementById("defaultHourlyWageInput"),
  autoFillMissingInput: document.getElementById("autoFillMissingInput"),
  defaultMissingHoursInput: document.getElementById("defaultMissingHoursInput"),
  defaultStartTimeInput: document.getElementById("defaultStartTimeInput"),
  defaultNightStartInput: document.getElementById("defaultNightStartInput"),
  nightEndTimeInput: document.getElementById("nightEndTimeInput"),
  weeklyThresholdInput: document.getElementById("weeklyThresholdInput"),
  weeklyMultiplierInput: document.getElementById("weeklyMultiplierInput"),

  holidayDateInput: document.getElementById("holidayDateInput"),
  addHolidayBtn: document.getElementById("addHolidayBtn"),
  holidayList: document.getElementById("holidayList"),

  employeeSettingsBody: document.getElementById("employeeSettingsBody"),

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
  summaryTotalHours: document.getElementById("summaryTotalHours"),
  summaryTotalGross: document.getElementById("summaryTotalGross"),
  summaryTotalNet: document.getElementById("summaryTotalNet"),
};

init();

function init() {
  hydrateSettingInputs();
  bindEvents();
  renderHolidayList();
  renderMonthOptions();
  renderEmployeeSelectors();
  renderEmployeeSettingsTable();
  renderPreview();
  renderStatement();
  renderSummary();
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
    renderStatement();
    renderSummary();
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

  bindSettingListeners();

  dom.addHolidayBtn.addEventListener("click", addHoliday);
  dom.holidayList.addEventListener("click", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLButtonElement)) return;
    if (!target.dataset.date) return;
    removeHoliday(target.dataset.date);
  });

  dom.employeeSettingsBody.addEventListener("input", handleEmployeeSettingChange);
  dom.addMissingShiftBtn.addEventListener("click", addMissingShift);

  dom.statementEmployeeSelect.addEventListener("change", () => {
    state.selectedEmployee = dom.statementEmployeeSelect.value;
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
      const value = toNumber(input.value, NaN);
      if (!Number.isFinite(value) || value < 0) return;
      const entry = state.entries.find((item) => item.id === entryId);
      if (!entry) return;
      entry.adjustedHours = round2(value);
      entry.isAdjustedManually = true;
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
  const checkboxBindings = [
    ["useWeeklyHolidayInput", "useWeeklyHoliday"],
    ["useHolidayPremiumInput", "useHolidayPremium"],
    ["useOvertimePremiumInput", "useOvertimePremium"],
    ["useNightPremiumInput", "useNightPremium"],
    ["autoFillMissingInput", "autoFillMissing"],
  ];
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
    ["weeklyMultiplierInput", "weeklyMultiplier"],
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
  dom.useWeeklyHolidayInput.checked = Boolean(state.settings.useWeeklyHoliday);
  dom.useHolidayPremiumInput.checked = Boolean(state.settings.useHolidayPremium);
  dom.useOvertimePremiumInput.checked = Boolean(state.settings.useOvertimePremium);
  dom.useNightPremiumInput.checked = Boolean(state.settings.useNightPremium);
  dom.taxRateInput.value = String(state.settings.taxRate);
  dom.defaultHourlyWageInput.value = String(state.settings.defaultHourlyWage);
  dom.autoFillMissingInput.checked = Boolean(state.settings.autoFillMissing);
  dom.defaultMissingHoursInput.value = String(state.settings.defaultMissingHours);
  dom.defaultStartTimeInput.value = state.settings.defaultStartTime;
  dom.defaultNightStartInput.value = state.settings.defaultNightStart;
  dom.nightEndTimeInput.value = state.settings.nightEndTime;
  dom.weeklyThresholdInput.value = String(state.settings.weeklyThresholdHours);
  dom.weeklyMultiplierInput.value = String(state.settings.weeklyMultiplier);
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
    state.entries = parsedEntries;
    state.selectedMonth = getAvailableMonths().slice(-1)[0] || "";
    syncEmployeeSettings();
    state.selectedEmployee = getEmployeesInMonth(state.selectedMonth)[0] || "";
    renderEverything();
    setStatus(
      `${parsedEntries.length}건을 불러왔습니다. 정산시간은 근무자별 화면에서 직접 수정 가능합니다.`
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
  const name = customName || selectedExisting;
  const workDate = dom.missingDate.value;
  if (!name || !workDate) {
    setStatus("누락 근무 추가 시 근무자 이름과 근무일은 필수입니다.", true);
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
  persistConfig();
  renderStatement();
  renderSummary();
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

  renderMonthOptions();
  renderEmployeeSelectors();
  renderEmployeeSettingsTable();
  renderPreview();
  renderStatement();
  renderSummary();
  persistConfig();
}

function renderMonthOptions() {
  const months = getAvailableMonths();
  if (!months.length) {
    dom.monthSelect.innerHTML = '<option value="">월 선택</option>';
    dom.monthSelect.value = "";
    return;
  }
  dom.monthSelect.innerHTML = months
    .map((month) => `<option value="${month}">${formatMonthLabel(month)}</option>`)
    .join("");
  dom.monthSelect.value = months.includes(state.selectedMonth) ? state.selectedMonth : months[months.length - 1];
  state.selectedMonth = dom.monthSelect.value;
}

function renderEmployeeSelectors() {
  const employees = getEmployeesInMonth(state.selectedMonth);
  const placeholder = '<option value="">근무자 선택</option>';
  const optionHtml = employees
    .map((name) => `<option value="${escapeHtml(name)}">${escapeHtml(name)}</option>`)
    .join("");

  dom.statementEmployeeSelect.innerHTML = placeholder + optionHtml;
  dom.missingEmployeeSelect.innerHTML = placeholder + optionHtml;

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
      '<tr><td colspan="4" class="empty">업로드 후 자동 생성됩니다.</td></tr>';
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
      </tr>`;
    })
    .join("");
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
        <td>${formatHours(entry.adjustedHours)}시간</td>
        <td>${escapeHtml(entry.note || "-")}</td>
      </tr>`;
    })
    .join("");
}

function renderStatement() {
  if (!state.selectedMonth || !state.selectedEmployee) {
    dom.statementMeta.textContent = "근무자와 정산 월을 선택하세요.";
    dom.statementBody.innerHTML =
      '<tr><td colspan="13" class="empty">표시할 근무 내역이 없습니다.</td></tr>';
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
      '<tr><td colspan="13" class="empty">선택한 월의 근무기록이 없습니다.</td></tr>';
    dom.weeklySummaryBody.innerHTML = '<tr><td colspan="6" class="empty">데이터가 없습니다.</td></tr>';
    renderTotalsEmpty();
    return;
  }

  dom.statementBody.innerHTML = payroll.shifts
    .map((shift) => {
      const notes = [];
      if (shift.entry.missingStart || shift.entry.missingEnd) {
        notes.push('<span class="badge warn">누락 보정</span>');
      }
      if (shift.isHolidayDate) {
        notes.push('<span class="badge">공휴일</span>');
      }
      if (shift.entry.note) {
        notes.push(escapeHtml(shift.entry.note));
      }
      return `<tr>
        <td>${shift.weekIndex}주차</td>
        <td>${formatDateWithDay(shift.entry.workDate)}</td>
        <td>${shift.entry.startIso ? formatTimeFromIso(shift.entry.startIso) : "-"}</td>
        <td>${shift.entry.endIso ? formatTimeFromIso(shift.entry.endIso) : "-"}</td>
        <td>${formatHours(shift.clockHours)}시간</td>
        <td>
          <input class="hour-editor" data-id="${shift.entry.id}" type="number" min="0" step="0.1" value="${formatHourInput(
            shift.entry.adjustedHours
          )}" />
        </td>
        <td>${formatWon(shift.basePay)}</td>
        <td>${formatWon(shift.overtimePremium)}</td>
        <td>${formatWon(shift.nightPremium)}</td>
        <td>${formatWon(shift.holidayPremium)}</td>
        <td>${formatWon(shift.shiftPay)}</td>
        <td>${notes.join(" ") || "-"}</td>
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
          <td>${formatHours(week.hours)}시간</td>
          <td>${week.workDays}</td>
          <td>${formatWon(week.basePay)}</td>
          <td>${formatWon(premiumSum)}</td>
          <td>${formatWon(week.totalPay)}</td>
        </tr>`;
      })
      .join("");
  }

  dom.totalHoursCell.textContent = `${formatHours(payroll.totals.hours)}시간`;
  dom.basePayCell.textContent = formatWon(payroll.totals.basePay);
  dom.allowanceCell.textContent = formatWon(payroll.totals.allowances);
  dom.grossPayCell.textContent = formatWon(payroll.totals.grossPay);
  dom.netPayCell.textContent = formatWon(payroll.totals.netPay);
}

function renderTotalsEmpty() {
  dom.totalHoursCell.textContent = "0.00시간";
  dom.basePayCell.textContent = "₩0";
  dom.allowanceCell.textContent = "₩0";
  dom.grossPayCell.textContent = "₩0";
  dom.netPayCell.textContent = "₩0";
}

function renderSummary() {
  if (!state.selectedMonth) {
    dom.summaryBody.innerHTML = '<tr><td colspan="4" class="empty">정산할 데이터가 없습니다.</td></tr>';
    dom.summaryTotalHours.textContent = "0.00시간";
    dom.summaryTotalGross.textContent = "₩0";
    dom.summaryTotalNet.textContent = "₩0";
    return;
  }

  const employees = getEmployeesInMonth(state.selectedMonth);
  if (!employees.length) {
    dom.summaryBody.innerHTML = '<tr><td colspan="4" class="empty">정산할 데이터가 없습니다.</td></tr>';
    dom.summaryTotalHours.textContent = "0.00시간";
    dom.summaryTotalGross.textContent = "₩0";
    dom.summaryTotalNet.textContent = "₩0";
    return;
  }

  const rows = employees
    .map((name) => ({
      name,
      payroll: computeEmployeeMonthPayroll(name, state.selectedMonth),
    }))
    .sort((a, b) => b.payroll.totals.grossPay - a.payroll.totals.grossPay);

  dom.summaryBody.innerHTML = rows
    .map(
      ({ name, payroll }) => `<tr>
      <td>${escapeHtml(name)}</td>
      <td>${formatHours(payroll.totals.hours)}시간</td>
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

  dom.summaryTotalHours.textContent = `${formatHours(totals.hours)}시간`;
  dom.summaryTotalGross.textContent = formatWon(totals.gross);
  dom.summaryTotalNet.textContent = formatWon(totals.net);
}

function computeEmployeeMonthPayroll(name, month) {
  const config = ensureEmployeeSetting(name);
  const hourlyRate = Math.max(0, toNumber(config.hourlyRate, state.settings.defaultHourlyWage));

  const shifts = state.entries
    .filter((entry) => entry.name === name && entry.workDate.startsWith(month))
    .map((entry) => computeShiftPayroll(entry, hourlyRate, config.nightStart))
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

  if (state.settings.useWeeklyHoliday) {
    weekMap.forEach((week) => {
      const workDays = week.daySet.size;
      const threshold = Math.max(0, toNumber(state.settings.weeklyThresholdHours, 15));
      if (week.hours >= threshold && workDays > 0) {
        const multiplier = Math.max(0, toNumber(state.settings.weeklyMultiplier, 1));
        const weeklyHolidayHours = (week.hours / workDays) * multiplier;
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

function computeShiftPayroll(entry, hourlyRate, nightStart) {
  const adjustedHours = Math.max(0, toNumber(entry.adjustedHours, 0));
  const start = parseDateTimeKey(entry.startIso);
  const end = parseDateTimeKey(entry.endIso);
  const clockHours = start && end ? Math.max(0, diffHours(start, end)) : 0;
  const ratio = clockHours > 0 ? adjustedHours / clockHours : 1;

  const overtimeHours = state.settings.useOvertimePremium ? Math.max(0, adjustedHours - 8) : 0;
  const rawNightHours =
    state.settings.useNightPremium && start && end
      ? calculateNightHours(start, end, nightStart || state.settings.defaultNightStart, state.settings.nightEndTime)
      : 0;
  const nightHours = Math.min(adjustedHours, rawNightHours * ratio);
  const isHolidayDate = isOfficialHolidayDate(entry.workDate);
  const holidayHours = state.settings.useHolidayPremium && isHolidayDate ? adjustedHours : 0;

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
  return Array.from(new Set(state.entries.map((entry) => entry.name))).sort((a, b) =>
    a.localeCompare(b, "ko-KR")
  );
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
    };
  }
  if (!state.employeeSettings[name].nightStart) {
    state.employeeSettings[name].nightStart = state.settings.defaultNightStart;
  }
  if (!Number.isFinite(toNumber(state.employeeSettings[name].hourlyRate, NaN))) {
    state.employeeSettings[name].hourlyRate = state.settings.defaultHourlyWage;
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

function formatHourInput(value) {
  return round2(toNumber(value, 0)).toFixed(2);
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

function pad2(number) {
  return String(number).padStart(2, "0");
}

function setStatus(message, isError = false) {
  dom.uploadStatus.textContent = message;
  dom.uploadStatus.classList.toggle("text-danger", isError);
}

function persistConfig() {
  const payload = {
    settings: state.settings,
    holidays: state.holidays,
    employeeSettings: state.employeeSettings,
  };
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
