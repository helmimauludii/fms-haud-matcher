const state = {
  files: {
    fms: null,
    haud: null,
  },
  rows: {
    fms: [],
    haud: [],
  },
  workbook: {
    fms: null,
    haud: null,
  },
  output: [],
  outputBlob: null,
  outputName: "",
};

const els = {
  form: document.querySelector("#matchingForm"),
  fmsFile: document.querySelector("#fmsFile"),
  haudFile: document.querySelector("#haudFile"),
  fmsFileName: document.querySelector("#fmsFileName"),
  haudFileName: document.querySelector("#haudFileName"),
  fmsSheet: document.querySelector("#fmsSheet"),
  haudSheet: document.querySelector("#haudSheet"),
  fmsNumberColumn: document.querySelector("#fmsNumberColumn"),
  fmsDateColumn: document.querySelector("#fmsDateColumn"),
  haudSourceColumn: document.querySelector("#haudSourceColumn"),
  haudNumberColumn: document.querySelector("#haudNumberColumn"),
  haudDateColumn: document.querySelector("#haudDateColumn"),
  processStartDate: document.querySelector("#processStartDate"),
  processEndDate: document.querySelector("#processEndDate"),
  dateFormat: document.querySelector("#dateFormat"),
  outputFormat: document.querySelector("#outputFormat"),
  ignoreCase: document.querySelector("#ignoreCase"),
  stripSeparators: document.querySelector("#stripSeparators"),
  removeLeadingZero: document.querySelector("#removeLeadingZero"),
  includeUnmatched: document.querySelector("#includeUnmatched"),
  previewButton: document.querySelector("#previewButton"),
  resetButton: document.querySelector("#resetButton"),
  downloadButton: document.querySelector("#downloadButton"),
  processBox: document.querySelector("#processBox"),
  processState: document.querySelector("#processState"),
  processTitle: document.querySelector("#processTitle"),
  processPercent: document.querySelector("#processPercent"),
  processBar: document.querySelector("#processBar"),
  processMessage: document.querySelector("#processMessage"),
  processLog: document.querySelector("#processLog"),
  previewBox: document.querySelector("#previewBox"),
  readinessMeter: document.querySelector("#readinessMeter"),
  checklist: document.querySelector("#checklist"),
  fileMetric: document.querySelector("#fileMetric"),
  dateMetric: document.querySelector("#dateMetric"),
  matchMetric: document.querySelector("#matchMetric"),
  outputMetric: document.querySelector("#outputMetric"),
  matchedCount: document.querySelector("#matchedCount"),
  fmsOnlyCount: document.querySelector("#fmsOnlyCount"),
  haudOnlyCount: document.querySelector("#haudOnlyCount"),
  totalCount: document.querySelector("#totalCount"),
  resultBody: document.querySelector("#resultBody"),
};

const columnHints = {
  fms: {
    number: ["bnumberprocessed", "bnumber", "numberprocessed", "processed", "number", "nomor"],
    date: ["date", "tanggal", "tgl"],
  },
  haud: {
    source: ["sourceaddr", "sourceaddress", "source"],
    number: ["destinationaddr", "destinationaddress", "destination", "destaddr", "dest", "number", "nomor"],
    date: ["date", "tanggal", "tgl"],
  },
};

els.fmsFile.addEventListener("change", () => handleFile("fms", els.fmsFile.files[0]));
els.haudFile.addEventListener("change", () => handleFile("haud", els.haudFile.files[0]));
els.fmsSheet.addEventListener("change", () => loadSheet("fms", els.fmsSheet.value));
els.haudSheet.addEventListener("change", () => loadSheet("haud", els.haudSheet.value));
els.processStartDate.addEventListener("change", syncEndDate);
els.processStartDate.addEventListener("change", updateReadiness);
els.processEndDate.addEventListener("change", updateReadiness);
els.previewButton.addEventListener("click", showPreview);
els.resetButton.addEventListener("click", resetDashboard);
els.downloadButton.addEventListener("click", downloadOutput);
els.form.addEventListener("submit", runMatching);

[
  els.fmsNumberColumn,
  els.fmsDateColumn,
  els.haudSourceColumn,
  els.haudNumberColumn,
  els.haudDateColumn,
  els.dateFormat,
  els.ignoreCase,
  els.stripSeparators,
  els.removeLeadingZero,
  els.includeUnmatched,
].forEach((control) => control.addEventListener("change", updateReadiness));

async function handleFile(type, file) {
  if (!file) return;

  state.files[type] = file;
  document.querySelector(`[data-zone="${type}"]`).classList.add("ready");
  document.querySelector(`#${type}FileName`).textContent = file.name;

  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array", cellDates: false });
  state.workbook[type] = workbook;

  const sheetSelect = type === "fms" ? els.fmsSheet : els.haudSheet;
  sheetSelect.disabled = false;
  sheetSelect.innerHTML = workbook.SheetNames.map((name) => `<option value="${escapeHtml(name)}">${escapeHtml(name)}</option>`).join("");

  loadSheet(type, workbook.SheetNames[0]);
  updateReadiness();
}

function loadSheet(type, sheetName) {
  const workbook = state.workbook[type];
  if (!workbook || !sheetName) return;

  const worksheet = workbook.Sheets[sheetName];
  state.rows[type] = XLSX.utils.sheet_to_json(worksheet, { defval: "", raw: true });

  const columns = collectColumns(state.rows[type]);
  if (type === "haud") populateColumnSelect(type, "source", columns);
  populateColumnSelect(type, "number", columns);
  populateColumnSelect(type, "date", columns);
  updateReadiness();
}

function collectColumns(rows) {
  const columns = new Set();
  rows.slice(0, 25).forEach((row) => {
    Object.keys(row).forEach((key) => columns.add(key));
  });
  return Array.from(columns);
}

function populateColumnSelect(type, field, columns) {
  const select = getColumnSelect(type, field);
  select.disabled = columns.length === 0;

  if (!columns.length) {
    select.innerHTML = `<option value="">Kolom tidak terdeteksi</option>`;
    return;
  }

  const selected = guessColumn(type, field, columns);
  select.innerHTML = columns
    .map((column) => {
      const isSelected = column === selected ? "selected" : "";
      return `<option value="${escapeHtml(column)}" ${isSelected}>${escapeHtml(column)}</option>`;
    })
    .join("");
}

function getColumnSelect(type, field) {
  if (type === "fms" && field === "number") return els.fmsNumberColumn;
  if (type === "fms" && field === "date") return els.fmsDateColumn;
  if (type === "haud" && field === "source") return els.haudSourceColumn;
  if (type === "haud" && field === "number") return els.haudNumberColumn;
  return els.haudDateColumn;
}

function guessColumn(type, field, columns) {
  const hints = columnHints[type][field];
  return (
    columns.find((column) => {
      const normalized = normalizeColumnName(column);
      return hints.some((hint) => normalized.includes(hint));
    }) || columns[0]
  );
}

function normalizeColumnName(column) {
  return String(column).toLowerCase().replace(/[^a-z0-9]/g, "");
}

function updateReadiness() {
  const hasValidDateRange = Boolean(els.processStartDate.value && els.processEndDate.value && readDateInput(els.processStartDate.value) <= readDateInput(els.processEndDate.value));
  const checks = {
    fms: Boolean(state.files.fms),
    haud: Boolean(state.files.haud),
    date: hasValidDateRange,
    mapping: Boolean(
      els.fmsNumberColumn.value &&
        els.fmsDateColumn.value &&
        els.haudSourceColumn.value &&
        els.haudNumberColumn.value &&
        els.haudDateColumn.value,
    ),
  };

  Object.entries(checks).forEach(([key, isDone]) => {
    const item = els.checklist.querySelector(`[data-check="${key}"]`);
    item.classList.toggle("done", isDone);
    item.textContent = getChecklistText(key, isDone);
  });

  const readyCount = Object.values(checks).filter(Boolean).length;
  els.readinessMeter.style.setProperty("--ready", `${(readyCount / 4) * 100}%`);
  els.fileMetric.textContent = `${Number(checks.fms) + Number(checks.haud)}/2`;
  els.dateMetric.textContent =
    els.processStartDate.value && els.processEndDate.value
      ? `${els.processStartDate.value} s/d ${els.processEndDate.value}`
      : "Belum dipilih";
}

function getChecklistText(key, isDone) {
  const messages = {
    fms: ["FMS belum diupload", "FMS sudah diupload"],
    haud: ["HAUD belum diupload", "HAUD sudah diupload"],
    date: ["Range tanggal belum lengkap", "Range tanggal valid"],
    mapping: ["Mapping kolom belum lengkap", "Mapping kolom lengkap"],
  };
  return messages[key][isDone ? 1 : 0];
}

function showPreview() {
  const fmsSample = state.rows.fms.slice(0, 3);
  const haudSample = state.rows.haud.slice(0, 3);

  if (!fmsSample.length && !haudSample.length) {
    els.previewBox.textContent = "Upload file FMS dan HAUD untuk melihat preview.";
    return;
  }

  els.previewBox.textContent = JSON.stringify(
    {
      fms: {
        rows: state.rows.fms.length,
        sample: fmsSample,
      },
      haud: {
        rows: state.rows.haud.length,
        sample: haudSample,
      },
    },
    null,
    2,
  );
}

function syncEndDate() {
  if (!els.processEndDate.value) {
    els.processEndDate.value = els.processStartDate.value;
  }
}

async function runMatching(event) {
  event.preventDefault();

  if (!isReadyToRun()) {
    els.previewBox.textContent = "Lengkapi file, range tanggal, dan mapping kolom sebelum menjalankan matching.";
    setProcessStatus("warning", "Input belum lengkap", "Lengkapi file, range tanggal, dan mapping kolom sebelum menjalankan matching.", 0);
    setProcessLog(["Proses dibatalkan karena input belum lengkap."]);
    updateReadiness();
    return;
  }

  setRunningState(true);
  resetOutputState();
  setProcessLog([]);

  try {
    const config = readConfig();
    addProcessLog("Validasi input selesai. Mulai membaca range tanggal dan mapping kolom.");
    setProcessStatus("running", "Menyiapkan data", "Membaca file FMS dan HAUD dari browser.", 8);
    await yieldToUi();

    const fmsRows = await filterComparableRows("fms", config, config.processStartDate, config.processEndDate, 10, 28);
    fmsRows.sort(compareComparableRows);
    addProcessLog(`FMS masuk range: ${formatNumber(fmsRows.length)} dari ${formatNumber(state.rows.fms.length)} baris.`);

    const haudPool = await readComparableRows("haud", config, 30, 46);
    haudPool.sort(compareComparableRows);
    addProcessLog(`HAUD dibaca penuh: ${formatNumber(haudPool.length)} dari ${formatNumber(state.rows.haud.length)} baris valid.`);

    setProcessStatus("running", "Membuat index HAUD", "Mengelompokkan HAUD berdasarkan destinationAddr agar matching lebih cepat.", 50);
    await yieldToUi();
    const haudIndex = indexHaudRows(haudPool);
    addProcessLog(`Index HAUD selesai: ${formatNumber(haudIndex.size)} nomor unik.`);

    const output = await matchRowsInBatches(fmsRows, haudIndex, config);

    state.output = output;
    renderResults(output, { fmsRows: fmsRows.length, haudRows: haudPool.length });
    prepareDownload(output);

    const processSummary = buildProcessSummary(config, fmsRows, haudPool, output);
    els.previewBox.textContent = processSummary.previewText;

    if (fmsRows.length === 0 || output.length === 0) {
      setProcessStatus("warning", "Proses selesai tanpa output", processSummary.message, 100);
      addProcessLog(processSummary.message);
    } else if (processSummary.hasMissingDates) {
      setProcessStatus("warning", "Output parsial dibuat", processSummary.message, 100);
      addProcessLog(processSummary.message);
      addProcessLog(`Output parsial siap: ${formatNumber(output.length)} baris. File akan diunduh otomatis.`);
      downloadOutput();
    } else {
      setProcessStatus("success", "Proses berhasil", processSummary.message, 100);
      addProcessLog(`Output siap: ${formatNumber(output.length)} baris. File akan diunduh otomatis.`);
      downloadOutput();
    }
  } catch (error) {
    setProcessStatus("error", "Proses gagal", error.message || "Terjadi error saat matching.", 100);
    addProcessLog(`Error: ${error.message || error}`);
  } finally {
    setRunningState(false);
  }
}

function isReadyToRun() {
  return (
    state.files.fms &&
    state.files.haud &&
    els.processStartDate.value &&
    els.processEndDate.value &&
    readDateInput(els.processStartDate.value) <= readDateInput(els.processEndDate.value) &&
    els.fmsNumberColumn.value &&
    els.fmsDateColumn.value &&
    els.haudSourceColumn.value &&
    els.haudNumberColumn.value &&
    els.haudDateColumn.value
  );
}

async function filterComparableRows(type, config, startDate, endDate, progressStart, progressEnd) {
  const rows = state.rows[type];
  const filtered = [];
  const batchSize = 5000;

  for (let index = 0; index < rows.length; index += batchSize) {
    const batch = rows.slice(index, index + batchSize);
    batch.forEach((row, batchIndex) => {
      const comparable = buildComparableRow(row, index + batchIndex, type, config);
      if (isDateInRange(comparable.date, startDate, endDate)) filtered.push(comparable);
    });

    const ratio = rows.length ? Math.min((index + batch.length) / rows.length, 1) : 1;
    const progress = progressStart + (progressEnd - progressStart) * ratio;
    const label = type === "fms" ? "Filter FMS" : "Filter HAUD";
    setProcessStatus("running", label, `${label}: ${formatNumber(index + batch.length)} dari ${formatNumber(rows.length)} baris dibaca.`, progress);
    await yieldToUi();
  }

  return filtered;
}

async function readComparableRows(type, config, progressStart, progressEnd) {
  const rows = state.rows[type];
  const parsed = [];
  const batchSize = 5000;

  for (let index = 0; index < rows.length; index += batchSize) {
    const batch = rows.slice(index, index + batchSize);
    batch.forEach((row, batchIndex) => {
      const comparable = buildComparableRow(row, index + batchIndex, type, config);
      if (comparable.date) parsed.push(comparable);
    });

    const ratio = rows.length ? Math.min((index + batch.length) / rows.length, 1) : 1;
    const progress = progressStart + (progressEnd - progressStart) * ratio;
    const label = type === "haud" ? "Baca HAUD penuh" : "Baca data";
    setProcessStatus(
      "running",
      label,
      `${label}: ${formatNumber(index + batch.length)} dari ${formatNumber(rows.length)} baris dibaca.`,
      progress,
    );
    await yieldToUi();
  }

  return parsed;
}

function indexHaudRows(haudRows) {
  const index = new Map();

  haudRows.forEach((row) => {
    if (!row.key) return;
    if (!index.has(row.key)) index.set(row.key, []);
    index.get(row.key).push(row);
  });

  index.forEach((rows) => rows.sort(compareComparableRows));

  return index;
}

async function matchRowsInBatches(fmsRows, haudIndex, config) {
  const output = [];
  const batchSize = 1000;

  for (let index = 0; index < fmsRows.length; index += batchSize) {
    const batch = fmsRows.slice(index, index + batchSize);

    batch.forEach((fms) => {
      const candidates = haudIndex.get(fms.key) || [];
      const matches = candidates.filter((haud) => isDateInRange(haud.date, addDays(fms.date, -1), addDays(fms.date, 1)));

      if (matches.length) {
        matches.forEach((match) => {
          output.push({
            fms_row: fms.excelRow,
            fms_date: formatDateForDisplay(fms.date),
            "B-Number Processed": fms.originalNumber,
            haud_row: match.excelRow,
            haud_date: formatDateForDisplay(match.date),
            sourceAddr: match.sourceAddr,
            destinationAddr: match.originalNumber,
            match_type: getMatchType(fms.date, match.date),
          });
        });
        return;
      }

      if (config.includeUnmatched) {
        output.push({
          fms_row: fms.excelRow,
          fms_date: formatDateForDisplay(fms.date),
          "B-Number Processed": fms.originalNumber,
          haud_row: "",
          haud_date: "",
          sourceAddr: "",
          destinationAddr: "",
          match_type: "FMS_ONLY",
        });
      }
    });

    const ratio = fmsRows.length ? Math.min((index + batch.length) / fmsRows.length, 1) : 1;
    const progress = 55 + 40 * ratio;
    setProcessStatus(
      "running",
      "Matching data",
      `Mencocokkan FMS: ${formatNumber(index + batch.length)} dari ${formatNumber(fmsRows.length)} baris. Output sementara: ${formatNumber(output.length)} baris.`,
      progress,
    );
    await yieldToUi();
  }

  return output.sort(compareOutputRows);
}

function buildProcessSummary(config, fmsRows, haudPool, output) {
  const fmsDateStats = getDateStats(state.rows.fms, config.fmsDateColumn, config);
  const requestedDates = listDateRange(config.processStartDate, config.processEndDate);
  const coveredDateKeys = new Set(fmsRows.map((row) => dateKey(row.date)));
  const missingDates = requestedDates.filter((date) => !coveredDateKeys.has(dateKey(date)));
  const matchedRows = output.filter((row) => row.match_type !== "FMS_ONLY");
  const fmsOnlyRows = output.filter((row) => row.match_type === "FMS_ONLY");
  const previewLines = [
    `Range FMS: ${els.processStartDate.value} s/d ${els.processEndDate.value}`,
    `FMS diproses: ${formatNumber(fmsRows.length)} dari ${formatNumber(state.rows.fms.length)} baris`,
    `Distribusi tanggal FMS: ${formatDateDistribution(fmsRows) || "-"}`,
    `HAUD dibaca penuh: ${formatNumber(haudPool.length)} dari ${formatNumber(state.rows.haud.length)} baris valid`,
    `Matched output: ${formatNumber(matchedRows.length)} baris`,
    `FMS tanpa match: ${formatNumber(fmsOnlyRows.length)} baris`,
    "Rule HAUD: tidak difilter oleh range dashboard; hanya dicek H-1, H, H+1 terhadap masing-masing tanggal FMS",
    "Output: fms_row, fms_date, B-Number Processed, haud_row, haud_date, sourceAddr, destinationAddr, match_type.",
  ];

  if (missingDates.length) {
    previewLines.push("");
    previewLines.push(`PERHATIAN: Tidak ada baris FMS untuk tanggal: ${missingDates.map(formatDateForDisplay).join(", ")}.`);
  }

  if (fmsRows.length === 0) {
    previewLines.push("PERHATIAN: Tidak ada tanggal FMS yang masuk range.");
    previewLines.push(`Tanggal FMS terdeteksi: ${fmsDateStats.min || "-"} s/d ${fmsDateStats.max || "-"}`);
    previewLines.push(`Contoh tanggal FMS: ${fmsDateStats.samples.join(", ") || "-"}`);
    previewLines.push("Coba ubah range tanggal atau Format tanggal file.");
    return {
      previewText: previewLines.join("\n"),
      message: "Tidak ada baris FMS dalam range yang dipilih. Output tidak dibuat otomatis karena tidak ada data yang diproses.",
      hasMissingDates: true,
    };
  }

  if (output.length === 0) {
    return {
      previewText: previewLines.join("\n"),
      message: "Range FMS valid, tetapi tidak ada pasangan HAUD di window H-1/H/H+1 dan opsi unmatched tidak disertakan.",
      hasMissingDates: missingDates.length > 0,
    };
  }

  if (missingDates.length) {
    return {
      previewText: previewLines.join("\n"),
      message: `Output dibuat untuk tanggal FMS yang tersedia, tetapi tanggal ${missingDates.map(formatDateForDisplay).join(", ")} tidak ditemukan pada data FMS.`,
      hasMissingDates: true,
    };
  }

  return {
    previewText: previewLines.join("\n"),
    message: `Berhasil membuat ${formatNumber(output.length)} baris output dari ${formatNumber(fmsRows.length)} baris FMS yang diproses.`,
    hasMissingDates: false,
  };
}

function readConfig() {
  return {
    processStartDate: readDateInput(els.processStartDate.value),
    processEndDate: readDateInput(els.processEndDate.value),
    dateFormat: els.dateFormat.value,
    outputFormat: els.outputFormat.value,
    ignoreCase: els.ignoreCase.checked,
    stripSeparators: els.stripSeparators.checked,
    removeLeadingZero: els.removeLeadingZero.checked,
    includeUnmatched: els.includeUnmatched.checked,
    fmsNumberColumn: els.fmsNumberColumn.value,
    fmsDateColumn: els.fmsDateColumn.value,
    haudSourceColumn: els.haudSourceColumn.value,
    haudNumberColumn: els.haudNumberColumn.value,
    haudDateColumn: els.haudDateColumn.value,
  };
}

function buildComparableRow(row, index, type, config) {
  const numberColumn = type === "fms" ? config.fmsNumberColumn : config.haudNumberColumn;
  const dateColumn = type === "fms" ? config.fmsDateColumn : config.haudDateColumn;
  const originalNumber = String(row[numberColumn] ?? "").trim();
  const originalDate = String(row[dateColumn] ?? "").trim();

  return {
    index,
    excelRow: index + 2,
    originalNumber,
    originalDate,
    sourceAddr: type === "haud" ? String(row[config.haudSourceColumn] ?? "").trim() : "",
    key: normalizeNumber(originalNumber, config),
    date: parseDate(originalDate, config),
  };
}

function normalizeNumber(value, config) {
  let normalized = String(value ?? "").trim();
  if (config.ignoreCase) normalized = normalized.toUpperCase();
  if (config.stripSeparators) normalized = normalized.replace(/[\s./\\_-]+/g, "");
  if (config.removeLeadingZero) normalized = normalized.replace(/^0+/, "");
  return normalized;
}

function parseDate(value, config = { dateFormat: "auto" }) {
  if (!value) return null;
  if (value instanceof Date) return clearTime(value);

  const text = String(value).trim();
  const excelSerial = Number(text);
  if (!Number.isNaN(excelSerial) && excelSerial > 20000 && excelSerial < 90000) {
    const parsed = XLSX.SSF.parse_date_code(excelSerial);
    if (parsed) return new Date(parsed.y, parsed.m - 1, parsed.d);
  }

  const slashMatch = text.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})$/);
  if (slashMatch) {
    return parseSlashDate(Number(slashMatch[1]), Number(slashMatch[2]), Number(slashMatch[3]), config.dateFormat);
  }

  const isoLikeMatch = text.match(/^(\d{4})[/-](\d{1,2})[/-](\d{1,2})$/);
  if (isoLikeMatch) {
    const year = Number(isoLikeMatch[1]);
    const month = Number(isoLikeMatch[2]);
    const day = Number(isoLikeMatch[3]);
    return createValidDate(year, month, day);
  }

  const parsed = new Date(text);
  return Number.isNaN(parsed.getTime()) ? null : clearTime(parsed);
}

function parseSlashDate(first, second, yearValue, dateFormat) {
  const year = normalizeYear(yearValue);
  const formats =
    dateFormat === "dmy"
      ? ["dmy", "mdy"]
      : dateFormat === "mdy"
        ? ["mdy", "dmy"]
        : first > 12 && second <= 12
          ? ["dmy", "mdy"]
          : second > 12 && first <= 12
            ? ["mdy", "dmy"]
            : ["dmy", "mdy"];

  for (const format of formats) {
    const day = format === "dmy" ? first : second;
    const month = format === "dmy" ? second : first;
    const date = createValidDate(year, month, day);
    if (date) return date;
  }

  return null;
}

function createValidDate(year, month, day) {
  if (!year || !month || !day || month < 1 || month > 12 || day < 1 || day > 31) return null;
  const date = new Date(year, month - 1, day);
  if (date.getFullYear() !== year || date.getMonth() !== month - 1 || date.getDate() !== day) return null;
  return date;
}

function formatDateForDisplay(date) {
  if (!date) return "";
  return [
    String(date.getDate()).padStart(2, "0"),
    String(date.getMonth() + 1).padStart(2, "0"),
    String(date.getFullYear()).slice(-2),
  ].join("/");
}

function getDateStats(rows, dateColumn, config) {
  const dates = [];
  const samples = [];

  rows.forEach((row) => {
    const raw = row[dateColumn];
    const parsed = parseDate(raw, config);
    if (samples.length < 6 && raw !== undefined && raw !== "") samples.push(String(raw));
    if (parsed) dates.push(parsed);
  });

  dates.sort((left, right) => left - right);
  return {
    min: dates.length ? formatDateForDisplay(dates[0]) : "",
    max: dates.length ? formatDateForDisplay(dates[dates.length - 1]) : "",
    samples,
  };
}

function formatDateDistribution(rows) {
  const counts = new Map();

  rows.forEach((row) => {
    const label = formatDateForDisplay(row.date);
    if (!label) return;
    counts.set(label, (counts.get(label) || 0) + 1);
  });

  return Array.from(counts.entries())
    .sort(([left], [right]) => parseDisplayDate(left) - parseDisplayDate(right))
    .map(([date, count]) => `${date}: ${formatNumber(count)}`)
    .join(", ");
}

function parseDisplayDate(value) {
  const match = String(value).match(/^(\d{2})\/(\d{2})\/(\d{2})$/);
  if (!match) return 0;
  return new Date(normalizeYear(Number(match[3])), Number(match[2]) - 1, Number(match[1])).getTime();
}

function normalizeYear(year) {
  if (year < 100) return year + 2000;
  return year;
}

function clearTime(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function readDateInput(value) {
  return value ? new Date(`${value}T00:00:00`) : null;
}

function addDays(date, days) {
  const next = clearTime(date);
  next.setDate(next.getDate() + days);
  return next;
}

function listDateRange(startDate, endDate) {
  const dates = [];
  for (let current = clearTime(startDate); dateKey(current) <= dateKey(endDate); current = addDays(current, 1)) {
    dates.push(current);
  }
  return dates;
}

function isDateInRange(date, startDate, endDate) {
  if (!date || !startDate || !endDate) return false;
  const value = dateKey(date);
  return value >= dateKey(startDate) && value <= dateKey(endDate);
}

function datesEqual(leftDate, rightDate) {
  if (!leftDate || !rightDate) return false;
  return dateKey(leftDate) === dateKey(rightDate);
}

function dateKey(date) {
  const cleared = clearTime(date);
  return [
    cleared.getFullYear(),
    String(cleared.getMonth() + 1).padStart(2, "0"),
    String(cleared.getDate()).padStart(2, "0"),
  ].join("-");
}

function compareComparableRows(left, right) {
  const dateCompare = dateKey(left.date).localeCompare(dateKey(right.date));
  if (dateCompare !== 0) return dateCompare;
  return left.excelRow - right.excelRow;
}

function compareOutputRows(left, right) {
  const dateCompare = parseDisplayDate(left.fms_date) - parseDisplayDate(right.fms_date);
  if (dateCompare !== 0) return dateCompare;
  return Number(left.fms_row || 0) - Number(right.fms_row || 0);
}

function getMatchType(fmsDate, haudDate) {
  const diffDays = Math.round((clearTime(haudDate) - clearTime(fmsDate)) / 86400000);
  if (diffDays === -1) return "H-1";
  if (diffDays === 1) return "H+1";
  return "H";
}

function renderResults(output, processed = { fmsRows: state.rows.fms.length, haudRows: state.rows.haud.length }) {
  const matchedRows = output.filter((row) => row.match_type !== "FMS_ONLY");
  const matched = matchedRows.length;
  const matchedFmsRows = new Set(matchedRows.map((row) => row.fms_row)).size;
  const fmsOnly = output.filter((row) => row.match_type === "FMS_ONLY").length;
  const haudWindow = processed.haudRows;
  const total = output.length;
  const matchRate = processed.fmsRows ? Math.round((matchedFmsRows / processed.fmsRows) * 100) : 0;

  els.matchedCount.textContent = matched;
  els.fmsOnlyCount.textContent = fmsOnly;
  els.haudOnlyCount.textContent = haudWindow;
  els.totalCount.textContent = total;
  els.matchMetric.textContent = `${matchRate}%`;
  els.outputMetric.textContent = `${total} baris`;

  els.resultBody.innerHTML = output.length
    ? output
        .slice(0, 200)
        .map(
          (row) => `
            <tr>
              <td>${escapeHtml(row.fms_row)}</td>
              <td>${escapeHtml(row.fms_date)}</td>
              <td>${escapeHtml(row["B-Number Processed"])}</td>
              <td>${escapeHtml(row.haud_row)}</td>
              <td>${escapeHtml(row.haud_date)}</td>
              <td>${escapeHtml(row.sourceAddr)}</td>
              <td>${escapeHtml(row.destinationAddr)}</td>
              <td>${statusPill(row.match_type)}</td>
            </tr>
          `,
        )
        .join("")
    : `<tr><td colspan="8">Tidak ada baris output.</td></tr>`;

  if (output.length > 200) {
    els.resultBody.insertAdjacentHTML(
      "beforeend",
      `<tr><td colspan="8">Preview dibatasi 200 baris. Download untuk melihat seluruh data.</td></tr>`,
    );
  }
}

function statusPill(matchType) {
  const className = matchType === "FMS_ONLY" ? "fms" : matchType === "H" ? "match" : "haud";
  return `<span class="pill ${className}">${matchType}</span>`;
}

function setProcessStatus(stateName, title, message, progress) {
  const roundedProgress = Math.max(0, Math.min(100, Math.round(progress)));
  const stateLabels = {
    idle: "Belum dijalankan",
    running: "Sedang proses",
    success: "Berhasil",
    warning: "Perlu dicek",
    error: "Gagal",
  };

  els.processBox.dataset.state = stateName;
  els.processState.textContent = stateLabels[stateName] || stateLabels.idle;
  els.processTitle.textContent = title;
  els.processMessage.textContent = message;
  els.processPercent.textContent = `${roundedProgress}%`;
  els.processBar.style.setProperty("--progress", `${roundedProgress}%`);
}

function setProcessLog(items) {
  els.processLog.innerHTML = items.length ? items.map((item) => `<li>${escapeHtml(item)}</li>`).join("") : "";
}

function addProcessLog(message) {
  if (!els.processLog.children.length) {
    els.processLog.innerHTML = "";
  }
  els.processLog.insertAdjacentHTML("beforeend", `<li>${escapeHtml(message)}</li>`);
  els.processLog.scrollTop = els.processLog.scrollHeight;
}

function resetOutputState() {
  state.output = [];
  state.outputBlob = null;
  state.outputName = "";
  els.downloadButton.disabled = true;
  renderResults([], { fmsRows: 0, haudRows: 0 });
}

function setRunningState(isRunning) {
  els.form.querySelectorAll("button, input, select").forEach((control) => {
    if (control === els.downloadButton) return;
    control.disabled = isRunning;
  });
}

function yieldToUi() {
  return new Promise((resolve) => setTimeout(resolve, 0));
}

function formatNumber(value) {
  return new Intl.NumberFormat("id-ID").format(value);
}

function prepareDownload(output) {
  const format = els.outputFormat.value;
  const startStamp = els.processStartDate.value || new Date().toISOString().slice(0, 10);
  const endStamp = els.processEndDate.value || startStamp;

  if (format === "json") {
    state.outputBlob = new Blob([JSON.stringify(output, null, 2)], { type: "application/json" });
    state.outputName = `hasil_matching_${startStamp}_sd_${endStamp}.json`;
  } else {
    state.outputBlob = new Blob([toCsv(output)], { type: "text/csv;charset=utf-8" });
    state.outputName = `hasil_matching_${startStamp}_sd_${endStamp}.csv`;
  }

  els.downloadButton.disabled = false;
}

function toCsv(rows) {
  const headers = ["fms_row", "fms_date", "B-Number Processed", "haud_row", "haud_date", "sourceAddr", "destinationAddr", "match_type"];
  const lines = [headers.join(",")];
  rows.forEach((row) => {
    lines.push(headers.map((header) => csvEscape(row[header])).join(","));
  });
  return lines.join("\n");
}

function csvEscape(value) {
  const text = String(value ?? "");
  if (/[",\n]/.test(text)) return `"${text.replace(/"/g, '""')}"`;
  return text;
}

function downloadOutput() {
  if (!state.outputBlob) return;

  const url = URL.createObjectURL(state.outputBlob);
  const link = document.createElement("a");
  link.href = url;
  link.download = state.outputName;
  link.click();
  URL.revokeObjectURL(url);
}

function resetDashboard() {
  state.files.fms = null;
  state.files.haud = null;
  state.rows.fms = [];
  state.rows.haud = [];
  state.workbook.fms = null;
  state.workbook.haud = null;
  state.output = [];
  state.outputBlob = null;
  state.outputName = "";

  els.form.reset();
  els.fmsFileName.textContent = "Pilih file FMS";
  els.haudFileName.textContent = "Pilih file HAUD";
  document.querySelectorAll(".upload-zone").forEach((zone) => zone.classList.remove("ready"));

  [
    els.fmsSheet,
    els.haudSheet,
    els.fmsNumberColumn,
    els.fmsDateColumn,
    els.haudSourceColumn,
    els.haudNumberColumn,
    els.haudDateColumn,
  ].forEach((select) => {
    select.disabled = true;
    select.innerHTML = `<option value="">Upload file dulu</option>`;
  });

  els.previewBox.textContent = "Belum ada preview.";
  els.resultBody.innerHTML = `<tr><td colspan="8">Belum ada proses.</td></tr>`;
  els.downloadButton.disabled = true;
  els.matchMetric.textContent = "-";
  els.outputMetric.textContent = "Belum ada";
  els.matchedCount.textContent = "0";
  els.fmsOnlyCount.textContent = "0";
  els.haudOnlyCount.textContent = "0";
  els.totalCount.textContent = "0";
  setProcessStatus(
    "idle",
    "Menunggu proses matching",
    "Upload file, lengkapi range tanggal dan mapping, lalu klik Jalankan Matching.",
    0,
  );
  setProcessLog(["Belum ada aktivitas proses."]);
  updateReadiness();
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

updateReadiness();
