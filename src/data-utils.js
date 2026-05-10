(function initDataVizUtils(root, factory) {
  if (typeof module === "object" && module.exports) {
    module.exports = factory();
  } else {
    root.DataVizUtils = factory();
  }
})(typeof self !== "undefined" ? self : this, function createDataVizUtils() {
  const MAX_CHART_POINTS = 80;
  const MAX_PIE_SLICES = 12;
  const EMPTY_LABEL = "(空白)";

  function normalizeHeader(value, index) {
    const text = String(value ?? "").trim();
    return text || `字段${index + 1}`;
  }

  function toDisplayText(value) {
    if (value === null || value === undefined) return "";
    if (value instanceof Date) return value.toLocaleDateString();
    return String(value);
  }

  function normalizeRows(rawRows) {
    if (!Array.isArray(rawRows) || rawRows.length === 0) {
      return { headers: [], rows: [] };
    }

    const maxLength = rawRows.reduce((max, row) => {
      return Math.max(max, Array.isArray(row) ? row.length : 0);
    }, 0);

    const headers = Array.from({ length: maxLength }, (_, index) => {
      return normalizeHeader(rawRows[0]?.[index], index);
    });

    const rows = rawRows.slice(1).map((row) => {
      const record = {};
      headers.forEach((header, index) => {
        record[header] = row?.[index] ?? "";
      });
      return record;
    }).filter((record) => {
      return headers.some((header) => toDisplayText(record[header]).trim() !== "");
    });

    return { headers, rows };
  }

  function parseNumber(value) {
    if (typeof value === "number" && Number.isFinite(value)) return value;
    if (value instanceof Date) return Number.NaN;
    if (typeof value !== "string") return Number.NaN;

    const normalized = value
      .trim()
      .replace(/[%￥¥$,\s]/g, "")
      .replace(/^[(](.*)[)]$/, "-$1");

    if (!normalized) return Number.NaN;
    const parsed = Number(normalized);
    return Number.isFinite(parsed) ? parsed : Number.NaN;
  }

  function detectNumericColumns(rows, headers) {
    return headers.filter((header) => {
      let validCount = 0;
      let nonEmptyCount = 0;

      rows.forEach((row) => {
        const text = toDisplayText(row[header]).trim();
        if (!text) return;
        nonEmptyCount += 1;
        if (Number.isFinite(parseNumber(row[header]))) validCount += 1;
      });

      if (nonEmptyCount === 0) return false;
      return validCount / nonEmptyCount >= 0.8;
    });
  }

  function chooseDefaultCategory(headers, numericColumns) {
    return headers.find((header) => !numericColumns.includes(header)) || headers[0] || "";
  }

  function buildChartData(rows, categoryColumn, valueColumn, chartType) {
    const totals = new Map();

    rows.forEach((row) => {
      const value = parseNumber(row[valueColumn]);
      if (!Number.isFinite(value)) return;

      const label = toDisplayText(row[categoryColumn]).trim() || EMPTY_LABEL;
      totals.set(label, (totals.get(label) || 0) + value);
    });

    let entries = Array.from(totals.entries());
    if (chartType === "pie") {
      entries.sort((a, b) => Math.abs(b[1]) - Math.abs(a[1]));
      const visible = entries.slice(0, MAX_PIE_SLICES);
      const rest = entries.slice(MAX_PIE_SLICES).reduce((sum, item) => sum + item[1], 0);
      entries = rest ? [...visible, ["其他", rest]] : visible;
    } else {
      entries = entries.slice(0, MAX_CHART_POINTS);
    }

    return {
      labels: entries.map(([label]) => label),
      values: entries.map(([, value]) => Number(value.toFixed(4))),
      totalGroups: totals.size,
      shownGroups: entries.length
    };
  }

  function summarizeValues(values) {
    const numbers = values.filter((value) => Number.isFinite(value));
    if (numbers.length === 0) {
      return {
        count: 0,
        total: 0,
        average: 0,
        max: 0,
        min: 0
      };
    }

    const total = numbers.reduce((sum, value) => sum + value, 0);
    return {
      count: numbers.length,
      total: roundToFour(total),
      average: roundToFour(total / numbers.length),
      max: Math.max(...numbers),
      min: Math.min(...numbers)
    };
  }

  function roundToFour(value) {
    return Math.round((value + 1e-10) * 10000) / 10000;
  }

  function getPreviewRows(rows, limit = 50) {
    return rows.slice(0, limit);
  }

  return {
    MAX_CHART_POINTS,
    MAX_PIE_SLICES,
    normalizeRows,
    parseNumber,
    detectNumericColumns,
    chooseDefaultCategory,
    buildChartData,
    summarizeValues,
    getPreviewRows,
    toDisplayText
  };
});
