(function initApp() {
  const elements = {
    fileInput: document.getElementById("fileInput"),
    sheetSelect: document.getElementById("sheetSelect"),
    categorySelect: document.getElementById("categorySelect"),
    valueSelect: document.getElementById("valueSelect"),
    chartTypeSelect: document.getElementById("chartTypeSelect"),
    themeSelect: document.getElementById("themeSelect"),
    emptyState: document.getElementById("emptyState"),
    dashboard: document.getElementById("dashboard"),
    rowCount: document.getElementById("rowCount"),
    columnCount: document.getElementById("columnCount"),
    numericCount: document.getElementById("numericCount"),
    fileName: document.getElementById("fileName"),
    chartTitle: document.getElementById("chartTitle"),
    chartHint: document.getElementById("chartHint"),
    chartBadge: document.getElementById("chartBadge"),
    chartInsights: document.getElementById("chartInsights"),
    chartCanvas: document.getElementById("chartCanvas"),
    previewTable: document.getElementById("previewTable"),
    tableHint: document.getElementById("tableHint"),
    toast: document.getElementById("toast")
  };

  const state = {
    workbook: null,
    currentFileName: "",
    headers: [],
    rows: [],
    numericColumns: [],
    chart: null
  };

  const chartTypeNames = {
    bar: "柱状图",
    line: "折线图",
    pie: "饼图"
  };

  const chartThemes = {
    aurora: {
      primary: "#2563eb",
      secondary: "#14b8a6",
      soft: "rgba(37, 99, 235, 0.14)",
      grid: "rgba(37, 99, 235, 0.12)",
      colors: ["#2563eb", "#14b8a6", "#f97316", "#7c3aed", "#dc2626", "#059669", "#db2777", "#4f46e5", "#ca8a04", "#0891b2", "#9333ea", "#475569", "#ea580c"]
    },
    sunset: {
      primary: "#f97316",
      secondary: "#db2777",
      soft: "rgba(249, 115, 22, 0.16)",
      grid: "rgba(249, 115, 22, 0.13)",
      colors: ["#f97316", "#db2777", "#facc15", "#ef4444", "#a855f7", "#06b6d4", "#84cc16", "#fb7185", "#f59e0b", "#8b5cf6", "#22c55e", "#64748b", "#e11d48"]
    },
    forest: {
      primary: "#059669",
      secondary: "#0f766e",
      soft: "rgba(5, 150, 105, 0.15)",
      grid: "rgba(5, 150, 105, 0.13)",
      colors: ["#059669", "#0f766e", "#65a30d", "#0891b2", "#2563eb", "#ca8a04", "#16a34a", "#4f46e5", "#84cc16", "#14b8a6", "#9333ea", "#475569", "#f97316"]
    },
    ink: {
      primary: "#334155",
      secondary: "#2563eb",
      soft: "rgba(51, 65, 85, 0.14)",
      grid: "rgba(51, 65, 85, 0.13)",
      colors: ["#334155", "#2563eb", "#0f766e", "#7c3aed", "#ea580c", "#0891b2", "#475569", "#9333ea", "#059669", "#ca8a04", "#dc2626", "#64748b", "#14b8a6"]
    }
  };

  elements.fileInput.addEventListener("change", handleFileUpload);
  elements.sheetSelect.addEventListener("change", () => loadSheet(elements.sheetSelect.value));
  elements.categorySelect.addEventListener("change", renderChart);
  elements.valueSelect.addEventListener("change", renderChart);
  elements.chartTypeSelect.addEventListener("change", renderChart);
  elements.themeSelect.addEventListener("change", renderChart);

  async function handleFileUpload(event) {
    const file = event.target.files?.[0];
    if (!file) return;

    try {
      showToast("正在读取文件...");
      const buffer = await file.arrayBuffer();
      state.workbook = XLSX.read(buffer, {
        type: "array",
        cellDates: true,
        cellText: false
      });
      state.currentFileName = file.name;
      hydrateSheetOptions();
      loadSheet(state.workbook.SheetNames[0]);
      showToast("文件读取完成");
    } catch (error) {
      console.error(error);
      showToast("文件解析失败，请确认格式是否正确");
    }
  }

  function hydrateSheetOptions() {
    const sheetNames = state.workbook?.SheetNames || [];
    fillSelect(elements.sheetSelect, sheetNames, sheetNames[0]);
    elements.sheetSelect.disabled = sheetNames.length === 0;
  }

  function loadSheet(sheetName) {
    const sheet = state.workbook?.Sheets?.[sheetName];
    if (!sheet) return;

    const rawRows = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      blankrows: false,
      defval: ""
    });

    const normalized = DataVizUtils.normalizeRows(rawRows);
    state.headers = normalized.headers;
    state.rows = normalized.rows;
    state.numericColumns = DataVizUtils.detectNumericColumns(state.rows, state.headers);

    hydrateControls();
    renderSummary();
    renderTable();
    renderChart();
    elements.emptyState.classList.add("hidden");
    elements.dashboard.classList.remove("hidden");
  }

  function hydrateControls() {
    const defaultCategory = DataVizUtils.chooseDefaultCategory(state.headers, state.numericColumns);
    const defaultValue = state.numericColumns[0] || "";

    fillSelect(elements.categorySelect, state.headers, defaultCategory);
    fillSelect(elements.valueSelect, state.numericColumns, defaultValue);

    elements.categorySelect.disabled = state.headers.length === 0;
    elements.valueSelect.disabled = state.numericColumns.length === 0;
    elements.chartTypeSelect.disabled = state.numericColumns.length === 0;
    elements.themeSelect.disabled = state.numericColumns.length === 0;
  }

  function fillSelect(select, values, selectedValue) {
    select.innerHTML = "";

    if (values.length === 0) {
      const option = document.createElement("option");
      option.value = "";
      option.textContent = "无可用字段";
      select.appendChild(option);
      return;
    }

    values.forEach((value) => {
      const option = document.createElement("option");
      option.value = value;
      option.textContent = value;
      if (value === selectedValue) option.selected = true;
      select.appendChild(option);
    });
  }

  function renderSummary() {
    elements.rowCount.textContent = state.rows.length.toLocaleString("zh-CN");
    elements.columnCount.textContent = state.headers.length.toLocaleString("zh-CN");
    elements.numericCount.textContent = state.numericColumns.length.toLocaleString("zh-CN");
    elements.fileName.textContent = state.currentFileName || "-";
  }

  function renderTable() {
    const previewRows = DataVizUtils.getPreviewRows(state.rows);
    const fragment = document.createDocumentFragment();

    const thead = document.createElement("thead");
    const headRow = document.createElement("tr");
    state.headers.forEach((header) => {
      const th = document.createElement("th");
      th.textContent = header;
      headRow.appendChild(th);
    });
    thead.appendChild(headRow);

    const tbody = document.createElement("tbody");
    previewRows.forEach((row) => {
      const tr = document.createElement("tr");
      state.headers.forEach((header) => {
        const td = document.createElement("td");
        td.textContent = DataVizUtils.toDisplayText(row[header]);
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });

    fragment.appendChild(thead);
    fragment.appendChild(tbody);
    elements.previewTable.replaceChildren(fragment);
    elements.tableHint.textContent = `显示前 ${previewRows.length} 行，共 ${state.rows.length} 行。`;
  }

  function renderChart() {
    const categoryColumn = elements.categorySelect.value;
    const valueColumn = elements.valueSelect.value;
    const chartType = elements.chartTypeSelect.value;
    const theme = chartThemes[elements.themeSelect.value] || chartThemes.aurora;

    if (!categoryColumn || !valueColumn) {
      destroyChart();
      elements.chartInsights.replaceChildren();
      elements.chartTitle.textContent = "没有可用图表";
      elements.chartHint.textContent = "当前工作表未识别到数值字段。";
      return;
    }

    const chartData = DataVizUtils.buildChartData(state.rows, categoryColumn, valueColumn, chartType);
    const isPie = chartType === "pie";
    const isLine = chartType === "line";
    const isBar = chartType === "bar";

    if (chartData.values.length === 0) {
      destroyChart();
      elements.chartInsights.replaceChildren();
      elements.chartTitle.textContent = "没有可用图表";
      elements.chartHint.textContent = "当前字段没有可用数值。";
      return;
    }

    destroyChart();
    renderChartInsights(chartData, valueColumn);

    const context = elements.chartCanvas.getContext("2d");
    const chartGradient = createChartGradient(context, theme);
    const lineGradient = createLineGradient(context, theme);

    state.chart = new Chart(elements.chartCanvas, {
      type: chartType,
      data: {
        labels: chartData.labels,
        datasets: [
          {
            label: valueColumn,
            data: chartData.values,
            borderColor: isPie ? "#ffffff" : theme.primary,
            backgroundColor: isPie ? theme.colors : isLine ? lineGradient : chartGradient,
            hoverBackgroundColor: isPie ? theme.colors : theme.secondary,
            borderRadius: isBar ? 9 : 0,
            borderSkipped: false,
            maxBarThickness: 46,
            pointRadius: isLine ? 4 : 0,
            pointHoverRadius: isLine ? 7 : 0,
            pointBackgroundColor: "#ffffff",
            pointBorderColor: theme.primary,
            pointBorderWidth: isLine ? 2 : 0,
            borderWidth: isPie ? 3 : isLine ? 3 : 1.5,
            hoverOffset: isPie ? 9 : 0,
            spacing: isPie ? 2 : 0,
            fill: isLine,
            tension: isLine ? 0.34 : 0
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        animation: {
          duration: 720,
          easing: "easeOutQuart"
        },
        layout: {
          padding: {
            top: 4,
            right: 12,
            bottom: 4,
            left: 6
          }
        },
        plugins: {
          legend: {
            display: isPie,
            position: "bottom",
            labels: {
              boxWidth: 9,
              usePointStyle: true,
              color: "#334155",
              padding: 16,
              font: { size: 12, weight: "700" }
            }
          },
          tooltip: {
            backgroundColor: "rgba(15, 23, 42, 0.92)",
            borderColor: "rgba(255, 255, 255, 0.12)",
            borderWidth: 1,
            padding: 12,
            displayColors: true,
            titleFont: { size: 13, weight: "800" },
            bodyFont: { size: 13, weight: "600" },
            callbacks: {
              label(context) {
                const label = context.label || context.dataset.label || "";
                const value = Number(context.raw || 0).toLocaleString("zh-CN");
                return `${label}: ${value}`;
              }
            }
          }
        },
        scales: isPie ? {} : {
          x: {
            ticks: {
              color: "#64748b",
              font: { size: 12, weight: "700" },
              maxRotation: 45,
              minRotation: 0
            },
            grid: { display: false }
          },
          y: {
            beginAtZero: true,
            ticks: {
              color: "#64748b",
              callback(value) {
                return compactNumber(value);
              }
            },
            grid: {
              color: theme.grid,
              drawTicks: false
            },
            border: { display: false }
          }
        }
      }
    });

    elements.chartTitle.textContent = `${chartTypeNames[chartType]}：${valueColumn}`;
    elements.chartBadge.textContent = chartTypeNames[chartType];
    elements.chartHint.textContent = `${categoryColumn} 分组，展示 ${chartData.shownGroups} 项。`;
  }

  function createChartGradient(context, theme) {
    const height = elements.chartCanvas.parentElement?.clientHeight || 320;
    const gradient = context.createLinearGradient(0, 0, 0, height);
    gradient.addColorStop(0, theme.primary);
    gradient.addColorStop(0.62, theme.secondary);
    gradient.addColorStop(1, "rgba(255, 255, 255, 0.9)");
    return gradient;
  }

  function createLineGradient(context, theme) {
    const height = elements.chartCanvas.parentElement?.clientHeight || 320;
    const gradient = context.createLinearGradient(0, 0, 0, height);
    gradient.addColorStop(0, theme.soft);
    gradient.addColorStop(1, "rgba(255, 255, 255, 0)");
    return gradient;
  }

  function renderChartInsights(chartData, valueColumn) {
    const stats = DataVizUtils.summarizeValues(chartData.values);
    const items = [
      ["总计", formatNumber(stats.total)],
      ["均值", formatNumber(stats.average)],
      ["最大", formatNumber(stats.max)],
      ["分组", `${chartData.totalGroups} 组`]
    ];

    const fragment = document.createDocumentFragment();
    items.forEach(([label, value]) => {
      const card = document.createElement("article");
      card.className = "insight";
      const title = document.createElement("span");
      title.textContent = label;
      const number = document.createElement("strong");
      number.textContent = value;
      card.append(title, number);
      fragment.appendChild(card);
    });

    elements.chartInsights.replaceChildren(fragment);
    elements.chartInsights.setAttribute("data-value-column", valueColumn);
  }

  function compactNumber(value) {
    return new Intl.NumberFormat("zh-CN", {
      notation: "compact",
      maximumFractionDigits: 1
    }).format(value);
  }

  function formatNumber(value) {
    return new Intl.NumberFormat("zh-CN", {
      maximumFractionDigits: 2
    }).format(value);
  }

  function destroyChart() {
    if (!state.chart) return;
    state.chart.destroy();
    state.chart = null;
  }

  function showToast(message) {
    elements.toast.textContent = message;
    elements.toast.classList.remove("hidden");
    window.clearTimeout(showToast.timer);
    showToast.timer = window.setTimeout(() => {
      elements.toast.classList.add("hidden");
    }, 2200);
  }
})();
