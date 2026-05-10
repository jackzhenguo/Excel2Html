(function initApp() {
  const elements = {
    fileInput: document.getElementById("fileInput"),
    sheetSelect: document.getElementById("sheetSelect"),
    categorySelect: document.getElementById("categorySelect"),
    valueSelect: document.getElementById("valueSelect"),
    chartTypeSelect: document.getElementById("chartTypeSelect"),
    emptyState: document.getElementById("emptyState"),
    dashboard: document.getElementById("dashboard"),
    rowCount: document.getElementById("rowCount"),
    columnCount: document.getElementById("columnCount"),
    numericCount: document.getElementById("numericCount"),
    fileName: document.getElementById("fileName"),
    chartTitle: document.getElementById("chartTitle"),
    chartHint: document.getElementById("chartHint"),
    chartBadge: document.getElementById("chartBadge"),
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

  const palette = [
    "#2563eb",
    "#14b8a6",
    "#f97316",
    "#7c3aed",
    "#dc2626",
    "#059669",
    "#db2777",
    "#4f46e5",
    "#ca8a04",
    "#0891b2",
    "#9333ea",
    "#475569",
    "#ea580c"
  ];

  elements.fileInput.addEventListener("change", handleFileUpload);
  elements.sheetSelect.addEventListener("change", () => loadSheet(elements.sheetSelect.value));
  elements.categorySelect.addEventListener("change", renderChart);
  elements.valueSelect.addEventListener("change", renderChart);
  elements.chartTypeSelect.addEventListener("change", renderChart);

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

    if (!categoryColumn || !valueColumn) {
      destroyChart();
      elements.chartTitle.textContent = "没有可用图表";
      elements.chartHint.textContent = "当前工作表未识别到数值字段。";
      return;
    }

    const chartData = DataVizUtils.buildChartData(state.rows, categoryColumn, valueColumn, chartType);
    const isPie = chartType === "pie";

    destroyChart();
    state.chart = new Chart(elements.chartCanvas, {
      type: chartType,
      data: {
        labels: chartData.labels,
        datasets: [
          {
            label: valueColumn,
            data: chartData.values,
            borderColor: "#2563eb",
            backgroundColor: isPie ? palette : "rgba(37, 99, 235, 0.72)",
            pointRadius: chartType === "line" ? 3 : 0,
            pointHoverRadius: chartType === "line" ? 5 : 0,
            borderWidth: chartType === "line" ? 2 : 1.5,
            tension: chartType === "line" ? 0.25 : 0
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            display: isPie,
            position: "bottom",
            labels: {
              boxWidth: 12,
              color: "#334155",
              font: { size: 12 }
            }
          },
          tooltip: {
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
              maxRotation: 45,
              minRotation: 0
            },
            grid: { display: false }
          },
          y: {
            beginAtZero: true,
            ticks: { color: "#64748b" },
            grid: { color: "rgba(148, 163, 184, 0.2)" }
          }
        }
      }
    });

    elements.chartTitle.textContent = `${chartTypeNames[chartType]}：${valueColumn}`;
    elements.chartBadge.textContent = chartTypeNames[chartType];
    elements.chartHint.textContent = `${categoryColumn} 分组，展示 ${chartData.shownGroups} 项。`;
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
