const DEFAULT_SHEETS = ["供应数据", "供应商表", "设备零件表"];

const meta = document.getElementById("meta");
const tableHead = document.getElementById("table-head");
const tableBody = document.getElementById("table-body");
const sheetTabs = document.getElementById("sheet-tabs");
const csvLink = document.getElementById("download-csv");
const sqlLink = document.getElementById("download-sql");

const initialPayload = window.__INITIAL_DATA__ || null;
const availableSheets =
  Array.isArray(window.__AVAILABLE_SHEETS__) && window.__AVAILABLE_SHEETS__.length
    ? window.__AVAILABLE_SHEETS__
    : DEFAULT_SHEETS;

let currentSheet = initialPayload?.sheetName || availableSheets[0] || "";

function setDownloadLinks(sheetName) {
  if (!sheetName) return;
  const encodedSheet = encodeURIComponent(sheetName);
  csvLink.href = `/api/data.csv?sheet=${encodedSheet}`;
  sqlLink.href = `/api/mysql.sql?sheet=${encodedSheet}`;
}

function getCurrentTimeString() {
  const now = new Date();
  return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')} ${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}:${String(now.getSeconds()).padStart(2, '0')}`;
}

function updateMetaTime() {
  if (!currentSheet || !meta.dataset.rowCount) return;
  meta.textContent = `工作表：${currentSheet} | 记录数：${meta.dataset.rowCount} | 生成时间：${getCurrentTimeString()}`;
}

function renderTable(payload) {
  const headers = payload.headers || [];
  const rows = payload.rows || [];

  meta.dataset.rowCount = rows.length;
  updateMetaTime();
  meta.style.color = "";

  tableHead.innerHTML = "";
  tableBody.innerHTML = "";

  const trHead = document.createElement("tr");
  headers.forEach((header) => {
    const th = document.createElement("th");
    th.textContent = header;
    trHead.appendChild(th);
  });
  tableHead.appendChild(trHead);

  rows.forEach((row) => {
    const tr = document.createElement("tr");
    headers.forEach((header) => {
      const td = document.createElement("td");
      td.textContent = row[header] === null || row[header] === undefined ? "" : row[header];
      td.setAttribute("data-field", header);
      tr.appendChild(td);
    });
    tableBody.appendChild(tr);
  });
}

function renderSheetTabs(activeSheetName) {
  sheetTabs.innerHTML = "";

  availableSheets.forEach((sheetName) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "sheet-tab";
    if (sheetName === activeSheetName) {
      button.classList.add("active");
    }
    button.textContent = sheetName;
    button.addEventListener("click", () => {
      if (sheetName === currentSheet) return;
      loadData(sheetName);
    });
    sheetTabs.appendChild(button);
  });
}

async function loadData(sheetName, useInitial = false) {
  try {
    const targetSheet = sheetName || currentSheet;
    let payload = null;

    if (useInitial && initialPayload && initialPayload.sheetName === targetSheet) {
      payload = initialPayload;
    } else {
      const response = await fetch(`/api/data?sheet=${encodeURIComponent(targetSheet)}`);
      payload = await response.json();
    }

    if (!payload.ok) {
      meta.textContent = payload.message;
      meta.style.color = "#c0392b";
      return;
    }

    currentSheet = payload.sheetName;
    renderSheetTabs(currentSheet);
    setDownloadLinks(currentSheet);
    renderTable(payload);
  } catch (error) {
    meta.textContent = `数据加载失败：${error.message}`;
    meta.style.color = "#c0392b";
  }
}

renderSheetTabs(currentSheet);
setDownloadLinks(currentSheet);
loadData(currentSheet, true);

// 每秒更新时间
setInterval(updateMetaTime, 1000);
