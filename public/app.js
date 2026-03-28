async function loadData() {
  const meta = document.getElementById("meta");
  const tableHead = document.getElementById("table-head");
  const tableBody = document.getElementById("table-body");

  try {
    const response = await fetch("/api/data");
    const payload = await response.json();

    if (!payload.ok) {
      meta.textContent = payload.message;
      meta.style.color = "#c0392b";
      return;
    }

    const headers = payload.headers || [];
    const rows = payload.rows || [];

    meta.textContent = `工作表：${payload.sheetName} | 记录数：${rows.length} | 生成时间：${new Date(payload.generatedAt).toLocaleString()}`;

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
  } catch (error) {
    meta.textContent = `数据加载失败：${error.message}`;
    meta.style.color = "#c0392b";
  }
}

loadData();
