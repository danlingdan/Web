const express = require("express");
const path = require("path");
const fs = require("fs");
const XLSX = require("xlsx");

const app = express();
const PORT = process.env.PORT || 3000;
const EXCEL_FILE_CANDIDATES = ["供应明细.xlsx", "明细.xlsx"];
const TARGET_SHEET_INDEX = 1;
const TABLE_NAME = process.env.MYSQL_TABLE || "excel_items";

function resolveExcelFilePath() {
  for (const name of EXCEL_FILE_CANDIDATES) {
    const filePath = path.join(process.cwd(), name);
    if (fs.existsSync(filePath)) {
      return filePath;
    }
  }

  return path.join(process.cwd(), EXCEL_FILE_CANDIDATES[0]);
}

function randomDateInMarch2026() {
  const day = Math.floor(Math.random() * 31) + 1;
  return `2026-03-${String(day).padStart(2, "0")}`;
}

function formatDateValue(dateObj) {
  const y = dateObj.getFullYear();
  const m = String(dateObj.getMonth() + 1).padStart(2, "0");
  const d = String(dateObj.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

function normalizePurchaseDate(value) {
  if (value === undefined || value === null || value === "") {
    return randomDateInMarch2026();
  }

  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed && parsed.y && parsed.m && parsed.d) {
      return `${parsed.y}-${String(parsed.m).padStart(2, "0")}-${String(parsed.d).padStart(2, "0")}`;
    }
    return randomDateInMarch2026();
  }

  const text = String(value).trim();
  if (!text) {
    return randomDateInMarch2026();
  }

  const normalizedText = text.replace(/[./年]/g, "-").replace(/[月]/g, "-").replace(/[日]/g, "");
  const parsedTime = Date.parse(normalizedText);
  if (!Number.isNaN(parsedTime)) {
    return formatDateValue(new Date(parsedTime));
  }

  return randomDateInMarch2026();
}

function normalizeRows(rawRows) {
  if (!Array.isArray(rawRows)) return [];

  return rawRows.map((row) => {
    const safeRow = { ...row };

    if (safeRow["商品名称"] !== undefined && safeRow["生产设备/零件"] === undefined) {
      safeRow["生产设备/零件"] = safeRow["商品名称"];
      delete safeRow["商品名称"];
    }

    safeRow["采购日期"] = normalizePurchaseDate(safeRow["采购日期"]);

    if (safeRow["数量"] === undefined || safeRow["数量"] === null || safeRow["数量"] === "") {
      safeRow["数量"] = 0;
    }

    const orderedRow = {};
    const keyOrder = ["序号", "供应商", "生产设备/零件", "采购日期", "数量"];

    for (const key of keyOrder) {
      if (safeRow[key] !== undefined) {
        orderedRow[key] = safeRow[key];
      }
    }

    for (const key of Object.keys(safeRow)) {
      if (orderedRow[key] === undefined) {
        orderedRow[key] = safeRow[key];
      }
    }

    return orderedRow;
  });
}

function readExcelData() {
  const excelFilePath = resolveExcelFilePath();

  if (!fs.existsSync(excelFilePath)) {
    return {
      ok: false,
      message: `未找到 Excel 文件：${excelFilePath}`,
      sheetName: "",
      headers: [],
      rows: [],
      generatedAt: new Date().toISOString(),
    };
  }

  const workbook = XLSX.readFile(excelFilePath);
  const sheetName = workbook.SheetNames[TARGET_SHEET_INDEX] || workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const rawRows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
  const rows = normalizeRows(rawRows);
  const headers = rows.length ? Object.keys(rows[0]) : [];

  return {
    ok: true,
    message: "success",
    sheetName,
    headers,
    rows,
    generatedAt: new Date().toISOString(),
  };
}

function escapeSqlValue(value) {
  if (value === null || value === undefined || value === "") return "NULL";
  if (typeof value === "number") return String(value);

  const text = String(value).replace(/\\/g, "\\\\").replace(/'/g, "\\'");
  return `'${text}'`;
}

function buildCreateTableSQL(headers, tableName) {
  const cols = headers.map((h) => {
    if (h === "序号") return `\`${h}\` INT`;
    if (String(h).toLowerCase().includes("数量")) return `\`${h}\` INT`;
    if (String(h).includes("日期")) return `\`${h}\` DATE`;
    return `\`${h}\` TEXT`;
  });

  return [
    `CREATE TABLE IF NOT EXISTS \`${tableName}\` (`,
    `  id BIGINT PRIMARY KEY AUTO_INCREMENT,`,
    `  ${cols.join(",\n  ")}`,
    `);`,
  ].join("\n");
}

function buildInsertSQL(rows, headers, tableName) {
  if (!rows.length || !headers.length) return "";

  const columns = headers.map((h) => `\`${h}\``).join(", ");
  const valuesSql = rows
    .map((row) => {
      const values = headers.map((h) => escapeSqlValue(row[h]));
      return `(${values.join(", ")})`;
    })
    .join(",\n");

  return `INSERT INTO \`${tableName}\` (${columns}) VALUES\n${valuesSql};`;
}

function toCsv(rows, headers) {
  const csvEscape = (value) => {
    if (value === null || value === undefined) return "";
    const text = String(value);
    if (/[",\n]/.test(text)) {
      return `"${text.replace(/"/g, '""')}"`;
    }
    return text;
  };

  const lines = [headers.join(",")];
  for (const row of rows) {
    lines.push(headers.map((h) => csvEscape(row[h])).join(","));
  }

  return lines.join("\n");
}

app.use(express.static(path.join(process.cwd(), "public")));

app.get("/api/data", (req, res) => {
  const payload = readExcelData();
  res.json(payload);
});

app.get("/api/data.csv", (req, res) => {
  const payload = readExcelData();

  if (!payload.ok) {
    return res.status(404).send(payload.message);
  }

  const csv = toCsv(payload.rows, payload.headers);
  res.setHeader("Content-Type", "text/csv; charset=utf-8");
  res.setHeader("Content-Disposition", "attachment; filename=excel_data.csv");
  return res.send(`\uFEFF${csv}`);
});

app.get("/api/mysql.sql", (req, res) => {
  const payload = readExcelData();

  if (!payload.ok) {
    return res.status(404).send(payload.message);
  }

  const createTableSQL = buildCreateTableSQL(payload.headers, TABLE_NAME);
  const insertSql = buildInsertSQL(payload.rows, payload.headers, TABLE_NAME);
  const sql = `${createTableSQL}\n\n${insertSql}`;

  res.setHeader("Content-Type", "text/plain; charset=utf-8");
  return res.send(sql);
});

app.get("*", (req, res) => {
  res.sendFile(path.join(process.cwd(), "public", "index.html"));
});

app.listen(PORT, () => {
  console.log(`Server is running at http://localhost:${PORT}`);
  console.log(`Reading Excel from: ${resolveExcelFilePath()}`);
});
