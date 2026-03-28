const express = require("express");
const path = require("path");
const fs = require("fs");
const XLSX = require("xlsx");

const app = express();
const PORT = process.env.PORT || 3000;
const EXCEL_FILE_PATH = path.join(process.cwd(), "明细.xlsx");
const TABLE_NAME = process.env.MYSQL_TABLE || "excel_items";

function detectQuantityKey(row) {
  const keys = Object.keys(row || {});
  const keywordList = ["数量", "qty", "count", "数目", "件数", "数量(个)"];

  for (const key of keys) {
    const lowerKey = String(key).toLowerCase();
    if (keywordList.some((k) => lowerKey.includes(String(k).toLowerCase()))) {
      return key;
    }
  }

  return null;
}

function randomQuantity() {
  return Math.floor(Math.random() * 200) + 1;
}

function normalizeRows(rawRows) {
  if (!Array.isArray(rawRows)) return [];

  return rawRows.map((row, index) => {
    const safeRow = { ...row };
    const quantityKey = detectQuantityKey(safeRow);

    if (quantityKey) {
      safeRow[quantityKey] = randomQuantity();
    } else {
      safeRow["数量"] = randomQuantity();
    }

    return {
      序号: index + 1,
      ...safeRow,
    };
  });
}

function readExcelData() {
  if (!fs.existsSync(EXCEL_FILE_PATH)) {
    return {
      ok: false,
      message: `未找到 Excel 文件：${EXCEL_FILE_PATH}`,
      sheetName: "",
      headers: [],
      rows: [],
      generatedAt: new Date().toISOString(),
    };
  }

  const workbook = XLSX.readFile(EXCEL_FILE_PATH);
  const sheetName = workbook.SheetNames[0];
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
  console.log(`Reading Excel from: ${EXCEL_FILE_PATH}`);
});
