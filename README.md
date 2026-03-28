# Excel 测试网页（可爬取 + 可直接入 MySQL）

这是一个可部署到服务器的 Node.js 网页项目，特性如下：

- 页面数据读取自根目录 `明细.xlsx` 的第一个工作表
- 自动将 `数量` 列替换为随机数（不使用 Excel 原始数量）
- 网页美观展示表格，支持移动端
- 提供稳定接口，便于爬虫直接抓取并入库 MySQL

## 1. 目录结构

```text
Web/
  ├─ public/
  │   ├─ index.html
  │   ├─ styles.css
  │   └─ app.js
  ├─ src/
  │   └─ server.js
  ├─ package.json
  └─ 明细.xlsx   <-- 你需要放在项目根目录
```

## 2. 启动方式

```bash
npm install
npm run start
```

默认端口为 `3000`，可通过环境变量覆盖：

```bash
PORT=8080 npm run start
```

## 3. 可爬取接口

- 页面：`/`
- JSON：`/api/data`
- CSV：`/api/data.csv`
- MySQL SQL：`/api/mysql.sql`

推荐爬虫抓取 JSON 接口 `/api/data`。

JSON 返回结构示例：

```json
{
  "ok": true,
  "sheetName": "Sheet1",
  "headers": ["序号", "商品", "数量", "价格"],
  "rows": [
    { "序号": 1, "商品": "A", "数量": 137, "价格": 9.9 }
  ],
  "generatedAt": "2026-03-28T00:00:00.000Z"
}
```

## 4. 直接入 MySQL 的两种方式

### 方式 A：下载 SQL 文件直接执行

访问 `/api/mysql.sql`，拿到 `CREATE TABLE + INSERT` 语句后执行即可。

### 方式 B：爬虫抓 JSON，再写入 MySQL

伪代码流程：

1. GET `/api/data`
2. 读取 `headers` 和 `rows`
3. 将 `rows` 批量插入目标表

## 5. 注意事项

- 默认读取第一个工作表
- 如果 Excel 中没有“数量”列，系统会自动新增 `数量` 随机列
- 每次请求接口时都会重新生成随机数量值
- 若你希望“每次启动固定随机值”，我可以再给你改成启动时生成并缓存

## 6. 服务器部署建议

- 使用 `pm2` 守护进程：`pm2 start src/server.js --name excel-web`
- 通过 Nginx 反向代理到 Node 端口
- 如果你有域名和 HTTPS，我可以继续帮你补 Nginx 完整配置
