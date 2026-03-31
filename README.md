# 华瑞新材供应链数据管理系统

基于 Node.js 的供应链数据管理平台，主要功能包括：

- 供应商信息管理
- 设备零件采购记录查询
- 多维度数据视图切换
- 数据导出功能（支持 CSV、SQL 格式）

## 安装与启动

```bash
npm install
npm run start
```

默认端口为 `3000`，可通过环境变量覆盖：

```bash
PORT=8080 npm run start
```

## API 接口

- 主页：`/`
- 数据列表：`/api/sheets`
- 数据查询：`/api/data?sheet=供应数据`
- CSV 导出：`/api/data.csv?sheet=供应商表`
- SQL 导出：`/api/mysql.sql?sheet=设备零件表`

JSON 返回结构示例：

```json
{
  "ok": true,
  "sheetName": "供应数据",
  "headers": ["序号", "供应商", "生产设备/零件", "采购日期", "数量"],
  "rows": [
    { "序号": 1, "供应商": "XX公司", "生产设备/零件": "设备A", "采购日期": "2026-03-15", "数量": 100 }
  ],
  "generatedAt": "2026-03-31T00:00:00.000Z"
}
```

## 数据导出

系统支持两种方式导出数据：

### SQL 导出
访问 `/api/mysql.sql` 获取完整的建表和插入语句。

### API 集成
通过 `/api/data` 接口获取 JSON 格式数据，可集成到其他系统。

## 系统说明

- 支持多个数据视图切换（供应数据、供应商表、设备零件表）
- 数据自动格式化和规范化处理
- 响应式设计，支持移动端访问

## 部署建议

- 使用 PM2 进行进程管理：`pm2 start src/server.js --name supply-chain`
- 配置 Nginx 反向代理
- 建议启用 HTTPS
