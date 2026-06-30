# 企业授信报告系统 v2.0

## 架构说明

本系统**绕开 WorkBuddy**，直接通过后端代理调用企查查 QCC MCP API。

```
浏览器（前端）  ──→  Express 后端代理  ──→  企查查 QCC MCP (6大Server)
  index.html         server/index.js          agent.qcc.com/mcp/*/stream
```

### QCC MCP 6 大 Server / 12 个工具

| Server | 工具 | 用途 |
|--------|------|------|
| **company** | get_company_registration_info | 工商信息 |
| **company** | get_financial_data | 财务数据 |
| **company** | get_annual_reports | 企业年报 |
| **company** | get_shareholder_info | 股东信息 |
| **company** | get_external_investments | 对外投资 |
| **risk** | get_company_risk_scan | 风险扫描 |
| **ipr** | get_software_copyright_info | 软件著作权 |
| **ipr** | get_patent_info | 专利信息 |
| **operation** | get_bidding_info | 招投标信息 |
| **operation** | get_qualifications | 资质证书 |
| **executive** | get_executive_info | 高管信息 |
| **history** | get_company_change_history | 历史沿革 |

## 快速开始

### 前置条件
- Node.js >= 18
- 网络可访问 `agent.qcc.com`

### 安装 & 启动

```bash
cd credit-report
npm install
npm start
```

启动后访问: http://localhost:3000

### 使用流程
1. 输入银行支行/分行名称
2. 输入要查询的企业名称（完整全称）
3. 点击「生成授信报告」
4. 系统自动并行调用 12 个 QCC MCP 工具
5. 生成结构化授信报告，支持复制和 PDF 下载

## 文件结构

```
credit-report/
├── index.html          # 前端页面（含完整 UI 和报告生成逻辑）
├── package.json        # Node.js 项目配置
├── server/
│   └── index.js        # Express 后端代理（调用 QCC MCP）
├── server.py           # Python 备选后端（无需 Node.js）
└── README.md           # 本文件
```

## Python 备选方案

如果没有 Node.js，也可以用 Python 3 启动：

```bash
python server.py
```

Python 版本功能完全相同，无需安装任何第三方依赖。
