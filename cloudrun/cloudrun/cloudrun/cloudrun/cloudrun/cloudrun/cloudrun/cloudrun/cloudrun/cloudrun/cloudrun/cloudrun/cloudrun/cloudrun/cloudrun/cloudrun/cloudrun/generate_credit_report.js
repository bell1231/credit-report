/**
 * ICBC 普惠金融中小企业授信报告 — Docx 生成器
 * 
 * 用法: node generate_credit_report.js <企业名称> [输出文件名]
 * 
 * 数据来源: 企查查 QCC MCP 6大 Server（12个SKILL场景）
 * 报告模板: 严格遵循 ICBC_Credit_Report_Feature_Inventory.docx 的9章节结构
 */

const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, LevelFormat, BorderStyle, WidthType,
  ShadingType, HeadingLevel, PageNumber, PageBreak
} = require('docx');

// ============ QCC MCP 配置 ============
const QCC_SERVERS = {
  company:   { url: 'https://agent.qcc.com/mcp/company/stream',   token: 'Bearer MmDFP6lIn3akremGqfRhudPo0NON0LTrz92cBzCHtI3eA7IZ' },
  risk:      { url: 'https://agent.qcc.com/mcp/risk/stream',      token: 'Bearer MmDFP6lIn3akremGqfRhudPo0NON0LTrz92cBzCHtI3eA7IZ' },
  ipr:       { url: 'https://agent.qcc.com/mcp/ipr/stream',       token: 'Bearer MmDFP6lIn3akremGqfRhudPo0NON0LTrz92cBzCHtI3eA7IZ' },
  operation: { url: 'https://agent.qcc.com/mcp/operation/stream', token: 'Bearer MmDFP6lIn3akremGqfRhudPo0NON0LTrz92cBzCHtI3eA7IZ' },
  executive: { url: 'https://agent.qcc.com/mcp/executive/stream', token: 'Bearer MmDFP6lIn3akremGqfRhudPo0NON0LTrz92cBzCHtI3eA7IZ' },
  history:   { url: 'https://agent.qcc.com/mcp/history/stream',   token: 'Bearer MmDFP6lIn3akremGqfRhudPo0NON0LTrz92cBzCHtI3eA7IZ' }
};

// ============ QCC MCP 调用 ============
async function callQccMCP(server, toolName, args) {
  const config = QCC_SERVERS[server];
  if (!config) throw new Error(`未知 QCC 服务: ${server}`);
  const response = await fetch(config.url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'Authorization': config.token },
    body: JSON.stringify({
      jsonrpc: '2.0', id: Date.now(),
      method: 'tools/call',
      params: { name: toolName, arguments: args }
    }),
    signal: AbortSignal.timeout(30000)
  });
  if (!response.ok) throw new Error(`QCC ${server} HTTP ${response.status}`);
  const text = await response.text();
  let lastResult = null;
  for (const line of text.split('\n')) {
    if (line.startsWith('data: ')) {
      try {
        const sse = JSON.parse(line.substring(6));
        if (sse.error) throw new Error(`QCC ${server}/${toolName}: ${JSON.stringify(sse.error)}`);
        const r = sse.result;
        if (r?.content?.[0]?.text) {
          try { lastResult = JSON.parse(r.content[0].text); } catch { lastResult = r.content[0].text; }
        } else if (r) { lastResult = r; }
      } catch (e) { if (e.message?.includes('QCC ')) throw e; }
    }
  }
  return lastResult;
}

// ============ 数据收集 ============
async function collectAllData(companyName) {
  const data = { errors: [] };
  const a = () => ({ searchKey: companyName });
  
  console.log(`\n📡 开始采集企业数据: ${companyName}\n`);
  
  // 阶段1: 串行获取核心工商数据（避免限流）
  console.log('=== 阶段1: 核心工商数据（串行）===');
  
  const basicCalls = [
    { name: '工商注册信息', server: 'company', tool: 'get_company_registration_info', set: 'company' },
    { name: '财务数据', server: 'company', tool: 'get_financial_data', set: 'financial' },
    { name: '企业年报', server: 'company', tool: 'get_annual_reports', set: 'annualReports' },
    { name: '股东信息', server: 'company', tool: 'get_shareholder_info', set: 'shareholders' },
    { name: '对外投资', server: 'company', tool: 'get_external_investments', set: 'investments' },
  ];
  
  for (const c of basicCalls) {
    try {
      data[c.set] = await callQccMCP(c.server, c.tool, a());
      console.log(`  ✅ ${c.name}`);
    } catch (e) {
      console.log(`  ❌ ${c.name}: ${e.message}`);
      data.errors.push(`${c.name}: ${e.message}`);
    }
    await new Promise(r => setTimeout(r, 600)); // 延迟避免限流
  }
  
  // 阶段1.5: 工商变更记录
  console.log('\n=== 阶段1.5: 工商变更记录 ===');
  try {
    const uscc = data.company?.['统一社会信用代码'] || data.company?.KeyNo;
    data.changeRecords = await callQccMCP('company', 'get_change_records', { searchKey: uscc || companyName });
    data.history = data.changeRecords;
    console.log('  ✅ 工商变更记录');
  } catch (e) {
    console.log(`  ❌ 工商变更记录: ${e.message}`);
    data.errors.push(`工商变更记录: ${e.message}`);
  }
  
  // 阶段2: 并发获取其余数据
  console.log('\n=== 阶段2: 风险/司法/高管/经营/IP/历史（并发）===');
  
  const concurrentCalls = [
    // 风险扫描
    { name: '风险扫描', server: 'risk', tool: 'get_company_risk_scan', set: 'risk' },
    { name: '失信信息', server: 'risk', tool: 'get_dishonest_info', set: 'riskDishonest' },
    { name: '被执行信息', server: 'risk', tool: 'get_judgment_debtor_info', set: 'riskJudgmentDebtor' },
    { name: '限高消费', server: 'risk', tool: 'get_high_consumption_restriction', set: 'riskHighConsumption' },
    { name: '股权冻结', server: 'risk', tool: 'get_equity_freeze', set: 'riskEquityFreeze' },
    { name: '股权出质', server: 'risk', tool: 'get_equity_pledge_info', set: 'riskEquityPledge' },
    { name: '经营异常', server: 'risk', tool: 'get_business_exception', set: 'riskBizException' },
    { name: '行政处罚', server: 'risk', tool: 'get_administrative_penalty', set: 'riskAdminPenalty' },
    { name: '欠税公告', server: 'risk', tool: 'get_tax_arrears_notice', set: 'riskTaxArrears' },
    { name: '裁判文书', server: 'risk', tool: 'get_judicial_documents', set: 'riskJudicialDocs' },
    { name: '立案信息', server: 'risk', tool: 'get_case_filing_info', set: 'riskCaseFiling' },
    { name: '违约事项', server: 'risk', tool: 'get_default_info', set: 'riskDefaultInfo' },
    // 高管
    { name: '高管/主要人员', server: 'company', tool: 'get_key_personnel', set: 'keyPersonnel' },
    { name: '实际控制人', server: 'company', tool: 'get_actual_controller', set: 'actualController' },
    // 经营
    { name: '招投标信息', server: 'operation', tool: 'get_bidding_info', set: 'bidding' },
    { name: '资质证书', server: 'operation', tool: 'get_qualifications', set: 'qualifications' },
    { name: '荣誉信息', server: 'operation', tool: 'get_honor_info', set: 'honor' },
    { name: '招聘信息', server: 'operation', tool: 'get_recruitment_info', set: 'recruitment' },
    { name: '信用评估', server: 'operation', tool: 'get_credit_evaluation', set: 'creditEval' },
    { name: '进出口信用', server: 'operation', tool: 'get_import_export_credit', set: 'importExportCredit' },
    // IP
    { name: '软件著作权', server: 'ipr', tool: 'get_software_copyright_info', set: 'software' },
    { name: '专利信息', server: 'ipr', tool: 'get_patent_info', set: 'patents' },
    // UBO
    { name: '受益所有人', server: 'company', tool: 'get_beneficial_owners', set: 'beneficialOwners' },
    { name: '企业档案', server: 'company', tool: 'get_company_profile', set: 'companyProfile' },
    { name: '分支机构', server: 'company', tool: 'get_branches', set: 'branches' },
    // 历史（可选）
    { name: '历史法代', server: 'history', tool: 'get_historical_legal_rep', set: 'histLegalRep' },
    { name: '历史股东', server: 'history', tool: 'get_historical_shareholders', set: 'histShareholders' },
  ];
  
  const results = await Promise.all(concurrentCalls.map(c =>
    callQccMCP(c.server, c.tool, a())
      .then(r => ({ ...c, result: r, ok: true }))
      .catch(e => ({ ...c, error: e, ok: false }))
  ));
  
  for (const r of results) {
    if (r.ok) {
      data[r.set] = r.result;
      const hasData = r.result && (typeof r.result === 'object' ? Object.keys(r.result).length > 0 : true);
      console.log(`  ${hasData ? '✅' : '⚠️'} ${r.name}: ${hasData ? '已获取' : '无记录'}`);
    } else {
      console.log(`  ❌ ${r.name}: ${r.error?.message || String(r.error)}`);
    }
  }
  
  console.log(`\n📊 数据采集完成，共 ${Object.keys(data).length - 1} 个数据项（含 ${data.errors.length} 个错误）\n`);
  return data;
}

// ============ 辅助函数 ============
function safeStr(val, ...path) {
  try {
    if (!val) return '';
    let v = val;
    for (const key of path) {
      if (v[key] !== undefined) v = v[key];
      else return '';
    }
    return String(v);
  } catch { return ''; }
}

function safeArr(val, ...path) {
  try {
    if (!val) return [];
    if (Array.isArray(val)) return val;
    let v = val;
    for (const key of path) {
      if (v[key] !== undefined) v = v[key];
      else return [];
    }
    return Array.isArray(v) ? v : [];
  } catch { return []; }
}

function safeObj(val, ...path) {
  try {
    if (!val) return {};
    let v = val;
    for (const key of path) {
      if (v[key] !== undefined) v = v[key];
      else return {};
    }
    return typeof v === 'object' && !Array.isArray(v) ? v : {};
  } catch { return {}; }
}

function fmtMoney(v) {
  const n = parseFloat(String(v).replace(/[^0-9.\-]/g, ''));
  if (isNaN(n)) return '——';
  if (Math.abs(n) >= 1e8) return (n / 1e8).toFixed(2) + '亿';
  if (Math.abs(n) >= 1e4) return (n / 1e4).toFixed(2) + '万';
  return n.toLocaleString() + '元';
}

function fmtPct(v) {
  const n = parseFloat(String(v).replace(/[^0-9.\-]/g, ''));
  return isNaN(n) ? '——' : n.toFixed(1) + '%';
}

function orDash(v) { return v || '——'; }

// ============ Docx 样式常量 ============
const FONT = '仿宋';
const FONT_TITLE = '黑体';
const PAGE_WIDTH = 11906; // A4
const PAGE_HEIGHT = 16838;

const BORDER_LIGHT = { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' };
const BORDERS = { top: BORDER_LIGHT, bottom: BORDER_LIGHT, left: BORDER_LIGHT, right: BORDER_LIGHT };
const NO_BORDER = { style: BorderStyle.NONE, size: 0 };
const NO_BORDERS = { top: NO_BORDER, bottom: NO_BORDER, left: NO_BORDER, right: NO_BORDER };

const MARGIN_TOP = 1440;   // 2.54cm
const MARGIN_BOTTOM = 1440;
const MARGIN_LEFT = 1800;  // 3.17cm
const MARGIN_RIGHT = 1440;

const CONTENT_WIDTH = PAGE_WIDTH - MARGIN_LEFT - MARGIN_RIGHT; // 8666

function p(text, opts = {}) {
  const { bold, size, color, font, alignment, spacing, indent, heading } = opts;
  const runs = [];
  if (typeof text === 'string') {
    runs.push(new TextRun({
      text, bold: bold || false, size: size || 28, // 14pt = 28 half-pt
      color: color || '000000',
      font: font || FONT,
    }));
  } else if (Array.isArray(text)) {
    for (const t of text) {
      if (typeof t === 'string') {
        runs.push(new TextRun({ text: t, size: size || 28, font: font || FONT }));
      } else {
        runs.push(new TextRun({
          text: t.text || '', bold: t.bold || false, size: t.size || size || 28,
          color: t.color || color || '000000', font: t.font || font || FONT,
        }));
      }
    }
  }
  
  const paraOpts = {
    children: runs,
    spacing: spacing || { before: 60, after: 60, line: 360 }, // 1.5倍行距
  };
  if (alignment) paraOpts.alignment = alignment;
  if (indent) paraOpts.indent = indent;
  if (heading) paraOpts.heading = heading;
  return new Paragraph(paraOpts);
}

function heading1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    alignment: AlignmentType.CENTER,
    spacing: { before: 360, after: 240, line: 360 },
    children: [new TextRun({ text, bold: true, size: 44, font: FONT_TITLE })], // 二号22pt=44
  });
}

function heading2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 240, after: 180, line: 360 },
    children: [new TextRun({ text, bold: true, size: 32, font: FONT_TITLE })], // 三号16pt=32
  });
}

function heading3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 180, after: 120, line: 360 },
    children: [new TextRun({ text, bold: true, size: 28, font: FONT_TITLE })], // 四号14pt=28
  });
}

function makeTable(headers, rows, colWidths) {
  const totalWidth = colWidths ? colWidths.reduce((a, b) => a + b, 0) : CONTENT_WIDTH;
  const cw = colWidths || headers.map(() => Math.floor(CONTENT_WIDTH / headers.length));
  
  const headerRow = new TableRow({
    children: headers.map((h, i) => new TableCell({
      borders: BORDERS,
      width: { size: cw[i], type: WidthType.DXA },
      shading: { fill: '1A2D4E', type: ShadingType.CLEAR },
      margins: { top: 60, bottom: 60, left: 80, right: 80 },
      children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: h, bold: true, size: 20, color: 'FFFFFF', font: FONT_TITLE })]
      })]
    }))
  });
  
  const dataRows = rows.map((row, ri) => new TableRow({
    children: row.map((cell, ci) => new TableCell({
      borders: BORDERS,
      width: { size: cw[ci], type: WidthType.DXA },
      shading: ri % 2 === 0 ? { fill: 'F7F9FC', type: ShadingType.CLEAR } : undefined,
      margins: { top: 40, bottom: 40, left: 80, right: 80 },
      children: [new Paragraph({
        children: [new TextRun({ text: String(cell || '——'), size: 20, font: FONT })]
      })]
    }))
  }));
  
  return new Table({
    width: { size: totalWidth, type: WidthType.DXA },
    columnWidths: cw,
    rows: [headerRow, ...dataRows],
  });
}

function infoTable(pairs, colWidths) {
  const cw = colWidths || [1800, 2500, 1800, 2566];
  const rows = [];
  for (let i = 0; i < pairs.length; i += 2) {
    const row = [];
    row.push(pairs[i]?.[0] || '', pairs[i]?.[1] || '');
    if (i + 1 < pairs.length) {
      row.push(pairs[i+1]?.[0] || '', pairs[i+1]?.[1] || '');
    } else {
      row.push('', '');
    }
    rows.push(row);
  }
  
  const headerRow = new TableRow({
    children: (colWidths ? colWidths : [1800, 2500, 1800, 2566]).map((w, i) => new TableCell({
      borders: BORDERS,
      width: { size: w, type: WidthType.DXA },
      shading: { fill: 'EEF2F7', type: ShadingType.CLEAR },
      margins: { top: 40, bottom: 40, left: 80, right: 80 },
      children: [new Paragraph({
        children: [new TextRun({ text: i % 2 === 0 ? '项目' : '内容', bold: true, size: 20, font: FONT_TITLE, color: '1A2D4E' })]
      })]
    }))
  });
  
  const dataRows = rows.map((row, ri) => new TableRow({
    children: row.map((cell, ci) => new TableCell({
      borders: BORDERS,
      width: { size: (colWidths || [1800, 2500, 1800, 2566])[ci], type: WidthType.DXA },
      shading: ri % 2 === 0 ? { fill: 'F7F9FC', type: ShadingType.CLEAR } : undefined,
      margins: { top: 40, bottom: 40, left: 80, right: 80 },
      children: [new Paragraph({
        children: [new TextRun({
          text: String(cell || '——'),
          size: 20,
          font: FONT,
          bold: ci % 2 === 0,
          color: ci % 2 === 0 ? '1A2D4E' : '000000'
        })]
      })]
    }))
  }));
  
  return new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: colWidths || [1800, 2500, 1800, 2566],
    rows: [headerRow, ...dataRows],
  });
}

// ============ 报告生成 ============
async function buildReport(data, companyName, branch = 'XX支行') {
  const now = new Date();
  const dateStr = `${now.getFullYear()}年${now.getMonth()+1}月${now.getDate()}日`;
  const dateShort = now.toISOString().slice(0, 10);
  
  // ===== 数据提取 =====
  const comp = data.company || {};
  const companyFullName = safeStr(comp, '企业名称') || companyName;
  const creditCode = safeStr(comp, '统一社会信用代码');
  const legalPerson = safeStr(comp, '法定代表人');
  const regCapital = safeStr(comp, '注册资本');
  const estDate = safeStr(comp, '成立日期');
  const companyType = safeStr(comp, '企业类型');
  const regAddr = safeStr(comp, '注册地址');
  const businessScope = safeStr(comp, '经营范围');
  const status = safeStr(comp, '登记状态');
  const industry = safeStr(comp, '国标行业');
  const employees = safeStr(comp, '参保人数');
  const region = safeStr(comp, '所属地区');
  const registeredCapitalPaid = safeStr(comp, '实缴资本');
  const bizTerm = safeStr(comp, '营业期限');
  
  // 财务数据
  const finArr = safeArr(data.financial, '财务数据信息');
  let finRevenue = '', finNetProfit = '', finTotalAssets = '', finTotalLiabilities = '';
  let finEquity = '', finAssetRatio = '', finNetMargin = '', finRevenueGrowth = '';
  let finROE = '', financePeriod = '';
  if (finArr.length > 0) {
    const latest = finArr[0];
    financePeriod = latest['报告期'] || '';
    const indicators = latest['指标详情'] || {};
    const bs = safeObj(indicators, '财务报表', '资产负债表');
    const is = safeObj(indicators, '财务报表', '利润表');
    const analysis = safeObj(indicators, '分析数据');
    finRevenue = safeStr(is, '营业总收入');
    finNetProfit = safeStr(is, '净利润');
    finTotalAssets = safeStr(bs, '资产合计');
    finTotalLiabilities = safeStr(bs, '负债合计');
    finEquity = safeStr(bs, '所有者权益总计');
    finAssetRatio = safeStr(analysis, '偿还能力', '资产负债率');
    finNetMargin = safeStr(analysis, '盈利能力', '净利率');
    finRevenueGrowth = safeStr(analysis, '成长能力', '营业收入同比');
    finROE = safeStr(analysis, '盈利能力', '净资产收益率');
  }
  
  // 年报数据
  const arArr = safeArr(data.annualReports, '企业年报信息');
  let arRevenue = '', arNetProfit = '', arTotalAssets = '', arTotalLiabilities = '', arEquity = '', arYear = '';
  if (arArr.length > 0) {
    arYear = arArr[0]['年报年度'] || '';
    const assets = arArr[0]['企业资产状况信息'] || {};
    const getV = (k) => assets[k] && assets[k] !== '企业选择不公示' ? assets[k] : '';
    arRevenue = getV('营业总收入'); arNetProfit = getV('净利润');
    arTotalAssets = getV('资产总额'); arTotalLiabilities = getV('负债总额');
    arEquity = getV('所有者权益合计');
  }
  
  const displayRevenue = finRevenue || arRevenue;
  const displayNetProfit = finNetProfit || arNetProfit;
  const displayAssets = finTotalAssets || arTotalAssets;
  const displayLiabilities = finTotalLiabilities || arTotalLiabilities;
  const displayEquity = finEquity || arEquity;
  const displayAssetRatio = finAssetRatio;
  const displayNetMargin = finNetMargin;
  const displayRevenueGrowth = finRevenueGrowth;
  const displayROE = finROE;
  const hasFinance = !!(finRevenue || arRevenue);
  
  // 股东
  const shareholderList = safeArr(data.shareholders, '股东信息');
  
  // 高管
  const execRaw = data.keyPersonnel || {};
  let execList = [];
  for (const k of ['高管信息', '主要人员', '董监高', 'executives', 'personnel', 'key_personnel']) {
    const v = execRaw[k];
    if (Array.isArray(v) && v.length > 0) { execList = v; break; }
  }
  if (!execList.length && Array.isArray(execRaw)) execList = execRaw;
  
  // 风险
  const riskScanList = safeArr(data.risk, '风险因子扫描');
  const riskFactorCount = data.risk?.['有记录因子数'] || 0;
  const riskTotalCount = data.risk?.['有记录条目数'] || 0;
  
  function getRiskCount(name) {
    if (!data.risk?.['风险因子扫描']) return 0;
    const factor = data.risk['风险因子扫描'].find(f => f['风险因子'] === name);
    return factor ? (factor['条目数'] || 0) : 0;
  }
  
  const dishonestCount = getRiskCount('失信被执行人');
  const highConsumptionCount = getRiskCount('限制高消费');
  const bankruptcyCount = getRiskCount('破产重整');
  const equityFreezeCount = getRiskCount('股权冻结');
  const equityPledgeCount = getRiskCount('股权出质');
  const adminPenaltyCount = getRiskCount('行政处罚');
  const bizExceptionCount = getRiskCount('经营异常');
  const taxArrearsCount = getRiskCount('欠税公告');
  const judgmentDebtorCount = getRiskCount('被执行人');
  const judicialDocsCount = getRiskCount('裁判文书');
  const caseFilingCount = getRiskCount('立案信息');
  
  // 变更记录
  const changeList = (() => {
    const raw = data.changeRecords || data.history;
    if (!raw) return [];
    if (Array.isArray(raw)) return raw;
    for (const k of ['变更记录信息', '变更记录', 'change_records', 'records', 'items', 'data']) {
      if (Array.isArray(raw[k])) return raw[k];
    }
    for (const k of Object.keys(raw)) {
      if (Array.isArray(raw[k])) return raw[k];
    }
    return [];
  })();
  
  // 知识产权
  const patentRaw = data.patents;
  const swRaw = data.software;
  let patentList = [];
  if (Array.isArray(patentRaw)) patentList = patentRaw;
  else if (patentRaw && typeof patentRaw === 'object') {
    for (const k of ['专利信息', 'patents', 'records', 'items']) {
      if (Array.isArray(patentRaw[k])) { patentList = patentRaw[k]; break; }
    }
  }
  let swList = [];
  if (Array.isArray(swRaw)) swList = swRaw;
  else if (swRaw && typeof swRaw === 'object') {
    for (const k of ['软件著作权', 'software', 'records', 'items']) {
      if (Array.isArray(swRaw[k])) { swList = swRaw[k]; break; }
    }
  }
  
  // 经营
  const biddingList = safeArr(data.bidding, '招投标信息');
  const honorList = safeArr(data.honor, '荣誉信息');
  const qualificationList = safeArr(data.qualifications, '资质证书');
  
  // 实际控制人
  const ac = data.actualController;
  const acName = safeStr(ac, '实际控制人') || safeStr(ac, '姓名');
  
  // 对外投资
  const invList = safeArr(data.investments, '对外投资');
  
  // 分支机构
  const branchList = safeArr(data.branches, '分支机构');
  
  // ===== 评分 =====
  let score = 50;
  const fatalRisks = [dishonestCount, judgmentDebtorCount, highConsumptionCount, bankruptcyCount, equityFreezeCount];
  const fatalTotal = fatalRisks.reduce((a, b) => a + b, 0);
  if (fatalTotal > 0) score -= fatalTotal * 15;
  else score += 25;
  
  if (finAssetRatio) {
    const ar = parseFloat(finAssetRatio);
    if (!isNaN(ar)) {
      if (ar < 40) score += 20; else if (ar < 60) score += 14; else if (ar < 75) score += 8; else score += 2;
    }
  }
  if (finNetMargin) {
    const nm = parseFloat(finNetMargin);
    if (!isNaN(nm)) {
      if (nm > 20) score += 15; else if (nm > 10) score += 10; else score += 5;
    }
  }
  if (estDate) {
    const age = now.getFullYear() - parseInt(estDate.substring(0, 4));
    if (age >= 20) score += 12; else if (age >= 10) score += 9; else if (age >= 5) score += 6; else score += 3;
  }
  const emp = parseInt(employees);
  if (!isNaN(emp)) {
    if (emp >= 100) score += 7; else if (emp > 0) score += 4;
  }
  if (patentList.length + swList.length > 10) score += 5;
  else if (patentList.length + swList.length > 0) score += 3;
  if (shareholderList.length > 0) score += 5;
  
  score = Math.max(0, Math.min(100, Math.round(score)));
  
  let grade, ratingLabel;
  if (score >= 95) { grade = 'AAA'; ratingLabel = '极优'; }
  else if (score >= 88) { grade = 'AA+'; ratingLabel = '优良'; }
  else if (score >= 80) { grade = 'AA'; ratingLabel = '良好'; }
  else if (score >= 72) { grade = 'A'; ratingLabel = '良好'; }
  else if (score >= 65) { grade = 'BBB'; ratingLabel = '一般'; }
  else if (score >= 55) { grade = 'BB'; ratingLabel = '关注'; }
  else { grade = 'B'; ratingLabel = '谨慎'; }
  
  // ===== 构建文档 =====
  const children = [];
  
  // ====== 封面 ======
  children.push(
    new Paragraph({ spacing: { before: 3600 }, children: [] }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 240 },
      children: [new TextRun({ text: '中国工商银行', bold: true, size: 52, font: FONT_TITLE, color: 'C41E1E' })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 480 },
      children: [new TextRun({ text: '普惠金融业务（中小企业主客群）授信调查报告', bold: true, size: 36, font: FONT_TITLE, color: '1A2D4E' })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 120 },
      children: [new TextRun({ text: `—— ${companyFullName} ——`, size: 32, font: FONT_TITLE, color: '1D6FA4' })]
    }),
    new Paragraph({ spacing: { before: 600 }, children: [] }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 120 },
      children: [new TextRun({ text: `统一社会信用代码：${creditCode || '——'}`, size: 24, font: FONT, color: '555555' })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 120 },
      children: [new TextRun({ text: `申请机构：${branch}`, size: 24, font: FONT, color: '555555' })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 120 },
      children: [new TextRun({ text: `报告日期：${dateStr}`, size: 24, font: FONT, color: '555555' })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 360 },
      children: [new TextRun({ text: '密级：机密·仅限内部审批', size: 20, font: FONT, color: 'C41E1E' })]
    }),
    new Paragraph({ children: [new PageBreak()] })
  );
  
  // ====== 前言 ======
  children.push(
    heading1('前  言'),
    p(`本报告依据中国工商银行总行普惠金融业务授信政策，采用"三品三表三流"（人品/产品/押品 + 电表/水表/税表 + 人流/物流/现金流）尽调方法论，以企业主个人/家庭与企业经营实体一体化分析为核心，对 ${companyFullName} 进行全维度授信评估。`, { indent: { firstLine: 560 } }),
    p(`报告数据来源：企查查 QCC 工商数据平台（查询日期：${dateShort}），覆盖工商注册信息、股权结构、财务数据、司法风险、经营状态等维度。`, { indent: { firstLine: 560 } }),
    p(`声明：本文档由系统基于企查查公开数据自动生成，仅供参考，不构成任何投资建议或授信决策依据。实际授信应以客户经理现场尽调、最新征信报告及我行内部审批意见为准。`, { indent: { firstLine: 560 }, color: 'C41E1E' }),
    new Paragraph({ children: [new PageBreak()] })
  );
  
  // ====== 第一章：借款人概况 ======
  children.push(heading1('一、借款人概况'));
  children.push(heading2('（一）经营实体基本情况'));
  
  children.push(infoTable([
    ['企业名称', companyFullName, '统一社会信用代码', creditCode || '——'],
    ['法定代表人', legalPerson || '——', '企业类型', companyType || '——'],
    ['注册资本', regCapital || '——', '实缴资本', registeredCapitalPaid || '——'],
    ['成立日期', estDate || '——', '营业期限', bizTerm || '——'],
    ['登记状态', status || '——', '国标行业', industry || '——'],
    ['注册地址', regAddr || '——', '所属地区', region || '——'],
    ['参保人数', employees || '——', '经营范围', businessScope ? businessScope.substring(0, 60) + '...' : '——'],
  ]));
  
  children.push(p(''));
  children.push(heading2('（二）股权结构与实际控制人'));
  
  if (shareholderList.length > 0) {
    const shRows = shareholderList.slice(0, 10).map(sh => [
      safeStr(sh, '股东名称') || safeStr(sh, '股东'),
      safeStr(sh, '持股比例') || safeStr(sh, '比例'),
      safeStr(sh, '认缴出资额') || safeStr(sh, '认缴'),
      safeStr(sh, '出资方式') || safeStr(sh, '方式') || '——'
    ]);
    children.push(makeTable(['股东名称', '持股比例', '认缴出资额', '出资方式'], shRows, [3200, 1400, 2000, 2066]));
  } else {
    children.push(p('（暂无股东数据）', { color: '888888' }));
  }
  
  if (acName) {
    children.push(p(''));
    children.push(p([
      { text: '实际控制人：', bold: true },
      { text: `${acName}，经核查工商登记及企查查股权穿透数据，该自然人为企业最终受益人，对企业经营决策具有实际控制力。` }
    ]));
  }
  
  // ====== 第二章：行业与市场分析 ======
  children.push(new Paragraph({ children: [new PageBreak()] }));
  children.push(heading1('二、行业与市场分析'));
  children.push(heading2('（一）行业概况'));
  
  if (industry) {
    children.push(p([
      { text: `企业所属行业为「${industry}」。`, bold: true },
      { text: `本报告采用精简聚焦的行业分析框架，重点评估行业景气度、政策环境及周边竞争格局。` }
    ]));
  } else {
    children.push(p('企业所属行业信息暂未获取，建议客户经理在实地尽调中补充行业背景分析。'));
  }
  
  children.push(p(`经企查查数据库查询，该企业主营业务为：${businessScope ? businessScope.substring(0, 120) : '——'}。`, { indent: { firstLine: 560 } }));
  
  if (qualificationList.length > 0) {
    children.push(heading2('（二）资质证照'));
    const qualRows = qualificationList.slice(0, 10).map(q => [
      safeStr(q, '资质名称') || safeStr(q, '证书名称'),
      safeStr(q, '证书编号') || safeStr(q, '编号'),
      safeStr(q, '有效期') || safeStr(q, '有效期限') || '——'
    ]);
    children.push(makeTable(['资质名称', '证书编号', '有效期'], qualRows, [3200, 3000, 2466]));
  }
  
  // ====== 第三章：财务状况分析 ======
  children.push(new Paragraph({ children: [new PageBreak()] }));
  children.push(heading1('三、财务状况分析'));
  children.push(heading2('（一）主要财务指标'));
  
  if (hasFinance) {
    const finPairs = [
      ['营业收入', fmtMoney(displayRevenue), '净利润', fmtMoney(displayNetProfit)],
      ['总资产', fmtMoney(displayAssets), '总负债', fmtMoney(displayLiabilities)],
      ['所有者权益', fmtMoney(displayEquity), '资产负债率', fmtPct(displayAssetRatio)],
      ['净利率', fmtPct(displayNetMargin), '净资产收益率', fmtPct(finROE)],
      ['营业收入同比', fmtPct(finRevenueGrowth), '数据来源', finArr.length > 0 ? 'QCC财务数据库' : '工商年报'],
    ];
    children.push(infoTable(finPairs));
    
    if (financePeriod) {
      children.push(p(`数据期间：${financePeriod}`, { size: 20, color: '888888' }));
    }
    
    // 财务分析文字
    children.push(p(''));
    children.push(heading2('（二）财务分析摘要'));
    
    if (displayAssetRatio) {
      const ar = parseFloat(displayAssetRatio);
      if (!isNaN(ar)) {
        const level = ar < 40 ? '较低，长期偿债能力充足' : ar < 60 ? '处于合理区间' : ar < 75 ? '偏高，需关注偿债压力' : '处于较高水平，存在一定偿债风险';
        children.push(p(`资产负债率分析：${fmtPct(displayAssetRatio)}，${level}。`, { indent: { firstLine: 560 } }));
      }
    }
    if (displayNetMargin) {
      const nm = parseFloat(displayNetMargin);
      if (!isNaN(nm)) {
        const level = nm > 20 ? '盈利能力极优' : nm > 10 ? '盈利能力良好' : nm > 5 ? '盈利能力一般' : '盈利能力偏弱';
        children.push(p(`盈利能力分析：净利率${fmtPct(displayNetMargin)}，${level}。`, { indent: { firstLine: 560 } }));
      }
    }
    if (displayRevenue) {
      children.push(p(`营收规模：${fmtMoney(displayRevenue)}，${finRevenueGrowth ? '同比增长' + fmtPct(finRevenueGrowth) : ''}。`, { indent: { firstLine: 560 } }));
    }
  } else {
    children.push(p('该企业工商年报未公示财务数据，或企查查财务数据库中暂无该企业财务信息。建议客户经理在实地尽调中收集企业近2年银行流水、纳税申报表及自述财务报表进行交叉验证。', { indent: { firstLine: 560 }, color: 'C41E1E' }));
  }
  
  // ====== 第四章：司法风险分析 ======
  children.push(new Paragraph({ children: [new PageBreak()] }));
  children.push(heading1('四、司法风险分析'));
  children.push(heading2('（一）风险扫描概览'));
  
  const riskPairs = [
    ['失信被执行人', `${dishonestCount}条`, '被执行人', `${judgmentDebtorCount}条`],
    ['限制高消费', `${highConsumptionCount}条`, '破产重整', `${bankruptcyCount}条`],
    ['股权冻结', `${equityFreezeCount}条`, '股权出质', `${equityPledgeCount}条`],
    ['行政处罚', `${adminPenaltyCount}条`, '经营异常', `${bizExceptionCount}条`],
    ['欠税公告', `${taxArrearsCount}条`, '裁判文书', `${judicialDocsCount}条`],
    ['立案信息', `${caseFilingCount}条`, '风险因子数', `${riskFactorCount}个`],
  ];
  children.push(infoTable(riskPairs));
  
  children.push(p(''));
  children.push(heading2('（二）风险评估结论'));
  
  if (fatalTotal === 0) {
    children.push(p('经企查查风险扫描，该企业不存在失信被执行人、被执行人、限制高消费、破产重整、股权冻结等致命风险事项。信用状况良好。', { indent: { firstLine: 560 }, color: '1E8449' }));
  } else {
    children.push(p(`经企查查风险扫描，该企业存在 ${fatalTotal} 项致命/严重风险事项（含失信${dishonestCount}条、被执行${judgmentDebtorCount}条、限高${highConsumptionCount}条、破产${bankruptcyCount}条、股权冻结${equityFreezeCount}条），须重点关注，审慎评估授信可行性。`, { indent: { firstLine: 560 }, color: 'C41E1E' }));
  }
  
  // ====== 第五章：工商变更与历史沿革 ======
  children.push(new Paragraph({ children: [new PageBreak()] }));
  children.push(heading1('五、工商变更与历史沿革'));
  
  if (changeList.length > 0) {
    const recentChanges = changeList.slice(0, 15);
    const chRows = recentChanges.map(c => [
      safeStr(c, '变更日期') || safeStr(c, '变更时间') || safeStr(c, 'date') || '——',
      safeStr(c, '变更项目') || safeStr(c, '变更事项') || safeStr(c, 'item') || '——',
      (safeStr(c, '变更前内容') || safeStr(c, '变更前') || safeStr(c, 'before') || '——').substring(0, 30),
      (safeStr(c, '变更后内容') || safeStr(c, '变更后') || safeStr(c, 'after') || '——').substring(0, 30),
    ]);
    children.push(makeTable(['变更日期', '变更项目', '变更前', '变更后'], chRows, [1400, 1800, 2700, 2766]));
    children.push(p(`（共 ${changeList.length} 条工商变更记录，上表展示最近 ${Math.min(15, changeList.length)} 条）`, { size: 20, color: '888888' }));
  } else {
    children.push(p('暂无工商变更记录或数据未获取。', { color: '888888' }));
  }
  
  // ====== 第六章：知识产权与经营能力 ======
  children.push(new Paragraph({ children: [new PageBreak()] }));
  children.push(heading1('六、知识产权与经营能力'));
  children.push(heading2('（一）知识产权'));
  
  const ipPairs = [
    ['专利数量', `${patentList.length}件`, '软件著作权', `${swList.length}件`],
    ['资质证书', `${qualificationList.length}件`, '招投标项目', `${biddingList.length}个`],
    ['荣誉奖项', `${honorList.length}项`, '分支机构', `${branchList.length}家`],
    ['对外投资企业', `${invList.length}家`, '知识产权合计', `${patentList.length + swList.length}件`],
  ];
  children.push(infoTable(ipPairs));
  
  if (patentList.length + swList.length > 0) {
    children.push(p(`该企业拥有知识产权合计 ${patentList.length + swList.length} 件，体现了一定的技术积累和创新能力。`, { indent: { firstLine: 560 } }));
  } else {
    children.push(p('该企业暂无知识产权记录。', { indent: { firstLine: 560 }, color: '888888' }));
  }
  
  children.push(p(''));
  children.push(heading2('（二）经营能力评估'));
  
  if (invList.length > 0) {
    children.push(p(`对外投资：该企业对外投资 ${invList.length} 家企业，具有一定的资本运作和产业链布局能力。`, { indent: { firstLine: 560 } }));
  }
  if (biddingList.length > 0) {
    children.push(p(`招投标参与：近年在企查查可查的招投标项目 ${biddingList.length} 个，表明企业具有一定的市场开拓和项目承接能力。`, { indent: { firstLine: 560 } }));
  }
  if (honorList.length > 0) {
    children.push(p(`企业荣誉：获得 ${honorList.length} 项荣誉，体现了一定的行业认可度和社会信誉。`, { indent: { firstLine: 560 } }));
  }
  
  // ====== 第七章：高管团队 ======
  children.push(new Paragraph({ children: [new PageBreak()] }));
  children.push(heading1('七、高管团队'));
  
  if (execList.length > 0) {
    const execRows = execList.slice(0, 10).map(e => [
      safeStr(e, '姓名') || safeStr(e, 'name') || '——',
      safeStr(e, '职务') || safeStr(e, 'position') || safeStr(e, 'title') || '——',
      safeStr(e, '任职时间') || safeStr(e, 'term') || '——',
    ]);
    children.push(makeTable(['姓名', '职务', '任职时间'], execRows, [2200, 3600, 2866]));
  } else {
    children.push(p('暂无高管/主要人员数据。', { color: '888888' }));
  }
  
  children.push(p(''));
  children.push(heading2('高管风险提示'));
  const execRiskRaw = data.execRisk || {};
  const hasExecRisk = Object.values(execRiskRaw).some(v => 
    (Array.isArray(v) && v.length > 0) || (typeof v === 'object' && Object.keys(v).length > 0)
  );
  if (hasExecRisk) {
    children.push(p('经企查查高管风险扫描，该企业高管团队存在部分风险记录，建议客户经理进一步核实高管个人征信及涉诉情况。', { indent: { firstLine: 560 }, color: 'D68910' }));
  } else {
    children.push(p('经企查查公开数据核查，该企业高管团队暂无公开的失信、被执行、限高等重大风险记录。', { indent: { firstLine: 560 }, color: '1E8449' }));
  }
  
  // ====== 第八章：综合评估 ======
  children.push(new Paragraph({ children: [new PageBreak()] }));
  children.push(heading1('八、综合评估与授信建议'));
  children.push(heading2('（一）综合授信评级'));
  
  children.push(p([
    { text: `综合授信评级：${grade}（${ratingLabel}）`, bold: true, size: 32, color: score >= 72 ? '1E8449' : score >= 55 ? 'D68910' : 'C41E1E' }
  ], { alignment: AlignmentType.CENTER, spacing: { before: 200, after: 200 } }));
  
  children.push(p([
    { text: `综合评分：${score}/100`, bold: true, size: 28, color: '1A2D4E' }
  ], { alignment: AlignmentType.CENTER }));
  
  children.push(p(''));
  children.push(heading2('（二）优势与亮点'));
  
  const strengths = [];
  if (fatalTotal === 0) strengths.push('致命风险（失信/被执行/限高/破产/股权冻结）全部清零，信用基础良好');
  if (estDate) {
    const age = now.getFullYear() - parseInt(estDate.substring(0, 4));
    if (age >= 10) strengths.push(`企业经营年限 ${age} 年，行业经验丰富，经营稳定性强`);
    else if (age >= 5) strengths.push(`企业经营年限 ${age} 年，具备一定的经营稳定性`);
  }
  if (displayAssetRatio) {
    const ar = parseFloat(displayAssetRatio);
    if (!isNaN(ar) && ar < 60) strengths.push(`资产负债率 ${fmtPct(displayAssetRatio)}，财务结构健康`);
  }
  if (displayNetMargin) {
    const nm = parseFloat(displayNetMargin);
    if (!isNaN(nm) && nm > 10) strengths.push(`净利率 ${fmtPct(displayNetMargin)}，盈利能力强`);
  }
  if (patentList.length + swList.length > 0) strengths.push(`拥有知识产权 ${patentList.length + swList.length} 件，具备技术创新能力`);
  if (emp && !isNaN(parseInt(emp)) && parseInt(emp) >= 100) strengths.push(`参保员工 ${emp} 人，经营规模较大`);
  
  if (strengths.length > 0) {
    for (const s of strengths) {
      children.push(p(`• ${s}`, { indent: { left: 360 } }));
    }
  } else {
    children.push(p('暂无明显突出优势，建议客户经理在实地尽调中进一步挖掘企业经营亮点。', { color: '888888' }));
  }
  
  children.push(p(''));
  children.push(heading2('（三）风险关注点'));
  
  const risks = [];
  if (fatalTotal > 0) risks.push(`存在 ${fatalTotal} 项致命/严重风险事项，须重点核查并评估对授信安全的影响`);
  if (displayAssetRatio) {
    const ar = parseFloat(displayAssetRatio);
    if (!isNaN(ar) && ar >= 75) risks.push(`资产负债率 ${fmtPct(displayAssetRatio)}，处于较高水平，需关注偿债压力`);
  }
  if (adminPenaltyCount > 0) risks.push(`存在 ${adminPenaltyCount} 条行政处罚记录，需核实处罚原因及整改情况`);
  if (bizExceptionCount > 0) risks.push(`存在 ${bizExceptionCount} 条经营异常记录，需核实是否已移出异常名录`);
  if (taxArrearsCount > 0) risks.push(`存在 ${taxArrearsCount} 条欠税公告，需关注企业税务合规性`);
  if (judicialDocsCount > 0) risks.push(`存在 ${judicialDocsCount} 条裁判文书，需评估涉诉对企业经营的影响`);
  if (!hasFinance) risks.push('企业未公示财务数据，财务信息不透明，需通过银行流水和纳税申报交叉验证');
  if (changeList.length > 20) risks.push(`工商变更记录 ${changeList.length} 条，变更较为频繁，需关注经营稳定性`);
  
  if (risks.length > 0) {
    for (const r of risks) {
      children.push(p(`• ${r}`, { indent: { left: 360 }, color: 'C41E1E' }));
    }
  } else {
    children.push(p('经企查查公开数据核查，暂未发现重大风险关注事项。', { color: '1E8449' }));
  }
  
  children.push(p(''));
  children.push(heading2('（四）授信建议'));
  
  let recommendation;
  if (score >= 88) {
    recommendation = `建议同意授信。企业综合评级 ${grade}（${ratingLabel}），经营基本面扎实，信用记录良好，偿债能力充足。建议授信金额与期限结合客户实际资金需求和还款能力综合确定。`;
  } else if (score >= 72) {
    recommendation = `建议有条件同意授信。企业综合评级 ${grade}（${ratingLabel}），整体风险可控，但存在个别需关注事项。建议在落实相关授信条件后予以授信，并加强贷后监控频率。`;
  } else if (score >= 55) {
    recommendation = `建议审慎授信或压缩额度。企业综合评级 ${grade}（${ratingLabel}），存在若干风险关注点，整体信用水平一般。如确需授信，建议压缩额度、增加担保措施，并执行更严格的贷后管理。`;
  } else {
    recommendation = `建议暂缓/否决授信。企业综合评级 ${grade}（${ratingLabel}），存在重大风险事项，授信安全难以保障。建议退回补充材料或直接否决。`;
  }
  
  children.push(p(recommendation, { indent: { firstLine: 560 }, bold: true }));
  
  children.push(p(''));
  children.push(heading2('（五）贷后管理建议'));
  children.push(p('1. 按月监控还款账户还款情况，首次扣款失败须当日联系企业了解原因。', { indent: { left: 360 } }));
  children.push(p('2. 按季度查询企业及实控人个人征信报告，监控新增逾期、新增网贷、新增被执行等风险信号。', { indent: { left: 360 } }));
  children.push(p('3. 每季度对经营场所进行一次非现场/现场回访，重点关注经营是否持续、流水是否显著下滑。', { indent: { left: 360 } }));
  children.push(p('4. 如经营流水连续3个月下降超过30%，或出现新增30天以上逾期，须在5个工作日内启动风险排查。', { indent: { left: 360 } }));
  children.push(p('5. 授信到期前1个月综合评估是否续贷，如经营明显恶化应审慎续贷或增加担保条件。', { indent: { left: 360 } }));
  
  // ====== 第九章：数据来源与说明 ======
  children.push(new Paragraph({ children: [new PageBreak()] }));
  children.push(heading1('九、数据来源与说明'));
  children.push(heading2('（一）数据来源'));
  
  children.push(infoTable([
    ['工商注册信息', '企查查 QCC 工商数据库', '财务数据', finArr.length > 0 ? 'QCC 财务数据库' : '未获取'],
    ['股权结构', '企查查 QCC 股东数据库', '司法风险', '企查查 QCC 风险扫描'],
    ['知识产权', '企查查 QCC 知识产权数据库', '经营信息', '企查查 QCC 经营数据库'],
    ['工商变更', '企查查 QCC 历史沿革数据库', '查询日期', dateShort],
  ]));
  
  children.push(p(''));
  children.push(heading2('（二）免责声明'));
  children.push(p('1. 本报告数据来源于企查查 QCC 公开数据平台，数据的准确性、完整性和时效性受限于企查查数据库的更新频率和覆盖范围。', { indent: { firstLine: 560 } }));
  children.push(p('2. 本报告中的财务数据可能与企业最新实际经营情况存在差异，建议结合企业银行流水、纳税申报表及现场尽调结果进行交叉验证。', { indent: { firstLine: 560 } }));
  children.push(p('3. 本报告的综合评分和授信评级系基于公开数据的量化模型计算结果，不构成授信决策的直接依据。实际授信应以客户经理尽职调查、最新征信报告、担保评估及我行内部审批意见为准。', { indent: { firstLine: 560 } }));
  children.push(p('4. 本报告为系统自动生成，仅供内部审批参考，不得对外提供。', { indent: { firstLine: 560 } }));
  
  children.push(p(''));
  children.push(p(`报告生成时间：${now.toLocaleString('zh-CN')}`, { alignment: AlignmentType.RIGHT, size: 20, color: '888888' }));
  children.push(p(`数据来源：企查查 QCC MCP 平台（6 大 Server / 12 个信贷场景 Skill）`, { alignment: AlignmentType.RIGHT, size: 20, color: '888888' }));
  children.push(p('—— 报告完 ——', { alignment: AlignmentType.CENTER, spacing: { before: 400 }, color: '888888' }));
  
  // ====== 组装文档 ======
  const doc = new Document({
    styles: {
      default: {
        document: { run: { font: FONT, size: 28 } }
      },
      paragraphStyles: [
        { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { size: 44, bold: true, font: FONT_TITLE, color: '1A2D4E' },
          paragraph: { spacing: { before: 360, after: 240, line: 360 }, alignment: AlignmentType.CENTER, outlineLevel: 0 } },
        { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { size: 32, bold: true, font: FONT_TITLE, color: '1A2D4E' },
          paragraph: { spacing: { before: 240, after: 180, line: 360 }, outlineLevel: 1 } },
        { id: 'Heading3', name: 'Heading 3', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { size: 28, bold: true, font: FONT_TITLE, color: '1D6FA4' },
          paragraph: { spacing: { before: 180, after: 120, line: 360 }, outlineLevel: 2 } },
      ]
    },
    sections: [{
      properties: {
        page: {
          size: { width: PAGE_WIDTH, height: PAGE_HEIGHT },
          margin: { top: MARGIN_TOP, bottom: MARGIN_BOTTOM, left: MARGIN_LEFT, right: MARGIN_RIGHT }
        }
      },
      headers: {
        default: new Header({
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: `中国工商银行普惠金融授信调查报告 — ${companyFullName}`, size: 16, color: '999999', font: FONT })],
            border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC', space: 4 } }
          })]
        })
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({ text: '机密·仅供内部审批 — 第 ', size: 16, color: '999999', font: FONT }),
              new TextRun({ children: [PageNumber.CURRENT], size: 16, color: '999999', font: FONT }),
              new TextRun({ text: ' 页', size: 16, color: '999999', font: FONT }),
            ]
          })]
        })
      },
      children
    }]
  });
  
  return doc;
}

// ============ 主流程 ============
async function main() {
  const args = process.argv.slice(2);
  if (args.length < 1) {
    console.log('用法: node generate_credit_report.js <企业名称> [输出文件名] [支行名称]');
    console.log('示例: node generate_credit_report.js "腾讯科技(深圳)有限公司"');
    console.log('示例: node generate_credit_report.js "腾讯科技(深圳)有限公司" 腾讯科技_授信报告.docx "深圳分行南山支行"');
    process.exit(1);
  }
  
  const companyName = args[0];
  const outputFile = args[1] || `${companyName.replace(/[\\/:*?"<>|]/g, '_')}_授信调查报告.docx`;
  const branch = args[2] || 'XX分行XX支行';
  
  try {
    // 1. 收集数据
    const data = await collectAllData(companyName);
    
    // 2. 生成报告
    console.log(`\n📝 生成授信报告...`);
    const doc = await buildReport(data, companyName, branch);
    
    // 3. 输出文件
    const buffer = await Packer.toBuffer(doc);
    const outPath = path.resolve(outputFile);
    fs.writeFileSync(outPath, buffer);
    
    console.log(`\n✅ 报告已生成: ${outPath}`);
    console.log(`   文件大小: ${(buffer.length / 1024).toFixed(1)} KB`);
    
    if (data.errors.length > 0) {
      console.log(`\n⚠️ 数据采集中有 ${data.errors.length} 个错误:`);
      for (const e of data.errors) {
        console.log(`   - ${e}`);
      }
    }
  } catch (e) {
    console.error(`\n❌ 生成失败: ${e.message}`);
    console.error(e.stack);
    process.exit(1);
  }
}

main();
