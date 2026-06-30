/**
 * 企业授信报告系统 — Express 后端代理 (12 Skills Edition)
 * 
 * 整合企查查 QCC MCP 6 大 Server，覆盖 12 个银行信贷场景 Skills
 * 按专业授信报告格式（封面 + 评级横幅 + 9 章节）生成 HTML
 * 
 * QCC MCP 6 大 Server:
 *   qcc-company   — 工商、财务、年报、股东、投资、UBO、实控人
 *   qcc-risk      — 风险扫描、失信、被执行、限高、股权冻结等
 *   qcc-ipr       — 知识产权（软著、专利）
 *   qcc-operation — 经营信息（招投标、资质、进出口信用、招聘、荣誉等）
 *   qcc-executive — 高管/董监高个人风险
 *   qcc-history   — 历史沿革
 */

const express = require('express');
const cors = require('cors');
const path = require('path');
const { buildDocxReport, Packer } = require('./docx_builder');

const app = express();
app.use(cors({
  origin: function (origin, callback) {
    // Allow null origin (file://), localhost, and 127.0.0.1
    if (!origin || origin === 'null' || origin.startsWith('http://localhost') || origin.startsWith('http://127.0.0.1')) {
      callback(null, true);
    } else {
      callback(null, true); // Allow all for development
    }
  },
  credentials: false
}));
app.use(express.json({ limit: '10mb' }));
app.use(express.static(path.join(__dirname, '..')));

// ============ QCC MCP 配置 ============
const QCC_SERVERS = {
  company:   { url: 'https://agent.qcc.com/mcp/company/stream',   token: 'Bearer MmDFP6lIn3akremGqfRhudPo0NON0LTrz92cBzCHtI3eA7IZ' },
  risk:      { url: 'https://agent.qcc.com/mcp/risk/stream',      token: 'Bearer MmDFP6lIn3akremGqfRhudPo0NON0LTrz92cBzCHtI3eA7IZ' },
  ipr:       { url: 'https://agent.qcc.com/mcp/ipr/stream',       token: 'Bearer MmDFP6lIn3akremGqfRhudPo0NON0LTrz92cBzCHtI3eA7IZ' },
  operation: { url: 'https://agent.qcc.com/mcp/operation/stream', token: 'Bearer MmDFP6lIn3akremGqfRhudPo0NON0LTrz92cBzCHtI3eA7IZ' },
  executive: { url: 'https://agent.qcc.com/mcp/executive/stream', token: 'Bearer MmDFP6lIn3akremGqfRhudPo0NON0LTrz92cBzCHtI3eA7IZ' },
  history:   { url: 'https://agent.qcc.com/mcp/history/stream',   token: 'Bearer MmDFP6lIn3akremGqfRhudPo0NON0LTrz92cBzCHtI3eA7IZ' }
};

// ============ MCP 调用 ============
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

  // 处理 SSE 响应 — 取最后一个有效结果（某些接口可能返回多条 data）
  let lastResult = null;
  for (const line of text.split('\n')) {
    if (line.startsWith('data: ')) {
      try {
        const sse = JSON.parse(line.substring(6));
        if (sse.error) throw new Error(`QCC ${server}/${toolName}: ${JSON.stringify(sse.error)}`);
        const r = sse.result;
        if (r?.content?.[0]?.text) {
          try { lastResult = JSON.parse(r.content[0].text); } catch { lastResult = r.content[0].text; }
        } else if (r) {
          lastResult = r;
        }
      } catch (e) { if (e.message?.includes('QCC ')) throw e; }
    }
  }
  return lastResult;
}

async function executeTaskBatch(tasks, data) {
  const results = await Promise.all(tasks.map(t =>
    t.fn().then(r => ({ task: t, result: r, ok: true }))
         .catch(e => ({ task: t, error: e, ok: false }))
  ));
  for (const r of results) {
    if (r.ok) {
      if (r.task.set) r.task.set(r.result);
      const hasData = r.result && (typeof r.result === 'object' ? Object.keys(r.result).length > 0 : true);
      console.log(`  ✅ ${r.task.name}: ${hasData ? '已获取' : '无记录'}`);
    } else {
      const em = r.error?.message || String(r.error);
      // optional 任务失败不记录到 errors（如 history server 未授权的工具）
      if (!r.task.optional) {
        data.errors.push(`${r.task.name}: ${em}`);
      }
      console.log(`  ${r.task.optional ? '⚠️' : '❌'} ${r.task.name}: ${em}${r.task.optional ? ' (可选，已忽略)' : ''}`);
    }
  }
}

function task(name, server, toolName, argsFn, setFn, optional) {
  return { name, fn: () => callQccMCP(server, toolName, argsFn()), set: setFn, optional: !!optional };
}

// ============ 工具定义 ============
function basicCompanyTasks(data, companyName) {
  const a = () => ({ searchKey: companyName });
  return [
    task('工商信息', 'company', 'get_company_registration_info', a, r => data.company = r),
    task('财务数据', 'company', 'get_financial_data', a, r => data.financial = r),
    task('企业年报', 'company', 'get_annual_reports', a, r => data.annualReports = r),
    task('股东信息', 'company', 'get_shareholder_info', a, r => data.shareholders = r),
    task('对外投资', 'company', 'get_external_investments', a, r => data.investments = r),
    task('风险扫描', 'risk', 'get_company_risk_scan', a, r => data.risk = r),
  ];
}

function riskDetailTasks(data, companyName) {
  const a = () => ({ searchKey: companyName });
  const rd = () => { if (!data.riskDetail) data.riskDetail = {}; return data.riskDetail; };
  return [
    task('失信信息', 'risk', 'get_dishonest_info', a, r => rd().dishonest = r),
    task('被执行信息', 'risk', 'get_judgment_debtor_info', a, r => rd().judgmentDebtor = r),
    task('限高消费', 'risk', 'get_high_consumption_restriction', a, r => rd().highConsumption = r),
    task('股权冻结', 'risk', 'get_equity_freeze', a, r => rd().equityFreeze = r),
    task('股权出质', 'risk', 'get_equity_pledge_info', a, r => rd().equityPledge = r),
    task('经营异常', 'risk', 'get_business_exception', a, r => rd().businessException = r),
    task('严重违法', 'risk', 'get_serious_violation', a, r => rd().seriousViolation = r),
    task('行政处罚', 'risk', 'get_administrative_penalty', a, r => rd().adminPenalty = r),
    task('欠税公告', 'risk', 'get_tax_arrears_notice', a, r => rd().taxArrears = r),
    task('税务异常', 'risk', 'get_tax_abnormal', a, r => rd().taxAbnormal = r),
    task('税务违法', 'risk', 'get_tax_violation', a, r => rd().taxViolation = r),
    task('破产重整', 'risk', 'get_bankruptcy_reorganization', a, r => rd().bankruptcy = r),
    task('限出境', 'risk', 'get_exit_restriction', a, r => rd().exitRestriction = r),
    task('终本案件', 'risk', 'get_terminated_cases', a, r => rd().terminatedCases = r),
    task('违约事项', 'risk', 'get_default_info', a, r => rd().defaultInfo = r),
  ];
}

function litigationTasks(data, companyName) {
  const a = () => ({ searchKey: companyName });
  const lt = () => { if (!data.litigation) data.litigation = {}; return data.litigation; };
  return [
    task('裁判文书', 'risk', 'get_judicial_documents', a, r => lt().judicialDocs = r),
    task('立案信息', 'risk', 'get_case_filing_info', a, r => lt().caseFiling = r),
    task('开庭公告', 'risk', 'get_hearing_notice', a, r => lt().hearingNotice = r),
    task('法院公告', 'risk', 'get_court_notice', a, r => lt().courtNotice = r),
    task('送达公告', 'risk', 'get_service_notice', a, r => lt().serviceNotice = r),
    task('诉前调解', 'risk', 'get_pre_litigation_mediation', a, r => lt().mediation = r),
  ];
}

function executiveTasks(data, companyName) {
  const a = () => ({ searchKey: companyName });
  return [
    // executive server 的 get_executive_info 未启用，统一用 company server 获取人员信息
    task('高管/主要人员', 'company', 'get_key_personnel', a, r => { data.executives = r; data.keyPersonnel = r; }),
    task('实际控制人', 'company', 'get_actual_controller', a, r => data.actualController = r),
  ];
}

function historyTasks(data, companyName) {
  // 优先用统一社会信用代码查询（比公司名更精准），回退到公司名
  const searchKey = () => {
    const uscc = data.company?.['统一社会信用代码'] || data.company?.KeyNo;
    return { searchKey: uscc || companyName };
  };
  return [
    // get_change_records 已在阶段1.5单独串行执行，此处不再重复调用
    task('历史法代', 'history', 'get_historical_legal_rep', searchKey, r => data.histLegalRep = r, true),
    task('历史股东', 'history', 'get_historical_shareholders', searchKey, r => data.histShareholders = r, true),
    task('历史失信', 'history', 'get_historical_dishonest', searchKey, r => data.histDishonest = r, true),
    task('历史被执行', 'history', 'get_historical_judgment_debtor', searchKey, r => data.histJudgmentDebtor = r, true),
    task('历史注册变更', 'history', 'get_historical_registration', searchKey, r => data.histRegistration = r, true),
    task('历史高管', 'history', 'get_historical_executives', searchKey, r => data.histExecutives = r, true),
  ];
}

function uboTasks(data, companyName) {
  const a = () => ({ searchKey: companyName });
  return [
    task('受益所有人', 'company', 'get_beneficial_owners', a, r => data.beneficialOwners = r),
    task('企业档案', 'company', 'get_company_profile', a, r => data.companyProfile = r),
    task('分支机构', 'company', 'get_branches', a, r => data.branches = r),
    // get_change_records 已由 historyTasks 覆盖，不再重复调用（并发重复请求可能被 QCC 去重导致返回"无匹配项"）
  ];
}

function operationTasks(data, companyName) {
  const a = () => ({ searchKey: companyName });
  const op = () => { if (!data.operation) data.operation = {}; return data.operation; };
  return [
    task('招投标信息', 'operation', 'get_bidding_info', a, r => op().bidding = r),
    task('资质证书', 'operation', 'get_qualifications', a, r => op().qualifications = r),
    task('招聘信息', 'operation', 'get_recruitment_info', a, r => op().recruitment = r),
    task('信用评估', 'operation', 'get_credit_evaluation', a, r => op().creditEval = r),
    task('荣誉信息', 'operation', 'get_honor_info', a, r => op().honor = r),
  ];
}

function iprTasks(data, companyName) {
  const a = () => ({ searchKey: companyName });
  const ip = () => { if (!data.ipr) data.ipr = {}; return data.ipr; };
  return [
    task('软件著作权', 'ipr', 'get_software_copyright_info', a, r => ip().software = r),
    task('专利信息', 'ipr', 'get_patent_info', a, r => ip().patents = r),
  ];
}

function executiveRiskTasks(data, companyName, personName) {
  const a = () => ({ searchKey: companyName, personName: personName || '' });
  const er = () => { if (!data.execRisk) data.execRisk = {}; return data.execRisk; };
  return [
    task('高管风险扫描', 'executive', 'get_executive_risk_scan', a, r => er().riskScan = r),
    task('高管失信', 'executive', 'get_executive_dishonest', a, r => er().dishonest = r),
    task('高管限高', 'executive', 'get_executive_high_consumption_ban', a, r => er().highConsumption = r),
    task('高管限出境', 'executive', 'get_executive_exit_restriction', a, r => er().exitRestriction = r),
    task('高管被执行', 'executive', 'get_executive_judgment_debtor', a, r => er().judgmentDebtor = r),
  ];
}

function tradeFinanceTasks(data, companyName) {
  const a = () => ({ searchKey: companyName });
  return [
    task('进出口信用', 'operation', 'get_import_export_credit', a, r => { if (!data.tradeFinance) data.tradeFinance = {}; data.tradeFinance.importExportCredit = r; }),
  ];
}

// ============ 数据提取辅助 ============

function safeArr(val, path) {
  try {
    if (!val) return [];
    if (Array.isArray(val)) return val;
    for (const key of path) {
      if (val[key] !== undefined) val = val[key];
      else return [];
    }
    return Array.isArray(val) ? val : [];
  } catch { return []; }
}

function safeObj(val, path) {
  try {
    if (!val) return {};
    for (const key of path) {
      if (val[key] !== undefined) val = val[key];
      else return {};
    }
    return typeof val === 'object' && !Array.isArray(val) ? val : {};
  } catch { return {}; }
}

function safeStr(val, path) {
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

function countArr(val) {
  if (!val) return 0;
  if (Array.isArray(val)) return val.length;
  if (typeof val === 'object') return Object.keys(val).length;
  return 0;
}

function getRiskCount(data, factorName) {
  if (!data?.risk?.['风险因子扫描']) return 0;
  const factor = data.risk['风险因子扫描'].find(f => f['风险因子'] === factorName);
  return factor ? (factor['条目数'] || 0) : 0;
}

function hasRiskData(data) {
  return !!(data?.riskDetail);
}

function getRiskItemCount(data, key) {
  if (!data?.riskDetail?.[key]) return 0;
  const v = data.riskDetail[key];
  if (Array.isArray(v)) return v.length;
  if (typeof v === 'object') return Object.keys(v).length;
  return v ? 1 : 0;
}

// ============ 评分 ============
function calcRating(data) {
  let score = 50;
  const details = [];
  const fatalRisks = ['dishonest', 'judgmentDebtor', 'highConsumption', 'bankruptcy', 'equityFreeze'];

  // 致命风险检查
  let fatalCount = 0;
  for (const key of fatalRisks) {
    const c = getRiskItemCount(data, key);
    if (c > 0) { fatalCount += c; details.push(`⚠️ ${key}: ${c}条`); }
  }
  if (fatalCount > 0) {
    score -= fatalCount * 15;
  } else {
    score += 25;
    details.push('致命风险全部清零');
  }

  // 财务数据
  const fin = data?.financial?.['财务数据信息'];
  if (fin && fin.length > 0) {
    const latest = fin[0];
    const indicators = latest['指标详情'] || {};
    const analysis = indicators['分析数据'] || {};
    const ability = analysis['偿还能力'] || {};
    const profit = analysis['盈利能力'] || {};
    const growth = analysis['成长能力'] || {};

    const assetRatio = parseFloat(ability['资产负债率']);
    if (!isNaN(assetRatio)) {
      if (assetRatio < 40) { score += 20; details.push(`资产负债率 ${assetRatio}% — 优秀`); }
      else if (assetRatio < 60) { score += 14; details.push(`资产负债率 ${assetRatio}% — 良好`); }
      else if (assetRatio < 75) { score += 8; details.push(`资产负债率 ${assetRatio}% — 偏高`); }
      else { score += 2; details.push(`资产负债率 ${assetRatio}% — 较高`); }
    }

    const netMargin = parseFloat(profit['净利率']);
    if (!isNaN(netMargin)) {
      if (netMargin > 20) { score += 15; details.push(`净利率 ${netMargin}% — 极优`); }
      else if (netMargin > 10) { score += 10; details.push(`净利率 ${netMargin}% — 良好`); }
      else { score += 5; details.push(`净利率 ${netMargin}%`); }
    }
  } else {
    // 尝试年报数据
    const ar = data?.annualReports?.['企业年报信息'];
    if (ar && ar.length > 0) {
      const assets = ar[0]['企业资产状况信息'] || {};
      if (assets['资产总额'] && assets['资产总额'] !== '企业选择不公示') {
        score += 10;
        details.push('年报财务数据可用');
      }
    }
  }

  // 经营年限
  const estDate = safeStr(data?.company, ['成立日期']);
  if (estDate) {
    const y = parseInt(estDate.substring(0, 4));
    const age = new Date().getFullYear() - y;
    if (age >= 20) { score += 12; details.push(`经营 ${age} 年 — 成熟企业`); }
    else if (age >= 10) { score += 9; details.push(`经营 ${age} 年`); }
    else if (age >= 5) { score += 6; details.push(`经营 ${age} 年`); }
    else { score += 3; details.push(`经营 ${age} 年`); }
  }

  // 参保人数
  const emp = parseInt(safeStr(data?.company, ['参保人数']));
  if (!isNaN(emp)) {
    if (emp >= 5000) { score += 12; details.push(`参保 ${emp} 人 — 超大规模`); }
    else if (emp >= 1000) { score += 10; details.push(`参保 ${emp} 人 — 大型企业`); }
    else if (emp >= 100) { score += 7; details.push(`参保 ${emp} 人 — 中型企业`); }
    else if (emp > 0) { score += 4; details.push(`参保 ${emp} 人`); }
  }

  // 知识产权
  const patentCount = countArr(data?.ipr?.patents);
  const swCount = countArr(data?.ipr?.software);
  if (patentCount + swCount > 50) { score += 8; details.push('知识产权丰富'); }
  else if (patentCount + swCount > 10) { score += 5; details.push('有知识产权积累'); }

  // 股东
  if (data?.shareholders) {
    const sh = data.shareholders['股东信息'] || [];
    if (sh.length > 0) { score += 5; details.push('股东结构清晰'); }
  }

  score = Math.max(0, Math.min(100, Math.round(score)));

  let grade, label, color;
  if (score >= 95) { grade = 'AAA'; label = '极优'; color = '#145a2c'; }
  else if (score >= 88) { grade = 'AA+'; label = '优良'; color = '#1e8449'; }
  else if (score >= 80) { grade = 'AA'; label = '良好'; color = '#2a9d5c'; }
  else if (score >= 72) { grade = 'A'; label = '良好'; color = '#3ba86c'; }
  else if (score >= 65) { grade = 'BBB'; label = '一般'; color = '#d68910'; }
  else if (score >= 55) { grade = 'BB'; label = '关注'; color = '#c0392b'; }
  else { grade = 'B'; label = '谨慎'; color = '#c0392b'; }

  return { score, grade, label, color, details };
}

// ============ 报告生成 ============
function buildReportHTML(branch, companyName, data, skillId) {
  const now = new Date();
  const dateStr = now.toLocaleDateString('zh-CN', { year: 'numeric', month: 'long', day: 'numeric' });
  const dateShort = now.toISOString().slice(0, 10);
  const timeStr = now.toLocaleString('zh-CN');

  const rating = calcRating(data);
  const skillNames = {
    'credit-due-diligence': '授信尽调报告', 'counterparty-risk': '交易对手风险评估',
    'guarantor-check': '担保方资信核查', 'kyb-verification': 'KYB 企业核验报告',
    'ubo-screening': '受益所有人识别报告', 'business-health-scan': '经营健康度扫描',
    'credit-monitoring': '信贷风险监控报告', 'bankruptcy-monitor': '破产预警监控',
    'equity-structure': '股权结构穿透分析', 'executive-background': '高管背景核查报告',
    'litigation-analysis': '诉讼风险评估报告', 'trade-finance-compliance': '贸易融资合规核查'
  };
  const reportType = skillNames[skillId] || '综合授信评审报告';

  // ========== 数据提取 ==========
  const comp = data.company || {};
  const companyFullName = safeStr(comp, ['企业名称']) || companyName;
  const creditCode = safeStr(comp, ['统一社会信用代码']);
  const legalPerson = safeStr(comp, ['法定代表人']);
  const regCapital = safeStr(comp, ['注册资本']);
  const estDate = safeStr(comp, ['成立日期']);
  const companyType = safeStr(comp, ['企业类型']);
  const regAddr = safeStr(comp, ['注册地址']);
  const businessScope = safeStr(comp, ['经营范围']);
  const status = safeStr(comp, ['登记状态']);
  const industry = safeStr(comp, ['国标行业']);
  const employees = safeStr(comp, ['参保人数']);
  const region = safeStr(comp, ['所属地区']);

  // 财务
  let finRevenue = '', finNetProfit = '', finTotalAssets = '', finTotalLiabilities = '';
  let finEquity = '', finAssetRatio = '', finNetMargin = '', finRevenueGrowth = '';
  let finROE = '', finCashFlow = '';
  let financeSource = '', financePeriod = '';
  const finArr = safeArr(data.financial, ['财务数据信息']);
  if (finArr.length > 0) {
    financeSource = 'QCC财务数据库';
    const latest = finArr[0];
    financePeriod = latest['报告期'] || '';
    const indicators = latest['指标详情'] || {};
    const bs = safeObj(indicators, ['财务报表', '资产负债表']);
    const is = safeObj(indicators, ['财务报表', '利润表']);
    const analysis = safeObj(indicators, ['分析数据']);
    finRevenue = safeStr(is, ['营业总收入']);
    finNetProfit = safeStr(is, ['净利润']);
    finTotalAssets = safeStr(bs, ['资产合计']);
    finTotalLiabilities = safeStr(bs, ['负债合计']);
    finEquity = safeStr(bs, ['所有者权益总计']);
    finAssetRatio = safeStr(analysis, ['偿还能力', '资产负债率']);
    finNetMargin = safeStr(analysis, ['盈利能力', '净利率']);
    finRevenueGrowth = safeStr(analysis, ['成长能力', '营业收入同比']);
    finROE = safeStr(analysis, ['盈利能力', '净资产收益率']);
  }

  // 年报数据
  let arRevenue = '', arNetProfit = '', arTotalAssets = '', arTotalLiabilities = '', arEquity = '';
  let arYear = '';
  const arArr = safeArr(data.annualReports, ['企业年报信息']);
  if (arArr.length > 0) {
    arYear = arArr[0]['年报年度'] || '';
    const assets = arArr[0]['企业资产状况信息'] || {};
    const getV = (k) => assets[k] && assets[k] !== '企业选择不公示' ? assets[k] : '';
    arRevenue = getV('营业总收入'); arNetProfit = getV('净利润');
    arTotalAssets = getV('资产总额'); arTotalLiabilities = getV('负债总额');
    arEquity = getV('所有者权益合计');
    if (!financeSource) financeSource = '工商年报';
  }

  // 合并财务（优先 QCC 数据库）
  const hasFinance = !!(finRevenue || arRevenue);
  const displayRevenue = finRevenue || arRevenue;
  const displayNetProfit = finNetProfit || arNetProfit;
  const displayAssets = finTotalAssets || arTotalAssets;
  const displayLiabilities = finTotalLiabilities || arTotalLiabilities;
  const displayEquity = finEquity || arEquity;
  const displayAssetRatio = finAssetRatio;
  const displayNetMargin = finNetMargin;

  // 股东
  const shareholderList = safeArr(data.shareholders, ['股东信息']);
  const invList = safeArr(data.investments, ['对外投资']);

  // 高管 — 增强 fallback 逻辑：尝试多种可能的 key 名
  const execRaw = data.executives || {};
  const execList = (() => {
    // 尝试常见的 QCC API 返回 key
    const keys = ['高管信息', '主要人员', '董监高', 'executives', 'personnel', 'key_personnel', 'management'];
    for (const k of keys) {
      const v = execRaw[k];
      if (Array.isArray(v) && v.length > 0) return v;
    }
    // 如果 execRaw 本身是数组
    if (Array.isArray(execRaw)) return execRaw;
    // 遍历 execRaw 的所有 key，取第一个数组值
    for (const k of Object.keys(execRaw)) {
      const v = execRaw[k];
      if (Array.isArray(v) && v.length > 0) return v;
    }
    return [];
  })();
  const keyPersonnelList = safeArr(data.keyPersonnel, ['主要人员']);
  // 合并高管和主要人员（去重）
  const allPersonnelNames = new Set();
  const mergedPersonnelList = [];
  for (const p of [...execList, ...keyPersonnelList]) {
    const name = p['姓名'] || p['name'] || p['Name'] || '';
    if (name && !allPersonnelNames.has(name)) {
      allPersonnelNames.add(name);
      mergedPersonnelList.push(p);
    }
  }

  // 风险
  const riskScanList = safeArr(data.risk, ['风险因子扫描']);
  const riskFactorCount = data.risk?.['有记录因子数'] || 0;
  const riskTotalCount = data.risk?.['有记录条目数'] || 0;

  const dishonestCount = getRiskCount(data, '失信被执行人');
  const highConsumptionCount = getRiskCount(data, '限制高消费');
  const bankruptcyCount = getRiskCount(data, '破产重整');
  const equityFreezeCount = getRiskCount(data, '股权冻结');
  const equityPledgeCount = getRiskCount(data, '股权出质');
  const adminPenaltyCount = getRiskCount(data, '行政处罚');
  const bizExceptionCount = getRiskCount(data, '经营异常');
  const taxArrearsCount = getRiskCount(data, '欠税公告');
  const taxViolationCount = getRiskCount(data, '税务违法');
  const terminatedCount = getRiskCount(data, '终本案件');
  const judgmentDebtorCount = getRiskCount(data, '被执行人');
  const judicialDocsCount = getRiskCount(data, '裁判文书');
  const caseFilingCount = getRiskCount(data, '立案信息');
  const hearingNoticeCount = getRiskCount(data, '开庭公告');
  const defaultCount = getRiskCount(data, '违约事项');

  // 知识产权 — 增强解析：支持多种 QCC API 返回格式
  const patentRaw = data.ipr?.patents;
  const swRaw = data.ipr?.software;
  const patentList = (() => {
    if (Array.isArray(patentRaw)) return patentRaw;
    if (!patentRaw || typeof patentRaw !== 'object') return [];
    // 尝试常见 key
    for (const k of ['专利信息', 'patents', 'patent_list', 'records', 'items', 'data']) {
      const v = patentRaw[k];
      if (Array.isArray(v)) return v;
    }
    // 遍历所有 key 取第一个数组
    for (const k of Object.keys(patentRaw)) {
      if (Array.isArray(patentRaw[k])) return patentRaw[k];
    }
    return [];
  })();
  const swList = (() => {
    if (Array.isArray(swRaw)) return swRaw;
    if (!swRaw || typeof swRaw !== 'object') return [];
    for (const k of ['软件著作权', 'software', 'copyright_list', 'records', 'items', 'data']) {
      const v = swRaw[k];
      if (Array.isArray(v)) return v;
    }
    for (const k of Object.keys(swRaw)) {
      if (Array.isArray(swRaw[k])) return swRaw[k];
    }
    return [];
  })();
  const patentCount = patentList.length;
  const swCount = swList.length;

  // UBO
  const boList = safeArr(data.beneficialOwners, ['受益所有人']);
  const actualController = data.actualController;
  const acName = safeStr(actualController, ['实际控制人']) || safeStr(actualController, ['姓名']);

  // 历史 — 增强解析，支持 company/get_change_records 和 history server 多种返回格式
  const historyChanges = (() => {
    const raw = data.history || data.changeRecords;
    if (!raw) return [];
    if (Array.isArray(raw)) return raw;
    // 尝试常见 key
    for (const k of ['变更记录信息', '变更记录', 'change_records', 'records', 'items', 'data', '变更信息']) {
      const v = raw[k];
      if (Array.isArray(v)) return v;
    }
    for (const k of Object.keys(raw)) {
      if (Array.isArray(raw[k])) return raw[k];
    }
    return [];
  })();

  // 经营
  const biddingList = safeArr(data.operation?.bidding, ['招投标信息']);
  const honorList = safeArr(data.operation?.honor, ['荣誉信息']);
  const qualificationList = safeArr(data.operation?.qualifications, ['资质证书']);

  // ========== 数值格式化 ==========
  function fmtMoneyCN(v) {
    const n = parseFloat(String(v).replace(/[^0-9.\-]/g, ''));
    if (isNaN(n)) return '——';
    if (Math.abs(n) >= 1e12) return (n / 1e12).toFixed(2) + '万亿';
    if (Math.abs(n) >= 1e8) return (n / 1e8).toFixed(2) + '亿';
    if (Math.abs(n) >= 1e4) return (n / 1e4).toFixed(2) + '万';
    return n.toLocaleString() + '元';
  }

  function fmtMoneyHK(v) {
    const n = parseFloat(String(v).replace(/[^0-9.\-]/g, ''));
    if (isNaN(n)) return '——';
    if (Math.abs(n) >= 1e8) return '¥' + (n / 1e8).toFixed(2) + '亿';
    if (Math.abs(n) >= 1e4) return '¥' + (n / 1e4).toFixed(0) + '万';
    return '¥' + n.toLocaleString();
  }

  function fmtPct(v) {
    const n = parseFloat(String(v).replace(/[^0-9.\-]/g, ''));
    return isNaN(n) ? '——' : n.toFixed(1) + '%';
  }

  function tag(cls, text) { return `<span class="tag ${cls}">${text}</span>`; }

  // ========== 构建 HTML ==========
  let html = `<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>${reportType} — ${companyFullName}</title>
<style>
  :root {
    --primary: #1a2d4e; --accent: #1d6fa4; --gold: #c8941a;
    --bg: #f5f7fa; --card-bg: #ffffff; --text: #2c3e50; --text-light: #6b7d8e;
    --border: #dce3ec; --success: #1e8449; --warning: #d68910; --danger: #c0392b; --info: #1a6a9e;
  }
  * { margin:0; padding:0; box-sizing:border-box; }
  body { font-family:"PingFang SC","Microsoft YaHei","Hiragino Sans GB",sans-serif; background:var(--bg); color:var(--text); line-height:1.8; font-size:14px; }
  .cover { background: linear-gradient(145deg,#051827 0%,#0a2a4a 55%,#0d3561 100%); color:#fff; padding:72px 64px 60px; text-align:center; position:relative; overflow:hidden; }
  .cover::before { content:''; position:absolute; inset:0; background: radial-gradient(ellipse at 20% 15%,rgba(29,111,164,.18) 0%,transparent 55%), radial-gradient(ellipse at 75% 70%,rgba(200,148,26,.1) 0%,transparent 45%); }
  .cover * { position:relative; z-index:1; }
  .cover-tag { display:inline-block; border:1.5px solid var(--gold); color:var(--gold); padding:4px 20px; border-radius:2px; font-size:12px; letter-spacing:5px; margin-bottom:26px; }
  .cover h1 { font-size:33px; font-weight:700; letter-spacing:2px; margin-bottom:8px; }
  .cover .subj { font-size:20px; font-weight:400; opacity:.85; margin-bottom:32px; }
  .cover-meta { display:flex; justify-content:center; gap:48px; font-size:12px; opacity:.72; flex-wrap:wrap; }
  .cover-meta span { display:flex; flex-direction:column; }
  .cover-meta .lbl { font-size:10px; opacity:.6; margin-bottom:2px; letter-spacing:1px; }
  .container { max-width:980px; margin:0 auto; padding:0 18px; }
  .rating-banner { background:var(--card-bg); margin:-44px auto 28px; max-width:980px; border-radius:10px; box-shadow:0 6px 28px rgba(0,0,0,.10); display:flex; position:relative; z-index:10; }
  .rating-left { background:linear-gradient(135deg,${rating.color},${rating.color}); color:#fff; padding:28px 34px; border-radius:10px 0 0 10px; text-align:center; min-width:176px; display:flex; flex-direction:column; align-items:center; justify-content:center; }
  .rating-grade { font-size:58px; font-weight:700; line-height:1; margin-bottom:4px; }
  .rating-label { font-size:13px; opacity:.9; letter-spacing:3px; }
  .rating-right { flex:1; padding:22px 30px; display:flex; flex-wrap:wrap; gap:18px; align-items:center; }
  .r-item { display:flex; flex-direction:column; min-width:88px; }
  .r-item .val { font-size:19px; font-weight:700; color:var(--primary); }
  .r-item .lbl { font-size:11px; color:var(--text-light); }
  .r-div { width:1px; height:38px; background:var(--border); }
  section { margin-bottom:24px; }
  .card { background:var(--card-bg); border-radius:10px; box-shadow:0 2px 10px rgba(0,0,0,.055); padding:30px 34px; }
  .card h2 { font-size:17px; font-weight:700; color:var(--primary); padding-bottom:11px; border-bottom:2.5px solid var(--accent); margin-bottom:20px; display:flex; align-items:center; gap:10px; }
  .card h2 .num { display:inline-flex; align-items:center; justify-content:center; width:27px; height:27px; background:var(--accent); color:#fff; font-size:13px; border-radius:50%; flex-shrink:0; }
  .card h3 { font-size:14px; font-weight:700; color:var(--primary); margin:18px 0 10px; padding-left:8px; border-left:3px solid var(--accent); }
  table { width:100%; border-collapse:collapse; font-size:13px; }
  th,td { padding:9px 12px; text-align:left; border-bottom:1px solid var(--border); }
  th { background:#eef2f7; font-weight:700; color:var(--primary); font-size:12px; letter-spacing:.4px; }
  tr:last-child td { border-bottom:none; }
  .nr { text-align:right; font-variant-numeric:tabular-nums; }
  .good { color:var(--success); font-weight:700; }
  .warn { color:var(--warning); font-weight:700; }
  .bad { color:var(--danger); font-weight:700; }
  .tag { display:inline-block; padding:2px 9px; border-radius:3px; font-size:11px; font-weight:700; letter-spacing:.4px; }
  .tg { background:#e8f5e9; color:#1e6b30; }
  .ty { background:#fef9e7; color:#976009; }
  .tr { background:#fdedec; color:#9b2020; }
  .tb { background:#e8f2fb; color:#154c82; }
  .tgray { background:#f0f2f5; color:#555; }
  .info-grid { display:grid; grid-template-columns:1fr 1fr; gap:3px 28px; font-size:13px; }
  .info-grid .item { display:flex; padding:5px 0; border-bottom:1px dotted var(--border); }
  .info-grid .key { color:var(--text-light); min-width:108px; flex-shrink:0; }
  .info-grid .val { font-weight:600; }
  .kpi-row { display:grid; grid-template-columns:repeat(auto-fit,minmax(165px,1fr)); gap:14px; margin-bottom:20px; }
  .kpi-card { background:#f7f9fc; border-radius:7px; padding:16px 18px; text-align:center; border-left:3px solid var(--info); }
  .kpi-card.g { border-left-color:var(--success); }
  .kpi-card.y { border-left-color:var(--warning); }
  .kpi-card.r { border-left-color:var(--danger); }
  .kpi-card .kv { font-size:22px; font-weight:700; color:var(--primary); }
  .kpi-card .ks { font-size:12px; color:var(--text-light); margin-top:2px; }
  .kpi-card .kt { font-size:11px; margin-top:3px; }
  .risk-matrix { display:grid; gap:7px; }
  .risk-row { display:flex; align-items:center; padding:9px 14px; background:#f8fafc; border-radius:5px; font-size:13px; }
  .risk-row .rt { width:145px; font-weight:700; flex-shrink:0; }
  .risk-row .rbw { flex:1; height:8px; background:#dde4ec; border-radius:4px; margin:0 14px; overflow:hidden; }
  .risk-row .rb { height:100%; border-radius:4px; }
  .risk-row .rl { width:52px; flex-shrink:0; font-weight:700; font-size:12px; text-align:right; }
  .info-box { background:#f0f5fc; border-left:4px solid var(--accent); border-radius:6px; padding:16px 20px; margin-top:14px; font-size:13px; }
  .warn-box { background:#fef9ec; border-left:4px solid var(--warning); border-radius:6px; padding:16px 20px; margin-top:14px; font-size:13px; }
  .conclusion-box { background:linear-gradient(135deg,#eef3fb,#e4edf8); border-left:4px solid var(--primary); border-radius:7px; padding:22px 26px; margin-top:22px; }
  .conclusion-box h4 { font-size:15px; color:var(--primary); margin-bottom:10px; }
  .conclusion-box ul { list-style:none; }
  .conclusion-box li { padding:5px 0 5px 18px; position:relative; font-size:13px; line-height:1.75; }
  .conclusion-box li::before { content:''; position:absolute; left:0; top:12px; width:6px; height:6px; border-radius:50%; background:var(--accent); }
  .timeline { position:relative; padding-left:24px; }
  .timeline::before { content:''; position:absolute; left:7px; top:4px; bottom:4px; width:2px; background:var(--border); }
  .tl-item { position:relative; padding:3px 0 12px 22px; font-size:13px; }
  .tl-item::before { content:''; position:absolute; left:-19px; top:8px; width:10px; height:10px; border-radius:50%; background:var(--gold); }
  .tl-item .date { font-size:11px; color:var(--text-light); margin-bottom:1px; }
  .disclaimer-box { background:#f9f0e0; border:1px dashed #c8941a; border-radius:6px; padding:14px 20px; margin-top:20px; font-size:12px; color:#7a5a0a; line-height:1.8; }
  .disclaimer-box strong { color:#5a3d00; }
  .report-footer { text-align:center; padding:28px 0; color:var(--text-light); font-size:12px; border-top:1px solid var(--border); margin-top:34px; }
  .back-btn { display:inline-block; margin-top:16px; padding:8px 24px; background:var(--accent); color:#fff; text-decoration:none; border-radius:4px; font-size:13px; }
  .back-btn:hover { opacity:0.9; }
  @media print {
    body { background:#fff; }
    .card { box-shadow:none; border:1px solid #ddd; page-break-inside:avoid; }
    .cover,.rating-banner,.rating-left { -webkit-print-color-adjust:exact; print-color-adjust:exact; }
    .back-btn { display:none; }
  }
</style>
</head>
<body>

<!-- COVER -->
<div class="cover">
  <div class="cover-tag">授 信 评 审 报 告</div>
  <h1>${companyFullName}</h1>
  <div class="subj">${reportType}</div>
  <div class="cover-meta">
    <span><span class="lbl">报告日期</span>${dateStr}</span>
    <span><span class="lbl">数据截止</span>${dateShort}</span>
    <span><span class="lbl">密级</span>机密·仅限内部审批</span>
    <span><span class="lbl">申请机构</span>${branch}</span>
    ${creditCode ? `<span><span class="lbl">统一社会信用代码</span>${creditCode}</span>` : ''}
  </div>
</div>

<div class="container">

<!-- RATING BANNER -->
<div class="rating-banner">
  <div class="rating-left">
    <div class="rating-grade">${rating.grade}</div>
    <div class="rating-label">综合授信评级</div>
  </div>
  <div class="rating-right">
    ${displayAssets ? `<div class="r-item"><div class="val">${fmtMoneyCN(displayAssets)}</div><div class="lbl">总资产${financePeriod ? '（'+financePeriod+'）' : ''}</div></div><div class="r-div"></div>` : ''}
    ${displayNetMargin ? `<div class="r-item"><div class="val">${fmtPct(displayNetMargin)}</div><div class="lbl">净利率</div></div><div class="r-div"></div>` : ''}
    ${displayRevenue ? `<div class="r-item"><div class="val">${fmtMoneyCN(displayRevenue)}</div><div class="lbl">营业收入</div></div><div class="r-div"></div>` : ''}
    ${employees ? `<div class="r-item"><div class="val">${parseInt(employees).toLocaleString()}人</div><div class="lbl">参保员工</div></div><div class="r-div"></div>` : ''}
    <div class="r-item"><div class="val">${dishonestCount + highConsumptionCount + bankruptcyCount + defaultCount} 项</div><div class="lbl">失信/限高/破产/违约</div></div>
    <div class="r-div"></div>
    <div class="r-item"><div class="val">${rating.score >= 88 ? '强烈推荐' : rating.score >= 72 ? '推荐' : rating.score >= 55 ? '审慎推荐' : '不建议'}</div><div class="lbl">授信态度</div></div>
  </div>
</div>

<!-- 1. 授信概要 -->
<section>
  <div class="card">
    <h2><span class="num">1</span> 授信评审概要</h2>
    <p style="font-size:14px;margin-bottom:16px;">
      ${companyFullName}（下称"被评企业"）${industry ? '属于' + industry + '行业' : ''}。本报告通过企查查工商数据、财务数据及风险扫描数据，对被评企业的主体质量、偿债能力与授信价值进行全维度评估。${hasFinance ? '财务数据来源：' + (financeSource || 'QCC数据库') + (financePeriod ? '（'+financePeriod+'）' : '') + '。' : '该企业工商年报选择不公示财务数据，或暂未获取到财务信息。'}
    </p>
    <div class="kpi-row">
      <div class="kpi-card g"><div class="kv">${rating.grade}</div><div class="ks">综合授信评级</div><div class="kt" style="color:var(--success);">${rating.score >= 88 ? '强烈推荐授信' : rating.score >= 72 ? '推荐授信' : '审慎评估'}</div></div>
      <div class="kpi-card ${dishonestCount + highConsumptionCount + bankruptcyCount > 0 ? 'r' : 'g'}"><div class="kv">${dishonestCount + highConsumptionCount + bankruptcyCount}</div><div class="ks">失信/限高/破产</div><div class="kt" style="color:${dishonestCount + highConsumptionCount + bankruptcyCount > 0 ? 'var(--danger)' : 'var(--success)'};">${dishonestCount + highConsumptionCount + bankruptcyCount > 0 ? '⚠️ 存在风险' : '致命风险全部通过'}</div></div>
      ${displayRevenue ? `<div class="kpi-card g"><div class="kv">${fmtMoneyCN(displayRevenue)}</div><div class="ks">营业收入</div><div class="kt" style="color:var(--success);">${finRevenueGrowth ? '同比'+fmtPct(finRevenueGrowth) : ''}</div></div>` : ''}
      ${displayNetProfit ? `<div class="kpi-card g"><div class="kv">${fmtMoneyCN(displayNetProfit)}</div><div class="ks">净利润</div><div class="kt" style="color:var(--success);">净利率${fmtPct(displayNetMargin)}</div></div>` : ''}
      <div class="kpi-card g"><div class="kv">${rating.score}</div><div class="ks">综合评分（满分100）</div><div class="kt" style="color:var(--success);">${rating.label}</div></div>
    </div>
    ${!hasFinance ? `<div class="info-box"><strong>⚠️ 重要说明：</strong>该企业工商年报财务数据选择不公示，或QCC数据库中暂无该企业财务数据。建议通过企业自主提供的财务报表进行补充评估。</div>` : ''}
    ${data.errors.length > 0 ? `<div class="warn-box"><strong>⚠️ 查询提示：</strong>以下数据源查询异常：${data.errors.slice(0,5).join('；')}${data.errors.length > 5 ? '等' + data.errors.length + '项' : ''}。部分报告内容可能不完整。</div>` : ''}
  </div>
</section>

<!-- 2. 企业基本信息 -->
<section>
  <div class="card">
    <h2><span class="num">2</span> 企业基本信息</h2>
    <div class="info-grid">
      <div class="item"><span class="key">企业全称</span><span class="val">${companyFullName}</span></div>
      <div class="item"><span class="key">统一社会信用代码</span><span class="val">${creditCode || '——'}</span></div>
      <div class="item"><span class="key">法定代表人</span><span class="val">${legalPerson || '——'}</span></div>
      <div class="item"><span class="key">注册资本</span><span class="val">${regCapital || '——'}</span></div>
      <div class="item"><span class="key">成立日期</span><span class="val">${estDate || '——'}</span></div>
      <div class="item"><span class="key">企业类型</span><span class="val">${companyType || '——'}</span></div>
      <div class="item"><span class="key">注册地址</span><span class="val">${regAddr || '——'}</span></div>
      <div class="item"><span class="key">经营状态</span><span class="val">${status ? tag(status.includes('存续')||status.includes('在营')?'tg':'ty', status) : '——'}</span></div>
      <div class="item"><span class="key">所属行业</span><span class="val">${industry || '——'}</span></div>
      <div class="item"><span class="key">参保人数</span><span class="val">${employees || '——'}</span></div>
      <div class="item"><span class="key">所属地区</span><span class="val">${region || '——'}</span></div>
      ${patentCount + swCount > 0 ? `<div class="item"><span class="key">知识产权</span><span class="val">专利 ${patentCount} 条 · 软著 ${swCount} 条</span></div>` : ''}
    </div>
    ${businessScope ? `<h3>主要经营范围</h3><p style="font-size:13px;line-height:2;">${businessScope}</p>` : ''}
    ${historyChanges.length > 0 ? `
    <h3>工商变更记录</h3>
    <div class="timeline">
      ${historyChanges.slice(0, 8).map(c => {
        const cd = c['变更日期'] || c['date'] || '';
        const ci = c['变更项目'] || c['变更事项'] || c['changeItem'] || '';
        const cb = Array.isArray(c['变更前内容']) ? c['变更前内容'].join('，') : (c['变更前内容'] || c['变更前'] || c['before'] || '');
        const ca = Array.isArray(c['变更后内容']) ? c['变更后内容'].join('，') : (c['变更后内容'] || c['变更后'] || c['after'] || '');
        const arrow = cb && ca ? ' → ' : '';
        return `<div class="tl-item"><div class="date">${cd}</div>${ci}：${cb}${arrow}${ca}</div>`;
      }).join('')}
      ${historyChanges.length > 8 ? `<div class="tl-item"><div class="date">...</div>共 ${historyChanges.length} 条变更记录</div>` : ''}
    </div>` : ''}
  </div>
</section>

<!-- 3. 股权结构与主要人员 -->
<section>
  <div class="card">
    <h2><span class="num">3</span> 股权结构与主要人员</h2>
    ${shareholderList.length > 0 ? `
    <h3>股东信息</h3>
    <table>
      <thead><tr><th>股东名称</th><th>持股比例</th><th>认缴出资额</th><th>股东类型</th></tr></thead>
      <tbody>
        ${shareholderList.slice(0, 10).map(s => `
          <tr>
            <td>${s['股东名称'] || s['name'] || '——'}</td>
            <td class="nr">${s['持股比例'] || s['ratio'] || '——'}</td>
            <td class="nr">${s['认缴出资额'] || s['amount'] || '——'}${s['认缴出资额'] ? '万元' : ''}</td>
            <td>${s['股东类型'] || s['type'] || '——'}</td>
          </tr>
        `).join('')}
      </tbody>
    </table>` : '<p style="color:var(--text-light);">股东信息待查询</p>'}
    ${acName ? `
    <h3>实际控制人</h3>
    <div class="info-box"><strong>实际控制人：</strong>${acName}</div>` : ''}
    ${boList.length > 0 ? `
    <h3>受益所有人 (UBO)</h3>
    <table>
      <thead><tr><th>姓名</th><th>受益比例</th><th>识别方式</th></tr></thead>
      <tbody>
        ${boList.slice(0, 5).map(b => `
          <tr><td>${b['姓名'] || b['name'] || '——'}</td><td class="nr">${b['受益比例'] || b['ratio'] || '——'}</td><td>${b['识别方式'] || b['method'] || '——'}</td></tr>
        `).join('')}
      </tbody>
    </table>` : ''}
    ${mergedPersonnelList.length > 0 ? `
    <h3>主要管理人员</h3>
    <table>
      <thead><tr><th>姓名</th><th>职务</th></tr></thead>
      <tbody>
        ${mergedPersonnelList.slice(0, 10).map(e => `
          <tr><td>${e['姓名'] || e['name'] || e['Name'] || '——'}</td><td>${e['职务'] || e['position'] || e['职位'] || e['jobTitle'] || '——'}</td></tr>
        `).join('')}
      </tbody>
    </table>` : ''}
    ${invList.length > 0 ? `
    <h3>对外投资</h3>
    <table>
      <thead><tr><th>被投资企业</th><th>持股比例</th><th>状态</th></tr></thead>
      <tbody>
        ${invList.slice(0, 10).map(i => `
          <tr><td>${i['被投资企业名称'] || i['name'] || '——'}</td><td class="nr">${i['持股比例'] || i['ratio'] || '——'}</td><td>${i['状态'] || i['status'] || '——'}</td></tr>
        `).join('')}
      </tbody>
    </table>` : ''}
  </div>
</section>

<!-- 4. 财务分析 -->
<section>
  <div class="card">
    <h2><span class="num">4</span> 财务分析</h2>
    ${hasFinance ? `
    ${!finRevenue && arRevenue ? `<div class="info-box" style="margin-bottom:18px;">该企业财务数据来自工商年报（${arYear}年度），部分企业选择不公示全部财务指标，数据仅供参考。</div>` : ''}
    <h3>关键财务指标</h3>
    <table>
      <thead><tr><th>指标</th><th>金额</th><th>说明</th></tr></thead>
      <tbody>
        ${displayRevenue ? `<tr><td>营业收入</td><td class="nr">${fmtMoneyCN(displayRevenue)}</td><td>${finRevenue ? financePeriod : arYear + '年度年报'}</td></tr>` : ''}
        ${displayNetProfit ? `<tr><td>净利润</td><td class="nr good">${fmtMoneyCN(displayNetProfit)}</td><td>${displayNetMargin ? '净利率 ' + fmtPct(displayNetMargin) : ''}</td></tr>` : ''}
        ${displayAssets ? `<tr><td>总资产</td><td class="nr">${fmtMoneyCN(displayAssets)}</td><td></td></tr>` : ''}
        ${displayLiabilities ? `<tr><td>总负债</td><td class="nr">${fmtMoneyCN(displayLiabilities)}</td><td></td></tr>` : ''}
        ${displayEquity ? `<tr><td>净资产（所有者权益）</td><td class="nr">${fmtMoneyCN(displayEquity)}</td><td></td></tr>` : ''}
        ${displayAssetRatio ? `<tr><td>资产负债率</td><td class="nr ${parseFloat(displayAssetRatio)<50?'good':parseFloat(displayAssetRatio)<70?'warn':'bad'}">${fmtPct(displayAssetRatio)}</td><td>${parseFloat(displayAssetRatio)<40?'安全区间':parseFloat(displayAssetRatio)<60?'正常水平':parseFloat(displayAssetRatio)<75?'偏高':'需关注'}</td></tr>` : ''}
        ${finROE ? `<tr><td>净资产收益率 (ROE)</td><td class="nr good">${fmtPct(finROE)}</td><td></td></tr>` : ''}
        ${finRevenueGrowth ? `<tr><td>营收同比增长</td><td class="nr ${parseFloat(finRevenueGrowth)>0?'good':'bad'}">${fmtPct(finRevenueGrowth)}</td><td></td></tr>` : ''}
      </tbody>
    </table>
    <p style="font-size:12px;color:var(--text-light);margin-top:6px;">※ 数据来源：${financeSource || '工商年报'}${financePeriod ? '（'+financePeriod+'）' : ''}${arYear ? '（'+arYear+'年度）' : ''}</p>` : `
    <div class="info-box"><strong>⚠️ 财务数据未获取：</strong>该企业工商年报选择不公示财务数据，或QCC数据库中暂无该企业财务数据。建议向企业索取经审计的财务报表进行补充评估。</div>`}
  </div>
</section>

<!-- 5. 信用状况与风险扫描 -->
<section>
  <div class="card">
    <h2><span class="num">5</span> 信用状况与风险扫描</h2>
    ${riskScanList.length > 0 ? `
    <h3>风险扫描总览</h3>
    <table>
      <thead><tr><th>风险类型</th><th>条目数</th><th>严重程度</th><th>评审意见</th></tr></thead>
      <tbody>
        <tr><td>失信被执行人</td><td class="nr ${dishonestCount>0?'bad':'good'}">${dishonestCount}</td><td>${dishonestCount>0?tag('tr','一票否决项'):tag('tg','一票否决项')}</td><td class="${dishonestCount>0?'bad':'good'}">${dishonestCount>0?'⚠️ 存在记录':'通过'}</td></tr>
        <tr><td>限制高消费</td><td class="nr ${highConsumptionCount>0?'bad':'good'}">${highConsumptionCount}</td><td>${highConsumptionCount>0?tag('tr','一票否决项'):tag('tg','一票否决项')}</td><td class="${highConsumptionCount>0?'bad':'good'}">${highConsumptionCount>0?'⚠️ 存在记录':'通过'}</td></tr>
        <tr><td>破产重整</td><td class="nr ${bankruptcyCount>0?'bad':'good'}">${bankruptcyCount}</td><td>${bankruptcyCount>0?tag('tr','一票否决项'):tag('tg','一票否决项')}</td><td class="${bankruptcyCount>0?'bad':'good'}">${bankruptcyCount>0?'⚠️ 存在记录':'通过'}</td></tr>
        <tr><td>股权冻结</td><td class="nr ${equityFreezeCount>0?'bad':'good'}">${equityFreezeCount}</td><td>${equityFreezeCount>0?tag('tr','一票否决项'):tag('tg','一票否决项')}</td><td class="${equityFreezeCount>0?'bad':'good'}">${equityFreezeCount>0?'⚠️ 存在记录':'通过'}</td></tr>
        <tr><td>违约事项</td><td class="nr ${defaultCount>0?'bad':'good'}">${defaultCount}</td><td>${defaultCount>0?tag('tr','一票否决项'):tag('tg','一票否决项')}</td><td class="${defaultCount>0?'bad':'good'}">${defaultCount>0?'⚠️ 存在记录':'通过'}</td></tr>
        <tr><td>行政处罚</td><td class="nr ${adminPenaltyCount>0?'warn':'good'}">${adminPenaltyCount}</td><td>${adminPenaltyCount>0?tag('ty','关注项'):tag('tb','监控项')}</td><td class="${adminPenaltyCount>0?'warn':'good'}">${adminPenaltyCount>0?'需关注':'通过'}</td></tr>
        <tr><td>被执行人</td><td class="nr ${judgmentDebtorCount>0?'warn':'good'}">${judgmentDebtorCount}</td><td>${judgmentDebtorCount>0?tag('ty','关注项'):tag('tb','监控项')}</td><td class="${judgmentDebtorCount>0?'warn':'good'}">${judgmentDebtorCount>0?'需关注':'通过'}</td></tr>
        <tr><td>裁判文书</td><td class="nr ${judicialDocsCount>0?'warn':'good'}">${judicialDocsCount}</td><td>${judicialDocsCount>100?tag('ty','关注'):tag('tb','常规')}</td><td class="${judicialDocsCount>0?'warn':'good'}">${judicialDocsCount>100?'量级较大需关注':'通过'}</td></tr>
        <tr><td>立案信息</td><td class="nr ${caseFilingCount>0?'warn':'good'}">${caseFilingCount}</td><td>${caseFilingCount>100?tag('ty','关注'):tag('tb','常规')}</td><td class="${caseFilingCount>0?'warn':'good'}">${caseFilingCount>100?'量级较大需关注':'通过'}</td></tr>
        <tr><td>开庭公告</td><td class="nr ${hearingNoticeCount>0?'warn':'good'}">${hearingNoticeCount}</td><td>${hearingNoticeCount>100?tag('ty','关注'):tag('tb','常规')}</td><td class="${hearingNoticeCount>0?'warn':'good'}">${hearingNoticeCount>100?'量级较大需关注':'通过'}</td></tr>
        <tr><td>股权出质</td><td class="nr ${equityPledgeCount>0?'warn':'good'}">${equityPledgeCount}</td><td>${equityPledgeCount>0?tag('tb','正常'):tag('tb','正常')}</td><td>${equityPledgeCount>0?'正常担保安排':'通过'}</td></tr>
        <tr><td>经营异常</td><td class="nr ${bizExceptionCount>0?'bad':'good'}">${bizExceptionCount}</td><td>${bizExceptionCount>0?tag('tr','严重'):tag('tb','监控项')}</td><td class="${bizExceptionCount>0?'bad':'good'}">${bizExceptionCount>0?'⚠️ 存在记录':'通过'}</td></tr>
        <tr><td>税务违法/欠税</td><td class="nr ${taxViolationCount+taxArrearsCount>0?'bad':'good'}">${taxViolationCount+taxArrearsCount}</td><td>${taxViolationCount+taxArrearsCount>0?tag('tr','严重'):tag('tb','监控项')}</td><td class="${taxViolationCount+taxArrearsCount>0?'bad':'good'}">${taxViolationCount+taxArrearsCount>0?'⚠️ 存在记录':'通过'}</td></tr>
        <tr><td>终本案件</td><td class="nr ${terminatedCount>0?'warn':'good'}">${terminatedCount}</td><td>${terminatedCount>0?tag('ty','关注'):tag('tb','监控项')}</td><td class="${terminatedCount>0?'warn':'good'}">${terminatedCount>0?'需关注':'通过'}</td></tr>
      </tbody>
    </table>` : '<p style="color:var(--text-light);">风险扫描数据待查询</p>'}
    ${honorList.length > 0 ? `
    <h3>荣誉资质</h3>
    <div class="info-box">${honorList.slice(0,5).map(h => h['荣誉名称'] || h['name'] || h['奖项'] || '').filter(Boolean).join('；')}</div>` : ''}
    ${qualificationList.length > 0 ? `
    <h3>资质证书</h3>
    <p style="font-size:13px;">共 ${qualificationList.length} 项资质证书</p>` : ''}
  </div>
</section>

<!-- 6. 经营与知识产权 -->
<section>
  <div class="card">
    <h2><span class="num">6</span> 经营与知识产权</h2>
    ${patentCount + swCount > 0 ? `
    <h3>知识产权概况</h3>
    <table>
      <thead><tr><th>类型</th><th>数量</th><th>说明</th></tr></thead>
      <tbody>
        ${patentCount > 0 ? `<tr><td>专利</td><td class="nr good">${patentCount}</td><td>反映企业技术创新能力</td></tr>` : ''}
        ${swCount > 0 ? `<tr><td>软件著作权</td><td class="nr good">${swCount}</td><td>反映企业软件开发能力</td></tr>` : ''}
      </tbody>
    </table>` : '<p style="color:var(--text-light);">知识产权信息待查询</p>'}
    ${biddingList.length > 0 ? `
    <h3>招投标信息</h3>
    <p style="font-size:13px;">共 ${biddingList.length} 条招投标记录，反映企业经营活跃度</p>` : ''}
  </div>
</section>

<!-- 7. 风险因素分析 -->
<section>
  <div class="card">
    <h2><span class="num">7</span> 风险因素分析</h2>
    <div class="risk-matrix">
      <div class="risk-row"><div class="rt">致命风险（一票否决）</div><div class="rbw"><div class="rb" style="width:${Math.min(dishonestCount+highConsumptionCount+bankruptcyCount+equityFreezeCount*10,100)}%;background:${dishonestCount+highConsumptionCount+bankruptcyCount+equityFreezeCount>0?'var(--danger)':'var(--success)'};"></div></div><div class="rl ${dishonestCount+highConsumptionCount+bankruptcyCount+equityFreezeCount>0?'bad':'good'}">${dishonestCount+highConsumptionCount+bankruptcyCount+equityFreezeCount>0?'存在风险':'零风险'}</div></div>
      <div class="risk-row"><div class="rt">财务偿债风险</div><div class="rbw"><div class="rb" style="width:${displayAssetRatio?Math.min(parseFloat(displayAssetRatio),100):30}%;background:${displayAssetRatio&&parseFloat(displayAssetRatio)>60?'var(--warning)':'var(--success)'};"></div></div><div class="rl ${displayAssetRatio&&parseFloat(displayAssetRatio)>60?'warn':'good'}">${displayAssetRatio&&parseFloat(displayAssetRatio)>60?'偏高':displayAssetRatio?'可控':'待评估'}</div></div>
      <div class="risk-row"><div class="rt">司法诉讼风险</div><div class="rbw"><div class="rb" style="width:${Math.min(judicialDocsCount/20,100)}%;background:${judicialDocsCount>100?'var(--warning)':'var(--success)'};"></div></div><div class="rl ${judicialDocsCount>100?'warn':'good'}">${judicialDocsCount>500?'较高':judicialDocsCount>100?'可控':judicialDocsCount>0?'较低':'无记录'}</div></div>
      <div class="risk-row"><div class="rt">经营稳定性风险</div><div class="rbw"><div class="rb" style="width:${bizExceptionCount>0?40:5}%;background:${bizExceptionCount>0?'var(--warning)':'var(--success)'};"></div></div><div class="rl ${bizExceptionCount>0?'warn':'good'}">${bizExceptionCount>0?'需关注':'低'}</div></div>
      <div class="risk-row"><div class="rt">税务合规风险</div><div class="rbw"><div class="rb" style="width:${taxViolationCount+taxArrearsCount>0?50:5}%;background:${taxViolationCount+taxArrearsCount>0?'var(--danger)':'var(--success)'};"></div></div><div class="rl ${taxViolationCount+taxArrearsCount>0?'bad':'good'}">${taxViolationCount+taxArrearsCount>0?'存在风险':'低'}</div></div>
      <div class="risk-row"><div class="rt">行政处罚风险</div><div class="rbw"><div class="rb" style="width:${Math.min(adminPenaltyCount*10,100)}%;background:${adminPenaltyCount>3?'var(--warning)':'var(--success)'};"></div></div><div class="rl ${adminPenaltyCount>3?'warn':'good'}">${adminPenaltyCount>5?'较高':adminPenaltyCount>0?'较低':'无记录'}</div></div>
    </div>
  </div>
</section>

<!-- 8. 综合授信评级与建议 -->
<section>
  <div class="card">
    <h2><span class="num">8</span> 综合授信评级与建议</h2>
    <h3>授信评级：${rating.grade}（${rating.label}）</h3>
    <table>
      <thead><tr><th>评审维度</th><th>权重</th><th>核心依据</th></tr></thead>
      <tbody>
        <tr><td>主体真实性与合规性</td><td class="nr">15%</td><td>工商登记合规${creditCode?'，USCC核验通过':''}，经营状态${status||'正常'}</td></tr>
        <tr><td>偿债能力</td><td class="nr">25%</td><td>${displayAssetRatio?('资产负债率'+fmtPct(displayAssetRatio)):'财务数据未公开'}${displayAssets?'，总资产'+fmtMoneyCN(displayAssets):''}</td></tr>
        <tr><td>盈利能力与现金流</td><td class="nr">20%</td><td>${displayNetProfit?('净利润'+fmtMoneyCN(displayNetProfit)+'，净利率'+fmtPct(displayNetMargin)):'财务数据未公开'}</td></tr>
        <tr><td>经营稳定性</td><td class="nr">15%</td><td>${estDate?'自'+estDate.substring(0,4)+'年成立至今':'经营年限未知'}${employees?'，参保'+employees+'人':''}</td></tr>
        <tr><td>信用记录</td><td class="nr">15%</td><td>${dishonestCount+highConsumptionCount+bankruptcyCount>0?'存在不良信用记录':'零失信/限高/破产记录'}</td></tr>
        <tr><td>司法风险</td><td class="nr">5%</td><td>${judicialDocsCount>0?('裁判文书'+judicialDocsCount+'件'):'无诉讼记录'}</td></tr>
        <tr><td>治理与背景</td><td class="nr">5%</td><td>${acName?('实控人'+acName):''}${legalPerson?('，法定代表人'+legalPerson):''}</td></tr>
      </tbody>
    </table>

    <div class="conclusion-box" style="margin-top:18px;">
      <h4>综合评审结论</h4>
      <ul>
        <li><strong>授信态度：${rating.score >= 88 ? '强烈推荐授信' : rating.score >= 72 ? '推荐授信' : rating.score >= 55 ? '审慎推荐授信' : '不建议授信'}。</strong>${rating.score >= 88 ? '该企业各项指标表现优异，致命风险全部为零，属于银行授信意愿最强的优质客户群体。' : rating.score >= 72 ? '该企业整体表现良好，核心风险指标可控，具备较好的偿债能力和经营稳定性。' : rating.score >= 55 ? '该企业存在一定的风险因素，建议在落实风险缓释措施的前提下审慎开展授信。' : '该企业存在较为突出的风险因素，建议暂缓授信决策，待进一步尽调后再行评估。'}</li>
        <li><strong>建议授信额度：</strong>${rating.score >= 88 ? '根据企业资产规模和经营情况，建议给予较高授信额度，具体额度视业务需求、授信期限和担保方式协商确定。' : rating.score >= 72 ? '建议给予适当授信额度，控制在合理范围内，视企业提供的担保增信措施可适当上调。' : '建议以较低额度起步，要求提供足额担保或抵质押措施。'}</li>
        <li><strong>担保方式：</strong>${rating.score >= 88 ? '信用贷款优先，必要时辅以保证担保' : rating.score >= 72 ? '信用 + 保证担保' : '抵押/质押 + 保证担保'}</li>
        <li><strong>贷后管理：</strong>纳入客户定期监控，${rating.score >= 80 ? '年度审查一次' : '半年度审查一次'}，关注工商变更及风险信息动态。</li>
      </ul>
    </div>

    <div class="conclusion-box" style="margin-top:14px;background:linear-gradient(135deg,#fef9ec,#fdf2d0);border-left-color:var(--warning);">
      <h4>持续监测要点</h4>
      <ul>
        <li>定期关注企业工商信息变更，特别是法定代表人、注册资本、股东结构的重大变化。</li>
        <li>持续监控失信被执行人、限制高消费、破产重整等一票否决风险指标。</li>
        <li>关注企业经营异常、行政处罚、税务违法等合规性风险。</li>
        <li>如触发以下预警，立即启动贷后评估：（a）出现失信/限高/破产记录；（b）出现经营异常或严重行政处罚；（c）实际控制人或主要高管发生重大变更；（d）核心财务指标出现断崖式下滑。</li>
        <li>建议每半年更新一次企业信用报告，保持数据时效性。</li>
      </ul>
    </div>
  </div>
</section>

<!-- 9. 数据说明与免责声明 -->
<section>
  <div class="card">
    <h2><span class="num">9</span> 数据说明与免责声明</h2>
    <div class="disclaimer-box">
      <strong>数据来源说明：</strong><br>
      1. 本报告工商登记信息来源：企查查工商数据（企查查科技股份有限公司），数据截止${dateStr}。<br>
      2. 财务数据来源：${financeSource || '企查查工商年报数据'}${financePeriod ? '（'+financePeriod+'）' : ''}${arYear ? '（'+arYear+'年度年报）' : ''}。<br>
      3. 风险扫描数据来源：企查查风险监控数据库，涵盖失信被执行、限高消费、股权冻结、破产重整等34类风险因子。<br>
      4. 本报告数据基于企查查QCC MCP实时接口查询，部分数据可能因企业选择不公示或数据库覆盖限制而缺失。<br><br>
      <strong>免责声明：</strong>本报告基于截至报告日前可获取的公开信息及企查查数据库数据编制，不构成最终授信决策依据。具体授信条件以贷款行信贷审批委员会决议为准。如被评企业提供经审计财务报告，建议以审计报告为优先参考依据。
    </div>
  </div>
</section>

<div class="report-footer">
  <p>本报告由企查查 QCC MCP 数据驱动生成 · 数据截止日期：${dateStr}</p>
  <p>报告仅供内部参考，不构成最终授信决策依据 · 密级：机密</p>
  <p style="margin-top:6px;opacity:.7;">生成时间：${timeStr} · ${reportType} · ${companyFullName}</p>
  <a href="/index.html" class="back-btn">← 返回查询页面</a>
</div>

</div>
</body>
</html>`;

  return html;
}

// ============ API 路由 ============

// POST /api/query — 生成授信报告 HTML（直接返回完整报告页面）
app.post('/api/query', async (req, res) => {
  try {
  const { companyName, branch, skill } = req.body;
  if (!companyName) {
    return res.status(400).json({ success: false, error: '缺少企业名称' });
  }
  if (!branch) {
    return res.status(400).json({ success: false, error: '缺少银行支行名称' });
  }

  console.log(`\n${'='.repeat(50)}`);
  console.log(`[API] 查询企业: ${companyName} | 机构: ${branch}${skill ? ' | 技能: ' + skill : ''}`);
  console.log(`${'='.repeat(50)}`);

  const data = { errors: [] };

  // 阶段1: 串行获取工商注册信息（避免 QCC 并发限流导致"无匹配项"）
  // QCC company server 对快速连续请求有严格限流，加入延迟并支持重试
  console.log(`[API] 阶段1: 串行获取工商注册信息...`);
  const basicTasks = basicCompanyTasks(data, companyName);
  for (const t of basicTasks) {
    await executeTaskBatch([t], data);
    await new Promise(r => setTimeout(r, 600)); // 每个请求间隔 600ms
  }
  const uscc = data.company?.['统一社会信用代码'] || data.company?.KeyNo;
  if (uscc) console.log(`[API] 统一社会信用代码: ${uscc}`);
  else console.log(`[API] 未获取到统一社会信用代码，后续使用公司名查询`);

  // 阶段1.5: 获取工商变更记录（独立串行，避免限流）
  console.log(`[API] 阶段1.5: 获取工商变更记录...`);
  await new Promise(r => setTimeout(r, 800)); // 额外间隔
  await executeTaskBatch([task('工商变更记录', 'company', 'get_change_records',
    () => ({ searchKey: uscc || companyName }),
    r => { data.history = r; data.changeRecords = r; }
  )], data);

  // 阶段2: 其余任务并发执行
  let restTasks = [];
  switch (skill) {
    case 'credit-due-diligence':
      restTasks = [...executiveTasks(data, companyName), ...riskDetailTasks(data, companyName), ...historyTasks(data, companyName), ...uboTasks(data, companyName), ...iprTasks(data, companyName), ...operationTasks(data, companyName)];
      break;
    case 'counterparty-risk':
      restTasks = [...tradeFinanceTasks(data, companyName), ...riskDetailTasks(data, companyName), ...executiveRiskTasks(data, companyName), ...iprTasks(data, companyName)];
      break;
    case 'guarantor-check':
      restTasks = [...riskDetailTasks(data, companyName), ...executiveRiskTasks(data, companyName), ...iprTasks(data, companyName)];
      break;
    case 'kyb-verification':
      restTasks = [...riskDetailTasks(data, companyName), ...historyTasks(data, companyName), ...uboTasks(data, companyName), ...operationTasks(data, companyName), ...iprTasks(data, companyName)];
      break;
    case 'ubo-screening':
      restTasks = [...uboTasks(data, companyName), ...executiveTasks(data, companyName), ...iprTasks(data, companyName)];
      break;
    case 'business-health-scan':
      restTasks = [...operationTasks(data, companyName), ...historyTasks(data, companyName), ...riskDetailTasks(data, companyName), ...iprTasks(data, companyName)];
      break;
    case 'credit-monitoring':
      restTasks = [...riskDetailTasks(data, companyName), ...historyTasks(data, companyName), ...executiveRiskTasks(data, companyName), ...iprTasks(data, companyName)];
      break;
    case 'bankruptcy-monitor':
      restTasks = [...riskDetailTasks(data, companyName), ...historyTasks(data, companyName), ...executiveRiskTasks(data, companyName), ...iprTasks(data, companyName)];
      break;
    case 'equity-structure':
      restTasks = [...uboTasks(data, companyName), ...executiveTasks(data, companyName), ...historyTasks(data, companyName), ...iprTasks(data, companyName)];
      break;
    case 'executive-background':
      restTasks = [...executiveTasks(data, companyName), ...executiveRiskTasks(data, companyName), ...iprTasks(data, companyName)];
      break;
    case 'litigation-analysis':
      restTasks = [...riskDetailTasks(data, companyName), ...litigationTasks(data, companyName), ...historyTasks(data, companyName), ...executiveRiskTasks(data, companyName), ...iprTasks(data, companyName)];
      break;
    case 'trade-finance-compliance':
      restTasks = [...tradeFinanceTasks(data, companyName), ...riskDetailTasks(data, companyName), ...operationTasks(data, companyName), ...historyTasks(data, companyName), ...iprTasks(data, companyName)];
      break;
    default:
      // 全量查询
      restTasks = [...iprTasks(data, companyName), ...operationTasks(data, companyName), ...executiveTasks(data, companyName), ...historyTasks(data, companyName), ...riskDetailTasks(data, companyName), ...litigationTasks(data, companyName), ...uboTasks(data, companyName), ...tradeFinanceTasks(data, companyName)];
  }

  console.log(`[API] 阶段2: 执行 ${restTasks.length} 个查询任务...`);
  await executeTaskBatch(restTasks, data);

  console.log(`[API] 查询完成，${data.errors.length} 个错误，${6 + restTasks.length} 个工具\n`);

  // 生成报告 HTML 并直接返回
  const reportHTML = buildReportHTML(branch, companyName, data, skill);
  res.setHeader('Content-Type', 'text/html; charset=utf-8');
  res.send(reportHTML);

  } catch (err) {
    console.error(`[API] 致命错误:`, err.message);
    res.status(500).json({ success: false, error: '服务内部错误: ' + (err.message || '未知错误') });
  }
});

// POST /api/query-docx — 生成 docx 格式授信报告（使用相同的数据采集+不同的输出格式）
app.post('/api/query-docx', async (req, res) => {
  try {
    const { companyName, branch, skill } = req.body;
    if (!companyName) {
      return res.status(400).json({ success: false, error: '缺少企业名称' });
    }
    if (!branch) {
      return res.status(400).json({ success: false, error: '缺少银行支行名称' });
    }

    console.log(`\n${'='.repeat(50)}`);
    console.log(`[DOCX] 查询企业: ${companyName} | 机构: ${branch}${skill ? ' | 技能: ' + skill : ''}`);
    console.log(`${'='.repeat(50)}`);

    const data = { errors: [] };

    // 阶段1: 串行获取工商注册信息
    console.log(`[DOCX] 阶段1: 串行获取工商注册信息...`);
    const basicTasks = basicCompanyTasks(data, companyName);
    for (const t of basicTasks) {
      await executeTaskBatch([t], data);
      await new Promise(r => setTimeout(r, 600));
    }
    const uscc = data.company?.['统一社会信用代码'] || data.company?.KeyNo;

    // 阶段1.5: 工商变更记录
    console.log(`[DOCX] 阶段1.5: 获取工商变更记录...`);
    await new Promise(r => setTimeout(r, 800));
    await executeTaskBatch([task('工商变更记录', 'company', 'get_change_records',
      () => ({ searchKey: uscc || companyName }),
      r => { data.history = r; data.changeRecords = r; }
    )], data);

    // 阶段2: 并发
    let restTasks = [];
    switch (skill) {
      case 'credit-due-diligence':
        restTasks = [...executiveTasks(data, companyName), ...riskDetailTasks(data, companyName), ...historyTasks(data, companyName), ...uboTasks(data, companyName), ...iprTasks(data, companyName), ...operationTasks(data, companyName)];
        break;
      case 'counterparty-risk':
        restTasks = [...tradeFinanceTasks(data, companyName), ...riskDetailTasks(data, companyName), ...executiveRiskTasks(data, companyName), ...iprTasks(data, companyName)];
        break;
      case 'guarantor-check':
        restTasks = [...riskDetailTasks(data, companyName), ...executiveRiskTasks(data, companyName), ...iprTasks(data, companyName)];
        break;
      case 'kyb-verification':
        restTasks = [...riskDetailTasks(data, companyName), ...historyTasks(data, companyName), ...uboTasks(data, companyName), ...operationTasks(data, companyName), ...iprTasks(data, companyName)];
        break;
      case 'ubo-screening':
        restTasks = [...uboTasks(data, companyName), ...executiveTasks(data, companyName), ...iprTasks(data, companyName)];
        break;
      case 'business-health-scan':
        restTasks = [...operationTasks(data, companyName), ...historyTasks(data, companyName), ...riskDetailTasks(data, companyName), ...iprTasks(data, companyName)];
        break;
      case 'credit-monitoring':
        restTasks = [...riskDetailTasks(data, companyName), ...historyTasks(data, companyName), ...executiveRiskTasks(data, companyName), ...iprTasks(data, companyName)];
        break;
      case 'bankruptcy-monitor':
        restTasks = [...riskDetailTasks(data, companyName), ...historyTasks(data, companyName), ...executiveRiskTasks(data, companyName), ...iprTasks(data, companyName)];
        break;
      case 'equity-structure':
        restTasks = [...uboTasks(data, companyName), ...executiveTasks(data, companyName), ...historyTasks(data, companyName), ...iprTasks(data, companyName)];
        break;
      case 'executive-background':
        restTasks = [...executiveTasks(data, companyName), ...executiveRiskTasks(data, companyName), ...iprTasks(data, companyName)];
        break;
      case 'litigation-analysis':
        restTasks = [...riskDetailTasks(data, companyName), ...litigationTasks(data, companyName), ...historyTasks(data, companyName), ...executiveRiskTasks(data, companyName), ...iprTasks(data, companyName)];
        break;
      case 'trade-finance-compliance':
        restTasks = [...tradeFinanceTasks(data, companyName), ...riskDetailTasks(data, companyName), ...operationTasks(data, companyName), ...historyTasks(data, companyName), ...iprTasks(data, companyName)];
        break;
      default:
        restTasks = [...iprTasks(data, companyName), ...operationTasks(data, companyName), ...executiveTasks(data, companyName), ...historyTasks(data, companyName), ...riskDetailTasks(data, companyName), ...litigationTasks(data, companyName), ...uboTasks(data, companyName), ...tradeFinanceTasks(data, companyName)];
    }

    console.log(`[DOCX] 阶段2: 执行 ${restTasks.length} 个查询任务...`);
    await executeTaskBatch(restTasks, data);

    console.log(`[DOCX] 查询完成，开始生成 docx...\n`);

    // 生成 docx
    const doc = await buildDocxReport(data, companyName, branch);
    const buffer = await Packer.toBuffer(doc);

    // 生成安全文件名
    const safeName = companyName.replace(/[\\/:*?"<>|]/g, '_');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', 'attachment; filename="ICBC_Credit_Report.docx"');
    res.setHeader('Content-Length', buffer.length);
    res.send(Buffer.from(buffer));

    console.log(`[DOCX] docx 已发送，大小: ${(buffer.length / 1024).toFixed(1)} KB`);

  } catch (err) {
    console.error(`[DOCX] 致命错误:`, err.message);
    res.status(500).json({ success: false, error: '服务内部错误: ' + (err.message || '未知错误') });
  }
});

// GET /api/health
app.get('/api/health', (req, res) => {
  res.json({ status: 'ok', time: new Date().toISOString() });
});

// GET /api/skills
app.get('/api/skills', (req, res) => {
  res.json({
    skills: [
      { id: 'credit-due-diligence',     name: '授信尽调报告',       icon: '📋', desc: '信贷审批放款前全维度企业尽调' },
      { id: 'counterparty-risk',        name: '交易对手风险评估',   icon: '🤝', desc: '贸易融资/信用证/保理交易对手风险评估' },
      { id: 'guarantor-check',          name: '担保方资信核查',     icon: '🛡️', desc: '贷款担保审批前保证人偿付能力核查' },
      { id: 'kyb-verification',         name: 'KYB 企业核验',       icon: '✅', desc: '对公客户开户/授信主体自动化核验' },
      { id: 'ubo-screening',            name: '受益所有人识别',     icon: '👤', desc: '反洗钱合规场景 UBO 识别' },
      { id: 'business-health-scan',     name: '经营健康度扫描',     icon: '💊', desc: '企业经营活跃度与健康度动态跟踪' },
      { id: 'credit-monitoring',        name: '信贷风险定期监控',   icon: '📡', desc: '贷后存量借款客户持续风险监控' },
      { id: 'bankruptcy-monitor',       name: '企业破产预警监控',   icon: '⚠️', desc: '重要客户/债务人破产风险持续监控' },
      { id: 'equity-structure',         name: '股权结构穿透分析',   icon: '🔗', desc: '多层股权穿透 + UBO + 历史变迁' },
      { id: 'executive-background',     name: '高管背景核查',       icon: '🔍', desc: '董监高/实控人个人风险穿透与关联网络' },
      { id: 'litigation-analysis',      name: '诉讼风险评估',       icon: '⚖️', desc: '企业+历史+核心人员三层诉讼全景扫描' },
      { id: 'trade-finance-compliance', name: '贸易融资合规核查',   icon: '🚢', desc: '跨境贸易信用证/保理/福费廷合规准入' },
    ]
  });
});

// ============ 启动 ============
const PORT = process.env.PORT || 3000;
const HOST = process.env.HOST || '0.0.0.0';
app.listen(PORT, HOST, () => {
  console.log(`\n🚀 企业授信报告系统 (专业报告版) 已启动`);
  console.log(`   后端代理: http://${HOST}:${PORT}`);
  console.log(`   HTML报告: POST http://${HOST}:${PORT}/api/query`);
  console.log(`   DOCX下载: POST http://${HOST}:${PORT}/api/query-docx`);
  console.log(`   前端输入: http://${HOST}:${PORT}/index.html`);
  console.log(`\n📡 QCC MCP 6 大 Server | 12 Skills | HTML预览 + DOCX下载\n`);
});
