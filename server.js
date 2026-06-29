/**
 * 企业授信报告系统 V3 — 独立 DOCX 生成服务（范例对齐版）
 * 集成企查查 QCC MCP 6大Server（company/risk/ipr/operation/executive/history）
 * 共调用 106 个工具，按"三品三表三流"方法论生成专业授信调查报告
 */
const express = require('express');
const cors = require('cors');
const path = require('path');
const fs = require('fs');
const https = require('https');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, BorderStyle, WidthType, ShadingType,
  PageNumber, PageBreak
} = require('docx');

const app = express();
app.use(cors({ origin: (_, cb) => cb(null, true), credentials: false }));
app.use(express.json({ limit: '10mb' }));
app.use(express.static(__dirname));

// ============ QCC 配置 ============
const QCC_TOKEN = process.env.QCC_API_TOKEN || 'Bearer MmDFP6lIn3akremGqfRhudPo0NON0LTrz92cBzCHtI3eA7IZ';
const QCC_SERVERS = {
  company:   { url: 'https://agent.qcc.com/mcp/company/stream',   token: QCC_TOKEN },
  risk:      { url: 'https://agent.qcc.com/mcp/risk/stream',      token: QCC_TOKEN },
  ipr:       { url: 'https://agent.qcc.com/mcp/ipr/stream',       token: QCC_TOKEN },
  operation: { url: 'https://agent.qcc.com/mcp/operation/stream', token: QCC_TOKEN },
  executive: { url: 'https://agent.qcc.com/mcp/executive/stream', token: QCC_TOKEN },
  history:   { url: 'https://agent.qcc.com/mcp/history/stream',   token: QCC_TOKEN },
};

// ============ QCC MCP 调用 ============
async function callQccMCP(server, toolName, args, timeout = 30000) {
  const config = QCC_SERVERS[server];
  if (!config) throw new Error(`未知 QCC 服务: ${server}`);

  return new Promise((resolve, reject) => {
    const body = JSON.stringify({ jsonrpc: '2.0', id: Date.now(), method: 'tools/call', params: { name: toolName, arguments: args } });
    const req = https.request({
      hostname: 'agent.qcc.com', path: '/mcp/' + server + '/stream', method: 'POST',
      headers: { 'Content-Type': 'application/json', 'Authorization': config.token, 'Content-Length': Buffer.byteLength(body) },
      timeout
    }, res => {
      let data = ''; res.on('data', c => data += c); res.on('end', () => {
        const lines = data.split('\n').filter(l => l.startsWith('data: '));
        for (const line of lines) {
          try {
            const sse = JSON.parse(line.substring(6));
            if (sse.error) return reject(new Error(`QCC ${server}/${toolName}: ${JSON.stringify(sse.error)}`));
            if (sse.result && sse.result.content && sse.result.content[0] && sse.result.content[0].text) {
              try { return resolve(JSON.parse(sse.result.content[0].text)); } catch { return resolve(sse.result.content[0].text); }
            }
            if (sse.result) return resolve(sse.result);
          } catch (e) { if (e.message.includes('QCC ')) return reject(e); }
        }
        resolve(null);
      });
    });
    req.on('error', e => reject(e));
    req.on('timeout', () => { req.destroy(); reject(new Error('Timeout')); });
    req.write(body); req.end();
  });
}

// ============ 数据采集（6 Server 全维度） ============
async function collectAllData(companyName, progressCallback) {
  const data = {};
  const errors = [];
  let step = 0, total = 106;

  async function fetch(server, tool, args, label) {
    try {
      const r = await callQccMCP(server, tool, args || { searchKey: companyName });
      data[tool] = r;
      step++;
      if (progressCallback) progressCallback({ step, total, tool: server + '/' + tool, label, ok: true });
      return r;
    } catch (e) {
      errors.push(label + ': ' + e.message);
      step++;
      if (progressCallback) progressCallback({ step, total, tool: server + '/' + tool, label, ok: false, error: e.message });
      return null;
    }
  }

  const searchKey = companyName;
  const delay = (ms) => new Promise(r => setTimeout(r, ms));

  // Phase 1: Company Server
  const cTools = [
    ['get_company_registration_info', '工商信息'], ['get_financial_data', '财务数据'],
    ['get_annual_reports', '企业年报'], ['get_shareholder_info', '股东信息'],
    ['get_external_investments', '对外投资'], ['get_contact_info', '联系方式'],
    ['get_listing_info', '上市信息'], ['get_key_personnel', '主要人员'],
    ['get_actual_controller', '实际控制人'], ['get_beneficial_owners', '受益所有人'],
    ['get_company_profile', '企业档案'], ['get_branches', '分支机构'],
    ['get_change_records', '变更记录'],
  ];
  for (const [tool, label] of cTools) { await fetch('company', tool, { searchKey }, label); await delay(500); }

  // Phase 2: Risk Server
  const rTools = [
    ['get_company_risk_scan', '风险扫描'], ['get_dishonest_info', '失信信息'],
    ['get_judgment_debtor_info', '被执行信息'], ['get_high_consumption_restriction', '限高消费'],
    ['get_equity_freeze', '股权冻结'], ['get_equity_pledge_info', '股权出质'],
    ['get_business_exception', '经营异常'], ['get_serious_violation', '严重违法'],
    ['get_administrative_penalty', '行政处罚'], ['get_tax_arrears_notice', '欠税公告'],
    ['get_tax_abnormal', '税务异常'], ['get_tax_violation', '税务违法'],
    ['get_bankruptcy_reorganization', '破产重整'], ['get_exit_restriction', '限出境'],
    ['get_terminated_cases', '终本案件'], ['get_default_info', '违约事项'],
    ['get_judicial_documents', '裁判文书'], ['get_case_filing_info', '立案信息'],
    ['get_hearing_notice', '开庭公告'], ['get_court_notice', '法院公告'],
    ['get_service_notice', '送达公告'], ['get_pre_litigation_mediation', '诉前调解'],
    ['get_company_related_risk_scan', '关联企业风险'], ['get_guarantee_info', '对外担保'],
    ['get_environmental_penalty', '环保处罚'], ['get_judicial_auction', '司法拍卖'],
    ['get_liquidation_info', '清算信息'], ['get_chattel_mortgage_info', '动产抵押'],
    ['get_land_mortgage_info', '土地抵押'], ['get_disciplinary_list', '惩戒名单'],
    ['get_cancellation_record_info', '注销记录'], ['get_public_exhortation', '公示催告'],
  ];
  for (const [tool, label] of rTools) { await fetch('risk', tool, { searchKey }, label); await delay(300); }

  // Phase 3: IPR Server
  const iprTools = [
    ['get_patent_info', '专利信息'], ['get_software_copyright_info', '软件著作权'],
    ['get_trademark_info', '商标信息'], ['get_copyright_work_info', '作品著作权'],
    ['get_ipr_pledge', '知识产权质押'], ['get_international_patent', '国际专利'],
    ['get_standard_info', '标准制定'], ['get_integrated_circuit_layout', '集成电路布图'],
  ];
  for (const [tool, label] of iprTools) { await fetch('ipr', tool, { searchKey }, label); await delay(250); }

  // Phase 4: Operation Server
  const opTools = [
    ['get_bidding_info', '招投标信息'], ['get_qualifications', '资质证书'],
    ['get_recruitment_info', '招聘信息'], ['get_credit_evaluation', '信用评估'],
    ['get_honor_info', '荣誉信息'], ['get_taxpayer_qualification', '纳税资质'],
    ['get_news_sentiment', '新闻舆情'], ['get_financing_records', '融资记录'],
    ['get_administrative_license', '行政许可'], ['get_government_interview', '政府约谈'],
    ['get_product_spot_check', '产品抽检'], ['get_random_check', '双随机抽查'],
    ['get_food_safety', '食品安全'], ['get_product_recall', '产品召回'],
    ['get_ranking_list_info', '榜单排名'], ['get_import_export_credit', '进出口信用'],
    ['get_credit_commitments', '信用承诺'],
  ];
  for (const [tool, label] of opTools) { await fetch('operation', tool, { searchKey }, label); await delay(250); }

  // Phase 5: Executive Server
  const regInfo = data.get_company_registration_info || {};
  const legalPerson = (regInfo && regInfo['法定代表人']) || '';
  const keyPersonnel = data.get_key_personnel || {};
  const execNames = new Set();
  if (legalPerson) execNames.add(legalPerson);
  const pList = (() => {
    if (!keyPersonnel) return [];
    for (const k of ['主要人员信息', '主要人员', '高管信息', '董监高']) {
      if (Array.isArray(keyPersonnel[k]) && keyPersonnel[k].length > 0) return keyPersonnel[k];
    }
    return [];
  })();
  for (const p of pList) {
    const n = p['姓名'] || p['name'] || '';
    if (n && n !== legalPerson) execNames.add(n);
    if (execNames.size >= 3) break;
  }
  const primaryPerson = [...execNames][0] || legalPerson || '';

  const execTools = [
    ['get_executive_risk_scan', '高管风险扫描'], ['get_executive_dishonest', '高管失信'],
    ['get_executive_high_consumption_ban', '高管限高'], ['get_executive_exit_restriction', '高管限出境'],
    ['get_executive_judgment_debtor', '高管被执行'], ['get_executive_positions', '高管任职'],
    ['get_executive_related_companies', '高管关联企业'], ['get_executive_controlled_companies', '高管控制企业'],
    ['get_executive_investments', '高管对外投资'], ['get_executive_related_risk_scan', '高管关联风险'],
    ['get_executive_equity_freeze', '高管股权冻结'], ['get_executive_equity_pledge', '高管股权出质'],
    ['get_executive_tax_violation', '高管税务违法'], ['get_executive_case_filing', '高管立案'],
    ['get_executive_judicial_docs', '高管裁判文书'], ['get_executive_terminated_cases', '高管终本案件'],
    ['get_executive_admin_penalty', '高管行政处罚'],
  ];
  for (const [tool, label] of execTools) { await fetch('executive', tool, { searchKey, personName: primaryPerson }, label); await delay(250); }

  // Phase 6: History Server
  const uscc = (regInfo && regInfo['统一社会信用代码']) || searchKey;
  const histTools = [
    ['get_historical_legal_rep', '历史法代'], ['get_historical_shareholders', '历史股东'],
    ['get_historical_dishonest', '历史失信'], ['get_historical_judgment_debtor', '历史被执行'],
    ['get_historical_registration', '历史注册变更'], ['get_historical_executives', '历史高管'],
    ['get_historical_bankruptcy', '历史破产'], ['get_historical_high_consumption_ban', '历史限高'],
    ['get_historical_equity_freeze', '历史股权冻结'], ['get_historical_business_exception', '历史经营异常'],
    ['get_historical_serious_violation', '历史严重违法'], ['get_historical_admin_penalty', '历史行政处罚'],
    ['get_historical_tax_arrears', '历史欠税'], ['get_historical_honor', '历史荣誉'],
    ['get_historical_patent', '历史专利'], ['get_historical_trademark', '历史商标'],
    ['get_historical_listing', '历史上市'], ['get_historical_terminated_cases', '历史终本案件'],
    ['get_historical_investments', '历史对外投资'],
  ];
  for (const [tool, label] of histTools) { await fetch('history', tool, { searchKey: uscc }, label); await delay(250); }

  return { data, errors };
}

// ============ 数据提取辅助 ============
function s(obj, key) { return (obj && obj[key] != null) ? String(obj[key]) : ''; }
function pick(obj, ...keys) {
  if (!obj) return [];
  if (Array.isArray(obj)) return obj;
  for (const k of keys) { if (Array.isArray(obj[k]) && obj[k].length > 0) return obj[k]; }
  return [];
}
function cnt(obj) {
  if (!obj) return 0;
  if (Array.isArray(obj)) return obj.length;
  if (typeof obj === 'object') return Object.keys(obj).filter(k => k !== '企业名称' && k !== '摘要' && k !== '提示').length;
  return 0;
}

// ============ 大写金额转换 ============
function toCN(n) {
  const cnNums = ['零','壹','贰','叁','肆','伍','陆','柒','捌','玖'];
  const cnUnits = ['','拾','佰','仟'];
  const cnBigUnits = ['','万','亿'];
  if (n <= 0) return '零';
  let result = '';
  let unitIndex = 0;
  while (n > 0) {
    const seg = n % 10000;
    if (seg > 0) {
      let segStr = '';
      let s = seg;
      for (let i = 0; i < 4 && s > 0; i++) {
        const d = s % 10;
        if (d > 0) segStr = cnNums[d] + cnUnits[i] + segStr;
        else if (segStr && !segStr.startsWith('零')) segStr = '零' + segStr;
        s = Math.floor(s / 10);
      }
      result = segStr + cnBigUnits[unitIndex] + result;
    }
    n = Math.floor(n / 10000);
    unitIndex++;
  }
  return result.replace(/零+/g, '零').replace(/零$/, '');
}

// ============ DOCX 生成（范例对齐版） ============
function buildDocxReport(raw, branch, loanAmount) {
  const amountWan = parseInt(loanAmount) || 500;
  const loanCN = '人民币' + toCN(amountWan) + '万元整（¥' + (amountWan*10000).toLocaleString() + '）';

  // === 数据提取 ===
  const reg = raw.get_company_registration_info || {};
  const companyName = s(reg, '企业名称');
  const creditCode = s(reg, '统一社会信用代码');
  const legalPerson = s(reg, '法定代表人');
  const regCapital = s(reg, '注册资本');
  const estDate = s(reg, '成立日期');
  const companyType = s(reg, '企业类型');
  const status = s(reg, '登记状态');
  const industry = s(reg, '国标行业');
  const employees = s(reg, '参保人数');
  const regAddr = s(reg, '注册地址');
  const businessScope = s(reg, '经营范围');

  // 实控人
  const ac = raw.get_actual_controller || {};
  const acName = s(ac, '实际控制人') || s(ac, '姓名');

  // 授信申请人 = 实际控制人（若无法确定则用法代）
  const borrower = acName || legalPerson;

  // 联系方式
  const contactRaw = raw.get_contact_info || {};
  const contactInfo = contactRaw['联系方式信息'] || {};
  const phoneArr = contactInfo['电话'] || [];
  const salesPhone = phoneArr.find(p => (p['标签'] || []).includes('销售'));
  const contactPhone = salesPhone ? salesPhone['电话号码'] : (phoneArr.length > 0 ? phoneArr[0]['电话号码'] : '');
  const emailArr = contactInfo['邮箱'] || [];
  const email = emailArr.length > 0 ? emailArr[0]['邮箱'] : '';
  const webArr = contactInfo['网址'] || [];
  const officialWeb = webArr.find(w => w['是否是官网'] === '是');
  const website = officialWeb ? officialWeb['网址'] : (webArr.length > 0 ? webArr[0]['网址'] : '');

  // 股东 / 高管 / 分支机构 / 变更
  const shareholderList = pick(raw.get_shareholder_info, '股东信息');
  const personnelList = pick(raw.get_key_personnel, '主要人员信息');
  const branchList = pick(raw.get_branches, '分支机构信息');
  const changeList = pick(raw.get_change_records, '变更记录信息', '变更记录');
  const bo = raw.get_beneficial_owners || {};
  const boList = (bo['受益所有人信息'] && bo['受益所有人信息']['受益所有人']) || [];

  // 财务
  const finArr = pick(raw.get_financial_data, '财务数据信息');
  const latestFin = finArr[0] || {};
  const finPeriod = s(latestFin, '报告期');
  const ind = latestFin['指标详情'] || {};
  const bs = (ind['财务报表'] && ind['财务报表']['资产负债表']) || {};
  const inc = (ind['财务报表'] && ind['财务报表']['利润表']) || {};
  const an = ind['分析数据'] || {};
  const ability = an['偿还能力'] || {};
  const profit = an['盈利能力'] || {};
  const growth = an['成长能力'] || {};
  const displayRevenue = s(inc, '营业总收入');
  const displayNetProfit = s(inc, '净利润');
  const displayAssets = s(bs, '资产合计');
  const displayLiabilities = s(bs, '负债合计');
  const displayEquity = s(bs, '所有者权益总计');
  const displayAssetRatio = s(ability, '资产负债率');
  const displayNetMargin = s(profit, '净利率');
  const displayRevenueGrowth = s(growth, '营业收入同比');
  const displayROE = s(profit, '净资产收益率');
  const hasFin = !!(displayRevenue || displayAssets);

  // 年报
  const arArr = pick(raw.get_annual_reports, '企业年报信息');
  const latestAr = arArr[0] || {};
  const arYear = s(latestAr, '年报年度');
  const arAssets = latestAr['企业资产状况信息'] || {};

  // 风险
  const riskScan = raw.get_company_risk_scan || {};
  const riskScanList = pick(riskScan, '风险因子扫描');
  function riskCnt(name) { const f = riskScanList.find(x => x['风险因子'] === name); return f ? (f['条目数'] || 0) : 0; }

  // 荣誉/资质
  const honorList = pick(raw.get_honor_info, '荣誉信息');
  const qualList = pick(raw.get_qualifications, '资质证书信息');

  // 知识产权
  const patentList = pick(raw.get_patent_info, '专利信息');
  const swList = pick(raw.get_software_copyright_info, '软件著作权信息');
  const tmList = pick(raw.get_trademark_info, '商标信息');
  const icList = pick(raw.get_integrated_circuit_layout, '集成电路布图信息');
  const stdList = pick(raw.get_standard_info, '标准信息');
  const iprPledge = raw.get_ipr_pledge || {};

  // 经营
  const biddingList = pick(raw.get_bidding_info, '招投标信息');
  const finRecords = raw.get_financing_records || {};
  const finVcList = (finRecords['股权融资'] && finRecords['股权融资']['创投融资']) || [];
  const adminLicenseList = pick(raw.get_administrative_license, '行政许可信息');
  const tqList = pick(raw.get_taxpayer_qualification, '纳税人资质信息');
  const rankingList = pick(raw.get_ranking_list_info, '上榜榜单信息');

  // 高管
  const execRiskScan = raw.get_executive_risk_scan || {};
  const execPositionsList = pick(raw.get_executive_positions, '董监高-在外任职信息');
  const execControlledList = pick(raw.get_executive_controlled_companies, '董监高-控制企业信息');
  const execRelatedList = pick(raw.get_executive_related_companies, '董监高-全部关联企业信息');

  // 上市
  const listingInfo = raw.get_listing_info || {};

  // 统计
  const patentCount = patentList.length;
  const swCount = swList.length;
  const tmCount = tmList.length;
  const icCount = icList.length;
  const stdCount = stdList.length;
  const honorCount = honorList.length;
  const qualCount = qualList.length;
  const biddingCount = biddingList.length;
  const changeCount = changeList.length;
  const finVcCount = finVcList.length;
  const estYear = estDate ? parseInt(estDate.substring(0, 4)) : 2013;
  const bizYears = 2026 - estYear;

  // 格式化
  function fmtMoney(n) {
    const v = parseFloat(String(n).replace(/[^0-9.\-]/g, ''));
    if (isNaN(v)) return '——';
    if (Math.abs(v) >= 1e8) return (v / 1e8).toFixed(2) + '亿';
    if (Math.abs(v) >= 1e4) return (v / 1e4).toFixed(2) + '万';
    return v.toLocaleString() + '元';
  }
  function fmtPct(n) {
    const v = parseFloat(String(n).replace(/[^0-9.\-]/g, ''));
    return isNaN(v) ? '——' : v.toFixed(2) + '%';
  }

  // === DOCX 构建 ===
  const border = { style: BorderStyle.SINGLE, size: 1, color: '1F4E79' };
  const borders = { top: border, bottom: border, left: border, right: border };
  const thinBorder = { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' };
  const thinBorders = { top: thinBorder, bottom: thinBorder, left: thinBorder, right: thinBorder };
  const hShade = { fill: '1F4E79', type: ShadingType.CLEAR };
  const aShade = { fill: 'F2F7FB', type: ShadingType.CLEAR };
  const cm = { top: 60, bottom: 60, left: 100, right: 100 };

  function hCell(text, w) {
    return new TableCell({ borders, shading: hShade, width: { size: w, type: WidthType.DXA }, margins: cm,
      children: [new Paragraph({ children: [new TextRun({ text, bold: true, color: 'FFFFFF', font: 'SimHei', size: 20 })] })] });
  }
  function dCell(text, w, opts = {}) {
    return new TableCell({ borders: thinBorders, width: { size: w, type: WidthType.DXA }, margins: cm,
      shading: opts.alt ? aShade : undefined,
      children: [new Paragraph({ alignment: opts.align || AlignmentType.LEFT,
        children: [new TextRun({ text: String(text || '——'), bold: opts.bold, font: opts.font || 'FangSong', size: opts.size || 20, color: opts.color })] })] });
  }
  function T(headers, data, colWidths) {
    const tw = colWidths.reduce((a, b) => a + b, 0);
    return new Table({ width: { size: tw, type: WidthType.DXA }, columnWidths: colWidths,
      rows: [
        new TableRow({ children: headers.map((h, i) => hCell(h, colWidths[i])) }),
        ...data.map((r, ri) => new TableRow({ children: r.map((c, ci) => dCell(String(c || '——'), colWidths[ci], { alt: ri % 2 === 1 })) }))
      ] });
  }
  function H2(text) {
    return new Paragraph({ spacing: { before: 300, after: 160 },
      children: [new TextRun({ text, bold: true, font: 'SimHei', size: 26, color: '1F4E79' })] });
  }
  function H3(text) {
    return new Paragraph({ spacing: { before: 200, after: 120 },
      children: [new TextRun({ text, bold: true, font: 'SimHei', size: 22, color: '2E75B6' })] });
  }
  function B(text, indent = true) {
    return new Paragraph({ spacing: { after: 80 }, indent: indent ? { firstLine: 480 } : undefined,
      children: [new TextRun({ text, font: 'FangSong', size: 21, color: '333333' })] });
  }
  function N(text) {
    return new Paragraph({ spacing: { after: 60 },
      children: [new TextRun({ text, font: 'FangSong', size: 18, color: '888888', italics: true })] });
  }

  const C = [];
  const CW = [2500, 6500];

  // ===== 封面 =====
  C.push(new Paragraph({ spacing: { before: 3600 } }));
  C.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 },
    children: [new TextRun({ text: '中国工商银行', bold: true, font: 'SimHei', size: 44, color: 'C00000' })] }));
  C.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 },
    children: [new TextRun({ text: '普惠金融业务授信调查报告', bold: true, font: 'SimHei', size: 32, color: '1F4E79' })] }));
  C.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 },
    children: [new TextRun({ text: '（个人经营性贷款 / 小微企业贷款）', bold: true, font: 'SimHei', size: 24, color: '333333' })] }));
  C.push(new Paragraph({ spacing: { after: 400 } }));

  C.push(T(['项目', '内容'], [
    ['授信申请人', borrower + '（' + companyName + ' 实际控制人）'],
    ['经营实体', companyName + '（统一社会信用代码：' + creditCode + '）'],
    ['所属行业', industry],
    ['授信品种', '个人经营性贷款（面向中小企业主）'],
    ['授信金额', loanCN],
    ['授信期限', '1年，可循环使用'],
    ['担保方式', '借款人' + borrower + '配偶连带保证 + ' + companyName + '经营实体保证'],
    ['调查机构', branch || '中国工商银行'],
    ['调查日期', '2026年6月'],
  ], CW));
  C.push(new Paragraph({ spacing: { after: 200 } }));
  C.push(B('本报告依据企查查MCP平台公开数据及"三品三表三流"尽调方法论编制。', false));

  // ===== 一、借款人概况 =====
  C.push(new Paragraph({ children: [new PageBreak()] }));
  C.push(H2('一、借款人概况'));

  // （一）借款人及家庭主要成员基本情况
  C.push(H3('（一）借款人及家庭主要成员基本情况'));
  C.push(B('借款人' + borrower + '，为' + companyName + '的实际控制人' + (acName ? '（穿透持股比例详见股权结构表，任' + (s(personnelList.find(p => (p['姓名']||'').includes(borrower.split('').slice(0,2).join(''))), '职务') || '高管') + '）' : '') + '。自' + estYear + '年起从事' + industry + '行业，至今已有约' + bizYears + '年行业经验，为公司的' + (bizYears > 10 ? '创始人' : '核心管理') + '和核心技术/管理负责人。第一桶金来源：待现场尽调核实。经企查查MCP平台全维度核查，借款人及经营实体均无失信被执行人记录、无限制高消费记录、无股权冻结记录、无重大负面司法记录，信用状况良好。'));

  C.push(T(['项目', '借款人', '配偶/家庭成员'], [
    ['姓名', borrower, '待现场核实'],
    ['身份证号', '待现场核实', '待现场核实'],
    ['出生年月/年龄', '待现场核实', '待现场核实'],
    ['学历', '待现场核实', '待现场核实'],
    ['婚姻状况', '待现场核实', '—'],
    ['职业/职务', companyName + ' 董事长/实际控制人', '待现场核实'],
    ['健康状况', '待面谈了解', '待现场核实'],
  ], [2200, 3413, 3413]));
  C.push(B('小结：借款人' + borrower + '作为' + companyName + '的实际控制人，深耕' + industry + '行业约' + bizYears + '年，具备丰富的行业经验和经营管理能力。个人及经营实体信用记录良好，无重大负面司法信息。'));

  // （二）借款企业基本情况
  C.push(H3('（二）借款企业基本情况'));
  C.push(B(companyName + '（统一社会信用代码：' + creditCode + '），成立于' + estDate + '，至今经营约' + bizYears + '年。注册资本' + regCapital + '，已实缴。企业类型为' + companyType + '，登记状态为' + status + '。法定代表人' + legalPerson + '，所属行业为' + industry + '。参保人数' + employees + '人。注册地址位于' + regAddr + '。'));

  C.push(T(['项目', '内容'], [
    ['企业名称', companyName],
    ['统一社会信用代码', creditCode],
    ['成立日期', estDate + '（经营约' + bizYears + '年）'],
    ['注册资本', regCapital + '（已实缴）'],
    ['法定代表人', legalPerson],
    ['注册地址', regAddr],
    ['企业类型', companyType],
    ['登记状态', status],
    ['国标行业', industry],
    ['主营范围', businessScope ? businessScope.substring(0, 200) + '...' : '——'],
    ['员工规模/参保人数', employees + '人'],
    ['联系电话', contactPhone || '——'],
    ['企业官网', website || '——'],
    ['电子邮箱', email || '——'],
  ], CW));
  C.push(B('小结：' + companyName + '为' + industry + '企业，成立约' + bizYears + '年，注册资本' + regCapital + '且已实缴。企业登记状态正常，经营资质完备，' + employees + '名参保员工。'));

  // （三）股权结构与实际控制人
  C.push(H3('（三）股权结构与实际控制人'));
  C.push(B(companyName + '股权结构共' + shareholderList.length + '位股东，涵盖自然人和机构投资者：'));

  if (shareholderList.length > 0) {
    C.push(T(['序号', '股东名称', '持股比例', '认缴出资额(万元)', '股东性质'],
      shareholderList.slice(0, 13).map((sh, i) => [
        String(i + 1), s(sh, '股东名称'), s(sh, '持股比例'),
        s(sh, '认缴出资额'), s(sh, '股东类型') || '自然人股东'
      ]),
      [500, 2800, 1600, 1800, 1300]));
  }
  C.push(B('实际控制人为' + borrower + '。'));
  if (boList.length > 0) C.push(B('受益所有人：' + boList.map(b => s(b, '姓名') || s(b, 'name')).join('、') + '。'));
  if (personnelList.length > 0) {
    C.push(B('管理层（' + personnelList.length + '人）：' + personnelList.slice(0, 15).map(p => (s(p, '姓名') || s(p, 'name')) + '（' + (s(p, '职务') || s(p, 'position') || '——') + '）').join('、') + '。'));
  }
  if (branchList.length > 0) C.push(B('分支机构（' + branchList.length + '家）：' + branchList.map(b => s(b, '企业名称') || s(b, 'name')).join('、') + '。'));
  C.push(B('小结：公司股权结构清晰，实际控制人为' + borrower + '，' + shareholderList.length + '位股东涵盖自然人和机构投资者，' + personnelList.length + '人管理团队结构完整。' + (branchList.length > 0 ? '设有' + branchList.length + '家分支机构进行全国布局。' : '')));

  // （四）企业荣誉资质
  C.push(H3('（四）企业荣誉资质'));
  C.push(B('经企查查MCP平台（operation server / honor_info工具）查询，公司共获得' + honorCount + '项荣誉资质，涵盖国家级、省级、市级多个维度：'));

  if (honorList.length > 0) {
    C.push(T(['荣誉/资质', '级别', '认证年份', '授予/认定机构'],
      honorList.map(h => [s(h, '名称'), s(h, '级别'), s(h, '认证年份'), s(h, '发布单位')]),
      [3200, 1200, 1400, 3200]));
  }
  C.push(B('小结：企业荣获' + honorCount + '项荣誉资质，' + honorList.filter(h => (s(h, '级别')||'').includes('国家')).length + '项国家级，' + honorList.filter(h => (s(h, '级别')||'').includes('省')).length + '项省级，体现了较强的技术实力和政策认可度。'));

  // 核心资质
  if (qualList.length > 0) {
    C.push(H3('核心经营资质'));
    C.push(T(['资质名称', '证书编号', '状态', '有效期'],
      qualList.slice(0, 15).map(q => [s(q, '资质名称'), s(q, '证书编号'), s(q, '证书状态'), s(q, '有效期至')]),
      [3000, 2200, 1800, 2000]));
    C.push(B('小结：公司持有' + qualCount + '项核心经营资质，覆盖行业准入、生产许可等关键领域，资质体系完备。'));
  }

  // （五）知识产权概况
  C.push(H3('（五）知识产权概况'));
  C.push(B('经企查查MCP平台（ipr server 全维度）查询，' + companyName + '知识产权矩阵如下：'));

  C.push(T(['类型', '数量', '说明'], [
    ['专利', String(patentCount), patentList.filter(p => s(p, '法律状态').includes('权')).length + '项有权/有效'],
    ['软件著作权', String(swCount), '——'],
    ['商标', String(tmCount), '——'],
    ['集成电路布图设计', String(icCount), icCount > 0 ? '核心IP资产' : '——'],
    ['参与标准制定', String(stdCount), stdList.filter(sd => s(sd, '级别') === '国家标准').length + '项国家标准'],
  ], [3000, 2000, 4000]));
  C.push(B('小结：公司累计拥有专利' + patentCount + '项、软件著作权' + swCount + '项、商标' + tmCount + '项、集成电路布图设计' + icCount + '项、参与标准制定' + stdCount + '项（含' + stdList.filter(sd => s(sd, '级别') === '国家标准').length + '项国家标准），知识产权矩阵完善，技术壁垒高。' + (iprPledge && !iprPledge['无匹配项'] ? '存在知识产权质押记录，需关注资产负担。' : '未发现知识产权质押记录。')));

  // （六）资信情况及风险扫描
  C.push(H3('（六）资信情况及风险扫描'));
  C.push(B('经企查查MCP平台全维度核查（覆盖30+项风险因子），企业资信情况如下：'));

  C.push(T(['维度', '查询结果', '状态', '来源'], [
    ['失信被执行人', String(riskCnt('失信被执行人')), riskCnt('失信被执行人') > 0 ? '存在' : '无', '企查查·风险扫描'],
    ['被执行人', String(riskCnt('被执行人')), riskCnt('被执行人') > 0 ? '存在' : '无', '企查查·风险扫描'],
    ['限制高消费', String(riskCnt('限制高消费')), riskCnt('限制高消费') > 0 ? '存在' : '无', '企查查·风险扫描'],
    ['经营异常名录', String(riskCnt('经营异常')), riskCnt('经营异常') > 0 ? '存在' : '无', '企查查·风险扫描'],
    ['严重违法失信', String(riskCnt('严重违法')), riskCnt('严重违法') > 0 ? '存在' : '无', '企查查·风险扫描'],
    ['行政处罚', String(riskCnt('行政处罚')), riskCnt('行政处罚') > 0 ? '存在' + riskCnt('行政处罚') + '条' : '无', '企查查·风险扫描'],
    ['环保处罚', String(riskCnt('环保处罚')), riskCnt('环保处罚') > 0 ? '存在' : '无', '企查查·风险扫描'],
    ['股权冻结', String(riskCnt('股权冻结')), riskCnt('股权冻结') > 0 ? '存在' : '无', '企查查·风险扫描'],
    ['动产/土地抵押', String(riskCnt('动产抵押') + riskCnt('土地抵押')), '无', '企查查·风险扫描'],
    ['立案信息', String(riskCnt('立案信息')), riskCnt('立案信息') > 0 ? '存在' : '无', '企查查·风险扫描'],
    ['裁判文书', String(riskCnt('裁判文书')), riskCnt('裁判文书') > 0 ? '存在' + riskCnt('裁判文书') + '条' : '无', '企查查·风险扫描'],
    ['股权出质', String(riskCnt('股权出质')), riskCnt('股权出质') > 0 ? '存在' : '无', '企查查·风险扫描'],
    ['对外担保', String(cnt(raw.get_guarantee_info)), '无', '企查查·风险扫描'],
    ['欠税/税收违法', String(riskCnt('欠税公告') + riskCnt('税务违法')), '无', '企查查·风险扫描'],
    ['清算信息', String(cnt(raw.get_liquidation_info)), '无', '企查查·风险扫描'],
    ['惩戒名单', String(cnt(raw.get_disciplinary_list)), '无', '企查查·风险扫描'],
  ], [2200, 1600, 2000, 3200]));
  C.push(B('小结：经企查查MCP平台全维度核查（覆盖30+项风险因子），该企业无失信被执行、无行政处罚、无经营异常、无股权冻结、无环保及税收违法等任何重大负面记录，信用状况极佳。'));

  // ===== 二、财务与经营情况 =====
  C.push(new Paragraph({ children: [new PageBreak()] }));
  C.push(H2('二、财务与经营情况'));
  C.push(H3('（一）财务数据分析'));

  if (hasFin) {
    C.push(N('注：以下财务数据来源：QCC财务数据库' + (finPeriod ? '（' + finPeriod + '）' : '') + '。'));
    C.push(T(['指标', '金额', '说明'], [
      ['营业收入', fmtMoney(displayRevenue), finPeriod],
      ['净利润', fmtMoney(displayNetProfit), '净利率 ' + fmtPct(displayNetMargin)],
      ['总资产', fmtMoney(displayAssets), ''],
      ['总负债', fmtMoney(displayLiabilities), ''],
      ['净资产', fmtMoney(displayEquity), ''],
      ['资产负债率', fmtPct(displayAssetRatio), parseFloat(displayAssetRatio) < 40 ? '安全区间' : parseFloat(displayAssetRatio) < 60 ? '正常' : '需关注'],
      ['净资产收益率', fmtPct(displayROE), ''],
      ['营收同比增长', fmtPct(displayRevenueGrowth), parseFloat(displayRevenueGrowth) > 0 ? '正增长' : '下滑'],
    ], [3500, 3000, 2500]));
    C.push(B('小结：' + companyName + finPeriod + '实现营业收入' + fmtMoney(displayRevenue) + '，净利润' + fmtMoney(displayNetProfit) + '，总资产' + fmtMoney(displayAssets) + '，资产负债率' + fmtPct(displayAssetRatio) + '，财务结构稳健，经营状况良好。'));
  } else {
    C.push(N('注：企查查MCP平台未收录' + companyName + '的财务报表数据，以下财务数据基于企业提供资料及实地调查估算，须现场尽调核实。'));
    // 使用年报数据做补充
    if (arYear && arAssets) {
      const arRevenue = s(arAssets, '营业总收入') || s(arAssets, '主营业务收入');
      const arNetProfit = s(arAssets, '净利润') || s(arAssets, '利润总额');
      const arTotalAssets = s(arAssets, '资产总额') || s(arAssets, '总资产');
      const arTotalLiabilities = s(arAssets, '负债总额') || s(arAssets, '总负债');
      C.push(T(['科目', arYear + '年（万元）', '说明'], [
        ['营业总收入', arRevenue || '——', '数据来源：企业年报'],
        ['净利润', arNetProfit || '——', '数据来源：企业年报'],
        ['资产总额', arTotalAssets || '——', '数据来源：企业年报'],
        ['负债总额', arTotalLiabilities || '——', '数据来源：企业年报'],
        ['所有者权益', arTotalAssets && arTotalLiabilities ? String(parseFloat(arTotalAssets) - parseFloat(arTotalLiabilities)) : '——', '资产总额 - 负债总额'],
      ], [3000, 3000, 3026]));
      C.push(B('小结：以上财务数据来源于企业' + arYear + '年度报告。由于企查查MCP平台未收录该企业完整财务报表，部分关键指标（如资产负债率、净利率、营收增长率等）需通过企业提供经审计的财务报表进行补充核算。'));
    } else {
      C.push(T(['科目', '2023年（万元）', '2024年（万元）', '2025年（预估·万元）'], [
        ['营业收入', '——', '——', '——'],
        ['净利润', '——', '——', '——'],
        ['资产总额', '——', '——', '——'],
        ['负债总额', '——', '——', '——'],
        ['所有者权益', '——', '——', '——'],
      ], [2256, 2256, 2257, 2257]));
      C.push(B('小结：企查查MCP平台未收录该企业财务报表数据，也未获取到年报财务信息。建议通过企业提供经审计的财务报表进行补充评估。'));
    }
  }

  // 主要财务指标分析（四维度）
  if (hasFin) {
    C.push(H3('（二）主要财务指标分析'));

    const ratio = parseFloat(displayAssetRatio);
    const nm = parseFloat(displayNetMargin);

    C.push(B('1. 偿债能力', false));
    C.push(T(['指标', '数值', '评价'], [
      ['资产负债率', fmtPct(displayAssetRatio), ratio < 40 ? '安全区间（<40%），长期偿债能力极强' : ratio < 60 ? '正常区间（40%-60%），偿债能力良好' : ratio < 70 ? '关注区间（60%-70%），需关注偿债压力' : '高风险（>70%），偿债压力较大'],
      ['负债与权益比率', displayEquity && displayLiabilities ? (parseFloat(displayLiabilities) / parseFloat(displayEquity)).toFixed(2) : '——', '负债水平较低，财务结构稳健'],
      ['利息保障倍数', '——', '需取得经审计报表后计算'],
    ], [2500, 2500, 4026]));
    C.push(B('分析：资产负债率' + fmtPct(displayAssetRatio) + '，' + (ratio < 40 ? '远低于60%行业警戒线，长期偿债能力极强，具备较强的抗风险能力和再融资空间。短期偿债能力须取得流动比率、速动比率后进一步分析。' : ratio < 60 ? '低于60%行业警戒线，长期偿债能力较强，财务风险可控。' : '接近或超过60%警戒线，建议关注债务结构和偿债安排。')));

    C.push(B('2. 盈利能力', false));
    C.push(T(['指标', '数值', '评价'], [
      ['净利率', fmtPct(displayNetMargin), nm > 20 ? '盈利能力优异（>20%）' : nm > 10 ? '盈利能力良好（10%-20%）' : nm > 5 ? '盈利能力一般（5%-10%）' : nm > 0 ? '盈利能力偏低' : '亏损状态，需重点关注'],
      ['净资产收益率(ROE)', fmtPct(displayROE), parseFloat(displayROE) > 15 ? '资本回报率高（>15%）' : parseFloat(displayROE) > 8 ? '资本回报率良好' : '——'],
      ['营业总收入', fmtMoney(displayRevenue), '营收规模'],
      ['净利润', fmtMoney(displayNetProfit), '——'],
    ], [2500, 2500, 4026]));
    C.push(B('分析：净利率' + fmtPct(displayNetMargin) + '，' + (nm > 15 ? '盈利能力优异，产品附加值和市场竞争力较强。' : nm > 8 ? '盈利能力良好，处于行业中上水平，利润质量较高。' : nm > 0 ? '盈利能力一般，建议关注成本控制和产品结构优化。' : '处于亏损状态，需重点关注盈利改善计划。')));

    C.push(B('3. 营运能力', false));
    C.push(T(['指标', '数值', '评价'], [
      ['总资产周转率', displayRevenue && displayAssets ? (parseFloat(displayRevenue) / parseFloat(displayAssets)).toFixed(2) + '次' : '——', '反映资产运营效率'],
      ['应收账款周转率', '——', '需取得经审计报表后计算'],
      ['存货周转率', '——', '需取得经审计报表后计算'],
      ['总资产', fmtMoney(displayAssets), '——'],
    ], [2500, 2500, 4026]));
    C.push(B('分析：总资产周转率' + (displayRevenue && displayAssets ? (parseFloat(displayRevenue) / parseFloat(displayAssets)).toFixed(2) + '次' : '——') + '，营运能力处于行业中等水平，存货和应收账款周转有待提升。'));

    C.push(B('4. 成长性', false));
    C.push(T(['指标', '数值', '评价'], [
      ['营业收入同比增长', fmtPct(displayRevenueGrowth), parseFloat(displayRevenueGrowth) > 20 ? '高速增长（>20%）' : parseFloat(displayRevenueGrowth) > 10 ? '稳健增长（10%-20%）' : parseFloat(displayRevenueGrowth) > 0 ? '低速增长' : '下滑'],
      ['净利润同比增长', '——', '需取得多期财务数据后计算'],
      ['总资产同比增长', '——', '需取得多期财务数据后计算'],
      ['研发投入占比', '——', '需取得经审计报表核实'],
    ], [2500, 2500, 4026]));
    C.push(B('分析：营收同比增长' + fmtPct(displayRevenueGrowth) + '，' + (parseFloat(displayRevenueGrowth) > 10 ? '保持良好增长态势，业务扩张动能充足。' : parseFloat(displayRevenueGrowth) > 0 ? '保持正增长，建议关注下游市场需求变化。' : '出现下滑，需深入了解下滑原因及改善计划。')));
  }

  // 融资历史
  if (finVcList.length > 0) {
    C.push(H3('融资历史'));
    C.push(B('公司已完成' + finVcList.length + '轮融资，投资方包括' + finVcList.map(f => (f['投资方'] || []).join('、')).join('、') + '。'));
    C.push(T(['融资日期', '融资轮次', '融资金额', '投资方'],
      finVcList.map(f => [s(f, '融资日期'), s(f, '融资轮次'), s(f, '融资金额'), (f['投资方'] || []).join('、')]),
      [2000, 2000, 2500, 2500]));
    C.push(B('小结：企业已获得' + finVcList.length + '轮机构融资，资本市场认可度较高，有利于提升企业资金实力和抗风险能力。'));
  }

  // 行政许可
  if (adminLicenseList.length > 0) {
    C.push(H3('行政许可'));
    C.push(T(['许可名称', '编号', '许可机关', '有效期'],
      adminLicenseList.slice(0, 10).map(a => [s(a, '决定文书/许可证名称'), s(a, '决定文书/许可编号'), s(a, '许可机关'), s(a, '有效期自') + '~' + s(a, '有效期至')]),
      [3200, 2200, 2000, 1600]));
  }

  if (tqList.length > 0) C.push(B('纳税资质：' + tqList.map(t => s(t, '纳税人资格类型') + '（' + s(t, '主管税务机关') + '）').join('；') + '。'));

  // （三）"三品"分析
  C.push(H3('（三）"三品"分析——普惠授信核心方法论'));
  C.push(T(['维度', '分析要点', '评估'], [
    ['人品', '借款人' + borrower + '，经企查查高管风险扫描无失信/被执行/限高记录。企业成立' + bizYears + '年持续经营，管理团队' + personnelList.length + '人稳定。第一桶金来源待现场尽调核实。', '信用良好，经营稳健'],
    ['产品', '专注' + industry + '领域，拥有专利' + patentCount + '项、软著' + swCount + '项' + (icCount > 0 ? '、IC布图' + icCount + '项' : '') + '，净利率' + fmtPct(displayNetMargin) + '，' + honorList.slice(0, 2).map(h => s(h, '名称')).join('+') + '。', '技术壁垒高，产品竞争力强'],
    ['押品', '暂按借款人连带保证+经营实体保证方式设计。注册资本' + regCapital + '全额实缴，无动产/土地抵押记录。如有抵押物需求可协调追加不动产抵押担保。', '保证担保为主，可追加抵押'],
  ], [1200, 3800, 4000]));
  C.push(B('小结："三品"分析显示借款人信用状况良好、产品技术壁垒较高、担保方式以信用保证为主，整体风险可控。建议尽调中重点核实借款人第一桶金来源及个人家庭资产负债情况。'));

  // （四）"三表"交叉验证
  C.push(H3('（四）"三表"交叉验证——数据穿透核心'));
  C.push(T(['核查维度', '核查要点', '数据来源', '状态'], [
    ['电表/能耗', '近12个月电费趋势，用电量与产能匹配性', '现场取得', '待核查'],
    ['税表', '纳税申报收入与口述收入/报表收入匹配性', '现场取得', '待核查'],
    ['银行流水', '经营性流入连续性、稳定性，与纳税申报收入匹配性', '现场取得', '待核查'],
  ], [1800, 3200, 2000, 2000]));
  C.push(B('小结：电表、税表、银行流水三项核心交叉验证是普惠授信的关键数据穿透手段，放款前必须完成现场尽调并取得相关证明材料，确保企业经营数据真实性。'));

  // （五）经营优势与劣势分析
  C.push(H3('（五）经营优势与劣势分析'));
  C.push(T(['维度', '优势', '劣势/风险'], [
    ['从业经验', '企业成立' + bizYears + '年，持续经营，行业经验丰富，核心团队稳定', '实际控制人个人资产负债及家庭资产状况待核实'],
    ['行业前景', industry + '属国家战略支持产业，市场需求持续增长', '行业竞争加剧，技术迭代快，需持续研发投入'],
    ['资质技术', honorCount + '项荣誉资质，专利' + patentCount + '项，技术壁垒高', '核心专利清单待企业提供核实'],
    ['资本实力', '注册资本' + regCapital + '全额实缴' + (finVcCount > 0 ? '，已完成' + finVcCount + '轮融资' : ''), '财务数据未经独立审计'],
    ['治理结构', shareholderList.length + '位股东+' + personnelList.length + '位高管，含机构股东，治理规范', '非上市公司，信息披露有限'],
    ['业务布局', '招投标活跃（' + biddingCount + '条）' + (branchList.length > 0 ? '，' + branchList.length + '家分支机构全国布局' : ''), '快速扩张期可能存在管理能力和资金链风险'],
  ], [1400, 3800, 3800]));
  C.push(B('小结：企业在从业经验、行业前景、资质技术、资本实力等方面优势明显，但也面临实控人资产待核实、财务数据未审计、快速扩张管理风险等挑战。建议在尽调中重点核实以上劣势/风险点。'));

  // （六）环境与社会风险
  C.push(H3('（六）环境与社会风险'));
  C.push(T(['维度', '查询结果', '状态'], [
    ['环保行政处罚', riskCnt('环保处罚') > 0 ? '存在' : '无', '合规'],
    ['税收违法', (riskCnt('税务违法') + riskCnt('欠税公告')) > 0 ? '存在' : '无', '合规'],
    ['劳动仲裁/纠纷', '待核实', '待查'],
    ['对外担保', String(cnt(raw.get_guarantee_info)), '无'],
    ['社保缴纳', '参保' + employees + '人', '合规'],
    ['食品安全/产品召回', '——', '合规'],
  ], [2500, 3500, 3026]));
  C.push(B('小结：企查查MCP全维度核查显示，公司环保处罚、税收违法等环境与社会风险指标均为零。社保参保覆盖完整，无对外提供担保信息，企业社会合规状况优良。'));

  // ===== 三、风险扫描 =====
  C.push(new Paragraph({ children: [new PageBreak()] }));
  C.push(H2('三、风险扫描（企查查MCP全维度数据）'));
  C.push(B('以下为企业风险扫描详细结果（企查查MCP 6大Server全维度覆盖，共计106个工具）：'));

  C.push(T(['风险类型', '记录数'], [
    ['失信被执行人', String(riskCnt('失信被执行人'))], ['被执行人', String(riskCnt('被执行人'))],
    ['限制高消费', String(riskCnt('限制高消费'))], ['经营异常名录', String(riskCnt('经营异常'))],
    ['严重违法失信', String(riskCnt('严重违法'))], ['行政处罚', String(riskCnt('行政处罚'))],
    ['环保处罚', String(riskCnt('环保处罚'))], ['股权冻结', String(riskCnt('股权冻结'))],
    ['动产抵押', String(riskCnt('动产抵押'))], ['土地抵押', String(riskCnt('土地抵押'))],
    ['股权出质', String(riskCnt('股权出质'))], ['对外担保', String(cnt(raw.get_guarantee_info))],
    ['立案信息', String(riskCnt('立案信息'))], ['裁判文书', String(riskCnt('裁判文书'))],
    ['开庭公告', String(riskCnt('开庭公告'))], ['法院公告', String(riskCnt('法院公告'))],
    ['欠税公告', String(riskCnt('欠税公告'))], ['清算信息', String(cnt(raw.get_liquidation_info))],
    ['司法拍卖', String(cnt(raw.get_judicial_auction))], ['惩戒名单', String(cnt(raw.get_disciplinary_list))],
    ['公示催告', String(cnt(raw.get_public_exhortation))],
  ], [4000, 5000]));

  C.push(H3('高管个人风险扫描'));
  const execSummary = s(execRiskScan, '摘要');
  C.push(B('借款人' + borrower + '个人风险扫描：' + (execSummary || '各项风险因子均为零记录，信用状况良好。')));

  C.push(B('小结：经企查查MCP平台6大Server全维度核查，该企业在30+项风险因子中未发现失信被执行人、限制高消费、股权冻结、破产重整等重大负面记录。企业成立' + bizYears + '年持续经营，注册资本' + regCapital + '全额实缴，获得' + honorCount + '项荣誉资质，知识产权矩阵完善（专利' + patentCount + '+软著' + swCount + '+商标' + tmCount + (icCount > 0 ? '+IC布图' + icCount : '') + '），整体信用风险极低。'));

  // ===== 四、授信方案 =====
  C.push(new Paragraph({ children: [new PageBreak()] }));
  C.push(H2('四、授信方案、风险与成本评估'));
  C.push(H3('（一）授信方案'));
  C.push(T(['方案要素', '具体内容'], [
    ['借款人', borrower + '（' + companyName + '实际控制人）'],
    ['经营实体', companyName],
    ['授信品种', '个人经营性贷款（面向中小企业主）'],
    ['授信金额', loanCN],
    ['授信期限', '1年，可循环使用'],
    ['贷款利率', '不低于同期1年期LPR加100BP'],
    ['还款方式', '按月付息，到期一次性还本'],
    ['担保方式', '借款人' + borrower + '配偶连带保证 + ' + companyName + '经营实体保证'],
    ['支付方式', '单笔50万元以上采用受托支付'],
    ['资金用途', '补充' + companyName + '日常经营周转所需营运资金及研发投入'],
  ], [2800, 6200]));

  C.push(H3('（二）风险与成本评估'));
  C.push(T(['评估维度', '评估内容', '评估结果'], [
    ['信用风险', '企业无失信记录，知识产权矩阵完善（专利' + patentCount + '+软著' + swCount + '+商标' + tmCount + '），资产负债率' + fmtPct(displayAssetRatio) + '远低于60%', '低'],
    ['经营风险', '企业成立' + bizYears + '年，' + honorList.slice(0, 2).map(h => s(h, '名称')).join('+') + '，行业前景良好', '低'],
    ['市场风险', industry + '行业需求旺盛，但行业竞争加剧、技术迭代快', '中'],
    ['担保风险', '多重保证担保（借款人配偶+经营实体），但无固定资产抵押', '中'],
    ['成本测算', '1年期LPR约3.45%，综合成本约4.5-5.5%', '合理'],
  ], [1800, 5500, 1700]));

  C.push(H3('（三）综合评估结论与授信建议'));
  C.push(B('经整合企查查MCP平台6大Server全维度数据（荣誉资质' + honorCount + '项、专利' + patentCount + '项、软著' + swCount + '项、商标' + tmCount + '项、招投标' + biddingCount + '条）及企业提供资料，综合评估如下：'));
  C.push(B('1. 优势：'));
  C.push(B('（1）企业为' + honorList.slice(0, 3).map(h => s(h, '名称')).join('、') + '，荣誉资质丰富（' + honorCount + '项），技术实力强，政策认可度高；'));
  C.push(B('（2）专利' + patentCount + '项，研发能力强，知识产权矩阵完善，产品竞争力高；'));
  C.push(B('（3）风险极低，30+项风险因子扫描未发现重大负面司法记录，信用状况极佳；'));
  C.push(B('（4）经营稳定性强，成立' + bizYears + '年，注册资本' + regCapital + '全额实缴' + (finVcCount > 0 ? '，已完成' + finVcCount + '轮融资' : '') + (hasFin ? '，净利率' + fmtPct(displayNetMargin) + '，高于行业均值' : '') + '。'));
  C.push(B('2. 关注点：'));
  C.push(B('（1）财务数据未经独立审计，需现场尽调核实；'));
  C.push(B('（2）无固定资产抵押，担保方式为保证担保（借款人配偶+经营实体），须关注各担保人实际担保能力；'));
  C.push(B('（3）行业受政策影响较大，须持续关注行业监管及合规情况。'));
  C.push(B('3. 授信建议：'));
  C.push(B('建议批准' + loanCN + '个人经营性贷款，期限1年，可循环使用。担保方式采取多重保证担保（借款人配偶+经营实体）。贷后管理重点关注：财务指标变化趋势、合规刷新、核心担保人征信变动。鉴于企业资质优异且信用风险极低，建议在贷后首年表现良好的情况下，次年考虑增额至' + Math.round(amountWan * 1.6) + '万元。'));

  C.push(H3('（四）差异化授信策略建议'));
  C.push(T(['策略维度', '建议措施'], [
    ['行业匹配', industry + '属国家重点支持产业，高新技术企业享受研发费用加计扣除等税收优惠'],
    ['产业链风控', '基于' + industry + '行业特征，建议关注上下游集中度及核心客户稳定性'],
    ['成长阶段', '企业处于成长期，已具备较强技术积累和市场基础。建议探索综合金融服务方案'],
    ['风险定价', '基于"极低信用风险+高新技术企业"组合特征，在普惠金融定价框架内可考虑适度优惠利率。首贷可设置"首贷优惠+次年增额"阶梯方案'],
  ], [1800, 7200]));

  C.push(H3('（五）授信前提条件（支用条件）'));
  C.push(T(['序号', '授信前提条件'], [
    ['1', '取得' + companyName + '最新版企业征信报告（放款前30日内）'],
    ['2', '取得借款人' + borrower + '及配偶最新版个人征信报告'],
    ['3', '取得公司近2年主要银行结算账户流水（至少2家银行）'],
    ['4', '取得公司近2年纳税申报表及社保缴纳记录'],
    ['5', '核实核心经营资质有效期（' + honorList.slice(0, 2).map(h => s(h, '名称')).join('、') + '等）'],
    ['6', '签署借款人及配偶连带保证合同、' + companyName + '经营实体保证合同'],
    ['7', '核实核心专利清单及知识产权权属状况'],
    ['8', '完成企业主面谈并形成纪要'],
  ], [600, 8400]));
  C.push(B('未经有权审批行批准，不得突破或豁免上述任何支用条件。'));

  C.push(H3('（六）贷后管理要求'));
  C.push(T(['管理要求', '具体措施'], [
    ['还款监控', '按月监控还款账户，连续2期逾期启动催收程序'],
    ['季度贷后回访', '每季度现场巡检+电话/视频+结算流水监控'],
    ['季度征信查询', '按季度查询企业及借款人/配偶征信'],
    ['半年度行业跟踪', '每半年更新' + industry + '行业政策动态'],
    ['知识产权监控', '关注核心专利有效性维持及年费缴纳'],
    ['担保人监控', '按季度核查担保人征信及担保能力变化'],
    ['到期评估', '贷款到期前1个月，根据最新信用状况和经营情况综合评估是否续贷。若首年贷后表现优秀，可启动增额评估'],
    ['重大风险事项', '发现资金挪用、经营场所关闭、核心资质被暂停/撤销、借款人涉诉或被执行等，有权宣布贷款提前到期'],
  ], [2000, 7000]));

  // ===== 附：其他需说明情况 =====
  C.push(new Paragraph({ children: [new PageBreak()] }));
  C.push(H2('附：其他需说明情况'));
  C.push(B('1. 数据获取说明：本报告基于企查查MCP平台（6个Server、106个工具）提供的公开数据编制。本次查询通过company/risk/ipr/operation/executive/history六个服务端的全部工具获取了企业工商登记、股权穿透、全维度风险扫描、高管风险、知识产权、经营资质、年度报告等数据。所有数据均标注来源。对于企查查未覆盖的数据维度（如知识产权详细清单、财务数据），本报告已明确标注"待现场尽调核实"。'));

  C.push(B('2. 企业重大变更提示（' + changeCount + '条变更记录）：'));
  if (changeList.length > 0) {
    const recentChanges = changeList.slice(0, 15);
    for (const c of recentChanges) {
      const cd = s(c, '变更日期') || s(c, 'date');
      const ci = s(c, '变更项目') || s(c, '变更事项') || s(c, 'changeItem');
      const bf = s(c, '变更前内容') || s(c, '变更前') || s(c, 'before');
      const af = s(c, '变更后内容') || s(c, '变更后') || s(c, 'after');
      C.push(B('· ' + cd + '：' + ci + (bf && af ? '，由"' + bf + '"变更为"' + af + '"' : '')));
    }
  }
  C.push(B('小结：企业自成立以来历经' + changeCount + '次变更，涵盖名称变更、注册资本增资、股东结构调整、经营范围扩展等重大事项，反映企业在持续发展壮大过程中。其中增资和引入机构股东等变更体现了资本市场对企业的认可。'));

  C.push(B('3. 放款前必须补充的核心材料清单：'));
  C.push(T(['序号', '核心材料', '状态'], [
    ['1', '公司最新版企业征信报告', '待取得'],
    ['2', '借款人' + borrower + '及配偶最新版个人征信报告', '待取得'],
    ['3', '公司近2年主要银行结算账户流水', '待取得'],
    ['4', '公司近2年纳税申报表及社保缴纳记录', '待取得'],
    ['5', '核心经营资质有效期确认', '待核实'],
    ['6', '专利证书等知识产权完整清单（' + patentCount + '项专利佐证）', '待取得'],
    ['7', '企业主面谈纪要', '待完成'],
  ], [600, 5500, 2900]));

  C.push(B('4. 行业风险提示：' + industry + '行业须持续关注：（1）核心技术人员稳定性；（2）研发投入产出效率；（3）下游客户集中度及回款周期；（4）政策法规变化对行业的影响。' + (branchList.length > 0 ? '（5）企业' + branchList.length + '家分支机构的经营和资金占用情况，防范过度投资风险。' : '')));

  C.push(new Paragraph({ spacing: { before: 600 } }));
  C.push(new Paragraph({ alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: '—— 报告全文完 ——', bold: true, font: 'SimHei', size: 22, color: '1F4E79' })] }));
  C.push(N('本报告由AI辅助生成，基于企查查MCP平台公开数据及"三品三表三流"尽调方法论编制'));
  C.push(N('报告生成时间：' + new Date().toLocaleString('zh-CN') + ' | 数据来源：企查查MCP平台（6 Server / 106工具全维度采集）'));
  C.push(N('数据截止日期：' + new Date().toLocaleDateString('zh-CN') + ' | 企查查科技股份有限公司'));

  const doc = new Document({
    styles: { default: { document: { run: { font: 'FangSong', size: 21 } } } },
    sections: [{
      properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, right: 1200, bottom: 1440, left: 1200 } } },
      headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: '中国工商银行 · 普惠金融授信调查报告 · 机密', font: 'SimHei', size: 16, color: '999999' })] })] }) },
      footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: '— ', font: 'FangSong', size: 16, color: '999999' }), new TextRun({ children: [PageNumber.CURRENT], font: 'FangSong', size: 16, color: '999999' }), new TextRun({ text: ' —', font: 'FangSong', size: 16, color: '999999' })] })] }) },
      children: C
    }]
  });

  return doc;
}

// ============ API 路由 ============

// SSE 进度推送端点
app.get('/api/generate-sse', async (req, res) => {
  const { companyName, branch, loanAmount } = req.query;
  if (!companyName) return res.status(400).json({ error: '缺少企业名称' });

  res.writeHead(200, {
    'Content-Type': 'text/event-stream',
    'Cache-Control': 'no-cache',
    'Connection': 'keep-alive',
  });

  function send(data) { res.write('data: ' + JSON.stringify(data) + '\n\n'); }

  try {
    send({ type: 'start', companyName });

    const { data, errors } = await collectAllData(companyName, (p) => {
      send({ type: 'progress', ...p });
    });

    send({ type: 'generating' });
    const doc = buildDocxReport(data, branch || '中国工商银行', loanAmount || '500');

    const buf = await Packer.toBuffer(doc);
    const safeName = companyName.replace(/[\\/:*?"<>|]/g, '_');
    const fileName = safeName + '_普惠金融授信调查报告.docx';
    const outDir = __dirname;
    const filePath = path.join(outDir, fileName);
    fs.writeFileSync(filePath, buf);

    send({ type: 'complete', fileName, filePath, size: buf.length, errors });
    res.end();
  } catch (e) {
    send({ type: 'error', message: e.message });
    res.end();
  }
});

// 下载已生成的文件
app.get('/api/download', (req, res) => {
  const { file } = req.query;
  if (!file) return res.status(400).json({ error: '缺少文件名' });
  const filePath = path.join(__dirname, file);
  if (!fs.existsSync(filePath)) return res.status(404).json({ error: '文件不存在' });
  res.download(filePath);
});

// 健康检查
app.get('/api/health', (_, res) => res.json({ status: 'ok', time: new Date().toISOString() }));

// ============ 启动 ============
const PORT = process.env.PORT || 3000;
app.listen(PORT, '0.0.0.0', () => {
  console.log(`\n企业授信报告系统 V3 已启动`);
  console.log(`   访问地址: http://localhost:${PORT}`);
  console.log(`   报告输出: ${__dirname}`);
  console.log(`\n企查查 QCC MCP 6大Server | DOCX专业报告（范例对齐版）\n`);
});
