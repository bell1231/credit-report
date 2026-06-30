const fs = require('fs');
const raw = JSON.parse(fs.readFileSync('C:/Users/Administrator/CodeBuddy/20260625211524/qcc_jinwei_full.json'));

// ========== 数据提取 ==========
function s(obj, key) { return (obj && obj[key]) ? String(obj[key]) : ''; }
function arr(obj, ...keys) {
  if (!obj) return [];
  if (Array.isArray(obj)) return obj;
  for (const k of keys) { if (Array.isArray(obj[k])) return obj[k]; }
  return [];
}
function cnt(v) {
  if (!v) return 0;
  if (Array.isArray(v)) return v.length;
  if (typeof v === 'object') return Object.keys(v).length;
  return 0;
}
function pickList(obj, ...keys) {
  if (!obj) return [];
  if (Array.isArray(obj)) return obj;
  for (const k of keys) { if (Array.isArray(obj[k]) && obj[k].length > 0) return obj[k]; }
  for (const k of Object.keys(obj)) { if (Array.isArray(obj[k]) && obj[k].length > 0) return obj[k]; }
  return [];
}

// 工商
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
const contact = raw.get_contact_info || {};
const contactPhone = s(contact, '联系电话');
const contactEmail = s(contact, '电子邮箱');
const website = s(contact, '网站');

// 实控人
const ac = raw.get_actual_controller || {};
const acName = s(ac, '实际控制人') || s(ac, '姓名');

// 股东
const shareholders = pickList(raw.get_shareholder_info, '股东信息');
const investments = pickList(raw.get_external_investments, '对外投资');

// 高管
const keyPersonnel = pickList(raw.get_key_personnel, '主要人员', '高管信息', '董监高');

// 财务
const finArr = arr(raw.get_financial_data, '财务数据信息');
const latestFin = finArr[0] || {};
const finPeriod = s(latestFin, '报告期');
const indicators = latestFin['指标详情'] || {};
const bs = (indicators['财务报表'] && indicators['财务报表']['资产负债表']) || {};
const inc = (indicators['财务报表'] && indicators['财务报表']['利润表']) || {};
const analysis = indicators['分析数据'] || {};
const ability = analysis['偿还能力'] || {};
const profit = analysis['盈利能力'] || {};
const growth = analysis['成长能力'] || {};

// 年报
const arArr = arr(raw.get_annual_reports, '企业年报信息');
const latestAr = arArr[0] || {};
const arYear = s(latestAr, '年报年度');
const arAssets = latestAr['企业资产状况信息'] || {};

// 风险
const riskScan = raw.get_company_risk_scan || {};
const riskScanList = arr(riskScan, '风险因子扫描');
function riskCnt(name) { const f = riskScanList.find(x => x['风险因子'] === name); return f ? (f['条目数'] || 0) : 0; }

// 荣誉/资质
const honorList = arr(raw.get_honor_info, '荣誉信息');
const qualList = arr(raw.get_qualifications, '资质证书');

// 专利
const patentList = pickList(raw.get_patent_info, '专利信息', 'patents');
const swList = pickList(raw.get_software_copyright_info, '软件著作权', 'software');
const trademarkList = pickList(raw.get_trademark_info, '商标信息', 'trademark');
const icLayoutList = pickList(raw.get_integrated_circuit_layout, '集成电路布图');
const standardList = pickList(raw.get_standard_info, '标准制定', 'standard');

// 经营
const biddingList = arr(raw.get_bidding_info, '招投标信息');
const finRecordsList = arr(raw.get_financing_records, '融资记录');
const adminLicenseList = arr(raw.get_administrative_license, '行政许可');
const taxpayerQual = raw.get_taxpayer_qualification || {};
const newsSentiment = raw.get_news_sentiment || {};
const rankingList = arr(raw.get_ranking_list_info, '榜单信息');

// 高管
const execRiskScan = raw.get_executive_risk_scan || {};
const execPositions = pickList(raw.get_executive_positions, '任职信息', 'positions');
const execRelated = pickList(raw.get_executive_related_companies, '关联企业', 'relatedCompanies');
const execControlled = pickList(raw.get_executive_controlled_companies, '控制企业', 'controlledCompanies');

// 变更
const changeList = arr(raw.get_change_records, '变更记录信息', '变更记录');

// 分支机构
const branches = arr(raw.get_branches, '分支机构');

// 上市信息
const listingInfo = raw.get_listing_info || {};
const isListed = !!(listingInfo && Object.keys(listingInfo).length > 0 && !listingInfo['无匹配项']);

console.log('=== 长沙金维集成电路股份有限公司 数据摘要 ===');
console.log('企业名称:', companyName);
console.log('统一社会信用代码:', creditCode);
console.log('法定代表人:', legalPerson);
console.log('注册资本:', regCapital);
console.log('成立日期:', estDate);
console.log('企业类型:', companyType);
console.log('状态:', status);
console.log('行业:', industry);
console.log('参保人数:', employees);
console.log('地址:', regAddr);
console.log('实控人:', acName);
console.log('股东数:', shareholders.length);
console.log('高管数:', keyPersonnel.length);
console.log('对外投资数:', investments.length);
console.log('财务报告期:', finPeriod);
console.log('营收:', s(inc, '营业总收入'));
console.log('净利润:', s(inc, '净利润'));
console.log('总资产:', s(bs, '资产合计'));
console.log('负债:', s(bs, '负债合计'));
console.log('资产负债率:', s(ability, '资产负债率'));
console.log('净利率:', s(profit, '净利率'));
console.log('年报年度:', arYear);
console.log('专利数:', patentList.length);
console.log('软著数:', swList.length);
console.log('商标数:', trademarkList.length);
console.log('IC布图数:', icLayoutList.length);
console.log('标准数:', standardList.length);
console.log('荣誉数:', honorList.length);
console.log('资质数:', qualList.length);
console.log('融资记录数:', finRecordsList.length);
console.log('行政许可数:', adminLicenseList.length);
console.log('招投标数:', biddingList.length);
console.log('高管任职数:', execPositions.length);
console.log('高管关联企业数:', execRelated.length);
console.log('高管控制企业数:', execControlled.length);
console.log('分支机构数:', branches.length);
console.log('变更记录数:', changeList.length);
console.log('风险-失信:', riskCnt('失信被执行人'));
console.log('风险-被执行:', riskCnt('被执行人'));
console.log('风险-限高:', riskCnt('限制高消费'));
console.log('风险-股权冻结:', riskCnt('股权冻结'));
console.log('风险-破产:', riskCnt('破产重整'));
console.log('风险-裁判文书:', riskCnt('裁判文书'));
