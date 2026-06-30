// 从 QCC 全维度数据中提取关键信息，用于生成授信报告
const fs = require('fs');
const data = JSON.parse(fs.readFileSync('C:/Users/Administrator/CodeBuddy/20260625211524/qcc_data_full.json'));

function safeStr(obj, key) {
  if (!obj) return '';
  return obj[key] || '';
}
function safeArr(obj, key) {
  if (!obj) return [];
  if (Array.isArray(obj)) return obj;
  const v = obj[key];
  return Array.isArray(v) ? v : [];
}
function countArr(val) {
  if (!val) return 0;
  if (Array.isArray(val)) return val.length;
  if (typeof val === 'object') return Object.keys(val).length;
  return 0;
}

const reg = data.get_company_registration_info || {};
const fin = data.get_financial_data || {};
const annual = data.get_annual_reports || {};
const shareholders = data.get_shareholder_info || {};
const investments = data.get_external_investments || {};
const contact = data.get_contact_info || {};
const listing = data.get_listing_info || {};
const keyPersonnel = data.get_key_personnel || {};
const actualController = data.get_actual_controller || {};
const beneficialOwners = data.get_beneficial_owners || {};
const profile = data.get_company_profile || {};
const branches = data.get_branches || {};
const changes = data.get_change_records || {};

const riskScan = data.get_company_risk_scan || {};
const honorInfo = data.get_honor_info || {};
const qualifications = data.get_qualifications || {};
const patentInfo = data.get_patent_info || {};
const softwareInfo = data.get_software_copyright_info || {};
const trademarkInfo = data.get_trademark_info || {};
const biddingInfo = data.get_bidding_info || {};
const taxpayerQual = data.get_taxpayer_qualification || {};
const newsSentiment = data.get_news_sentiment || {};
const financingRecords = data.get_financing_records || {};
const adminLicense = data.get_administrative_license || {};
const rankingList = data.get_ranking_list_info || {};
const execRiskScan = data.get_executive_risk_scan || {};
const execPositions = data.get_executive_positions || {};
const execRelatedCompanies = data.get_executive_related_companies || {};
const execControlledCompanies = data.get_executive_controlled_companies || {};
const histRegistration = data.get_historical_registration || {};
const histExecutives = data.get_historical_executives || {};
const histShareholders = data.get_historical_shareholders || {};
const histHonor = data.get_historical_honor || {};
const histPatent = data.get_historical_patent || {};
const histTrademark = data.get_historical_trademark || {};
const histInvestments = data.get_historical_investments || {};

// 提取核心数据
const companyName = safeStr(reg, '企业名称');
const creditCode = safeStr(reg, '统一社会信用代码');
const legalPerson = safeStr(reg, '法定代表人');
const regCapital = safeStr(reg, '注册资本');
const estDate = safeStr(reg, '成立日期');
const companyType = safeStr(reg, '企业类型');
const status = safeStr(reg, '登记状态');
const industry = safeStr(reg, '国标行业');
const employees = safeStr(reg, '参保人数');
const regAddr = safeStr(reg, '注册地址');
const businessScope = safeStr(reg, '经营范围');
const contactPhone = safeStr(contact, '联系电话');
const contactEmail = safeStr(contact, '电子邮箱');
const website = safeStr(contact, '网站');

// 实控人
const acName = safeStr(actualController, '实际控制人') || safeStr(actualController, '姓名');
const boList = safeArr(beneficialOwners, '受益所有人');

// 股东
const shareholderList = safeArr(shareholders, '股东信息');
const invList = safeArr(investments, '对外投资');

// 高管
const personnelList = (() => {
  const keys = ['主要人员', '高管信息', '董监高'];
  for (const k of keys) { if (Array.isArray(keyPersonnel[k]) && keyPersonnel[k].length > 0) return keyPersonnel[k]; }
  if (Array.isArray(keyPersonnel)) return keyPersonnel;
  return [];
})();

// 财务
const finArr = safeArr(fin, '财务数据信息');
const latestFin = finArr.length > 0 ? finArr[0] : null;
const finPeriod = latestFin ? safeStr(latestFin, '报告期') : '';
const indicators = latestFin ? (latestFin['指标详情'] || {}) : {};
const bs = indicators['财务报表'] ? (indicators['财务报表']['资产负债表'] || {}) : {};
const is = indicators['财务报表'] ? (indicators['财务报表']['利润表'] || {}) : {};
const analysis = indicators['分析数据'] || {};
const ability = analysis['偿还能力'] || {};
const profit = analysis['盈利能力'] || {};
const growth = analysis['成长能力'] || {};

// 年报
const arArr = safeArr(annual, '企业年报信息');
const latestAr = arArr.length > 0 ? arArr[0] : null;
const arYear = latestAr ? safeStr(latestAr, '年报年度') : '';
const arAssets = latestAr ? (latestAr['企业资产状况信息'] || {}) : {};

// 风险
const riskScanList = safeArr(riskScan, '风险因子扫描');
function getRiskCount(factorName) {
  const factor = riskScanList.find(f => f['风险因子'] === factorName);
  return factor ? (factor['条目数'] || 0) : 0;
}

// 荣誉
const honorList = safeArr(honorInfo, '荣誉信息');
// 资质
const qualList = safeArr(qualifications, '资质证书');

// 专利
const patentList = (() => {
  if (Array.isArray(patentInfo)) return patentInfo;
  for (const k of ['专利信息', 'patents', 'patent_list', 'records']) {
    if (Array.isArray(patentInfo[k])) return patentInfo[k];
  }
  return [];
})();

// 软著
const swList = (() => {
  if (Array.isArray(softwareInfo)) return softwareInfo;
  for (const k of ['软件著作权', 'software', 'records']) {
    if (Array.isArray(softwareInfo[k])) return softwareInfo[k];
  }
  return [];
})();

// 商标
const trademarkList = (() => {
  if (Array.isArray(trademarkInfo)) return trademarkInfo;
  for (const k of ['商标信息', 'trademark', 'records']) {
    if (Array.isArray(trademarkInfo[k])) return trademarkInfo[k];
  }
  return [];
})();

// 融资
const finRecordsList = safeArr(financingRecords, '融资记录');

// 行政许可
const adminLicenseList = safeArr(adminLicense, '行政许可');

// 招投标
const biddingList = safeArr(biddingInfo, '招投标信息');

// 高管全景
const execPositionsList = (() => {
  if (Array.isArray(execPositions)) return execPositions;
  for (const k of ['任职信息', 'positions', 'records']) {
    if (Array.isArray(execPositions[k])) return execPositions[k];
  }
  return [];
})();

// 变更记录
const changeList = safeArr(changes, '变更记录信息') || safeArr(changes, '变更记录');

// 历史
const histRegList = safeArr(histRegistration, '历史注册');
const histExecList = safeArr(histExecutives, '历史高管');
const histShareholderList = safeArr(histShareholders, '历史股东');

// 计算统计
const patentCount = patentList.length;
const swCount = swList.length;
const trademarkCount = trademarkList.length;

const report = {
  companyName,
  creditCode,
  legalPerson,
  regCapital,
  estDate,
  companyType,
  status,
  industry,
  employees,
  regAddr,
  businessScope,
  contactPhone,
  contactEmail,
  website,
  acName,
  boList,
  shareholderList,
  invList,
  personnelList,
  // 财务
  finPeriod,
  revenue: safeStr(is, '营业总收入'),
  netProfit: safeStr(is, '净利润'),
  totalAssets: safeStr(bs, '资产合计'),
  totalLiabilities: safeStr(bs, '负债合计'),
  equity: safeStr(bs, '所有者权益总计'),
  assetRatio: safeStr(ability, '资产负债率'),
  netMargin: safeStr(profit, '净利率'),
  revenueGrowth: safeStr(growth, '营业收入同比'),
  roe: safeStr(profit, '净资产收益率'),
  // 年报
  arYear,
  arRevenue: safeStr(arAssets, '营业总收入'),
  arNetProfit: safeStr(arAssets, '净利润'),
  arTotalAssets: safeStr(arAssets, '资产总额'),
  arTotalLiabilities: safeStr(arAssets, '负债总额'),
  // 风险
  riskScanList,
  getRiskCount,
  // 荣誉/资质
  honorList,
  qualList,
  // 知识产权
  patentList, patentCount,
  swList, swCount,
  trademarkList, trademarkCount,
  // 经营
  biddingList,
  finRecordsList,
  adminLicenseList,
  // 高管
  execPositionsList,
  execRiskScan,
  // 变更
  changeList,
  histRegList,
  histExecList,
  histShareholderList,
  // 其他Server数据
  taxpayerQual,
  newsSentiment,
  rankingList,
  listing,
  profile,
};

fs.writeFileSync('C:/Users/Administrator/CodeBuddy/20260625211524/extracted_data.json', JSON.stringify(report, null, 2));
console.log('Data extracted. Key stats:');
console.log('  Company:', companyName);
console.log('  CreditCode:', creditCode);
console.log('  LegalPerson:', legalPerson);
console.log('  EstDate:', estDate);
console.log('  RegCapital:', regCapital);
console.log('  Industry:', industry);
console.log('  Employees:', employees);
console.log('  Status:', status);
console.log('  AC:', acName);
console.log('  Shareholders:', shareholderList.length);
console.log('  Personnel:', personnelList.length);
console.log('  Patents:', patentCount);
console.log('  Software:', swCount);
console.log('  Trademarks:', trademarkCount);
console.log('  Honors:', honorList.length);
console.log('  Qualifications:', qualList.length);
console.log('  Financing records:', finRecordsList.length);
console.log('  Admin licenses:', adminLicenseList.length);
console.log('  Bidding:', biddingList.length);
console.log('  Changes:', changeList.length);
console.log('  Revenue:', safeStr(is, '营业总收入'));
console.log('  NetProfit:', safeStr(is, '净利润'));
console.log('  TotalAssets:', safeStr(bs, '资产合计'));
console.log('  AssetRatio:', safeStr(ability, '资产负债率'));
