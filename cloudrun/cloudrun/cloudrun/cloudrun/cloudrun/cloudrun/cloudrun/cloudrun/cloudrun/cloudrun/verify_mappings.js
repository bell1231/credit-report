const d = require('./qcc_jinwei_full.json');

// Correct field mappings based on actual QCC API responses
const mappings = {
  // Company
  'get_company_registration_info': { key: null, type: 'object' },
  'get_financial_data': { key: '财务数据信息', type: 'array' },
  'get_annual_reports': { key: '企业年报信息', type: 'array' },
  'get_shareholder_info': { key: '股东信息', type: 'array' },
  'get_external_investments': { key: '对外投资', type: 'array' },
  'get_contact_info': { key: null, type: 'object' },
  'get_listing_info': { key: null, type: 'object' },
  'get_key_personnel': { key: '主要人员信息', type: 'array' },
  'get_actual_controller': { key: null, type: 'object' },
  'get_beneficial_owners': { key: '受益所有人', type: 'array' },
  'get_company_profile': { key: null, type: 'object' },
  'get_branches': { key: '分支机构', type: 'array' },
  'get_change_records': { key: '变更记录信息', type: 'array' },
  
  // Risk
  'get_company_risk_scan': { key: '风险因子扫描', type: 'array' },
  'get_dishonest_info': { key: null, type: 'object' },
  'get_judgment_debtor_info': { key: null, type: 'object' },
  'get_high_consumption_restriction': { key: null, type: 'object' },
  'get_equity_freeze': { key: null, type: 'object' },
  'get_equity_pledge_info': { key: null, type: 'object' },
  'get_business_exception': { key: null, type: 'object' },
  'get_serious_violation': { key: null, type: 'object' },
  'get_administrative_penalty': { key: null, type: 'object' },
  'get_tax_arrears_notice': { key: null, type: 'object' },
  'get_tax_abnormal': { key: null, type: 'object' },
  'get_tax_violation': { key: null, type: 'object' },
  'get_bankruptcy_reorganization': { key: null, type: 'object' },
  'get_exit_restriction': { key: null, type: 'object' },
  'get_terminated_cases': { key: null, type: 'object' },
  'get_default_info': { key: null, type: 'object' },
  'get_judicial_documents': { key: null, type: 'object' },
  'get_case_filing_info': { key: null, type: 'object' },
  'get_hearing_notice': { key: null, type: 'object' },
  'get_court_notice': { key: null, type: 'object' },
  'get_service_notice': { key: null, type: 'object' },
  'get_pre_litigation_mediation': { key: null, type: 'object' },
  'get_company_related_risk_scan': { key: null, type: 'object' },
  'get_guarantee_info': { key: null, type: 'object' },
  'get_environmental_penalty': { key: null, type: 'object' },
  'get_judicial_auction': { key: null, type: 'object' },
  'get_liquidation_info': { key: null, type: 'object' },
  'get_chattel_mortgage_info': { key: null, type: 'object' },
  'get_land_mortgage_info': { key: null, type: 'object' },
  'get_disciplinary_list': { key: null, type: 'object' },
  'get_cancellation_record_info': { key: null, type: 'object' },
  'get_public_exhortation': { key: null, type: 'object' },

  // IPR
  'get_patent_info': { key: '专利信息', type: 'array' },
  'get_software_copyright_info': { key: '软件著作权信息', type: 'array' },
  'get_trademark_info': { key: '商标信息', type: 'array' },
  'get_copyright_work_info': { key: '作品著作权信息', type: 'array' },
  'get_ipr_pledge': { key: null, type: 'object' },
  'get_international_patent': { key: '国际专利信息', type: 'array' },
  'get_standard_info': { key: '标准信息', type: 'array' },
  'get_integrated_circuit_layout': { key: '集成电路布图信息', type: 'array' },

  // Operation
  'get_bidding_info': { key: '招投标信息', type: 'array' },
  'get_qualifications': { key: '资质证书信息', type: 'array' },
  'get_recruitment_info': { key: null, type: 'object' },
  'get_credit_evaluation': { key: null, type: 'object' },
  'get_honor_info': { key: '荣誉信息', type: 'array' },
  'get_taxpayer_qualification': { key: '纳税人资质信息', type: 'array' },
  'get_news_sentiment': { key: null, type: 'object' },
  'get_financing_records': { key: '股权融资', type: 'object_nested' },
  'get_administrative_license': { key: '行政许可信息', type: 'array' },
  'get_government_interview': { key: null, type: 'object' },
  'get_product_spot_check': { key: null, type: 'object' },
  'get_random_check': { key: null, type: 'object' },
  'get_food_safety': { key: null, type: 'object' },
  'get_product_recall': { key: null, type: 'object' },
  'get_ranking_list_info': { key: '榜单信息', type: 'array' },
  'get_import_export_credit': { key: null, type: 'object' },
  'get_credit_commitments': { key: null, type: 'object' },

  // Executive
  'get_executive_risk_scan': { key: null, type: 'object' },
  'get_executive_dishonest': { key: null, type: 'object' },
  'get_executive_high_consumption_ban': { key: null, type: 'object' },
  'get_executive_exit_restriction': { key: null, type: 'object' },
  'get_executive_judgment_debtor': { key: null, type: 'object' },
  'get_executive_positions': { key: '董监高-在外任职信息', type: 'array' },
  'get_executive_related_companies': { key: '董监高-全部关联企业信息', type: 'array' },
  'get_executive_controlled_companies': { key: '董监高-控制企业信息', type: 'array' },
  'get_executive_investments': { key: null, type: 'object' },
  'get_executive_related_risk_scan': { key: null, type: 'object' },
  'get_executive_equity_freeze': { key: null, type: 'object' },
  'get_executive_equity_pledge': { key: null, type: 'object' },
  'get_executive_tax_violation': { key: null, type: 'object' },
  'get_executive_case_filing': { key: null, type: 'object' },
  'get_executive_judicial_docs': { key: null, type: 'object' },
  'get_executive_terminated_cases': { key: null, type: 'object' },
  'get_executive_admin_penalty': { key: null, type: 'object' },

  // History
  'get_historical_legal_rep': { key: null, type: 'object' },
  'get_historical_shareholders': { key: '历史股东信息', type: 'array' },
  'get_historical_dishonest': { key: null, type: 'object' },
  'get_historical_judgment_debtor': { key: null, type: 'object' },
  'get_historical_registration': { key: null, type: 'object' },
  'get_historical_executives': { key: null, type: 'object' },
  'get_historical_bankruptcy': { key: null, type: 'object' },
  'get_historical_high_consumption_ban': { key: null, type: 'object' },
  'get_historical_equity_freeze': { key: null, type: 'object' },
  'get_historical_business_exception': { key: null, type: 'object' },
  'get_historical_serious_violation': { key: null, type: 'object' },
  'get_historical_admin_penalty': { key: null, type: 'object' },
  'get_historical_tax_arrears': { key: null, type: 'object' },
  'get_historical_honor': { key: null, type: 'object' },
  'get_historical_patent': { key: null, type: 'object' },
  'get_historical_trademark': { key: null, type: 'object' },
  'get_historical_listing': { key: null, type: 'object' },
  'get_historical_terminated_cases': { key: null, type: 'object' },
  'get_historical_investments': { key: null, type: 'object' },
};

function getList(toolName) {
  const m = mappings[toolName];
  const obj = d[toolName];
  if (!obj || !m) return [];
  if (!m.key) {
    // Object type - count as non-empty if has data
    if (typeof obj === 'object' && Object.keys(obj).length > 0 && !obj['无匹配项']) {
      return [obj]; // Return as single-item array for counting
    }
    return [];
  }
  if (m.type === 'object_nested') {
    // Special: financing records nested in 股权融资.创投融资
    if (obj[m.key] && obj[m.key]['创投融资']) return obj[m.key]['创投融资'];
    return [];
  }
  const arr = obj[m.key];
  return Array.isArray(arr) ? arr : [];
}

// Verify all tools
let totalOK = 0;
let totalNull = 0;
const nullTools = [];
const okTools = [];

for (const [tool, m] of Object.entries(mappings)) {
  const list = getList(tool);
  if (list.length > 0) {
    okTools.push(tool);
    totalOK++;
  } else {
    nullTools.push(tool);
    totalNull++;
  }
}

console.log('=== 数据提取验证 ===');
console.log('OK tools:', totalOK);
console.log('NULL/EMPTY tools:', totalNull);
console.log('\nNULL tools list:');
for (const t of nullTools) {
  const raw = d[t];
  const rawStr = raw ? JSON.stringify(raw).substring(0, 80) : 'undefined';
  console.log('  ' + t + ': ' + rawStr);
}

// Print specific data counts
console.log('\n=== 关键数据计数 ===');
console.log('荣誉:', getList('get_honor_info').length);
console.log('专利:', getList('get_patent_info').length);
console.log('软著:', getList('get_software_copyright_info').length);
console.log('商标:', getList('get_trademark_info').length);
console.log('IC布图:', getList('get_integrated_circuit_layout').length);
console.log('标准:', getList('get_standard_info').length);
console.log('资质:', getList('get_qualifications').length);
console.log('融资:', getList('get_financing_records').length);
console.log('行政许可:', getList('get_administrative_license').length);
console.log('纳税资质:', getList('get_taxpayer_qualification').length);
console.log('招投标:', getList('get_bidding_info').length);
console.log('高管任职:', getList('get_executive_positions').length);
console.log('高管控制:', getList('get_executive_controlled_companies').length);
console.log('高管关联:', getList('get_executive_related_companies').length);
console.log('股东:', getList('get_shareholder_info').length);
console.log('主要人员:', getList('get_key_personnel').length);
console.log('变更记录:', getList('get_change_records').length);
console.log('财务:', getList('get_financial_data').length);
console.log('年报:', getList('get_annual_reports').length);

// Honor details
const honors = getList('get_honor_info');
if (honors.length > 0) {
  console.log('\n=== 荣誉详情 ===');
  for (const h of honors) {
    console.log('  - ' + (h['名称'] || h['name']) + ' | ' + (h['级别'] || '') + ' | ' + (h['发布单位'] || '') + ' | ' + (h['认证年份'] || ''));
  }
}
