const https = require('https');
const fs = require('fs');
const token = 'Bearer MmDFP6lIn3akremGqfRhudPo0NON0LTrz92cBzCHtI3eA7IZ';
const COMPANY = '长沙金维集成电路股份有限公司';

async function callQcc(server, toolName, args, timeout = 30000) {
  return new Promise((resolve) => {
    const body = JSON.stringify({ jsonrpc: '2.0', id: Date.now(), method: 'tools/call', params: { name: toolName, arguments: args } });
    const req = https.request({
      hostname: 'agent.qcc.com', path: '/mcp/' + server + '/stream', method: 'POST',
      headers: { 'Content-Type': 'application/json', 'Authorization': token, 'Content-Length': Buffer.byteLength(body) },
      timeout
    }, res => {
      let d = ''; res.on('data', c => d += c); res.on('end', () => {
        const lines = d.split('\n').filter(l => l.startsWith('data: '));
        for (const line of lines) {
          try {
            const sse = JSON.parse(line.substring(6));
            if (sse.result && sse.result.content && sse.result.content[0] && sse.result.content[0].text) {
              try { resolve(JSON.parse(sse.result.content[0].text)); } catch { resolve(sse.result.content[0].text); }
              return;
            }
            if (sse.result) { resolve(sse.result); return; }
          } catch { }
        }
        resolve(null);
      });
    });
    req.on('error', () => resolve(null));
    req.on('timeout', () => { req.destroy(); resolve(null); });
    req.write(body); req.end();
  });
}

async function main() {
  const data = {};
  let count = 0;
  const uscc = '91430100070591256B'; // Known from first query
  
  async function doTask(server, tool, label, args) {
    const r = await callQcc(server, tool, args || { searchKey: COMPANY });
    data[tool] = r;
    count++;
    const ok = r && (typeof r === 'object' ? Object.keys(r).length > 0 : true);
    console.log(`  [${count}] ${server}/${tool}: ${ok ? 'OK' : 'NULL'}`);
    await new Promise(r => setTimeout(r, 300));
  }

  // Phase 1: Company Server (串行)
  console.log('=== Company Server ===');
  const cTools = [
    'get_company_registration_info', 'get_financial_data', 'get_annual_reports',
    'get_shareholder_info', 'get_external_investments', 'get_contact_info',
    'get_listing_info', 'get_key_personnel', 'get_actual_controller',
    'get_beneficial_owners', 'get_company_profile', 'get_branches', 'get_change_records',
  ];
  for (const t of cTools) {
    await doTask('company', t, t);
    await new Promise(r => setTimeout(r, 500));
  }

  // Phase 2: Risk Server
  console.log('\n=== Risk Server ===');
  const rTools = [
    'get_company_risk_scan', 'get_dishonest_info', 'get_judgment_debtor_info',
    'get_high_consumption_restriction', 'get_equity_freeze', 'get_equity_pledge_info',
    'get_business_exception', 'get_serious_violation', 'get_administrative_penalty',
    'get_tax_arrears_notice', 'get_tax_abnormal', 'get_tax_violation',
    'get_bankruptcy_reorganization', 'get_exit_restriction', 'get_terminated_cases',
    'get_default_info', 'get_judicial_documents', 'get_case_filing_info',
    'get_hearing_notice', 'get_court_notice', 'get_service_notice',
    'get_pre_litigation_mediation', 'get_company_related_risk_scan',
    'get_guarantee_info', 'get_environmental_penalty', 'get_judicial_auction',
    'get_liquidation_info', 'get_chattel_mortgage_info', 'get_land_mortgage_info',
    'get_disciplinary_list', 'get_cancellation_record_info', 'get_public_exhortation',
  ];
  for (const t of rTools) await doTask('risk', t, t);

  // Phase 3: IPR Server
  console.log('\n=== IPR Server ===');
  const iprTools = [
    'get_patent_info', 'get_software_copyright_info', 'get_trademark_info',
    'get_copyright_work_info', 'get_ipr_pledge', 'get_international_patent',
    'get_standard_info', 'get_integrated_circuit_layout',
  ];
  for (const t of iprTools) await doTask('ipr', t, t);

  // Phase 4: Operation Server
  console.log('\n=== Operation Server ===');
  const opTools = [
    'get_bidding_info', 'get_qualifications', 'get_recruitment_info',
    'get_credit_evaluation', 'get_honor_info', 'get_taxpayer_qualification',
    'get_news_sentiment', 'get_financing_records', 'get_administrative_license',
    'get_government_interview', 'get_product_spot_check', 'get_random_check',
    'get_food_safety', 'get_product_recall', 'get_ranking_list_info',
    'get_import_export_credit', 'get_credit_commitments',
  ];
  for (const t of opTools) await doTask('operation', t, t);

  // Get legal person for executive server
  const regInfo = data.get_company_registration_info;
  const legalPerson = (regInfo && regInfo['法定代表人']) || '';
  const keyPersonnel = data.get_key_personnel;
  const execNames = new Set();
  if (legalPerson) execNames.add(legalPerson);
  const pList = (() => {
    if (!keyPersonnel) return [];
    const keys = ['主要人员', '高管信息', '董监高'];
    for (const k of keys) { if (Array.isArray(keyPersonnel[k]) && keyPersonnel[k].length > 0) return keyPersonnel[k]; }
    return [];
  })();
  for (const p of pList) {
    const n = p['姓名'] || p['name'] || '';
    if (n && n !== legalPerson) execNames.add(n);
    if (execNames.size >= 3) break;
  }
  const primaryPerson = [...execNames][0] || legalPerson || '';
  console.log('\nCore execs:', [...execNames].join(', '));

  // Phase 5: Executive Server
  console.log('\n=== Executive Server ===');
  const execTools = [
    'get_executive_risk_scan', 'get_executive_dishonest', 'get_executive_high_consumption_ban',
    'get_executive_exit_restriction', 'get_executive_judgment_debtor', 'get_executive_positions',
    'get_executive_related_companies', 'get_executive_controlled_companies', 'get_executive_investments',
    'get_executive_related_risk_scan', 'get_executive_equity_freeze', 'get_executive_equity_pledge',
    'get_executive_tax_violation', 'get_executive_case_filing', 'get_executive_judicial_docs',
    'get_executive_terminated_cases', 'get_executive_admin_penalty',
  ];
  for (const t of execTools) {
    await doTask('executive', t, t, { searchKey: COMPANY, personName: primaryPerson });
  }

  // Phase 6: History Server
  console.log('\n=== History Server ===');
  const histTools = [
    'get_historical_legal_rep', 'get_historical_shareholders', 'get_historical_dishonest',
    'get_historical_judgment_debtor', 'get_historical_registration', 'get_historical_executives',
    'get_historical_bankruptcy', 'get_historical_high_consumption_ban', 'get_historical_equity_freeze',
    'get_historical_business_exception', 'get_historical_serious_violation',
    'get_historical_admin_penalty', 'get_historical_tax_arrears', 'get_historical_honor',
    'get_historical_patent', 'get_historical_trademark', 'get_historical_listing',
    'get_historical_terminated_cases', 'get_historical_investments',
  ];
  for (const t of histTools) {
    await doTask('history', t, t, { searchKey: uscc });
  }

  fs.writeFileSync('C:/Users/Administrator/CodeBuddy/20260625211524/qcc_jinwei_full.json', JSON.stringify(data, null, 2));
  console.log(`\n=== Done! ${count} tools called ===`);
}
main().catch(console.error);
