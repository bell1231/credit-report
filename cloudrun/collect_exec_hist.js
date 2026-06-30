const https = require('https');
const fs = require('fs');
const token = 'Bearer MmDFP6lIn3akremGqfRhudPo0NON0LTrz92cBzCHtI3eA7IZ';

async function callQcc(server, toolName, args, timeout = 30000) {
  return new Promise((resolve) => {
    const body = JSON.stringify({ jsonrpc: '2.0', id: Date.now(), method: 'tools/call', params: { name: toolName, arguments: args } });
    const req = https.request({
      hostname: 'agent.qcc.com', path: '/mcp/' + server + '/stream', method: 'POST',
      headers: { 'Content-Type': 'application/json', 'Authorization': token, 'Content-Length': Buffer.byteLength(body) },
      timeout: timeout
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
    req.on('error', e => resolve(null));
    req.on('timeout', () => { req.destroy(); resolve(null); });
    req.write(body); req.end();
  });
}

async function main() {
  const company = '长沙金维集成电路股份有限公司';
  const existing = JSON.parse(fs.readFileSync('C:/Users/Administrator/CodeBuddy/20260625211524/qcc_data_phase34.json'));

  const keyPersonnel = existing.get_key_personnel;
  const regInfo = existing.get_company_registration_info;
  const legalPerson = (regInfo && regInfo['法定代表人']) || '';

  const execNames = new Set();
  if (legalPerson) execNames.add(legalPerson);
  const personnelList = (() => {
    if (!keyPersonnel) return [];
    const keys = ['主要人员', '高管信息', '董监高'];
    for (const k of keys) { if (Array.isArray(keyPersonnel[k]) && keyPersonnel[k].length > 0) return keyPersonnel[k]; }
    return [];
  })();
  for (const p of personnelList) {
    const n = p['姓名'] || p['name'] || '';
    if (n && n !== legalPerson) execNames.add(n);
    if (execNames.size >= 3) break;
  }
  const primaryPerson = [...execNames][0] || legalPerson || '';
  console.log('Core executives:', [...execNames].join(', '));

  console.log('\n=== Phase 5: Executive Server ===');
  const execTools = [
    ['executive', 'get_executive_risk_scan'],
    ['executive', 'get_executive_dishonest'],
    ['executive', 'get_executive_high_consumption_ban'],
    ['executive', 'get_executive_exit_restriction'],
    ['executive', 'get_executive_judgment_debtor'],
    ['executive', 'get_executive_positions'],
    ['executive', 'get_executive_related_companies'],
    ['executive', 'get_executive_controlled_companies'],
    ['executive', 'get_executive_investments'],
    ['executive', 'get_executive_related_risk_scan'],
    ['executive', 'get_executive_equity_freeze'],
    ['executive', 'get_executive_equity_pledge'],
    ['executive', 'get_executive_tax_violation'],
    ['executive', 'get_executive_case_filing'],
    ['executive', 'get_executive_judicial_docs'],
    ['executive', 'get_executive_terminated_cases'],
    ['executive', 'get_executive_admin_penalty'],
  ];
  for (const [s, t] of execTools) {
    const r = await callQcc(s, t, { searchKey: company, personName: primaryPerson });
    console.log('  ' + t + ': ' + (r ? 'OK' : 'NULL'));
    existing[t] = r;
    await new Promise(r => setTimeout(r, 300));
  }

  console.log('\n=== Phase 6: History Server ===');
  const uscc = (regInfo && regInfo['统一社会信用代码']) || company;
  const histTools = [
    ['history', 'get_historical_legal_rep'],
    ['history', 'get_historical_shareholders'],
    ['history', 'get_historical_dishonest'],
    ['history', 'get_historical_judgment_debtor'],
    ['history', 'get_historical_registration'],
    ['history', 'get_historical_executives'],
    ['history', 'get_historical_bankruptcy'],
    ['history', 'get_historical_high_consumption_ban'],
    ['history', 'get_historical_equity_freeze'],
    ['history', 'get_historical_business_exception'],
    ['history', 'get_historical_serious_violation'],
    ['history', 'get_historical_admin_penalty'],
    ['history', 'get_historical_tax_arrears'],
    ['history', 'get_historical_honor'],
    ['history', 'get_historical_patent'],
    ['history', 'get_historical_trademark'],
    ['history', 'get_historical_listing'],
    ['history', 'get_historical_terminated_cases'],
    ['history', 'get_historical_investments'],
  ];
  for (const [s, t] of histTools) {
    const r = await callQcc(s, t, { searchKey: uscc });
    console.log('  ' + t + ': ' + (r ? 'OK' : 'NULL'));
    existing[t] = r;
    await new Promise(r => setTimeout(r, 250));
  }

  fs.writeFileSync('C:/Users/Administrator/CodeBuddy/20260625211524/qcc_data_full.json', JSON.stringify(existing, null, 2));
  console.log('\n=== All data collected ===');
  console.log('Total tools called:', Object.keys(existing).length);
}
main();
