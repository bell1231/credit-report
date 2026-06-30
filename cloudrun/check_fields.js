const d = require('./qcc_jinwei_full.json');

// Honor
const h = d.get_honor_info;
if (h && h['荣誉信息'] && h['荣誉信息'].length > 0) {
  console.log('=== Honor fields ===');
  console.log(Object.keys(h['荣誉信息'][0]));
  console.log('First honor:', JSON.stringify(h['荣誉信息'][0]));
} else {
  console.log('Honor: no data or wrong key');
  console.log('Keys:', Object.keys(h || {}));
}

// Patent
const p = d.get_patent_info;
if (Array.isArray(p) && p.length > 0) {
  console.log('\n=== Patent (array) fields ===');
  console.log(Object.keys(p[0]));
  console.log('First patent:', JSON.stringify(p[0]));
} else {
  for (const k of Object.keys(p || {})) {
    if (Array.isArray(p[k]) && p[k].length > 0) {
      console.log('\n=== Patent key:', k, 'fields ===');
      console.log(Object.keys(p[k][0]));
      console.log('First:', JSON.stringify(p[k][0]));
      break;
    }
  }
}

// Software
const sw = d.get_software_copyright_info;
if (Array.isArray(sw) && sw.length > 0) {
  console.log('\n=== SW (array) fields ===');
  console.log(Object.keys(sw[0]));
} else {
  for (const k of Object.keys(sw || {})) {
    if (Array.isArray(sw[k]) && sw[k].length > 0) {
      console.log('\n=== SW key:', k, 'fields ===');
      console.log(Object.keys(sw[k][0]));
      break;
    }
  }
}

// Trademark
const tm = d.get_trademark_info;
if (Array.isArray(tm) && tm.length > 0) {
  console.log('\n=== Trademark (array) fields ===');
  console.log(Object.keys(tm[0]));
} else {
  for (const k of Object.keys(tm || {})) {
    if (Array.isArray(tm[k]) && tm[k].length > 0) {
      console.log('\n=== Trademark key:', k, 'fields ===');
      console.log(Object.keys(tm[k][0]));
      break;
    }
  }
}

// IC Layout
const ic = d.get_integrated_circuit_layout;
if (Array.isArray(ic) && ic.length > 0) {
  console.log('\n=== IC Layout (array) fields ===');
  console.log(Object.keys(ic[0]));
} else {
  for (const k of Object.keys(ic || {})) {
    if (Array.isArray(ic[k]) && ic[k].length > 0) {
      console.log('\n=== IC Layout key:', k, 'fields ===');
      console.log(Object.keys(ic[k][0]));
      break;
    }
  }
}

// Standard
const st = d.get_standard_info;
console.log('\n=== Standard ===');
console.log(JSON.stringify(st).substring(0, 300));

// Qualifications
const q = d.get_qualifications;
console.log('\n=== Qualifications ===');
console.log(JSON.stringify(q).substring(0, 500));

// Financing
const f = d.get_financing_records;
console.log('\n=== Financing records ===');
console.log(JSON.stringify(f).substring(0, 500));

// Admin license
const al = d.get_administrative_license;
console.log('\n=== Admin license ===');
console.log(JSON.stringify(al).substring(0, 500));

// Taxpayer qual
const tq = d.get_taxpayer_qualification;
console.log('\n=== Taxpayer qual ===');
console.log(JSON.stringify(tq).substring(0, 500));

// Bidding
const bi = d.get_bidding_info;
console.log('\n=== Bidding keys ===');
console.log(Object.keys(bi || {}));
if (bi['招投标信息'] && bi['招投标信息'].length > 0) {
  console.log('First:', JSON.stringify(bi['招投标信息'][0]));
}

// Executive positions
const ep = d.get_executive_positions;
console.log('\n=== Exec positions ===');
console.log(JSON.stringify(ep).substring(0, 500));

// Executive controlled
const ec = d.get_executive_controlled_companies;
console.log('\n=== Exec controlled ===');
console.log(JSON.stringify(ec).substring(0, 500));

// Executive related
const er = d.get_executive_related_companies;
console.log('\n=== Exec related ===');
console.log(JSON.stringify(er).substring(0, 500));

// Shareholders
const sh = d.get_shareholder_info;
console.log('\n=== Shareholder keys ===');
console.log(Object.keys(sh || {}));
if (sh['股东信息'] && sh['股东信息'].length > 0) {
  console.log('First:', JSON.stringify(sh['股东信息'][0]));
}

// Key personnel
const kp = d.get_key_personnel;
console.log('\n=== Key personnel keys ===');
console.log(Object.keys(kp || {}));
if (kp['主要人员'] && kp['主要人员'].length > 0) {
  console.log('First:', JSON.stringify(kp['主要人员'][0]));
}
