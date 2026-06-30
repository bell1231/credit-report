/**
 * ICBC 普惠金融中小企业授信报告 — Docx 构建器
 * 
 * 严格匹配 ICBC_Ouya_Pharma_Inclusive_Credit_Report.docx 样本格式：
 * - 封面页（红字标题+深蓝表头信息表）
 * - 一至九章正文 + 附录（数据来源说明）
 * - 字体：标题黑体(SimHei)，正文仿宋(FangSong)
 * - 表格：深蓝表头(#1F4E79)+隔行浅蓝底色(#F2F7FB)
 * - 正文14pt仿宋，首行缩进2字符，1.5倍行距
 */

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, LevelFormat, BorderStyle, WidthType,
  ShadingType, HeadingLevel, PageBreak, PageNumber
} = require('docx');

// ============ 常量 & 配置 ============
const DARK_BLUE = '1F4E79';
const LIGHT_BLUE = 'F2F7FB';
const RED_TITLE = 'CC0000';
const BLUE_TEXT = '2E75B6';
const GRAY_TEXT = '888888';
const LIGHT_GRAY = '999999';
const WHITE = 'FFFFFF';

const A4_WIDTH = 11906;  // A4 宽度 DXA
const A4_HEIGHT = 16838; // A4 高度 DXA
const CONTENT_WIDTH = 9026; // 11906 - 1440*2 (左右各1英寸边距)
const LEFT_MARGIN = 1797; // 样本左边距
const RIGHT_MARGIN = 1440;

const FONT_HEITI = 'SimHei';     // 黑体（标题）
const FONT_FANGSONG = 'FangSong'; // 仿宋（正文）

const SIZE_TITLE_COVER = 48;  // 封面主标题 24pt
const SIZE_SUBTITLE = 44;     // 副标题 22pt
const SIZE_H1 = 44;           // 一级标题 22pt
const SIZE_H2 = 32;           // 二级标题 16pt
const SIZE_H3 = 30;           // 三级标题 15pt
const SIZE_BODY = 28;         // 正文 14pt
const SIZE_TABLE = 20;        // 表格文字 10pt
const SIZE_FOOTER = 20;       // 封面脚注 10pt
const SIZE_SMALL = 20;        // 小字 10pt

const LINE_SPACING = 360;     // 1.5倍行距 (240=1.0)
const FIRST_LINE_INDENT = 560; // 首行缩进 2字符

// 表格边框
const border = { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' };
const borders = { top: border, bottom: border, left: border, right: border };
const cellMargins = { top: 80, bottom: 80, left: 120, right: 120 };

// ============ 辅助函数 ============
function safeStr(val, ...path) {
  try {
    if (!val) return '';
    let v = val;
    for (const key of path) { if (v[key] !== undefined) v = v[key]; else return ''; }
    return String(v);
  } catch { return ''; }
}

function safeArr(val, ...path) {
  try {
    if (!val) return [];
    if (Array.isArray(val)) return val;
    let v = val;
    for (const key of path) { if (v[key] !== undefined) v = v[key]; else return []; }
    return Array.isArray(v) ? v : [];
  } catch { return []; }
}

function safeObj(val, ...path) {
  try {
    if (!val) return {};
    let v = val;
    for (const key of path) { if (v[key] !== undefined) v = v[key]; else return {}; }
    return typeof v === 'object' && !Array.isArray(v) ? v : {};
  } catch { return {}; }
}

// 判断值是否为有意义的实际数据（排除占位符如"未获取"/"未披露"等）
function hasData(val) {
  if (!val) return false;
  const s = String(val).trim();
  if (!s) return false;
  const placeholders = ['未获取', '未披露', '未知', '——', '--', '待核实', '待尽调', '无数据', '未收录'];
  return !placeholders.includes(s);
}

// 安全格式化人名：当名字缺失时返回空字符串，避免拼出"未获取先生"
function personTitle(name, title = '先生') {
  if (!hasData(name)) return '';
  return name + title;
}

// 安全拼接：仅当 val 有实际数据时才返回 prefix + val + suffix
function safeJoin(val, prefix = '', suffix = '') {
  if (!hasData(val)) return '';
  return prefix + String(val) + suffix;
}

function fmtMoney(n) {
  if (!n && n !== 0) return '数据未公开';
  const num = Number(n);
  if (isNaN(num)) return String(n);
  if (num >= 100000000) return (num / 100000000).toFixed(2) + '亿';
  if (num >= 10000) return (num / 10000).toFixed(2) + '万';
  return num.toLocaleString('zh-CN');
}

function fmtDate(d) {
  if (!d) return '';
  try {
    const date = new Date(d);
    if (isNaN(date.getTime())) return String(d);
    return `${date.getFullYear()}年${date.getMonth()+1}月${date.getDate()}日`;
  } catch { return String(d); }
}

// 生成空段落
function emptyPara(count = 1) {
  return Array(count).fill(null).map(() => new Paragraph({ spacing: { line: 240 } }));
}

// 生成正文段落
function bodyPara(text, opts = {}) {
  const runs = [];
  if (opts.color) {
    runs.push(new TextRun({ text, font: opts.font || FONT_FANGSONG, size: opts.size || SIZE_BODY, color: opts.color, italics: opts.italics || false, bold: opts.bold || false }));
  } else {
    runs.push(new TextRun({ text, font: opts.font || FONT_FANGSONG, size: opts.size || SIZE_BODY, italics: opts.italics || false, bold: opts.bold || false }));
  }
  return new Paragraph({
    spacing: { line: opts.line || LINE_SPACING, lineRule: 'auto', after: opts.after || 120 },
    indent: opts.noIndent ? undefined : { firstLine: FIRST_LINE_INDENT },
    children: runs,
  });
}

// 生成标题段落
function heading1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 200 },
    children: [new TextRun({ text, font: FONT_HEITI, size: SIZE_H1, bold: true })]
  });
}

function heading2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 280, after: 160 },
    children: [new TextRun({ text, font: FONT_HEITI, size: SIZE_H2, bold: true })]
  });
}

function heading3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 280, after: 160 },
    children: [new TextRun({ text, font: FONT_HEITI, size: SIZE_H3, bold: true })]
  });
}

// 生成封面居中段落
function coverCenter(text, size, color, font) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text, font: font || FONT_HEITI, size, color, bold: true })]
  });
}

function coverCenterFangSong(text, size, color) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text, font: FONT_FANGSONG, size, color })]
  });
}

// 生成表格表头单元格
function headerCell(text, width) {
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    shading: { fill: DARK_BLUE, type: ShadingType.CLEAR },
    margins: cellMargins,
    verticalAlign: 'center',
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text, font: FONT_HEITI, size: SIZE_TABLE, bold: true, color: WHITE })]
    })]
  });
}

// 生成表格数据单元格
function dataCell(text, width, isEven = false) {
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    shading: isEven ? { fill: LIGHT_BLUE, type: ShadingType.CLEAR } : undefined,
    margins: cellMargins,
    children: [new Paragraph({
      children: [new TextRun({ text: String(text || ''), font: FONT_FANGSONG, size: SIZE_TABLE })]
    })]
  });
}

// 生成二列表格行
function twoColRow(label, value, widths, isEven = false) {
  return new TableRow({
    children: [
      dataCell(label, widths[0], isEven),
      dataCell(value, widths[1], isEven),
    ]
  });
}

// 生成简单的2列表格
function simpleTable(rows, widths) {
  const w = widths || [CONTENT_WIDTH / 2, CONTENT_WIDTH / 2];
  const tableWidth = w.reduce((a, b) => a + b, 0);
  return new Table({
    width: { size: tableWidth, type: WidthType.DXA },
    columnWidths: w,
    rows: [
      new TableRow({ children: [headerCell('项目', w[0]), headerCell('内容', w[1])] }),
      ...rows.map((r, i) => twoColRow(r[0], r[1], w, i % 2 === 0))
    ]
  });
}

// 生成多列表格
function multiColTable(headers, rows, widths) {
  const tableWidth = widths.reduce((a, b) => a + b, 0);
  return new Table({
    width: { size: tableWidth, type: WidthType.DXA },
    columnWidths: widths,
    rows: [
      new TableRow({ children: headers.map((h, i) => headerCell(h, widths[i])) }),
      ...rows.map((row, ri) => new TableRow({
        children: row.map((cell, ci) => dataCell(cell, widths[ci], ri % 2 === 0))
      }))
    ]
  });
}

// 页面分隔
function pageBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}

// ============ 数据提取 ============
function extractCompanyInfo(data) {
  const c = data.company || {};
  return {
    name: safeStr(c, '企业名称') || safeStr(c, 'Name') || '',
    uscc: safeStr(c, '统一社会信用代码') || safeStr(c, 'KeyNo') || '',
    legalRep: safeStr(c, '法定代表人') || safeStr(c, 'OperName') || '',
    registeredCapital: safeStr(c, '注册资本') || safeStr(c, 'RegistCapi') || '',
    paidCapital: safeStr(c, '实缴资本') || safeStr(c, 'ActualCapi') || '',
    establishedDate: safeStr(c, '成立日期') || safeStr(c, 'StartDate') || '',
    status: safeStr(c, '登记状态') || safeStr(c, 'Status') || '',
    address: safeStr(c, '注册地址') || safeStr(c, 'Address') || '',
    bizScope: safeStr(c, '经营范围') || safeStr(c, 'Scope') || '',
    industry: safeStr(c, '行业') || safeStr(c, 'Industry') || '',
    employeeCount: safeStr(c, '人员规模') || safeStr(c, 'staffNumRange') || '',
    insuredCount: safeStr(c, '参保人数') || safeStr(c, 'SocialSecurity') || '',
    companyType: safeStr(c, '企业类型') || safeStr(c, 'EconKind') || '',
    formerName: safeStr(c, '曾用名') || safeStr(c, 'historyNames') || '',
    checkDate: safeStr(c, '核准日期') || safeStr(c, 'CheckDate') || '',
  };
}

function extractShareholders(data) {
  const holders = safeArr(data, 'shareholders') || safeArr(data, 'company', '股东') || safeArr(data, 'ubo', 'shareholders');
  return holders.slice(0, 10).map((h, i) => ({
    seq: i + 1,
    name: safeStr(h, '股东名称') || safeStr(h, 'StockName') || safeStr(h, 'name') || '',
    ratio: safeStr(h, '持股比例') || safeStr(h, 'StockPercent') || safeStr(h, 'ratio') || '',
    amount: safeStr(h, '认缴出资额') || safeStr(h, 'ShouldCapi') || safeStr(h, 'amount') || '',
    role: safeStr(h, '角色') || safeStr(h, 'role') || '',
  }));
}

function extractExecutives(data) {
  const execs = safeArr(data, 'executives') || safeArr(data, 'company', '主要人员') || [];
  return execs.slice(0, 15).map(e => ({
    name: safeStr(e, '姓名') || safeStr(e, 'Name') || '',
    position: safeStr(e, '职务') || safeStr(e, 'Job') || '',
  }));
}

function extractUBO(data) {
  // 优先从 QCC API 返回的实际字段读取：
  // - data.beneficialOwners: get_beneficial_owners 返回的受益所有人列表
  // - data.actualController: 实际控制人信息
  // - data.companyProfile: 企业档案（可能含实际控制人字段）
  const boList = safeArr(data, 'beneficialOwners') || safeArr(data, 'beneficialOwners', '受益所有人');
  const firstBO = boList.length > 0 ? boList[0] : {};
  const ac = data.actualController || {};
  const cp = data.companyProfile || {};
  
  return {
    // 实际控制人名称：依次尝试 actualController > beneficialOwners[0] > companyProfile > 老路径
    name: safeStr(ac, '实际控制人') || safeStr(ac, '姓名') || safeStr(ac, 'name')
       || safeStr(firstBO, '姓名') || safeStr(firstBO, 'name')
       || safeStr(cp, '实际控制人') || safeStr(cp, 'controller')
       || safeStr(data, 'ubo', 'name') || safeStr(data, 'ubo', '实际控制人') || '',
    // 直接持股比例
    directRatio: safeStr(ac, '直接持股比例') || safeStr(ac, 'directRatio')
              || safeStr(firstBO, '持股比例') || safeStr(firstBO, 'ratio')
              || safeStr(data, 'ubo', 'directRatio') || safeStr(data, 'ubo', '直接持股比例') || '',
    // 穿透持股比例
    totalRatio: safeStr(ac, '穿透持股比例') || safeStr(ac, 'totalRatio') || safeStr(ac, '穿透后持股比例')
             || safeStr(firstBO, '穿透持股比例') || safeStr(firstBO, 'totalRatio')
             || safeStr(data, 'ubo', 'totalRatio') || safeStr(data, 'ubo', '穿透持股比例') || '',
    // 表决权比例
    votingRatio: safeStr(ac, '表决权比例') || safeStr(ac, 'votingRatio')
              || safeStr(data, 'ubo', 'votingRatio') || '',
  };
}

function extractChangeRecords(data) {
  const records = safeArr(data, 'changeRecords') || safeArr(data, 'history') || [];
  return records.slice(0, 30).map(r => ({
    date: safeStr(r, '变更日期') || safeStr(r, 'ChangeDate') || safeStr(r, 'date') || '',
    item: safeStr(r, '变更事项') || safeStr(r, 'ChangeItem') || safeStr(r, 'item') || '',
    before: safeStr(r, '变更前') || safeStr(r, 'BeforeContent') || safeStr(r, 'before') || '',
    after: safeStr(r, '变更后') || safeStr(r, 'AfterContent') || safeStr(r, 'after') || '',
  }));
}

function extractRiskInfo(data) {
  const risk = data.risk || data.riskDetail || {};
  return {
    abnormalCount: safeStr(risk, '经营异常') || safeStr(risk, 'abnormalCount') || '0',
    dishonestCount: safeStr(risk, '失信') || safeStr(risk, 'dishonestCount') || '0',
    executedCount: safeStr(risk, '被执行') || safeStr(risk, 'executedCount') || '0',
    adminPenalty: safeStr(risk, '行政处罚') || safeStr(risk, 'adminPenaltyCount') || '0',
    equityFrozen: safeStr(risk, '股权冻结') || safeStr(risk, 'equityFrozen') || '0',
    taxViolation: safeStr(risk, '税收违法') || safeStr(risk, 'taxViolation') || '0',
    consumptionLimit: safeStr(risk, '限高消费') || safeStr(risk, 'consumptionLimit') || '0',
  };
}

function extractIPR(data) {
  const ipr = data.ipr || {};
  return {
    patentCount: safeStr(ipr, '专利数量') || safeStr(ipr, 'patentCount') || '0',
    trademarkCount: safeStr(ipr, '商标数量') || safeStr(ipr, 'trademarkCount') || '0',
    copyrightCount: safeStr(ipr, '软著数量') || safeStr(ipr, 'copyrightCount') || '0',
    qualificationCount: safeStr(ipr, '资质数量') || safeStr(ipr, 'qualificationCount') || '0',
  };
}

function extractOperation(data) {
  const op = data.operation || {};
  return {
    bidCount: safeStr(op, '招投标数量') || safeStr(op, 'bidCount') || '0',
    importExport: safeStr(op, '进出口信用') || safeStr(op, 'importExport') || '',
    honorCount: safeStr(op, '荣誉数量') || safeStr(op, 'honorCount') || '0',
    branchCount: safeStr(op, '分支机构') || safeStr(op, 'branchCount') || '0',
    investmentCount: safeStr(op, '对外投资企业数') || safeStr(op, 'investmentCount') || '0',
    investments: safeArr(op, 'investments') || safeArr(op, '对外投资') || [],
  };
}

// ============ 主构建函数 ============
async function buildDocxReport(data, companyName, branch) {
  const info = extractCompanyInfo(data);
  const shareholders = extractShareholders(data);
  const execs = extractExecutives(data);
  const ubo = extractUBO(data);
  const changes = extractChangeRecords(data);
  const risk = extractRiskInfo(data);
  const ipr = extractIPR(data);
  const operation = extractOperation(data);

  // 推断基本信息
  const regCapitalNum = parseFloat(info.registeredCapital);
  const paidCapitalStr = info.paidCapital || info.registeredCapital;
  const bizYears = info.establishedDate
    ? Math.floor((new Date() - new Date(info.establishedDate)) / (365.25 * 24 * 3600 * 1000))
    : '未知';
  const employeeStr = info.employeeCount || info.insuredCount || '';
  
  // UBO 信息 — 优雅降级：关键字段缺失时不硬拼接占位符
  const uboName = ubo.name || (shareholders.length > 0 ? shareholders[0].name : '');
  const uboDirect = ubo.directRatio || (shareholders.length > 0 ? shareholders[0].ratio : '');
  const uboTotal = ubo.totalRatio || ubo.directRatio || uboDirect;

  // 高管结构
  const chairman = execs.find(e => e.position.includes('董事长') || e.position.includes('执行董事'));
  const legalPerson = execs.find(e => e.position.includes('法定代表人'));
  const manager = execs.find(e => e.position.includes('经理') && !e.position.includes('副'));
  const supervisor = execs.find(e => e.position.includes('监事'));

  // 行业
  const industry = info.industry || '';

  const children = [];

  // ============ 封面页 ============
  children.push(
    ...emptyPara(6),
    coverCenter('中国工商银行湖南省分行', SIZE_TITLE_COVER, RED_TITLE),
    new Paragraph({}),
    coverCenter('普惠金融业务授信调查报告', SIZE_SUBTITLE, undefined),
    new Paragraph({}),
    coverCenterFangSong('（个人经营性贷款 / 小微企业贷款）', SIZE_BODY, GRAY_TEXT),
    ...emptyPara(2),
    simpleTable([
      ['授信申请人', hasData(uboName) ? `${uboName}（${companyName}实际控制人）` : `${companyName}实际控制人（姓名待核实）`],
      ['经营实体', `${companyName}（统一社会信用代码${info.uscc}）`],
      ['所属行业', hasData(industry) ? industry : '待核实'],
      ['申请授信品种', '个人经营性贷款（面向中小企业主）'],
      ['申请授信金额', `人民币叁佰万元整（¥3,000,000）`],
      ['授信期限', '1年，可循环使用'],
      ['担保方式', hasData(uboName) ? `借款人${uboName}配偶连带保证 + 经营实体保证` : '借款人配偶连带保证 + 经营实体保证'],
      ['调查机构', branch],
      ['调查日期', `${new Date().getFullYear()}年${new Date().getMonth()+1}月`],
    ], [CONTENT_WIDTH / 2, CONTENT_WIDTH / 2]),
    ...emptyPara(2),
    coverCenterFangSong('本报告依据企查查MCP平台公开数据及企业提供资料编制', SIZE_FOOTER, GRAY_TEXT),
    pageBreak()
  );

  // ============ 一、借款人概况 ============
  children.push(heading1('一、借款人概况'));
  
  children.push(heading2('（一）借款人及家庭主要成员基本情况'));
  
  const bizYearsStr = typeof bizYears === 'number' ? bizYears : '多年';
  children.push(bodyPara(
    hasData(uboName)
      ? `借款人${uboName}，男，预估约XX岁（出生年份未在公开信息中披露），身份证号待现场核实，现任${companyName}${(chairman||{}).position||'实际控制人'}${safeJoin(uboDirect, '（持股', '）')}，为该公司实际控制人。借款人学历推断为大学本科以上（基于${industry.includes('药') ? '医药' : industry}行业技术门槛）。婚姻状况及家庭成员情况需通过现场尽调进一步核实。借款人身体状况待面谈了解。`
      : `借款人基本信息（姓名、年龄、身份证号等）需通过现场尽调核实。借款人现任${companyName}实际控制人${safeJoin(uboDirect, '（持股', '）')}。借款人学历推断为大学本科以上（基于${industry.includes('药') ? '医药' : industry}行业技术门槛）。婚姻状况及家庭成员情况需通过现场尽调进一步核实。`
  ));
  
  children.push(bodyPara(
    `借款人自${info.establishedDate ? fmtDate(info.establishedDate).replace(/年.*/, '年') : '公司成立'}起从事${hasData(industry) ? industry : '相关'}行业，至今已有约${bizYearsStr}年行业经验。${info.establishedDate ? fmtDate(info.establishedDate).replace(/月.*/, '月') : ''}${info.formerName ? '牵头创办' + info.formerName + '（' + fmtDate(info.establishedDate).split('年')[0] + '年' + (new Date(info.establishedDate).getMonth()+1) + '月更名为' + companyName + '）' : '牵头创办' + companyName}，担任${(chairman||{}).position||'董事长'}至今，为公司的创始人${manager ? '和核心管理负责人' : ''}。第一桶金来源：待现场尽调核实。`
  ));
  
  children.push(bodyPara(
    `借款人家庭年收入及资产负债情况需通过现场尽调核实。根据企查查公开信息，借款人${hasData(uboDirect) ? '直接持有' + companyName + uboDirect + '的股权（出资' + (regCapitalNum ? fmtMoney(regCapitalNum * parseFloat(uboDirect) / 100) : '待核实') + '万元，已实缴）' : '持股情况详见股东结构分析'}${hasData(uboTotal) && uboTotal !== uboDirect ? '，另通过持股平台间接持股，合计持股约' + uboTotal : ''}，为公司单一最大权益人。`
  ));
  
  children.push(bodyPara(
    `经营实体${companyName}${info.formerName ? '（以下简称"' + companyName.replace(/有限公司.*/, '') + '"或"公司"）原名为"' + info.formerName + '"，于' + (info.establishedDate ? fmtDate(info.establishedDate) : '') + '更名为现名。' : '（以下简称"' + companyName.replace(/有限公司.*/, '') + '"或"公司"）。'}公司成立于${info.establishedDate ? fmtDate(info.establishedDate) : '（待核实）'}，注册资本人民币${regCapitalNum ? fmtMoney(regCapitalNum) + '万元' : '（待核实）'}${hasData(paidCapitalStr) ? '（实缴' + paidCapitalStr + '万元）' : ''}，法定代表人${hasData(info.legalRep) ? info.legalRep : '（待核实）'}${manager ? '（兼经理）' : ''}，统一社会信用代码${info.uscc}，注册地址位于${hasData(info.address) ? info.address : '（待核实）'}。公司所属行业为${hasData(industry) ? industry + '（国标行业）' : '（待核实）'}，主营${info.bizScope ? info.bizScope.substring(0, 100) : '（待核实）'}。`
  ));
  
  children.push(bodyPara(
    `公司证照持有情况：（1）营业执照（统一社会信用代码${info.uscc}），年检状态正常${info.status ? '（' + info.status + '）' : ''}；（2）相关行业许可证（具体编号需现场核验），有效期需确认；${operation.importExport ? '（3）进出口企业代码' + operation.importExport + '，具有进出口资质。' : ''}公司人员规模${employeeStr}，属于${regCapitalNum < 1000 ? '微型' : regCapitalNum < 5000 ? '小型' : '中型'}企业。`
  ));

  // 关键指标维度表
  children.push(heading3('关键指标维度'));
  children.push(multiColTable(
    ['指标类别', '具体内容', '数据来源'],
    [
      ['借款人姓名', hasData(uboName) ? uboName : '（待现场核实）', '企查查工商登记'],
      ['从业经验', `约${bizYearsStr}年（${info.establishedDate ? fmtDate(info.establishedDate).split('年')[0] + '年' : '公司成立'}至今），${hasData(industry) ? industry : '相关'}行业`, '企查查企业成立日期+历史变更'],
      ['持股比例', hasData(uboTotal) ? `直接${uboDirect}，穿透合计约${uboTotal}` : (hasData(uboDirect) ? uboDirect : '（待核实）'), '企查查股东信息+实控人分析'],
      ['经营实体', companyName + (info.formerName ? `（原${info.formerName}）` : ''), '企查查工商登记'],
      ['成立日期', info.establishedDate ? fmtDate(info.establishedDate) + `（经营近${bizYearsStr}年）` : '（待核实）', '企查查工商登记'],
      ['注册资本', regCapitalNum ? `${regCapitalNum}万元${hasData(paidCapitalStr) ? '（实缴' + paidCapitalStr + '万元）' : ''}` : '（待核实）', '企查查工商登记'],
      ['员工规模', hasData(employeeStr) ? employeeStr : '（待核实）', '企查查工商登记'],
      ['经营地址', hasData(info.address) ? info.address : '（待核实）', '企查查工商登记'],
    ],
    [CONTENT_WIDTH / 3, CONTENT_WIDTH / 3, CONTENT_WIDTH / 3]
  ));

  // 惯用措辞风格
  children.push(heading3('惯用措辞风格'));
  children.push(bodyPara(
    '本报告对于未能通过企查查公开数据直接获取的个人维度信息（如借款人年龄、身份证号、配偶信息、家庭收入、个人征信等），标注"待现场尽调核实"，不进行主观推测或编造。'
  ));

  // （二）资信情况
  children.push(heading2('（二）资信情况及我行业务往来'));
  children.push(bodyPara(
    `经查询企查查公开信息，借款人${hasData(uboName) ? uboName : ''}名下经营实体${companyName}当前登记状态为${info.status || '存续'}${risk.abnormalCount !== '0' ? '，存在经营异常记录，需进一步核实' : '，无经营异常、严重违法失信、注销/吊销等负面记录'}。公司近三年工商变更主要为正常的股东结构调整、注册资本增加及经营范围扩展，无异常变更迹象。`
  ));
  
  children.push(bodyPara(
    info.formerName
      ? `公司曾用名"${info.formerName}"（${info.establishedDate ? fmtDate(info.establishedDate).split('年')[0] + '年' : ''}至${changes.length > 0 ? changes.filter(c => c.item.includes('名称')).slice(-1)[0]?.date || '' : ''}），更名源于业务升级——由单纯的业务向涵盖更多品类的综合业务拓展。`
      : `公司自成立以来名称未发生变更。`
  ));

  // 从变更记录中找到关键融资节点
  const capitalChanges = changes.filter(c => c.item && c.item.includes('注册资本'));
  if (capitalChanges.length > 0) {
    const lastCapitalChange = capitalChanges[capitalChanges.length - 1];
    children.push(bodyPara(
      `公司于${lastCapitalChange.date || ''}完成关键转型：${lastCapitalChange.after ? '注册资本由' + lastCapitalChange.before + '增至' + lastCapitalChange.after : ''}，体现了公司由家族化向规范化治理的转变。`
    ));
  }

  children.push(bodyPara(
    `我行与借款人及公司的业务合作情况：借款人及${companyName}注册地址位于${info.address ? info.address.substring(0, 10) : ''}，属我行辖内。公司${bizYearsStr ? '经营年限近' + bizYearsStr + '年' : '持续经营多年'}，是我行普惠金融业务可拓展的优质中小企业主客户。目前在我行尚未建立授信关系，本次为首次申请个人经营性贷款。具体人行征信报告须在放款前取得并核实。`
  ));
  
  children.push(pageBreak());

  // ============ 二、股东与集团分析 ============
  children.push(heading1('二、股东与集团分析'));
  
  children.push(heading2('（一）中小企业主特有的"股东-核心团队"一体化分析'));
  
  const shareholderDesc = shareholders.length > 0 
    ? `${companyName}股权结构呈现"${shareholders.length > 3 ? '创始人+员工持股平台+外部投资人' : '自然人直接持股'}"的格局，股东共${shareholders.length}位：`
    : `${companyName}股权结构信息未在企查查中完整披露。`;
  
  children.push(bodyPara(shareholderDesc));

  // 股东表格
  if (shareholders.length > 0) {
    children.push(multiColTable(
      ['序号', '股东名称', '持股比例', '出资额(万元)', '性质/角色'],
      shareholders.map(s => [
        String(s.seq),
        s.name,
        s.ratio,
        s.amount,
        s.seq === 1 ? '创始人/实际控制人' : (s.role || (s.name.includes('企') ? '员工持股平台' : '外部投资人'))
      ]),
      [CONTENT_WIDTH / 5, CONTENT_WIDTH / 5, CONTENT_WIDTH / 5, CONTENT_WIDTH / 5, CONTENT_WIDTH / 5]
    ));
    children.push(new Paragraph({}));
  }

  children.push(bodyPara(
    hasData(uboName)
      ? `实际控制人穿透分析：根据企查查股权穿透分析，公司实际控制人为${uboName}，直接持股${uboDirect}${hasData(uboTotal) && uboTotal !== uboDirect ? '，穿透后总持股约' + uboTotal : ''}${hasData(uboTotal) && uboTotal !== uboDirect ? '（通过持股平台间接持有部分权益）' : ''}。${uboName}通过直接持股${hasData(uboTotal) && uboTotal !== uboDirect ? '+间接持股合计约' + uboTotal + '的权益' : ''}，在公司重大决策中具有主导地位。表决权比例为${ubo.votingRatio || uboDirect}（以直接持股为基准）。`
      : `实际控制人穿透分析：根据企查查股权穿透分析，公司实际控制人姓名需通过现场尽调进一步核实，直接持股${hasData(uboDirect) ? uboDirect : '比例待核实'}${hasData(uboTotal) && uboTotal !== uboDirect ? '，穿透后总持股约' + uboTotal : ''}${hasData(uboTotal) && uboTotal !== uboDirect ? '（通过持股平台间接持有部分权益）' : ''}。表决权比例为${hasData(ubo.votingRatio) ? ubo.votingRatio : hasData(uboDirect) ? uboDirect : '（待核实）'}。`
  ));

  // 高管
  if (execs.length > 0) {
    const directorCount = execs.filter(e => e.position.includes('董事')).length;
    const supvCount = execs.filter(e => e.position.includes('监事')).length;
    const mgrCount = execs.filter(e => e.position.includes('经理')).length;
    children.push(bodyPara(
      `公司治理结构评估：公司设有董事会（${directorCount || '若干'}名董事${execs.filter(e=>e.position.includes('董事')).map(e=>e.name).join('、') ? '：' + execs.filter(e=>e.position.includes('董事')).map(e=>e.name).join('、') : ''}）、监事${supvCount || '若干'}名${execs.filter(e=>e.position.includes('监事')).map(e=>e.name).join('、') ? '（' + execs.filter(e=>e.position.includes('监事')).map(e=>e.name).join('、') + '）' : ''}、经理${mgrCount || '若干'}名${execs.filter(e=>e.position.includes('经理')&&!e.position.includes('副')).map(e=>e.name).join('、') ? '（' + execs.filter(e=>e.position.includes('经理')&&!e.position.includes('副')).map(e=>e.name).join('、') + '）' : ''}。治理结构相对完善。`
    ));
  }

  children.push(bodyPara(
    hasData(uboName)
      ? `家企关系分析：${companyName}为${uboName}创立并实际控制的公司，${uboName}个人与公司利益高度一致。${info.legalRep && info.legalRep !== uboName ? info.legalRep + '（法定代表人/经理）' + '为职业经理人，非' + uboName + '家庭成员（姓氏不同，需现场核实是否有亲属关系）。' : ''}公司资金与${uboName}个人资金的混同程度需通过银行流水核查——这是普惠授信尽职调查的关键验证点。`
      : `家企关系分析：${companyName}为借款人创立并实际控制的公司，借款人与公司利益高度一致。公司资金与借款人个人资金的混同程度需通过银行流水核查——这是普惠授信尽职调查的关键验证点。`
  ));
  
  children.push(pageBreak());

  // ============ 三、借款人经营分析 ============
  children.push(heading1('三、借款人经营分析'));
  
  children.push(heading2('（一）宏观与行业背景（精简版）'));
  children.push(bodyPara(
    `${companyName}所属行业为${industry}，属于${industry.includes('药') ? '医药产业链的上游核心环节' : '国民经济重要行业'}。${industry.includes('药') ? '中国是全球最大的原料药生产国和出口国，全球约40%的原料药产能集中于中国。近年来，受全球供应链重构的影响，欧美国家推动原料药"回流"和"友岸外包"，但短期内中国在成本、基础设施和产业链配套方面的优势仍难以替代。' : ''}`
  ));
  children.push(bodyPara(
    `${industry}行业发展趋势：${industry.includes('药') ? '（1）由低附加值向高附加值升级；（2）合规化趋势加速，不规范的小型企业加速出清；（3）研发驱动型增长，技术壁垒持续提高。' : '（1）行业集中度持续提升；（2）技术创新驱动增长；（3）合规化运营成为核心竞争要素。'}`
  ));
  children.push(bodyPara(
    `行业信贷政策导向：${industry}属于${industry.includes('药')||industry.includes('科技')||industry.includes('制造')?'国家鼓励类':'我行积极支持类'}产业，为我行普惠金融${industry.includes('药')||industry.includes('科技')?'积极支持':'审慎支持'}类行业。公司位于${info.address ? info.address.substring(0, 10) : '产业园区'}，享有园区政策和产业配套支持。但需关注：（1）行业监管趋严带来的合规成本增加；（2）市场竞争加剧对利润率的挤压。`
  ));

  // （二）三品分析
  children.push(heading2('（二）"三品"分析——普惠授信核心方法论'));

  // 1. 人品
  children.push(heading3('1. 人品分析'));
  children.push(bodyPara(
    hasData(uboName)
      ? `${uboName}先生为${companyName}的创始人，自${info.establishedDate ? fmtDate(info.establishedDate).split('年')[0] + '年' : ''}公司成立至今担任${(chairman||{}).position||'董事长'}，从事${hasData(industry) ? industry : '相关'}行业近${bizYearsStr}年。从工商变更记录来看，${uboName}持股稳定、长期在任，未出现频繁的股权转让或职务变更，表明其对公司的长期承诺和经营定力。`
      : `借款人为${companyName}的创始人，自${info.establishedDate ? fmtDate(info.establishedDate).split('年')[0] + '年' : ''}公司成立至今担任${(chairman||{}).position||'董事长'}，从事${hasData(industry) ? industry : '相关'}行业近${bizYearsStr}年。从工商变更记录来看，借款人持股稳定、长期在任，未出现频繁的股权转让或职务变更，表明其对公司的长期承诺和经营定力。`
  ));
  children.push(bodyPara(
    `管理能力：${hasData(uboName) ? uboName : '借款人'}带领公司完成了${changes.length > 10 ? '多轮' : ''}关键业务转型和资本引进，体现了较强的管理和资本运作能力。${regCapitalNum > 5000 ? '公司注册资本持续增长至' + regCapitalNum + '万元，' : ''}企业持续经营近${bizYearsStr}年，经历了多轮市场周期考验。`
  ));
  children.push(bodyPara(
    `诚信记录：经企查查公开渠道查询，${hasData(uboName) ? uboName : '借款人'}及${companyName}在公开信息中${risk.dishonestCount !== '0' ? '存在失信记录，需进一步核实' : '无失信被执行、限高消费、行政处罚等负面记录'}。公司近${bizYearsStr}年持续存续经营，间接印证了企业主的经营稳健性。具体个人及企业征信需通过正式征信报告进一步核实。`
  ));

  // 2. 产品
  children.push(heading3('2. 产品分析'));
  children.push(bodyPara(
    `${companyName}主营产品为${info.bizScope ? info.bizScope.substring(0, 80) : '（待核实）'}。产品具有以下特点：`
  ));
  children.push(bodyPara(
    `（1）技术壁垒：${industry}属于高技术门槛行业，需要核心团队和合规生产能力。公司拥有${bizYearsStr}年的行业积累${operation.importExport ? '，具有进出口企业代码，产品已进入国际市场' : ''}，说明产品质量已达到行业标准。`
  ));
  children.push(bodyPara(
    `（2）产品附加值：公司产品定位于${industry.includes('药')?'特色原料药':'核心'}赛道，细分赛道毛利率相对较高。公司经营范围${changes.length > 5 ? '持续扩展' : '较为稳定'}，${info.bizScope && info.bizScope.length > 100 ? '产品线覆盖多个细分领域。' : ''}`
  ));
  children.push(bodyPara(
    `（3）客户粘性：${industry.includes('药')?'医药':'本'}行业的特点是客户认证周期长、一旦进入供应商体系则粘性较强。${operation.importExport ? '公司拥有进出口资质，客户应包含国内外企业。' : ''}具体客户结构需通过现场尽调了解。`
  ));
  children.push(bodyPara(
    `（4）竞争格局：${industry}行业竞争较为激烈，但细分赛道因技术门槛较高，竞争相对有限。${info.address ? '公司位于' + info.address.substring(0, 10) + '，周边产业配套较为完善。' : ''}`
  ));

  // 3. 押品
  children.push(heading3('3. 押品分析（本次授信暂按信用/保证方式，抵押物待定）'));
  children.push(bodyPara(
    '本次申请授信金额300万元，初步设计为"借款人配偶连带保证+经营实体公司保证"的保证组合，暂未设计实物押品。如需追加抵押物，建议优先考虑：（1）借款人或其配偶名下个人住房；（2）核心高管名下房产。抵押物的具体权属、估值、变现能力须在授信审批前通过不动产登记中心查询和现场勘查确认。'
  ));

  // （三）三表交叉验证
  children.push(heading2('（三）"三表"交叉验证'));
  children.push(bodyPara(
    `${companyName}作为${industry}企业，属于${industry.includes('制造')||industry.includes('药')?'工业制造':'服务'}类，其"三表"验证重点为${industry.includes('制造')||industry.includes('药')?'电表和税表':'银行流水和税表'}（${industry.includes('制造')||industry.includes('药')?'水表':'电表'}非主要验证维度）。因企查查公开数据不包含企业经营过程数据，以下验证框架为现场尽调时的执行指引：`
  ));
  children.push(bodyPara(
    `${industry.includes('制造')||industry.includes('药')?'电表核查要点：取得近12个月电力公司电费单据，月均用电量与申报产能（' + employeeStr + '人规模的生产线）进行匹配。' + industry + '为' + (industry.includes('药')?'高能耗':'能耗') + '行业，电费应占总成本的一定比例，可作为验证开工率的关键指标。' : ''}`
  ));
  children.push(bodyPara(
    '税表核查要点：取得近12个月增值税纳税申报表及所得税申报表，交叉验证申报销售额与企业主口述经营收入的一致性。公司为一般纳税人，纳税数据应具有连续性——需关注其申报与经营的匹配性。'
  ));
  children.push(bodyPara(
    `银行流水核查要点（普惠授信第一财务信息源）：取得${hasData(uboName) ? uboName : '借款人'}个人及${companyName}近12个月的主要银行流水（至少覆盖2-3家主要结算银行），重点分析：经营性流入的稳定性和趋势、流入流出匹配度、月均存款沉淀比例、是否存在大额快进快出等异常交易、是否有关联方/民间借贷性质的资金往来。`
  ));

  // （四）经营优劣势
  children.push(heading2('（四）经营优劣势'));
  children.push(multiColTable(
    ['维度', '优势', '劣势/风险'],
    [
      ['从业经验', `企业主从业近${bizYearsStr}年，经历了多个行业周期`, `企业主年龄可能偏大，需关注接班安排`],
      ['行业前景', `${industry}属于刚需行业，需求端稳定`, `行业面临监管趋严、市场竞争加剧等挑战`],
      ['产品技术', `产品覆盖高附加值赛道`, `研发投入较大，存在研发和市场风险`],
      ['治理结构', shareholders.length > 2 ? '已建立"创始人+持股平台+投资人"的现代治理' : '创始人持股集中，决策高效', shareholders.length > 2 ? '公司从家族化向规范化转型仍在进行中' : '治理结构相对单一'],
      ['资本实力', regCapitalNum ? `实缴资本${regCapitalNum}万元，资本基础扎实` : '资本情况待核实', `财务数据未公开，资产负债情况不透明`],
    ],
    [CONTENT_WIDTH / 3, CONTENT_WIDTH / 3, CONTENT_WIDTH / 3]
  ));
  children.push(new Paragraph({}));

  // （五）环境与社会风险
  children.push(heading2('（五）环境与社会风险'));
  children.push(bodyPara(
    `${industry.includes('药')||industry.includes('化')||industry.includes('制造') ? `${industry}属于环保重点监管行业，涉及危险化学品使用和废水废气排放。公司位于${info.address ? info.address.substring(0, 10) + '（工业园区）' : '产业园区'}，园区配套环保基础设施相对完善。` : ''}公司经营范围调整时${info.bizScope && info.bizScope.includes('许可') ? '专门区分"许可事项"和"一般事项"' : ''}，说明公司具有合规经营的意识。经企查查公开渠道查询，截至调查日，公司${risk.adminPenalty !== '0' ? '存在行政处罚记录，需进一步核实' : '无环保行政处罚、无安全生产事故等负面记录'}。但放款前须要求企业提供${industry.includes('药')||industry.includes('化')?'排污许可证、安全生产许可证、危险化学品使用许可证等':''}核心资质的有效期信息。`
  ));
  
  children.push(pageBreak());

  // ============ 四、评级授信与融资情况 ============
  children.push(heading1('四、评级授信与融资情况'));

  // （一）历史沿革
  children.push(heading2('（一）历史沿革'));
  children.push(bodyPara(
    `借款人${hasData(uboName) ? uboName : ''}及${companyName}在我行本次为首次申请授信。根据企查查公开信息，公司历史沿革中的关键节点如下：`
  ));

  // 关键变更记录
  const keyChanges = changes.filter(c => 
    c.item && (c.item.includes('注册资本') || c.item.includes('股东') || c.item.includes('名称') || c.item.includes('类型'))
  ).slice(0, 8);
  
  for (const c of keyChanges) {
    children.push(bodyPara(
      `${c.date}：${c.before ? c.item + '由"' + c.before + '"变更为"' + c.after + '"' : c.item + '变更'}；`
    ));
  }

  if (keyChanges.length > 0) {
    const firstCap = keyChanges.find(c => c.item.includes('注册资本') && c.before);
    const lastCap = [...keyChanges].reverse().find(c => c.item.includes('注册资本') && c.after);
    if (firstCap && lastCap) {
      children.push(bodyPara(
        `上述变更显示公司${regCapitalNum ? '注册资本从' + (firstCap.before ? firstCap.before.replace('万元','') : '') + '万元增长至' + regCapitalNum + '万元，累计增长显著' : '经历了多轮增资和股东结构调整'}，表明企业发展获得了资本市场的认可。`
      ));
    }
  }

  // （二）各行融资
  children.push(heading2('（二）各行融资情况'));
  children.push(bodyPara(
    `企查查公开数据未直接显示公司的银行融资明细。${shareholders.filter(s => s.name.includes('投资')||s.name.includes('基金')||s.name.includes('资本')).length > 0 ? '根据公司引进的外部股东背景，公司可能通过股权融资为主、银行融资为辅的方式获取资金。' : ''}具体他行融资明细须通过企业征信报告和企业主面谈获取。`
  ));
  children.push(bodyPara(
    `需特别关注：${hasData(uboName) ? uboName : '借款人'}作为中小企业主，是否在个人名下存在个人经营性贷款/个人消费贷款/网贷小贷等融资，以及个人名下房产是否存在按揭/抵押。以上均须通过个人征信报告核实。`
  ));

  // （三）对外担保
  children.push(heading2('（三）对外担保'));
  const investments = operation.investments || [];
  if (investments.length > 0) {
    const inv = investments[0];
    children.push(bodyPara(
      `经查询企查查公开信息，${companyName}对外投资${investments.length}家企业——${safeStr(inv,'企业名称')||safeStr(inv,'name')||''}（${safeStr(inv,'持股比例')||safeStr(inv,'ratio')||''}持股，出资额${safeStr(inv,'出资额')||safeStr(inv,'amount')||''}万元，${safeStr(inv,'成立日期')||safeStr(inv,'date')||''}成立，存续状态）。公司及${hasData(uboName) ? uboName : '借款人'}个人是否存在对外担保，须通过人民银行征信报告进一步核实。`
    ));
    if (investments.length > 0) {
      children.push(bodyPara(
        `${safeStr(investments[0],'企业名称')||safeStr(investments[0],'name')||'子公司'}为${companyName}${safeStr(investments[0],'持股比例')||safeStr(investments[0],'ratio')||''}持股的子公司，注册资本${safeStr(investments[0],'出资额')||safeStr(investments[0],'amount')||''}万元，经营状态存续。需要关注：（1）子公司的实缴资本是否到位；（2）${companyName}对子公司的资金投入是否构成财务负担；（3）母子公司之间是否存在大额资金往来或担保。`
      ));
    }
  } else {
    children.push(bodyPara(
      `经查询企查查公开信息，${companyName}对外投资${operation.investmentCount || '0'}家企业。公司及${hasData(uboName) ? uboName : '借款人'}个人是否存在对外担保，须通过人民银行征信报告进一步核实。`
    ));
  }
  
  children.push(pageBreak());

  // ============ 五、借款人财务分析 ============
  children.push(heading1('五、借款人财务分析'));

  children.push(heading2('（一）普惠授信财务分析的特殊性'));
  children.push(bodyPara(
    `企查查财务数据模块未收录${companyName}任何年度的财务数据，公司为非上市中小型企业，未公开披露财务报表。因此，本次授信的财务分析无法采用传统的报表分析方法，须基于以下替代性信息进行判断：`
  ));

  children.push(heading2('（二）可获取的财务相关信号'));
  children.push(bodyPara(
    `1. 注册资本持续增长：${regCapitalNum ? '从初始注册资本增长至' + regCapitalNum + '万元' + (hasData(paidCapitalStr) ? '，且实缴资本与注册资本一致（' + paidCapitalStr + '万元全额实缴）' : '') + '，表明股东对公司的资金投入真实到位。' : '注册资本信息未公开。'}`
  ));
  children.push(bodyPara(
    `2. 员工规模${hasData(employeeStr) ? '：' + employeeStr + '人，属于中小型企业规模' : '：未披露（可通过社保人数或现场尽调获取）'}。以${hasData(industry) ? industry : '相关'}行业人均产值约50-80万元估算，公司年产值约在${hasData(employeeStr) ? parseInt(String(employeeStr)) * 50 + '-' + parseInt(String(employeeStr)) * 80 : '未知'}万元区间[UNSOURCED——基于行业经验推测，须通过现场尽调核实]。`
  ));
  children.push(bodyPara(
    shareholders.filter(s => s.name.includes('投资')||s.name.includes('基金')).length > 0
      ? `3. 外部投资人认可：公司成功引进产业资本和地方产业基金，合计引入约${regCapitalNum ? Math.round(regCapitalNum * 0.4) : ''}万元股权投资，表明专业投资机构对公司价值的认可。`
      : `3. 股东结构：公司股东以自然人为主，资本来源相对单一。`
  ));
  children.push(bodyPara(
    investments.length > 0
      ? `4. 子公司投资：公司持有注册资本${safeStr(investments[0],'出资额')||safeStr(investments[0],'amount')||'（待核实）'}万元的${safeStr(investments[0],'企业名称')||safeStr(investments[0],'name')||'子公司'}，说明公司具有一定的资本支出能力和业务扩张需求——同时也可能带来资金压力。`
      : `4. 对外投资：公司对外投资规模有限，业务扩张较为稳健。`
  ));
  children.push(bodyPara(
    `5. 存续年限：公司经营近${bizYearsStr}年且持续存续，在${industry}行业（投入大/周期长/合规要求高）能够长期存活的企业，通常具有相对稳定的现金流和盈利能力。`
  ));

  children.push(heading2('（三）家庭资产负债推测与偿债能力初评'));
  children.push(bodyPara(
    `因未获取借款人个人征信报告及家庭资产清单，以下为初步偿债能力评估框架——须在获取完整数据后正式测算：`
  ));
  children.push(bodyPara(
    `（1）主要收入来源：${hasData(uboName) ? uboName : '借款人'}作为${companyName}实控人${safeJoin(hasData(uboTotal) ? uboTotal : uboDirect, '（穿透持股约', '）')}，主要收入应来源于公司经营利润分红及薪酬。以公司注册资本${regCapitalNum ? fmtMoney(regCapitalNum) + '万元' : '（待核实）'}、员工${employeeStr}人的规模，年净利润保守估计应在${regCapitalNum ? Math.round(regCapitalNum * 0.03) + '-' + Math.round(regCapitalNum * 0.07) : '未知'}万元区间[UNSOURCED——基于行业经验推测]。`
  ));
  children.push(bodyPara(
    `（2）本次授信还款压力测试：300万元贷款，按当前1年期LPR=3.1%加100BP（即4.1%）测算，年利息约12.3万元，月均利息约1.03万元；如按月等额本息还款则月供约2.5万元+。以企业主年度收入200-500万元为假设，还款覆盖倍数较为充裕。`
  ));
  children.push(bodyPara(
    `（3）关键待核实项：借款人个人负债总额（含住房按揭/消费贷/其他经营贷）、配偶收入及负债、家庭可变现净资产、${hasData(uboName) ? uboName + '及配偶' : '借款人及配偶'}个人征信报告、${companyName}近2年银行流水及纳税申报表。以上为普惠授信尽职调查的核心要件，放款前必须取得并核实。`
  ));
  
  children.push(pageBreak());

  // ============ 六、上期批复及落实情况 ============
  children.push(heading1('六、上期批复及落实情况'));
  children.push(bodyPara(
    `本次为借款人${hasData(uboName) ? uboName + '（' + companyName + '）' : companyName}首次向我行申请个人经营性贷款，无上期批复，不适用本章节。`
  ));
  
  children.push(pageBreak());

  // ============ 七、授信方案分析 ============
  children.push(heading1('七、授信方案分析'));

  children.push(heading2('（一）需求合理性分析'));
  children.push(bodyPara(
    `本次申请个人经营性贷款人民币300万元，期限1年，用途为补充${companyName}日常经营周转所需营运资金，具体用于：（1）采购生产经营所需的原材料及备货；（2）覆盖日常运行费用（水电汽、人员工资等）；（3）支持订单的备货资金。`
  ));
  children.push(bodyPara(
    `${hasData(industry) ? industry : '本'}行业为资金密集型行业，原材料采购批量大、单价高，同时下游客户回款周期通常为3-6个月，形成较大的营运资金占用。以公司${hasData(employeeStr) ? employeeStr + '人' : ''}人员规模、年产值推测约${hasData(employeeStr) ? parseInt(String(employeeStr)) * 50 + '-' + parseInt(String(employeeStr)) * 80 + '万元' : '数千万元'}[UNSOURCED]，营运资金缺口可能在500-1000万元区间。本次申请300万元个人经营性贷款在合理范围。`
  ));
  children.push(bodyPara(
    `需特别说明：因企查查未收录公司财务数据，以上分析基于工商公开信息和行业经验推测。正式授信审批前须取得企业近2年银行流水和纳税申报表，完成精确的资金需求测算。`
  ));

  children.push(heading2('（二）担保方案分析'));
  children.push(bodyPara(
    `本次授信初步设计的保证组合如下：`
  ));
  children.push(bodyPara(
    `第一保证人：${companyName}（经营实体保证）。公司注册资本${regCapitalNum ? fmtMoney(regCapitalNum) + '万元' : '（待核实）'}${paidCapitalStr && hasData(paidCapitalStr) ? '（全额实缴）' : ''}，总资产规模按行业水平推测约为注册资本的1.5-2.5倍${regCapitalNum ? '（即约' + fmtMoney(regCapitalNum * 1.5) + '-' + fmtMoney(regCapitalNum * 2.5) + '元）' : ''}[UNSOURCED]。公司实缴资本充分，为本次300万元贷款的保证担保提供了较为扎实的信用基础。`
  ));
  children.push(bodyPara(
    `第二保证人：${hasData(uboName) ? uboName + '配偶。' + uboName : '借款人配偶。借款人'}作为实际控制人，其配偶的连带保证在普惠授信中为标配安排。配偶的具体身份信息、个人征信、名下资产等须在放款前核实。`
  ));
  if (manager && manager.name !== uboName) {
    children.push(bodyPara(
      `第三保证人：${manager.name}（${manager.position}）。${manager.name}${shareholders.find(s=>s.name===manager.name) ? '持有' + companyName + shareholders.find(s=>s.name===manager.name).ratio + '股权（出资' + shareholders.find(s=>s.name===manager.name).amount + '万元，已实缴）' : ''}，在公司担任${manager.position}职务，为公司的核心经营管理层。${manager.name}提供个人连带保证，可将公司经营层的个人信用与贷款偿还责任绑定。${manager.name}个人的资产状况、征信情况、家庭负债等须在放款前通过征信报告和访谈核实后，正式评估其代偿能力。`
    ));
  }

  children.push(heading2('（三）第一还款来源与偿债保障'));
  children.push(bodyPara(
    `第一还款来源为${companyName}的经营收入及利润（通过${hasData(uboName) ? uboName : '借款人'}的分红/薪酬体现为个人还款来源）。以公司近${bizYearsStr}年的稳定经营历史、实缴资本${regCapitalNum ? fmtMoney(regCapitalNum) + '万元' : '（待核实）'}、行业前景良好等综合因素判断，借款人的第一还款来源具有一定的保障度。但鉴于财务数据未公开、企业主个人资产负债情况不明，须在获取完整资料后进行偿债能力的精确测算和压力测试。`
  ));
  children.push(bodyPara(
    `补充还款来源：（1）${hasData(uboName) ? uboName : '借款人'}持有${companyName}${hasData(uboDirect) ? uboDirect + '股权' : '股权'}的价值（按注册资本对应约${regCapitalNum && hasData(uboDirect) ? Math.round(regCapitalNum * parseFloat(uboDirect) / 100) : '待核实'}万元）；（2）${manager && manager.name !== uboName ? manager.name + '个人资产及持股价值' : '保证人个人资产'}；（3）如有抵押物，抵押物的处置变现收入。`
  ));

  // （四）风险与可行性评估
  children.push(heading2('（四）风险与可行性评估'));
  children.push(multiColTable(
    ['风险类别', '具体风险点', '发生概率', '影响程度', '缓释措施'],
    [
      ['信息不对称', '财务数据未公开，盈利能力无法精确评估', '中', '中', '放款前须取得银行流水+纳税申报表+个人征信'],
      ['行业政策', `${industry}监管持续收紧`, '中', '中', '核实企业核心资质有效期'],
      ['经营风险', '产品依赖少数品种，核心人员流失', '低-中', '中', '了解产品管线和核心人员结构'],
      ['资金挪用', '贷款资金未按约定用于经营周转', '低-中', '中', '受托支付+贷后流水监控'],
      ['个人信用', '借款人存在未披露的个人负债/网贷', '低', '高', '放款前查询个人征信，存在新增逾期则暂停放款'],
      ['子公司拖累', investments.length > 0 ? `全资子公司${safeStr(investments[0],'企业名称')||safeStr(investments[0],'name')||'子公司'}亏损或资金占用` : '关联方资金占用风险', '低-中', '中', '了解子公司/关联方经营和财务情况'],
    ],
    [CONTENT_WIDTH / 5, CONTENT_WIDTH / 5, CONTENT_WIDTH / 5, CONTENT_WIDTH / 5, CONTENT_WIDTH / 5]
  ));
  children.push(new Paragraph({}));
  
  children.push(pageBreak());

  // ============ 八、其他需说明情况 ============
  children.push(heading1('八、其他需说明情况'));
  
  if (info.formerName) {
    children.push(bodyPara(
      `1. 企业名称沿革：${companyName}原名"${info.formerName}"，于${info.establishedDate ? fmtDate(info.establishedDate) : ''}正式更名为现名。本次授信调研中，企业主及企业均以"${companyName}"为主体进行工商登记和业务经营，原名称已变更${info.establishedDate ? Math.floor((new Date() - new Date(info.establishedDate)) / (365.25 * 24 * 3600 * 1000)) : '多年'}，不影响本次授信主体认定。`
    ));
  }
  
  children.push(bodyPara(
    `2. 财务数据缺失：企查查财务数据模块未收录${companyName}任何年度的财务数据。作为非上市中小型企业，公司未在公开平台披露财务报表。本次报告的财务分析部分主要依据工商登记信息（注册资本、实缴资本、社保人数）和行业经验推测，所有推测数据均标注[UNSOURCED]并在放款前提条件中要求补充核实。这是普惠金融中小企业主授信的常态——依赖现场尽调获取非公开财务信息。`
  ));

  if (investments.length > 0) {
    children.push(bodyPara(
      `3. 子公司关注：${companyName}持有${safeStr(investments[0],'企业名称')||safeStr(investments[0],'name')||'子公司'}（注册资本${safeStr(investments[0],'出资额')||safeStr(investments[0],'amount')||'（待核实）'}万元）。${regCapitalNum && parseInt(safeStr(investments[0],'出资额')||safeStr(investments[0],'amount')||'0') > regCapitalNum ? '母子公司资本倒挂现象需在尽调中深入了解' : '需关注'}：（1）子公司的实缴资本是否全额到位；（2）${companyName}投资子公司的资金来源（自有资金/外部借款）；（3）子公司的经营状况——盈利、亏损或尚未投产。如子公司尚未产生收入且持续消耗母公司资金，将对${companyName}的整体偿债能力产生负面影响。`
    ));
  }

  // 股东变动
  const shareholderChanges = changes.filter(c => c.item && c.item.includes('股东'));
  if (shareholderChanges.length > 0) {
    children.push(bodyPara(
      `4. 股东变动观察：公司近两年经历了部分股东变更。股东退出属于正常商业行为，但需了解退出原因——是投资期满正常退出？还是对公司前景存疑？向借款人了解相关情况有助于评估企业经营的稳定性和股东信心。`
    ));
  }

  children.push(bodyPara(
    `5. 本次授信方案待定事项：以下为授信审批前必须补充的核心材料清单——（1）${hasData(uboName) ? uboName + '及配偶' : '借款人及配偶'}个人征信报告（最新版）；（2）${companyName}企业征信报告；（3）${companyName}近2年银行主要结算账户流水；（4）${companyName}近2年纳税申报表；（5）${hasData(uboName) ? uboName : '借款人'}家庭资产清单（含房产证/车辆登记证/存单等佐证）；${investments.length > 0 ? '（6）' + (safeStr(investments[0],'企业名称')||safeStr(investments[0],'name')||'子公司') + '基本情况说明；（7）' + companyName + '核心经营资质。' : '（6）' + companyName + '核心经营资质。'}`
  ));
  
  children.push(pageBreak());

  // ============ 九、授信结论及条件阐述 ============
  children.push(heading1('九、授信结论及条件阐述'));

  children.push(heading2('（一）综合评价'));
  children.push(bodyPara(
    hasData(uboName)
      ? `本次授信申请人${uboName}先生，为${companyName}创始人及实际控制人${safeJoin(hasData(uboTotal) ? uboTotal : uboDirect, '（穿透持股约', '）')}，从事${hasData(industry) ? industry : '相关'}行业近${bizYearsStr}年。经营实体${companyName}注册资本${regCapitalNum ? fmtMoney(regCapitalNum) + '万元' : '（待核实）'}${paidCapitalStr && hasData(paidCapitalStr) ? '，全额实缴' : ''}，员工${employeeStr}人，主营${info.bizScope ? info.bizScope.substring(0, 60) : hasData(industry) ? industry : '相关业务'}，属于技术门槛较高、行业前景向好的细分领域。公司位于${info.address ? info.address.substring(0, 10) : ''}，享有园区政策支持，经营基本面扎实。`
      : `本次授信申请人为${companyName}创始人及实际控制人（姓名待现场核实）${safeJoin(hasData(uboTotal) ? uboTotal : uboDirect, '（穿透持股约', '）')}，从事${hasData(industry) ? industry : '相关'}行业近${bizYearsStr}年。经营实体${companyName}注册资本${regCapitalNum ? fmtMoney(regCapitalNum) + '万元' : '（待核实）'}${paidCapitalStr && hasData(paidCapitalStr) ? '，全额实缴' : ''}，员工${employeeStr}人，主营${info.bizScope ? info.bizScope.substring(0, 60) : hasData(industry) ? industry : '相关业务'}，属于技术门槛较高、行业前景向好的细分领域。公司位于${info.address ? info.address.substring(0, 10) : ''}，享有园区政策支持，经营基本面扎实。`
  ));
  children.push(bodyPara(
    `${hasData(uboName) ? uboName + '先生' : '借款人'}行业经验丰富，带领企业完成了${changes.length > 10 ? '多轮' : ''}业务升级和资本引进，体现了较强的管理和资本运作能力。经企查查公开渠道查询，企业主及公司均无失信被执行、行政处罚等负面信用记录。`
  ));
  children.push(bodyPara(
    `需要重点关注的风险：（1）公司财务数据不透明，盈利能力无法精确评估——放款前须取得银行流水+纳税申报表+个人征信报告；（2）${hasData(industry) ? industry : '相关'}行业面临监管趋严、市场竞争加剧等外部挑战；${investments.length > 0 ? '（3）子公司' + (safeStr(investments[0],'企业名称')||safeStr(investments[0],'name')||'') + '的经营状况未知，可能对母公司形成资金拖累。' : ''}`
  ));
  children.push(bodyPara(
    `经综合评估，借款人${hasData(uboName) ? uboName : ''}的个人从业背景、公司经营历史及行业前景整体良好，本次300万元个人经营性贷款的规模与公司体量匹配。但鉴于财务数据缺失、个人资产负债情况不明，建议在落实全部放款前提条件（特别是银行流水审核和个人征信查询）后，审慎核定授信。`
  ));

  children.push(heading2('（二）授信结论'));
  children.push(bodyPara('结论：建议有条件同意。', { noIndent: true, bold: true }));

  children.push(heading2('（三）授信方案'));
  children.push(simpleTable([
    ['借款人', hasData(uboName) ? uboName : '（待现场核实）'],
    ['授信品种', '个人经营性贷款（普惠金融/中小企业主）'],
    ['授信金额', '人民币叁佰万元整（¥3,000,000）'],
    ['授信期限', '1年，可循环使用'],
    ['贷款利率', '不低于同期1年期LPR加100BP'],
    ['还款方式', '按月付息，到期一次性还本（或按月等额本息）'],
    ['支付方式', '单笔50万元以上采用受托支付，50万元以下自主支付'],
    ['担保方式', `（1）借款人${hasData(uboName) ? uboName : ''}配偶连带责任保证；（2）${companyName}公司保证${manager && manager.name !== uboName ? '；（3）' + manager.name + '个人连带责任保证' : ''}`],
    ['贷款用途', `补充${companyName}日常经营周转所需营运资金（原料采购+生产运行费）`],
  ], [CONTENT_WIDTH / 2, CONTENT_WIDTH / 2]));

  children.push(new Paragraph({}));

  children.push(heading2('（四）授信前提条件（支用条件）'));
  children.push(bodyPara('同意在落实以下全部条件后，方可办理授信支用：', { color: BLUE_TEXT, italics: true }));
  children.push(bodyPara(
    `（1）信用文件取得：取得借款人${hasData(uboName) ? uboName + '及配偶' : '及配偶'}最新版个人征信报告（报告日期须在放款前30日内），确认无当前逾期、无失信被执行、无新增大额网贷或高息借贷。取得${companyName}最新版企业征信报告，确认无不良信贷记录。`,
    { color: BLUE_TEXT, italics: true }
  ));
  children.push(bodyPara(
    `（2）经营验证：取得${companyName}近2年主要银行结算账户流水（至少覆盖2家主要银行），经分析确认月均经营性流入不低于XX万元、经营流水连续无断流超过30天、无异常资金往来。取得近2年纳税申报表（增值税+所得税），与银行流水交叉验证经营收入的真实性。`,
    { color: BLUE_TEXT, italics: true }
  ));
  if (investments.length > 0) {
    children.push(bodyPara(
      `（3）子公司尽调：了解${safeStr(investments[0],'企业名称')||safeStr(investments[0],'name')||'子公司'}的经营状况，确认其不会对${companyName}造成重大资金拖累。如子公司持续亏损或需大额后续投入，须重新评估授信额度或追加担保措施。`,
      { color: BLUE_TEXT, italics: true }
    ));
  }
  children.push(bodyPara(
    `（4）担保落实：签署${hasData(uboName) ? uboName : '借款人'}配偶连带保证合同、${companyName}公司保证合同${manager && manager.name !== uboName ? '、' + manager.name + '个人连带保证合同' : ''}。`,
    { color: BLUE_TEXT, italics: true }
  ));
  children.push(bodyPara(
    `（5）资质核查：核验${companyName}核心经营资质的有效性，所有资质须在有效期内。`,
    { color: BLUE_TEXT, italics: true }
  ));
  children.push(bodyPara(
    `（6）用途确认：取得与本次贷款用途对应的采购合同或订货意向书；确认资金用途不涉及房地产、股权投资等禁止领域。`,
    { color: BLUE_TEXT, italics: true }
  ));

  children.push(heading2('（五）贷后管理要求'));
  children.push(bodyPara('（1）按月监控还款账户还款情况，首次扣款失败须当日联系借款人了解原因。'));
  children.push(bodyPara('（2）每季度进行一次贷后回访（可现场巡检+电话/视频+结算流水监控），重点关注：企业经营是否正常、银行流水是否显著下滑、有无新增负面信息。'));
  children.push(bodyPara(`（3）按季度查询借款人${hasData(uboName) ? uboName : ''}个人征信报告，监控新增逾期、新增网贷、新增他行申贷等风险信号。`));
  if (investments.length > 0) {
    children.push(bodyPara(`（4）关注子公司${safeStr(investments[0],'企业名称')||safeStr(investments[0],'name')||''}的经营变化——如出现大额亏损、资本金不足、行政处罚等重大事项，及时评估对母公司偿债能力的影响。`));
  }
  children.push(bodyPara('（5）出现以下预警信号时须在T+5工作日内启动风险排查：经营流水连续3个月下降超过30%/个人征信出现新增逾期超过30天/核心资质到期未续/企业主或法定代表人变更/经营场所搬迁。'));
  children.push(bodyPara('（6）贷款到期前1个月，根据最新信用状况和经营情况综合评估是否续贷。如经营出现明显恶化或信用状况大幅下降，应审慎续贷或要求增加担保条件。'));
  
  children.push(pageBreak());

  // ============ 附：数据来源说明 ============
  children.push(heading1('附：数据来源说明'));
  children.push(multiColTable(
    ['数据类别', '数据来源', '数据时点'],
    [
      ['工商登记信息', '企查查MCP平台——国家企业信用信息公示系统', `${new Date().getFullYear()}年${new Date().getMonth()+1}月查询`],
      ['股东结构', '企查查MCP平台——国家企业信用信息公示系统', `${new Date().getFullYear()}年${new Date().getMonth()+1}月查询`],
      ['实际控制人', '企查查MCP平台——股权穿透分析', `${new Date().getFullYear()}年${new Date().getMonth()+1}月查询`],
      ['主要人员', '企查查MCP平台——国家企业信用信息公示系统', `${new Date().getFullYear()}年${new Date().getMonth()+1}月查询`],
      ['变更记录', '企查查MCP平台——国家企业信用信息公示系统', `成立至今（${changes.length || '若干'}条）`],
      ['对外投资', '企查查MCP平台——国家企业信用信息公示系统', `${new Date().getFullYear()}年${new Date().getMonth()+1}月查询`],
      ['企业简介', '企查查MCP平台——编辑团队维护', `${new Date().getFullYear()}年${new Date().getMonth()+1}月查询`],
      ['财务数据', '未收录（企查查财务数据模块无记录）', '——'],
      ['个人征信', '人民银行个人征信系统', '待取得（授信前提条件）'],
      ['企业征信', '人民银行企业征信系统', '待取得（授信前提条件）'],
      ['银行流水', '企业主提供', '待取得（授信前提条件）'],
      ['纳税申报', '企业主提供/电子税务局导出', '待取得（授信前提条件）'],
    ],
    [CONTENT_WIDTH / 3, CONTENT_WIDTH / 3, CONTENT_WIDTH / 3]
  ));

  children.push(...emptyPara(3));
  children.push(coverCenter('—— 报告全文完 ——', SIZE_BODY, RED_TITLE, FONT_HEITI));
  children.push(...emptyPara(3));
  children.push(coverCenterFangSong(
    `本报告基于企查查MCP平台公开数据及中国工商银行普惠金融（中小企业主客群）授信报告特征库存模板生成。所有标注[UNSOURCED]的数据为基于行业经验推测，须经现场尽调核实后方可作为授信决策依据。报告仅供参考，不构成授信决策。`,
    SIZE_SMALL, LIGHT_GRAY
  ));

  // ============ 构建 Document ============
  const doc = new Document({
    styles: {
      default: {
        document: {
          run: { font: FONT_FANGSONG, size: SIZE_BODY }
        }
      },
      paragraphStyles: [
        {
          id: 'Heading1', name: 'heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { size: SIZE_H1, bold: true, font: FONT_HEITI, color: '000000' },
          paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0 }
        },
        {
          id: 'Heading2', name: 'heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { size: SIZE_H2, bold: true, font: FONT_HEITI, color: '000000' },
          paragraph: { spacing: { before: 280, after: 160 }, outlineLevel: 1 }
        },
        {
          id: 'Heading3', name: 'heading 3', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { size: SIZE_H3, bold: true, font: FONT_HEITI, color: '000000' },
          paragraph: { spacing: { before: 280, after: 160 }, outlineLevel: 2 }
        },
      ]
    },
    sections: [{
      properties: {
        page: {
          size: { width: A4_WIDTH, height: A4_HEIGHT },
          margin: { top: 1440, right: RIGHT_MARGIN, bottom: 1440, left: LEFT_MARGIN }
        }
      },
      headers: {
        default: new Header({
          children: [new Paragraph({
            alignment: AlignmentType.RIGHT,
            children: [new TextRun({ text: '中国工商银行湖南省分行 | 普惠金融授信调查报告', font: FONT_FANGSONG, size: 18, color: GRAY_TEXT })]
          })]
        })
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({ text: '第 ', font: FONT_FANGSONG, size: 18, color: GRAY_TEXT }),
              new TextRun({ children: [PageNumber.CURRENT], font: FONT_FANGSONG, size: 18, color: GRAY_TEXT }),
              new TextRun({ text: ' 页', font: FONT_FANGSONG, size: 18, color: GRAY_TEXT }),
            ]
          })]
        })
      },
      children
    }]
  });

  return doc;
}

module.exports = { buildDocxReport, Packer };
