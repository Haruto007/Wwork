// ============================================================
// 副業案件管理アプリ - GASサーバーサイド
// ============================================================

const PROPS_SS_ID = '1LNeAbzn1QnaiqRNu-dgE51giQ0lfOz6cgOxZnnGVJIM';
const SHEET_PROJECTS = '案件マスタ';
const SHEET_FINANCE = '収支設定';
const SHEET_MASTER = 'マスタ';

// ---- エントリポイント ----

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('副業案件管理')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ---- スプレッドシート取得 ----

function getSpreadsheet() {
  return SpreadsheetApp.openById(PROPS_SS_ID);
}

function getSheet(sheetName) {
  return getSpreadsheet().getSheetByName(sheetName);
}

// ---- セットアップ（GASエディタから1回実行） ----

function setup() {
  const ss = SpreadsheetApp.openById(PROPS_SS_ID);

  _initProjectSheet(ss);
  _initFinanceSheet(ss);
  _initMasterSheet(ss);

  // デフォルトシートを削除
  const defaultSheet = ss.getSheetByName('シート1');
  if (defaultSheet && ss.getSheets().length > 1) ss.deleteSheet(defaultSheet);

  Logger.log('セットアップ完了: ' + ss.getUrl());
  return { spreadsheetId: ss.getId(), url: ss.getUrl() };
}

function _initProjectSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_PROJECTS);
  if (!sheet) sheet = ss.insertSheet(SHEET_PROJECTS);
  if (sheet.getLastRow() === 0) {
    const headers = [
      '案件ID', '案件名', '担当者', 'プラットフォーム', '案件URL', '連絡先URL',
      '種別', '提案日', 'ステータス', '納品日', '検収日', '報酬', '稼働時間',
      '時給', 'タスク', 'メモ', '関連資料URL', '請求書対象', '削除フラグ',
      '作成日時', '更新日時'
    ];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold').setBackground('#3B5BDB').setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1, headers.length, 120);
    sheet.setColumnWidth(2, 200);
    sheet.setColumnWidth(5, 200);
    sheet.setColumnWidth(16, 200);
    sheet.setColumnWidth(17, 200);
  }
}

function _initFinanceSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_FINANCE);
  if (!sheet) sheet = ss.insertSheet(SHEET_FINANCE);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['年月', '目標収入', '支出']);
    sheet.getRange(1, 1, 1, 3)
      .setFontWeight('bold').setBackground('#3B5BDB').setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
  }
}

function _initMasterSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_MASTER);
  if (!sheet) sheet = ss.insertSheet(SHEET_MASTER);
  if (sheet.getLastRow() === 0) {
    const platforms = ['ランサーズ', 'クラウドワークス', 'Upwork', 'SHIFT AI', 'ココナラ', 'クラウドリンクス', 'LINE', 'メール', 'Slack', 'Teams', 'その他'];
    const types = ['提案', 'スカウト', '相談', 'タスク', '追加依頼'];
    const statuses = ['完了', 'ステイ', '連絡なし', '不採用', '辞退'];

    sheet.getRange(1, 1, 1, 3).setValues([['プラットフォーム', '種別', 'ステータス']])
      .setFontWeight('bold').setBackground('#3B5BDB').setFontColor('#FFFFFF');
    platforms.forEach((v, i) => sheet.getRange(i + 2, 1).setValue(v));
    types.forEach((v, i) => sheet.getRange(i + 2, 2).setValue(v));
    statuses.forEach((v, i) => sheet.getRange(i + 2, 3).setValue(v));
    sheet.setFrozenRows(1);
  }
}

// ---- 案件ID採番 ----

function _generateProjectId(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 'PRJ-001';
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues()
    .map(r => r[0]).filter(id => /^PRJ-\d+$/.test(String(id)));
  if (ids.length === 0) return 'PRJ-001';
  const maxNum = Math.max(...ids.map(id => parseInt(String(id).replace('PRJ-', ''))));
  return 'PRJ-' + String(maxNum + 1).padStart(3, '0');
}

// ---- 行オブジェクト変換 ----

function _rowToProject(row) {
  if (!row || !row[0]) return null;

  function fmtDate(v) {
    if (!v) return '';
    try { return Utilities.formatDate(new Date(v), 'Asia/Tokyo', 'yyyy-MM-dd'); } catch (e) { return ''; }
  }
  function fmtDatetime(v) {
    if (!v) return '';
    try { return Utilities.formatDate(new Date(v), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss'); } catch (e) { return String(v); }
  }

  return {
    id:             String(row[0]),
    name:           row[1] || '',
    contact:        row[2] || '',
    platform:       row[3] || '',
    url:            row[4] || '',
    contactUrl:     row[5] || '',
    type:           row[6] || '',
    proposalDate:   fmtDate(row[7]),
    status:         row[8] || '',
    deliveryDate:   fmtDate(row[9]),
    acceptanceDate: fmtDate(row[10]),
    reward:         Number(row[11]) || 0,
    hours:          Number(row[12]) || 0,
    hourly:         Number(row[13]) || 0,
    tasks:          row[14] || '',
    memo:           row[15] || '',
    docs:           row[16] || '',
    isInvoice:      row[17] === true || row[17] === 'TRUE',
    isDeleted:      row[18] === true || row[18] === 'TRUE',
    createdAt:      fmtDatetime(row[19]),
    updatedAt:      fmtDatetime(row[20])
  };
}

// ---- 案件一覧取得 ----

function getProjects(filterJson) {
  try {
    const sheet = getSheet(SHEET_PROJECTS);
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: true, data: [] };

    const data = sheet.getRange(2, 1, lastRow - 1, 21).getValues();
    let projects = data.map(_rowToProject).filter(p => p && !p.isDeleted);

    if (filterJson) {
      const f = JSON.parse(filterJson);
      if (f.tab && f.tab !== 'all') {
        if (f.tab === 'active')  projects = projects.filter(p => p.status === 'ステイ');
        if (f.tab === 'done')    projects = projects.filter(p => p.status === '完了');
        if (f.tab === 'closed')  projects = projects.filter(p => ['不採用', '辞退', '連絡なし'].includes(p.status));
      }
      if (f.platform) projects = projects.filter(p => p.platform === f.platform);
      if (f.type)     projects = projects.filter(p => p.type === f.type);
      if (f.contact)  projects = projects.filter(p => p.contact.includes(f.contact));
      if (f.keyword) {
        const kw = f.keyword.toLowerCase();
        projects = projects.filter(p =>
          p.name.toLowerCase().includes(kw) ||
          p.contact.toLowerCase().includes(kw) ||
          p.memo.toLowerCase().includes(kw) ||
          p.tasks.toLowerCase().includes(kw)
        );
      }
      if (f.sortBy) {
        const dir = f.sortDir === 'asc' ? 1 : -1;
        projects.sort((a, b) => {
          const av = a[f.sortBy] ?? '';
          const bv = b[f.sortBy] ?? '';
          if (typeof av === 'number') return (av - bv) * dir;
          return String(av).localeCompare(String(bv), 'ja') * dir;
        });
      }
    }
    return { success: true, data: projects };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ---- 案件1件取得 ----

function getProject(projectId) {
  try {
    const sheet = getSheet(SHEET_PROJECTS);
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: false, error: '案件が見つかりません' };
    const data = sheet.getRange(2, 1, lastRow - 1, 21).getValues();
    for (const row of data) {
      if (String(row[0]) === String(projectId)) {
        return { success: true, data: _rowToProject(row) };
      }
    }
    return { success: false, error: '案件が見つかりません' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ---- 案件保存（新規・更新） ----

function saveProject(dataJson) {
  try {
    const p = JSON.parse(dataJson);
    const sheet = getSheet(SHEET_PROJECTS);
    const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');

    const reward  = parseFloat(p.reward)  || 0;
    const hours   = parseFloat(p.hours)   || 0;
    const hourly  = hours > 0 ? Math.round(reward / hours) : 0;

    const rowData = [
      p.id || '',
      p.name         || '',
      p.contact      || '',
      p.platform     || '',
      p.url          || '',
      p.contactUrl   || '',
      p.type         || '',
      p.proposalDate || '',
      p.status       || '',
      p.deliveryDate  || '',
      p.acceptanceDate || '',
      reward,
      hours,
      hourly,
      p.tasks  || '',
      p.memo   || '',
      p.docs   || '',
      p.isInvoice === true || p.isInvoice === 'true',
      false,
      p.createdAt || now,
      now
    ];

    if (p.id) {
      const lastRow = sheet.getLastRow();
      const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < ids.length; i++) {
        if (String(ids[i][0]) === String(p.id)) {
          const createdAt = sheet.getRange(i + 2, 20).getValue();
          rowData[19] = createdAt || now;
          sheet.getRange(i + 2, 1, 1, rowData.length).setValues([rowData]);
          return { success: true, data: { id: p.id } };
        }
      }
      return { success: false, error: '更新対象が見つかりません' };
    } else {
      rowData[0] = _generateProjectId(sheet);
      sheet.appendRow(rowData);
      return { success: true, data: { id: rowData[0] } };
    }
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ---- 案件削除（論理削除） ----

function deleteProject(projectId) {
  try {
    const sheet = getSheet(SHEET_PROJECTS);
    const lastRow = sheet.getLastRow();
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(projectId)) {
        sheet.getRange(i + 2, 19).setValue(true);
        sheet.getRange(i + 2, 21).setValue(now);
        return { success: true };
      }
    }
    return { success: false, error: '削除対象が見つかりません' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ---- マスタデータ取得 ----

function getMasterData() {
  try {
    const sheet = getSheet(SHEET_MASTER);
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: true, data: { platforms: [], types: [], statuses: [] } };
    const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    return {
      success: true,
      data: {
        platforms: data.map(r => r[0]).filter(v => v),
        types:     data.map(r => r[1]).filter(v => v),
        statuses:  data.map(r => r[2]).filter(v => v)
      }
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ---- 収支設定取得 ----

function getFinanceSettings() {
  try {
    const sheet = getSheet(SHEET_FINANCE);
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: true, data: [] };
    const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    return {
      success: true,
      data: data.filter(r => r[0]).map(r => ({
        yearMonth:     String(r[0]),
        targetRevenue: Number(r[1]) || 0,
        expense:       Number(r[2]) || 0
      }))
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ---- 収支設定保存 ----

function saveFinanceSetting(dataJson) {
  try {
    const s = JSON.parse(dataJson);
    const sheet = getSheet(SHEET_FINANCE);
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const months = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < months.length; i++) {
        if (String(months[i][0]) === String(s.yearMonth)) {
          sheet.getRange(i + 2, 1, 1, 3).setValues([[s.yearMonth, s.targetRevenue || 0, s.expense || 0]]);
          return { success: true };
        }
      }
    }
    sheet.appendRow([s.yearMonth, s.targetRevenue || 0, s.expense || 0]);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ---- ダッシュボードデータ取得 ----

function getDashboardData(year, month) {
  try {
    const projectsResult = getProjects(null);
    if (!projectsResult.success) throw new Error(projectsResult.error);
    const allProjects = projectsResult.data;

    const financeResult = getFinanceSettings();
    const financeSettings = financeResult.success ? financeResult.data : [];

    const targetYM = String(year) + '-' + String(month).padStart(2, '0');
    const fSetting = financeSettings.find(f => f.yearMonth === targetYM) || { targetRevenue: 0, expense: 0 };

    const monthProjects = allProjects.filter(p => p.acceptanceDate && p.acceptanceDate.startsWith(targetYM));
    const monthRevenue  = monthProjects.reduce((s, p) => s + p.reward, 0);
    const monthHours    = monthProjects.reduce((s, p) => s + p.hours, 0);
    const monthHourly   = monthHours > 0 ? Math.round(monthRevenue / monthHours) : 0;
    const achieveRate   = fSetting.targetRevenue > 0
      ? Math.min(100, Math.round(monthRevenue / fSetting.targetRevenue * 100)) : 0;

    const yearStr       = String(year);
    const yearProjects  = allProjects.filter(p => p.acceptanceDate && p.acceptanceDate.startsWith(yearStr));
    const yearRevenue   = yearProjects.reduce((s, p) => s + p.reward, 0);
    const yearHours     = yearProjects.reduce((s, p) => s + p.hours, 0);
    const yearHourly    = yearHours > 0 ? Math.round(yearRevenue / yearHours) : 0;
    const yearExpense   = financeSettings.filter(f => f.yearMonth.startsWith(yearStr))
      .reduce((s, f) => s + f.expense, 0);

    const totalRevenue  = allProjects.reduce((s, p) => s + p.reward, 0);
    const activeProjects = allProjects.filter(p => p.status === 'ステイ');

    const monthlyChart = [];
    for (let m = 1; m <= 12; m++) {
      const ym  = yearStr + '-' + String(m).padStart(2, '0');
      const mps = allProjects.filter(p => p.acceptanceDate && p.acceptanceDate.startsWith(ym));
      const rev = mps.reduce((s, p) => s + p.reward, 0);
      const fs  = financeSettings.find(f => f.yearMonth === ym) || { targetRevenue: 0 };
      monthlyChart.push({ month: m, revenue: rev, target: fs.targetRevenue });
    }

    const platformMap = {};
    allProjects.filter(p => p.acceptanceDate && p.reward > 0).forEach(p => {
      const k = p.platform || 'その他';
      platformMap[k] = (platformMap[k] || 0) + p.reward;
    });
    const platformChart = Object.entries(platformMap).sort((a, b) => b[1] - a[1])
      .map(([name, revenue]) => ({ name, revenue }));

    const clientMap = {};
    allProjects.filter(p => p.acceptanceDate && p.reward > 0).forEach(p => {
      const k = p.contact || '不明';
      clientMap[k] = (clientMap[k] || 0) + p.reward;
    });
    const clientChart = Object.entries(clientMap).sort((a, b) => b[1] - a[1]).slice(0, 10)
      .map(([name, revenue]) => ({ name, revenue }));

    return {
      success: true,
      data: {
        monthly: {
          yearMonth: targetYM,
          revenue: monthRevenue,
          expense: fSetting.expense,
          profit:  monthRevenue - fSetting.expense,
          hours:   monthHours,
          hourly:  monthHourly,
          targetRevenue: fSetting.targetRevenue,
          achieveRate:   achieveRate
        },
        yearly: {
          year:    year,
          revenue: yearRevenue,
          expense: yearExpense,
          profit:  yearRevenue - yearExpense,
          hours:   yearHours,
          hourly:  yearHourly
        },
        totalRevenue:    totalRevenue,
        activeProjects:  activeProjects.slice(0, 10),
        monthlyChart:    monthlyChart,
        platformChart:   platformChart,
        clientChart:     clientChart
      }
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ---- スプレッドシートURL取得 ----

function getSpreadsheetUrl() {
  try {
    return { success: true, data: getSpreadsheet().getUrl() };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ---- CSVエクスポート（請求書対象） ----

function exportInvoiceCSV() {
  try {
    const result = getProjects(JSON.stringify({ tab: 'done' }));
    if (!result.success) return result;
    const projects = result.data.filter(p => p.isInvoice);
    const headers = ['案件ID', '案件名', '担当者', 'プラットフォーム', '検収日', '報酬', '稼働時間', '時給'];
    const rows = [headers, ...projects.map(p =>
      [p.id, p.name, p.contact, p.platform, p.acceptanceDate, p.reward, p.hours, p.hourly]
    )];
    const csv = rows.map(r => r.map(v => '"' + String(v ?? '').replace(/"/g, '""') + '"').join(',')).join('\n');
    return { success: true, data: csv };
  } catch (e) {
    return { success: false, error: e.message };
  }
}
