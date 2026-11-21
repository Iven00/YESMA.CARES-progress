/** YESMA.CARES CG01 — 合併腳本（build from modules；函式勿重複定義） */
/** 來源目錄 google-sheet-channel；部署請貼本檔為 程式碼.gs 或 clasp push */
/** 對齊 04A／04B：一票一列；拆票＝新列＋原票封存。 */

// ===== FILE: Code.gs =====
/**
 * YESMA.CARES CG01｜Ticket 通道（Apps Script）
 * 對齊 04A／04B：一票一列；拆票→新列＋原票封存
 */

var SHEET_TICKETS = 'Tickets';
var SHEET_EVENTS = 'Events';
var SHEET_META = 'Meta';

var HEADERS = [
  'ticket_id', 'task_id', 'epic_id', 'story_id', 'objective',
  'status', 'depends', 'split_from', 'owner', 'evidence_url',
  'blocked_reason', 'updated_at', 'notes'
];

var LEGAL_STATUS = {
  READY: true,
  RUNNING: true,
  REVIEW: true,
  BLOCKED: true,
  DONE: true,
  '封存': true,
  CANCELLED: true
};

/** 通道自檢：能讀試算表＝權限大致可通 */
function pingChannel() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var out = {
    ok: true,
    spreadsheetId: ss.getId(),
    spreadsheetUrl: ss.getUrl(),
    name: ss.getName(),
    tz: Session.getScriptTimeZone(),
    user: Session.getActiveUser().getEmail() || '(容器綁定／可能空白)',
    at: new Date().toISOString()
  };
  Logger.log(JSON.stringify(out, null, 2));
  ensureMeta_();
  var meta = ss.getSheetByName(SHEET_META);
  meta.getRange('A1:B4').setValues([
    ['key', 'value'],
    ['last_ping_at', out.at],
    ['spreadsheet_id', out.spreadsheetId],
    ['channel', 'SHEET-CARES-CG01-TICKETS']
  ]);
  return out;
}

/** 建立 Tickets／Events／Meta 表頭 */
function bootstrapSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureMeta_();
  var tickets = ensureSheet_(ss, SHEET_TICKETS);
  tickets.clear();
  tickets.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  tickets.setFrozenRows(1);

  var events = ensureSheet_(ss, SHEET_EVENTS);
  events.clear();
  events.getRange(1, 1, 1, 5).setValues([['ts', 'type', 'ticket_id', 'detail', 'refs']]);
  events.setFrozenRows(1);

  ensurePoSheet_();
  ensureScheduleLog_();
  try { applyTicketStatusValidation(); } catch (e) {}

  pingChannel();
  logEvent_('NOTE', '', 'bootstrapSheet 完成', '');
  try {
    SpreadsheetApp.getUi().alert('bootstrapSheet 完成：Tickets／Events／Meta／PO_Actions／Schedule_Log');
  } catch (e) {
    Logger.log('bootstrapSheet 完成（無 UI 環境，略過 alert）');
  }
}

/**
 * 拆票：原票封存，寫入新票列
 * @param {string} originalId
 * @param {Array<Object>} newTickets [{ticket_id,task_id,epic_id,story_id,objective,depends,owner,notes}]
 */
function splitTicket(originalId, newTickets) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_TICKETS);
  if (!sheet) throw new Error('缺少 Tickets 工作表，先跑 bootstrapSheet');

  var data = sheet.getDataRange().getValues();
  var header = data[0];
  var idx = indexMap_(header);
  var rowNum = -1;
  for (var r = 1; r < data.length; r++) {
    if (String(data[r][idx.ticket_id]) === originalId) {
      rowNum = r + 1;
      break;
    }
  }
  if (rowNum < 0) throw new Error('找不到原票 ' + originalId);

  var status = String(data[rowNum - 1][idx.status]);
  if (status === '封存') throw new Error('原票已封存，不可再拆');
  if (status === 'DONE') throw new Error('已 DONE 的票不應用拆票代替；請開新票');

  sheet.getRange(rowNum, idx.status + 1).setValue('封存');
  sheet.getRange(rowNum, idx.updated_at + 1).setValue(new Date().toISOString());
  sheet.getRange(rowNum, idx.notes + 1).setValue(
    appendNote_(data[rowNum - 1][idx.notes], 'split → 封存')
  );

  var start = sheet.getLastRow() + 1;
  var rows = [];
  (newTickets || []).forEach(function (t) {
    rows.push([
      t.ticket_id, t.task_id || '', t.epic_id || '', t.story_id || '', t.objective || '',
      t.status || 'READY', t.depends || '', originalId, t.owner || 'Cursor Executor',
      t.evidence_url || '', '', new Date().toISOString(), t.notes || ''
    ]);
  });
  if (rows.length) {
    sheet.getRange(start, 1, start + rows.length - 1, HEADERS.length).setValues(rows);
  }

  logEvent_('SPLIT', originalId, '封存原票；新票 ' + rows.length + ' 筆', JSON.stringify(newTickets));
}

/** 範例：選單示範拆票（請改 id） */
function demoSplitSample() {
  splitTicket('TK-DEMO', [
    { ticket_id: 'TK-DEMO-a', task_id: 'T-demo-a', epic_id: 'EP-00', objective: '拆票A' },
    { ticket_id: 'TK-DEMO-b', task_id: 'T-demo-b', epic_id: 'EP-00', objective: '拆票B' }
  ]);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('CG01 Tickets')
    .addItem('1. 權限自檢 pingChannel', 'pingChannel')
    .addItem('2. 建表 bootstrapSheet', 'bootstrapSheet')
    .addSeparator()
    .addItem('3. 立即排程一輪（PM+SM+30分催票）', 'halfHourOrchestrate')
    .addItem('4. 安裝每30分鐘觸發', 'installHalfHourTrigger')
    .addItem('4b. 暫停排程（升級凍結）', 'pauseOrchestrateForUpgrade')
    .addSeparator()
    .addItem('5. 儲存異常頁 Web App URL', 'saveWebAppUrl')
    .addSeparator()
    .addItem('6. 單票制重整 applySequentialQueueFix', 'applySequentialQueueFix')
    .addItem('7. 票狀態重灌 resetTicketQueueFromSeed', 'resetTicketQueueFromSeed')
    .addItem('7b. 灌入補充票 appendSupplementTickets', 'appendSupplementTickets')
    .addItem('8. 套用 status 下拉（含 REVIEW）', 'applyTicketStatusValidation')
    .addItem('9. 確保 channel_token', 'ensureChannelToken')
    .addToUi();
}

// ——— helpers ———

function ensureSheet_(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

function ensureMeta_() {
  ensureSheet_(SpreadsheetApp.getActiveSpreadsheet(), SHEET_META);
}

function indexMap_(header) {
  var m = {};
  header.forEach(function (h, i) { m[String(h)] = i; });
  HEADERS.forEach(function (h) {
    if (m[h] === undefined) throw new Error('缺欄位: ' + h);
  });
  return m;
}

function appendNote_(old, extra) {
  var s = String(old || '').trim();
  return s ? s + ' | ' + extra : extra;
}

function logEvent_(type, ticketId, detail, refs) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ensureSheet_(ss, SHEET_EVENTS);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, 5).setValues([['ts', 'type', 'ticket_id', 'detail', 'refs']]);
  }
  sh.appendRow([new Date().toISOString(), type, ticketId, detail, refs || '']);
}

// ===== FILE: Transitions.gs =====
/**
 * 票狀態合法轉換（Agent／選單／pmTick 共用）
 *
 * role: pm | executor | reviewer
 */

var TRANSITIONS_ = {
  'READY>RUNNING': { roles: { pm: true }, event: 'DISPATCH' },
  'RUNNING>REVIEW': { roles: { executor: true }, needEvidence: true, event: 'REVIEW' },
  'REVIEW>DONE': { roles: { reviewer: true }, needEvidence: true, event: 'DONE' },
  'REVIEW>RUNNING': { roles: { reviewer: true }, event: 'STATUS' },
  'RUNNING>BLOCKED': { roles: { executor: true, pm: true }, needBlockedReason: true, event: 'STATUS' },
  'BLOCKED>READY': { roles: { executor: true, pm: true, reviewer: true }, event: 'STATUS' },
  'RUNNING>READY': { roles: { pm: true }, event: 'STATUS' }
};

/**
 * @param {string} ticketId
 * @param {string} toStatus
 * @param {Object=} opts { role, evidence_url, notes, blocked_reason }
 * @returns {{ok:boolean, ticket_id:string, from:string, to:string, error?:string}}
 */
function transitionTicket(ticketId, toStatus, opts) {
  opts = opts || {};
  var role = String(opts.role || '').toLowerCase();
  var to = String(toStatus || '').trim();
  var id = String(ticketId || '').trim();
  if (!id) return failTransition_('', '', 'missing ticket_id');
  if (!LEGAL_STATUS[to]) return failTransition_(id, '', 'illegal to status: ' + to);
  if (!role) return failTransition_(id, '', 'missing role');

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_TICKETS);
  if (!sheet || sheet.getLastRow() < 2) return failTransition_(id, '', 'Tickets empty');

  var data = sheet.getDataRange().getValues();
  var idx = indexMap_(data[0]);
  var rowNum = -1;
  for (var r = 1; r < data.length; r++) {
    if (String(data[r][idx.ticket_id]) === id) {
      rowNum = r + 1;
      break;
    }
  }
  if (rowNum < 0) return failTransition_(id, '', 'ticket not found');

  var from = String(data[rowNum - 1][idx.status] || '');
  var key = from + '>' + to;
  var rule = TRANSITIONS_[key];
  if (!rule) return failTransition_(id, from, 'illegal transition ' + key);
  if (!rule.roles[role]) return failTransition_(id, from, 'role ' + role + ' cannot ' + key);

  // Executor 禁止直寫 DONE（雙重保險）
  if (role === 'executor' && to === 'DONE') {
    return failTransition_(id, from, 'executor cannot set DONE');
  }

  var evidence = String(opts.evidence_url != null ? opts.evidence_url : data[rowNum - 1][idx.evidence_url] || '').trim();
  if (rule.needEvidence && !evidence) {
    return failTransition_(id, from, 'evidence_url required for ' + key);
  }

  var blockedReason = String(opts.blocked_reason != null ? opts.blocked_reason : data[rowNum - 1][idx.blocked_reason] || '').trim();
  if (rule.needBlockedReason && !blockedReason) {
    return failTransition_(id, from, 'blocked_reason required for ' + key);
  }

  var now = new Date().toISOString();
  sheet.getRange(rowNum, idx.status + 1).setValue(to);
  sheet.getRange(rowNum, idx.updated_at + 1).setValue(now);

  if (opts.evidence_url != null) {
    sheet.getRange(rowNum, idx.evidence_url + 1).setValue(String(opts.evidence_url).trim());
  } else if (rule.needEvidence) {
    sheet.getRange(rowNum, idx.evidence_url + 1).setValue(evidence);
  }

  if (to === 'BLOCKED') {
    sheet.getRange(rowNum, idx.blocked_reason + 1).setValue(blockedReason);
  } else if (from === 'BLOCKED' || to === 'READY' || to === 'RUNNING' || to === 'REVIEW' || to === 'DONE') {
    if (to !== 'BLOCKED') sheet.getRange(rowNum, idx.blocked_reason + 1).setValue('');
  }

  if (opts.notes != null && String(opts.notes).trim()) {
    var prev = String(data[rowNum - 1][idx.notes] || '');
    sheet.getRange(rowNum, idx.notes + 1).setValue(appendNote_(prev, String(opts.notes).trim()));
  }

  var evType = rule.event || 'STATUS';
  logEvent_(evType, id, from + '→' + to + ' by ' + role, evidence || '');

  var result = { ok: true, ticket_id: id, from: from, to: to, role: role, at: now };
  if (!opts.skip_pm_tick) {
    var pm = afterDeliveryTransition_(from, to, id);
    if (pm) result.pm = pm;
  }
  return result;
}

/**
 * 真正交付／焦點相關轉換後立刻 pmTick（事件驅動）。
 * READY→RUNNING 不觸發（那是 pmTick 認領結果）。
 */
function afterDeliveryTransition_(from, to, id) {
  var trigger =
    to === 'REVIEW' ||
    to === 'DONE' ||
    to === 'BLOCKED' ||
    (from === 'REVIEW' && to === 'RUNNING') ||
    (from === 'BLOCKED' && to === 'READY');
  if (!trigger) return null;
  if (typeof _inPmTick_ !== 'undefined' && _inPmTick_) return null;
  try {
    return pmTick();
  } catch (e) {
    logEvent_('NOTE', id, 'afterDeliveryTransition_ pmTick failed: ' + e, '');
    return { ok: false, error: String(e) };
  }
}

function failTransition_(id, from, msg) {
  return { ok: false, ticket_id: id, from: from || '', to: '', error: msg };
}

/** 僅套用 status 下拉（含 REVIEW），不清資料 */
function applyTicketStatusValidation() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_TICKETS);
  if (!sheet) throw new Error('缺少 Tickets');
  var idx = indexMap_(sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]);
  var last = Math.max(sheet.getLastRow(), 2);
  var range = sheet.getRange(2, idx.status + 1, last, 1);
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(Object.keys(LEGAL_STATUS), true)
    .setAllowInvalid(false)
    .build();
  range.setDataValidation(rule);
  logEvent_('NOTE', '', 'applyTicketStatusValidation keys=' + Object.keys(LEGAL_STATUS).join(','), '');
  return { ok: true, statuses: Object.keys(LEGAL_STATUS) };
}

// ===== FILE: Seed.gs =====
/** CSV 一行解析（支援雙引號欄位） */
function parseCsvLine_(l) {
  var cols = [];
  var cur = '';
  var q = false;
  for (var i = 0; i < l.length; i++) {
    var c = l.charAt(i);
    if (c === '"') { q = !q; continue; }
    if (c === ',' && !q) { cols.push(cur); cur = ''; continue; }
    cur += c;
  }
  cols.push(cur);
  return cols;
}

/** 從 CSV 文字解析成資料列（略過表頭） */
function parseSeedCsvText_(text) {
  var lines = String(text || '').replace(/^\uFEFF/, '').trim().split(/\r?\n/);
  if (lines.length < 2) throw new Error('seed 空白');
  var rows = [];
  for (var r = 1; r < lines.length; r++) {
    if (!String(lines[r]).trim()) continue;
    var cols = parseCsvLine_(lines[r]);
    while (cols.length < HEADERS.length) cols.push('');
    rows.push(cols.slice(0, HEADERS.length));
  }
  return rows;
}

/** 從公開 seed CSV 灌入 Tickets（覆蓋資料列，保留表頭） */
function seedTicketsFromUrl() {
  var url = 'https://raw.githubusercontent.com/Iven00/YESMA.CARES-progress/main/setup/seed_from_backlog.csv';
  var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (resp.getResponseCode() >= 400) throw new Error('seed fetch HTTP ' + resp.getResponseCode());
  var rows = parseSeedCsvText_(resp.getContentText('UTF-8'));
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_TICKETS) || ensureSheet_(ss, SHEET_TICKETS);
  if (sh.getLastRow() < 1) {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sh.setFrozenRows(1);
  } else {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  }
  if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, HEADERS.length).clearContent();
  }
  if (rows.length) sh.getRange(2, 1, rows.length, HEADERS.length).setValues(rows);
  var data = sh.getDataRange().getValues();
  var idx = indexMap_(data[0]);
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idx.ticket_id]) === 'TK-0002' && String(data[i][idx.status]) === 'BLOCKED') {
      sh.getRange(i + 1, idx.status + 1).setValue('DONE');
      sh.getRange(i + 1, idx.updated_at + 1).setValue(new Date().toISOString());
    }
  }
  try { applyTicketStatusValidation(); } catch (e) {}
  logEvent_('NOTE', '', 'seedTicketsFromUrl rows=' + rows.length, url);
  Logger.log('seeded ' + rows.length);
  return { rows: rows.length };
}

/**
 * EP-08／EP-09 補充票：增量寫入 Tickets（不覆蓋既有列）
 * - 已存在的 ticket_id → skip
 * - 缺列 → append READY
 * 來源：progress setup/seed_from_backlog_supplement.csv；抓不到則用內建列
 */
function appendSupplementTicketsFromUrl() {
  var url = 'https://raw.githubusercontent.com/Iven00/YESMA.CARES-progress/main/setup/seed_from_backlog_supplement.csv';
  var rows;
  var source = url;
  try {
    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (resp.getResponseCode() >= 400) throw new Error('HTTP ' + resp.getResponseCode());
    rows = parseSeedCsvText_(resp.getContentText('UTF-8'));
  } catch (e) {
    rows = supplementSeedRowsInline_();
    source = 'inline:' + e;
  }
  return appendTicketRows_(rows, source);
}

/** 不依賴 URL：直接灌內建 EP-08＋EP-09 補充票 */
function appendSupplementTickets() {
  return appendTicketRows_(supplementSeedRowsInline_(), 'inline');
}

/** 將列增量 append 到 Tickets；已有 ticket_id 則略過 */
function appendTicketRows_(rows, source) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_TICKETS) || ensureSheet_(ss, SHEET_TICKETS);
  if (sh.getLastRow() < 1) {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sh.setFrozenRows(1);
  }
  var data = sh.getDataRange().getValues();
  var idx = indexMap_(data[0]);
  var existing = {};
  for (var i = 1; i < data.length; i++) {
    var id = String(data[i][idx.ticket_id] || '').trim();
    if (id) existing[id] = true;
  }
  var now = new Date().toISOString();
  var toAdd = [];
  var skipped = [];
  (rows || []).forEach(function (cols) {
    var tid = String(cols[0] || '').trim();
    if (!tid) return;
    if (existing[tid]) {
      skipped.push(tid);
      return;
    }
    var row = cols.slice(0, HEADERS.length);
    while (row.length < HEADERS.length) row.push('');
    if (!row[5]) row[5] = 'READY';
    if (!row[8]) row[8] = 'Cursor Executor';
    row[11] = now;
    toAdd.push(row);
    existing[tid] = true;
  });
  if (toAdd.length) {
    var start = sh.getLastRow() + 1;
    sh.getRange(start, 1, start + toAdd.length - 1, HEADERS.length).setValues(toAdd);
  }
  try { applyTicketStatusValidation(); } catch (e2) {}
  var out = {
    added: toAdd.length,
    skipped: skipped.length,
    skipped_ids: skipped,
    added_ids: toAdd.map(function (r) { return r[0]; }),
    source: String(source || '')
  };
  logEvent_('NOTE', '', 'appendSupplementTickets added=' + out.added + ' skipped=' + out.skipped, JSON.stringify(out));
  Logger.log(JSON.stringify(out));
  try {
    SpreadsheetApp.getUi().alert(
      '補充票已灌入\n新增：' + out.added +
      '\n略過（已存在）：' + out.skipped +
      '\n新增 id：' + (out.added_ids.join(', ') || '—')
    );
  } catch (e3) {}
  return out;
}

/** 內建補充列（與 seed_from_backlog_supplement.csv 對齊） */
function supplementSeedRowsInline_() {
  var now = '2026-11-21';
  var note = 'from Ticket_Backlog_補充';
  var owner = 'Cursor Executor';
  return [
    ['TK-0801', 'T-08-01a', 'EP-08', 'ST-08-01', 'UI設計 vs 現況差距盤點（首頁／案件大廳／個人大廳＋檔期）', 'READY', '', '', owner, '', '', now, note],
    ['TK-0802', 'T-08-01b', 'EP-08', 'ST-08-01', '套用首頁（手機＋電腦）視覺／IA；主 CTA＝月嫂登入', 'READY', 'TK-0801', '', owner, '', '', now, note],
    ['TK-0803', 'T-08-01c', 'EP-08', 'ST-08-01', '套用案件大廳（手機＋電腦）視覺／IA；僅可派案語意', 'READY', 'TK-0801,TK-0802', '', owner, '', '', now, note],
    ['TK-0804', 'T-08-01d', 'EP-08', 'ST-08-01', '套用個人大廳＋檔期編輯（UI月嫂大廳＝個人大廳）', 'READY', 'TK-0801,TK-0803', '', owner, '', '', now, note],
    ['TK-0805', 'T-08-02a', 'EP-08', 'ST-08-02', '示意非 IN：履約／定金／線上合約／產媽自助 mock 不得當本期交付', 'READY', 'TK-0801', '', owner, '', '', now, note],
    ['TK-0806', 'T-08-02b', 'EP-08', 'ST-08-02', '負向驗收：網站地圖願景＋flows 覆寫勿誤驗', 'READY', 'TK-0805', '', owner, '', '', now, note],
    ['TK-0901', 'T-09-01a', 'EP-09', 'ST-09-01', 'S1 節點表定稿＋EP-09 目錄／交叉引用就緒', 'READY', '', '', owner, '', '', now, note],
    ['TK-0902', 'T-09-01b', 'EP-09', 'ST-09-01', 'S1 AC／方法矩陣定稿；缺口只補 SRS 補充', 'READY', 'TK-0901', '', owner, '', '', now, note],
    ['TK-0903', 'T-09-01c', 'EP-09', 'ST-09-01', 'S1 STAFF／負向／大廳 Web 劇本 L1 可勾', 'READY', 'TK-0902', '', owner, '', '', now, note],
    ['TK-0904', 'T-09-02a', 'EP-09', 'ST-09-02', 'S2 節點表＋方法矩陣定稿', 'READY', 'TK-0901', '', owner, '', '', now, note],
    ['TK-0905', 'T-09-02b', 'EP-09', 'ST-09-02', 'S2 月嫂 Web 劇本 L1～L2（:8787）可勾', 'READY', 'TK-0904', '', owner, '', '', now, note],
    ['TK-0906', 'T-09-02c', 'EP-09', 'ST-09-02', '種子／mock 說明＋可選 L4 驗收台；OUT 總負向', 'READY', 'TK-0903,TK-0905', '', owner, '', '', now, note]
  ];
}

// ===== FILE: PM.gs =====
/**
 * 程式 PM：單票推進、焦點認領、真阻塞才升級 PO_Actions
 *
 * 規則（自驗 → 獨立 Review → DONE）：
 * 1. 同時最多一張 RUNNING＝執行焦點
 * 2. 有 REVIEW 時焦點＝該票（二次驗證中），不認領下一張
 * 3. 無 RUNNING／REVIEW 時，認領下一張可開工 READY→RUNNING
 * 4. DONE 僅獨立 Review 通過後由 Reviewer 回寫；排程不發明 DONE／REVIEW
 * 5. 等依賴＝READY；BLOCKED 僅真卡點且應有 blocked_reason
 */

var PO_HEADERS = [
  'action_id', 'created_at', 'source_role', 'type', 'priority',
  'title', 'detail', 'related_tickets', 'status', 'po_resolution',
  'resolved_at', 'web_visible'
];

var _inPmTick_ = false;

function pmTick() {
  if (_inPmTick_) {
    return { ok: false, reason: 'reentrant_pmTick' };
  }
  _inPmTick_ = true;
  try {
    return pmTickBody_();
  } finally {
    _inPmTick_ = false;
  }
}

function pmTickBody_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ensurePoSheet_();
  ensureScheduleLog_();
  var tickets = ss.getSheetByName(SHEET_TICKETS);
  if (!tickets || tickets.getLastRow() < 2) {
    logSchedule_('PM', 'SKIP', 'Tickets 無資料');
    return { ok: false, reason: 'no_tickets' };
  }

  var rows = tickets.getDataRange().getValues();
  var idx = indexMap_(rows[0]);
  var ready = [];
  var blocked = [];
  var running = [];
  var inReview = [];
  var done = 0;
  var sealed = 0;

  for (var r = 1; r < rows.length; r++) {
    var st = String(rows[r][idx.status] || '');
    var id = String(rows[r][idx.ticket_id] || '');
    if (!id) continue;
    if (st === 'DONE') done++;
    else if (st === '封存' || st === 'CANCELLED') sealed++;
    else if (st === 'READY') ready.push(rowObj_(rows[r], idx));
    else if (st === 'RUNNING') running.push(rowObj_(rows[r], idx));
    else if (st === 'REVIEW') inReview.push(rowObj_(rows[r], idx));
    else if (st === 'BLOCKED') blocked.push(rowObj_(rows[r], idx));
  }

  var focus = [];
  var phase = 'idle';
  var multiRunning = false;
  var claimed = null;
  if (running.length > 0) {
    running.sort(function (a, b) { return String(a.ticket_id).localeCompare(String(b.ticket_id)); });
    focus = [running[0].ticket_id];
    phase = 'execute';
    multiRunning = running.length > 1;
  } else if (inReview.length > 0) {
    inReview.sort(function (a, b) { return String(a.ticket_id).localeCompare(String(b.ticket_id)); });
    focus = [inReview[0].ticket_id];
    phase = 'review';
  } else {
    var eligible = [];
    ready.forEach(function (t) {
      if (depsSatisfied_(t.depends, rows, idx)) eligible.push(t.ticket_id);
    });
    eligible.sort();
    if (eligible.length) {
      focus = [eligible[0]];
      claimed = claimTicketRunning_(tickets, rows, idx, focus[0]);
      if (claimed) {
        running = [claimed];
        ready = ready.filter(function (t) { return t.ticket_id !== focus[0]; });
        phase = 'execute';
      }
    }
  }

  var realBlocked = [];
  blocked.forEach(function (t) {
    var reason = String(t.blocked_reason || '').trim();
    if (!reason) return;
    realBlocked.push(t);
    upsertPoAction_({
      source_role: 'PM',
      type: looksLikeDecision_(reason) ? 'DECISION' : 'HANDLE',
      priority: 'P1',
      title: '票阻塞：' + t.ticket_id,
      detail: reason + '｜' + (t.objective || ''),
      related_tickets: t.ticket_id
    });
  });

  if (multiRunning) {
    upsertPoAction_({
      source_role: 'PM',
      type: 'IMPEDIMENT',
      priority: 'P0',
      title: '同時多張 RUNNING（違反單票制）',
      detail: running.map(function (t) { return t.ticket_id; }).join(',') + '｜請只留一張 RUNNING',
      related_tickets: running.map(function (t) { return t.ticket_id; }).join(',')
    });
  }

  if (focus.length === 0 && running.length === 0 && inReview.length === 0 && realBlocked.length > 0) {
    upsertPoAction_({
      source_role: 'PM',
      type: 'IMPEDIMENT',
      priority: 'P0',
      title: '無可開工票：全線受阻',
      detail: '真 BLOCKED=' + realBlocked.length + '；請清卡點後再推下一張',
      related_tickets: realBlocked.map(function (b) { return b.ticket_id; }).join(',')
    });
  }

  var summary = {
    role: 'PM',
    policy: 'sequential_one_ticket_auto_claim',
    delivery: 'self_verify_then_independent_review',
    phase: phase,
    ready: ready.length,
    focus: focus,
    running: running.length,
    review: inReview.length,
    claimed: claimed ? claimed.ticket_id : '',
    blocked: blocked.length,
    realBlocked: realBlocked.length,
    done: done,
    sealed: sealed,
    multiRunning: multiRunning
  };
  logSchedule_('PM', 'OK', JSON.stringify(summary));
  setMeta_('last_pm_tick', new Date().toISOString());
  setMeta_('pm_focus', focus.join(','));
  setMeta_('pm_phase', phase);
  Logger.log(JSON.stringify(summary));
  return summary;
}

/**
 * 認領焦點票：READY → RUNNING（寫 updated_at＋Events DISPATCH）
 * @returns {Object|null}
 */
function claimTicketRunning_(ticketsSheet, rows, idx, ticketId) {
  var tr = transitionTicket(ticketId, 'RUNNING', { role: 'pm', skip_pm_tick: true });
  if (!tr || !tr.ok) return null;
  for (var r = 1; r < rows.length; r++) {
    if (String(rows[r][idx.ticket_id]) !== ticketId) continue;
    rows[r][idx.status] = 'RUNNING';
    rows[r][idx.updated_at] = tr.at || new Date().toISOString();
    return rowObj_(rows[r], idx);
  }
  return null;
}

function looksLikeDecision_(text) {
  var s = String(text || '');
  return /範圍|核准|例外|是否|決策|OUT|放行|02B|01B/.test(s);
}

function depsSatisfied_(depends, rows, idx) {
  var d = String(depends || '').trim();
  if (!d) return true;
  var parts = d.split(/[,，]/).map(function (x) { return x.trim(); }).filter(Boolean);
  var statusById = {};
  for (var r = 1; r < rows.length; r++) {
    statusById[String(rows[r][idx.ticket_id])] = String(rows[r][idx.status]);
  }
  for (var i = 0; i < parts.length; i++) {
    var p = parts[i];
    if (/^EP-/.test(p)) continue;
    var st = statusById[p];
    if (!(st === 'DONE' || st === '封存')) return false;
  }
  return true;
}

function rowObj_(row, idx) {
  return {
    ticket_id: String(row[idx.ticket_id] || ''),
    objective: String(row[idx.objective] || ''),
    status: String(row[idx.status] || ''),
    depends: String(row[idx.depends] || ''),
    blocked_reason: String(row[idx.blocked_reason] || ''),
    evidence_url: String(row[idx.evidence_url] || ''),
    notes: String(row[idx.notes] || ''),
    updated_at: String(row[idx.updated_at] || '')
  };
}

/**
 * 一次重整：等依賴 BLOCKED→READY、對齊 TK-0002 DONE、關閉洗版 PO_Actions、再 pmTick
 */
function applySequentialQueueFix() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tickets = ss.getSheetByName(SHEET_TICKETS);
  if (!tickets || tickets.getLastRow() < 2) throw new Error('Tickets 無資料');
  var rows = tickets.getDataRange().getValues();
  var idx = indexMap_(rows[0]);
  var converted = 0;
  for (var r = 1; r < rows.length; r++) {
    var id = String(rows[r][idx.ticket_id] || '');
    var st = String(rows[r][idx.status] || '');
    var reason = String(rows[r][idx.blocked_reason] || '').trim();
    var evidence = String(rows[r][idx.evidence_url] || '').trim();
    if (id === 'TK-0002' && st !== 'DONE' && evidence) {
      tickets.getRange(r + 1, idx.status + 1).setValue('DONE');
      tickets.getRange(r + 1, idx.blocked_reason + 1).setValue('');
      tickets.getRange(r + 1, idx.updated_at + 1).setValue(new Date().toISOString());
      converted++;
      continue;
    }
    if (st === 'BLOCKED' && !reason) {
      tickets.getRange(r + 1, idx.status + 1).setValue('READY');
      tickets.getRange(r + 1, idx.updated_at + 1).setValue(new Date().toISOString());
      converted++;
    }
  }

  ensurePoSheet_();
  var po = ss.getSheetByName('PO_Actions');
  var closed = 0;
  if (po && po.getLastRow() >= 2) {
    var pdata = po.getDataRange().getValues();
    var pmap = {};
    pdata[0].forEach(function (n, i) { pmap[String(n)] = i; });
    for (var p = 1; p < pdata.length; p++) {
      if (String(pdata[p][pmap.status]) !== 'OPEN') continue;
      var title = String(pdata[p][pmap.title] || '');
      if (/票阻塞：|BLOCKED 未寫原因：|無可開工票：全線受阻/.test(title)) {
        po.getRange(p + 1, pmap.status + 1).setValue('DONE');
        po.getRange(p + 1, pmap.po_resolution + 1).setValue('單票制重整：等依賴≠BLOCKED');
        po.getRange(p + 1, pmap.resolved_at + 1).setValue(new Date().toISOString());
        closed++;
      }
    }
  }

  var summary = pmTick();
  summary.converted_to_ready_or_done = converted;
  summary.closed_po_noise = closed;
  try {
    SpreadsheetApp.getUi().alert(
      '單票制重整完成\n轉狀態 ' + converted + ' 列\n關閉洗版異常 ' + closed +
      '\n焦點：' + (summary.focus || []).join(',')
    );
  } catch (e) {}
  Logger.log(JSON.stringify(summary));
  return summary;
}

/** 票狀態恢復重來：重灌 seed → 關閉全部 OPEN PO_Actions → pmTick */
function resetTicketQueueFromSeed() {
  var seeded = seedTicketsFromUrl();
  ensurePoSheet_();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var po = ss.getSheetByName('PO_Actions');
  var closed = 0;
  if (po && po.getLastRow() >= 2) {
    var pdata = po.getDataRange().getValues();
    var pmap = {};
    pdata[0].forEach(function (n, i) { pmap[String(n)] = i; });
    for (var p = 1; p < pdata.length; p++) {
      if (String(pdata[p][pmap.status]) !== 'OPEN') continue;
      po.getRange(p + 1, pmap.status + 1).setValue('DONE');
      po.getRange(p + 1, pmap.po_resolution + 1).setValue('票狀態重灌 resetTicketQueueFromSeed');
      po.getRange(p + 1, pmap.resolved_at + 1).setValue(new Date().toISOString());
      closed++;
    }
  }
  var summary = pmTick();
  summary.seeded = seeded;
  summary.closed_all_open_po = closed;
  try {
    SpreadsheetApp.getUi().alert(
      '票狀態已重灌\nseed：' + JSON.stringify(seeded) +
      '\n關閉 OPEN 異常：' + closed +
      '\n焦點：' + (summary.focus || []).join(',')
    );
  } catch (e) {}
  Logger.log(JSON.stringify(summary));
  return summary;
}

// ===== FILE: SM.gs =====
/**
 * 程式 SM：阻塞衛生、逾時 RUNNING／REVIEW、缺證據
 */

var STALE_HOURS = 24;

function smTick() {
  ensurePoSheet_();
  ensureScheduleLog_();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tickets = ss.getSheetByName(SHEET_TICKETS);
  if (!tickets || tickets.getLastRow() < 2) {
    logSchedule_('SM', 'SKIP', 'Tickets 無資料');
    return { ok: false };
  }

  var rows = tickets.getDataRange().getValues();
  var idx = indexMap_(rows[0]);
  var now = new Date().getTime();
  var stale = [];
  var missingEvidence = [];
  var blockedNoReason = [];
  var staleReview = [];

  for (var r = 1; r < rows.length; r++) {
    var t = rowObj_(rows[r], idx);
    if (!t.ticket_id) continue;
    var updated = Date.parse(t.updated_at);
    var staleAge = !isNaN(updated) && (now - updated) > STALE_HOURS * 3600 * 1000;

    if (t.status === 'RUNNING' && staleAge) {
      stale.push(t.ticket_id);
      upsertPoAction_({
        source_role: 'SM',
        type: 'STALE',
        priority: 'P1',
        title: 'RUNNING 逾時：' + t.ticket_id,
        detail: '超過 ' + STALE_HOURS + 'h 無 updated_at 更新｜' + t.objective,
        related_tickets: t.ticket_id
      });
    }

    if (t.status === 'REVIEW') {
      var evR = String(t.evidence_url || '').trim();
      if (!evR) {
        missingEvidence.push(t.ticket_id);
        upsertPoAction_({
          source_role: 'SM',
          type: 'IMPEDIMENT',
          priority: 'P1',
          title: 'REVIEW 缺 evidence：' + t.ticket_id,
          detail: '自驗後進 REVIEW 必須有 evidence_url；請補件或改回 RUNNING',
          related_tickets: t.ticket_id
        });
      } else if (staleAge) {
        staleReview.push(t.ticket_id);
        upsertPoAction_({
          source_role: 'SM',
          type: 'STALE',
          priority: 'P1',
          title: 'REVIEW 逾時未結：' + t.ticket_id,
          detail: '超過 ' + STALE_HOURS + 'h 仍在 REVIEW；請獨立 Reviewer 通過→DONE 或退回 RUNNING｜' + t.objective,
          related_tickets: t.ticket_id
        });
      }
    }

    if (t.status === 'DONE') {
      var ev = String(t.evidence_url || rows[r][idx.evidence_url] || '').trim();
      if (!ev) {
        missingEvidence.push(t.ticket_id);
        upsertPoAction_({
          source_role: 'SM',
          type: 'IMPEDIMENT',
          priority: 'P2',
          title: 'DONE 缺 evidence：' + t.ticket_id,
          detail: '請補 evidence_url 或改回 RUNNING',
          related_tickets: t.ticket_id
        });
      }
    }

    if (t.status === 'BLOCKED' && !String(t.blocked_reason || '').trim()) {
      blockedNoReason.push(t.ticket_id);
      upsertPoAction_({
        source_role: 'SM',
        type: 'IMPEDIMENT',
        priority: 'P2',
        title: 'BLOCKED 未寫原因：' + t.ticket_id,
        detail: t.objective || '',
        related_tickets: t.ticket_id
      });
    }
  }

  var summary = {
    role: 'SM',
    stale: stale,
    staleReview: staleReview,
    missingEvidence: missingEvidence,
    blockedNoReason: blockedNoReason
  };
  logSchedule_('SM', 'OK', JSON.stringify(summary));
  setMeta_('last_sm_tick', new Date().toISOString());
  Logger.log(JSON.stringify(summary));
  return summary;
}

// ===== FILE: Scheduler.gs =====
/**
 * 排程總控：讀 Tickets 推進＋30 分鐘閒置催票規則
 *
 * halfHourOrchestrate（建議每 30 分）：
 *   1) pmTick — 依 Sheet 排出焦點、升級阻塞
 *   2) smTick — 長逾時／缺證據等衛生
 *   3) idleNudgeTick — 30 分規則（有近期 DONE 則本輪不催）
 */

function halfHourOrchestrate() {
  var pm = pmTick();
  var sm = smTick();
  var idle = idleNudgeTick();
  logSchedule_('ORCH30', 'OK', JSON.stringify({
    focus: pm && pm.focus,
    skippedNudge: idle && idle.skippedNudge,
    reason: idle && idle.reason
  }));
  return { pm: pm, sm: sm, idle: idle };
}

/** 相容舊名 */
function hourlyOrchestrate() {
  return halfHourOrchestrate();
}

/** 安裝每 30 分鐘觸發（執行一次即可） */
function installHalfHourTrigger() {
  removeOrchestrateTriggers_();
  ScriptApp.newTrigger('halfHourOrchestrate')
    .timeBased()
    .everyMinutes(30)
    .create();
  setMeta_('trigger', 'halfHourOrchestrate every 30m');
  try {
    SpreadsheetApp.getUi().alert('已安裝每 30 分鐘觸發：halfHourOrchestrate（含 30 分閒置催票規則）');
  } catch (e) {
    Logger.log('installHalfHourTrigger 完成（無 UI 環境，略過 alert）');
  }
}

/** @deprecated 改用 installHalfHourTrigger */
function installHourlyTrigger() {
  installHalfHourTrigger();
}

function removeOrchestrateTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    var fn = t.getHandlerFunction();
    if (fn === 'hourlyOrchestrate' || fn === 'halfHourOrchestrate') {
      ScriptApp.deleteTrigger(t);
    }
  });
}

/** 升級前凍結：刪除 30 分排程觸發，寫 Meta／Events */
function pauseOrchestrateForUpgrade() {
  removeOrchestrateTriggers_();
  setMeta_('trigger', 'paused_for_event_driven_upgrade');
  logEvent_('NOTE', '', 'event_driven_upgrade_schedule_paused', '');
  try {
    SpreadsheetApp.getUi().alert('已暫停 halfHourOrchestrate 觸發（事件驅動升級凍結）');
  } catch (e) {
    Logger.log('pauseOrchestrateForUpgrade 完成');
  }
  return { ok: true, trigger: 'paused_for_event_driven_upgrade' };
}

function ensureScheduleLog_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Schedule_Log');
  if (!sh) {
    sh = ss.insertSheet('Schedule_Log');
    sh.getRange(1, 1, 1, 4).setValues([['ts', 'role', 'result', 'detail']]);
    sh.setFrozenRows(1);
  } else if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, 4).setValues([['ts', 'role', 'result', 'detail']]);
  }
}

function logSchedule_(role, result, detail) {
  ensureScheduleLog_();
  SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('Schedule_Log')
    .appendRow([new Date().toISOString(), role, result, detail || '']);
}

function ensurePoSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('PO_Actions');
  if (!sh) {
    sh = ss.insertSheet('PO_Actions');
    sh.getRange(1, 1, 1, PO_HEADERS.length).setValues([PO_HEADERS]);
    sh.setFrozenRows(1);
  } else if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, PO_HEADERS.length).setValues([PO_HEADERS]);
  }
}

/**
 * 同 related_tickets+type+OPEN 不重複開單
 * related_tickets 空字串時：同 type+title 去重
 */
function upsertPoAction_(opts) {
  ensurePoSheet_();
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PO_Actions');
  var data = sh.getDataRange().getValues();
  if (data.length < 1) {
    sh.getRange(1, 1, 1, PO_HEADERS.length).setValues([PO_HEADERS]);
    data = [PO_HEADERS];
  }
  var h = data[0];
  var map = {};
  h.forEach(function (name, i) { map[String(name)] = i; });

  for (var r = 1; r < data.length; r++) {
    var st = String(data[r][map.status] || '');
    var rel = String(data[r][map.related_tickets] || '');
    var typ = String(data[r][map.type] || '');
    var title = String(data[r][map.title] || '');
    if (st !== 'OPEN' || typ !== opts.type) continue;
    if (opts.related_tickets) {
      if (rel === opts.related_tickets) return;
    } else if (title === (opts.title || '')) {
      return;
    }
  }

  var id = 'POA-' + Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyyMMdd-HHmmss') + '-' + String(Math.floor(Math.random() * 90) + 10);
  sh.appendRow([
    id,
    new Date().toISOString(),
    opts.source_role || 'PM',
    opts.type || 'HANDLE',
    opts.priority || 'P1',
    opts.title || '',
    opts.detail || '',
    opts.related_tickets || '',
    'OPEN',
    '',
    '',
    'TRUE'
  ]);
  logEvent_('NOTE', opts.related_tickets || '', 'PO_Actions: ' + opts.title, id);
}

function setMeta_(key, value) {
  ensureMeta_();
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_META);
  var data = sh.getDataRange().getValues();
  for (var r = 1; r < data.length; r++) {
    if (String(data[r][0]) === key) {
      sh.getRange(r + 1, 2).setValue(value);
      return;
    }
  }
  sh.appendRow([key, value]);
}

// ===== FILE: IdleNudge.gs =====
/**
 * 30 分鐘節奏：看 Tickets 推進／催票
 *
 * 規則：
 * 1. 排程讀 Google Sheet `Tickets` 推進（配合 pmTick 焦點）
 * 2. 若過去 30 分鐘內已有票改為 DONE（updated_at），本輪【不催票】
 * 3. 否則檢查 RUNNING：若 30 分鐘無更新 → 視為可能閒置，寫 PO_Actions(IDLE_CHECK)
 */

var IDLE_WINDOW_MS = 30 * 60 * 1000;

/**
 * @returns {{skippedNudge:boolean, recentDone:string[], idleRunning:string[], reason:string}}
 */
function idleNudgeTick() {
  ensurePoSheet_();
  ensureScheduleLog_();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tickets = ss.getSheetByName(SHEET_TICKETS);
  if (!tickets || tickets.getLastRow() < 2) {
    logSchedule_('IDLE', 'SKIP', 'Tickets 無資料');
    return { skippedNudge: true, recentDone: [], idleRunning: [], reason: 'no_tickets' };
  }

  var rows = tickets.getDataRange().getValues();
  var idx = indexMap_(rows[0]);
  var now = Date.now();
  var recentDone = [];
  var idleRunning = [];
  var runningAll = [];
  var reviewAll = [];
  var idleReview = [];

  for (var r = 1; r < rows.length; r++) {
    var t = rowObj_(rows[r], idx);
    if (!t.ticket_id) continue;
    var updated = Date.parse(t.updated_at);
    var age = isNaN(updated) ? null : (now - updated);

    if (t.status === 'DONE' && age !== null && age <= IDLE_WINDOW_MS) {
      recentDone.push(t.ticket_id);
    }
    if (t.status === 'RUNNING') {
      runningAll.push(t.ticket_id);
      if (age === null || age > IDLE_WINDOW_MS) idleRunning.push(t.ticket_id);
    }
    if (t.status === 'REVIEW') {
      reviewAll.push(t.ticket_id);
      if (age === null || age > IDLE_WINDOW_MS) idleReview.push(t.ticket_id);
    }
  }

  if (recentDone.length > 0) {
    var skipDetail = {
      rule: 'recent_DONE_within_30m_skip_nudge',
      recentDone: recentDone,
      running: runningAll,
      review: reviewAll
    };
    logSchedule_('IDLE', 'SKIP_NUDGE', JSON.stringify(skipDetail));
    setMeta_('last_idle_nudge', 'SKIP_NUDGE@' + new Date().toISOString());
    setMeta_('last_progress_pct', computeProgressPct_(rows, idx));
    return {
      skippedNudge: true,
      recentDone: recentDone,
      idleRunning: [],
      reason: 'recent_done'
    };
  }

  idleRunning.forEach(function (id) {
    var obj = '';
    for (var r2 = 1; r2 < rows.length; r2++) {
      if (String(rows[r2][idx.ticket_id]) === id) {
        obj = String(rows[r2][idx.objective] || '');
        break;
      }
    }
    upsertPoAction_({
      source_role: 'SM',
      type: 'IDLE_CHECK',
      priority: 'P1',
      title: '30分鐘閒置檢查：' + id,
      detail: 'RUNNING 超過 30 分鐘無更新｜' + obj,
      related_tickets: id
    });
  });

  idleReview.forEach(function (id) {
    upsertPoAction_({
      source_role: 'SM',
      type: 'IDLE_CHECK',
      priority: 'P1',
      title: '30分鐘 Review 未結：' + id,
      detail: 'REVIEW 超過 30 分鐘無更新；請獨立 Reviewer 對照 AC／evidence 通過→DONE 或退回 RUNNING',
      related_tickets: id
    });
  });

  if (runningAll.length > 0 && idleRunning.length === 0) {
    logSchedule_('IDLE', 'OK_ACTIVE', JSON.stringify({ running: runningAll }));
    setMeta_('last_idle_nudge', 'OK_ACTIVE@' + new Date().toISOString());
  } else if (reviewAll.length > 0 && idleReview.length === 0) {
    logSchedule_('IDLE', 'OK_REVIEW', JSON.stringify({ review: reviewAll }));
    setMeta_('last_idle_nudge', 'OK_REVIEW@' + new Date().toISOString());
  } else if (idleRunning.length > 0 || idleReview.length > 0) {
    logSchedule_('IDLE', 'NUDGE', JSON.stringify({ idleRunning: idleRunning, idleReview: idleReview }));
    setMeta_('last_idle_nudge', 'NUDGE@' + new Date().toISOString());
  } else {
    upsertPoAction_({
      source_role: 'PM',
      type: 'HANDLE',
      priority: 'P2',
      title: '30分鐘無完成且無 RUNNING／REVIEW：請確認佇列',
      detail: 'pmTick 應已自動認領可開工 READY；執行完成→REVIEW＋evidence；獨立 Review 通過→DONE',
      related_tickets: ''
    });
    logSchedule_('IDLE', 'NUDGE_NO_RUNNING', 'prompt_next_ready');
    setMeta_('last_idle_nudge', 'NUDGE_NO_RUNNING@' + new Date().toISOString());
  }

  setMeta_('last_progress_pct', computeProgressPct_(rows, idx));
  return {
    skippedNudge: false,
    recentDone: [],
    idleRunning: idleRunning,
    idleReview: idleReview,
    reason: (idleRunning.length || idleReview.length)
      ? 'nudge_idle'
      : (runningAll.length || reviewAll.length ? 'active' : 'nudge_no_running')
  };
}

/** 進度% = DONE / (總票 - 封存 - CANCELLED) */
function computeProgressPct_(rows, idx) {
  var done = 0;
  var denom = 0;
  for (var r = 1; r < rows.length; r++) {
    var st = String(rows[r][idx.status] || '');
    var id = String(rows[r][idx.ticket_id] || '');
    if (!id) continue;
    if (st === '封存' || st === 'CANCELLED') continue;
    denom++;
    if (st === 'DONE') done++;
  }
  if (!denom) return '0';
  return String(Math.round((done / denom) * 1000) / 10);
}

// ===== FILE: WebApp.gs =====
/**
 * 異常頁 Web App＋Agent JSON 橋
 * GET 無參數 → HTML 異常頁
 * GET ?format=json&action=board → 公開進度板 JSON（免 token；供 Pages 刷新）
 * GET ?format=json&action=open_actions → 公開 OPEN PO_Actions（免 token）
 * GET ?format=json&action=focus&token=… → 焦點票
 * POST JSON { action:transition, token, ticket_id, to, role, evidence_url?, notes?, blocked_reason? }
 * POST JSON { action:seed_from_url, token } → 重灌 seed CSV（覆蓋 Tickets 資料列）後 pmTick
 * POST JSON { action:pm_tick, token } → 手動觸發 pmTick
 */

function doGet(e) {
  e = e || {};
  var p = e.parameter || {};
  if (String(p.format || '') === 'json' || String(p.action || '')) {
    return jsonOut_(handleApiGet_(p));
  }
  return HtmlService.createHtmlOutput(buildExceptionsHtml_())
    .setTitle('CG01 異常／需 PO 處理')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  e = e || {};
  var body = {};
  try {
    body = JSON.parse(e.postData && e.postData.contents ? e.postData.contents : '{}');
  } catch (err) {
    return jsonOut_({ ok: false, error: 'invalid JSON body' });
  }
  var p = e.parameter || {};
  if (!body.token && p.token) body.token = p.token;
  return jsonOut_(handleApiPost_(body));
}

function handleApiGet_(p) {
  var action = String(p.action || 'focus');
  if (action === 'board') return apiBoard_();
  if (action === 'open_actions') return apiOpenActions_();
  var auth = assertChannelToken_(p.token);
  if (!auth.ok) return auth;
  if (action === 'focus') return apiFocus_();
  if (action === 'ping') return { ok: true, channel: 'SHEET-CARES-CG01-TICKETS', at: new Date().toISOString() };
  return { ok: false, error: 'unknown action: ' + action };
}

function handleApiPost_(body) {
  var auth = assertChannelToken_(body.token);
  if (!auth.ok) return auth;
  var action = String(body.action || '');
  if (action === 'transition') {
    var tr = transitionTicket(body.ticket_id, body.to, {
      role: body.role,
      evidence_url: body.evidence_url,
      notes: body.notes,
      blocked_reason: body.blocked_reason
    });
    if (tr && tr.ok) {
      var focus = apiFocus_();
      tr.pm_focus = focus.pm_focus;
      tr.pm_phase = focus.pm_phase;
      tr.ticket = focus.ticket;
    }
    return tr;
  }
  if (action === 'ensure_token') {
    return { ok: true, token: ensureChannelToken_() };
  }
  if (action === 'seed_from_url') {
    var seeded = seedTicketsFromUrl();
    var afterSeed = pmTick();
    return { ok: true, seeded: seeded, pm: afterSeed };
  }
  if (action === 'pm_tick') {
    return { ok: true, pm: pmTick() };
  }
  return { ok: false, error: 'unknown action: ' + action };
}

function apiFocus_() {
  var focus = getMeta_('pm_focus');
  var phase = getMeta_('pm_phase');
  var ticket = null;
  if (focus) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(SHEET_TICKETS);
    if (sh && sh.getLastRow() >= 2) {
      var data = sh.getDataRange().getValues();
      var idx = indexMap_(data[0]);
      for (var r = 1; r < data.length; r++) {
        if (String(data[r][idx.ticket_id]) === focus) {
          ticket = rowObj_(data[r], idx);
          break;
        }
      }
    }
  }
  return {
    ok: true,
    pm_focus: focus || '',
    pm_phase: phase || '',
    ticket: ticket,
    at: new Date().toISOString()
  };
}

/** 公開進度板：從 Sheet Tickets 即時彙總（Pages／Sync 用） */
function apiBoard_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_TICKETS);
  var counts = { total: 0, DONE: 0, RUNNING: 0, REVIEW: 0, BLOCKED: 0, READY: 0, CANCELLED: 0, other: 0 };
  var byEpic = {};
  var epicOrder = ['EP-00', 'EP-01', 'EP-02', 'EP-03', 'EP-04', 'EP-05', 'EP-06', 'EP-07', 'EP-08', 'EP-09'];
  epicOrder.forEach(function (k) {
    byEpic[k] = { total: 0, DONE: 0, RUNNING: 0, REVIEW: 0, READY: 0, BLOCKED: 0, running_ids: [], review_ids: [] };
  });

  if (sh && sh.getLastRow() >= 2) {
    var data = sh.getDataRange().getValues();
    var idx = indexMap_(data[0]);
    for (var r = 1; r < data.length; r++) {
      var id = String(data[r][idx.ticket_id] || '');
      if (!id) continue;
      var st = String(data[r][idx.status] || '');
      if (st === '封存' || st === 'CANCELLED') {
        counts.CANCELLED++;
        continue;
      }
      counts.total++;
      if (counts[st] !== undefined) counts[st]++;
      else counts.other++;

      var epic = 'OTHER';
      if (/^TK-00/.test(id)) epic = 'EP-00';
      else if (/^TK-01/.test(id)) epic = 'EP-01';
      else if (/^TK-02/.test(id)) epic = 'EP-02';
      else if (/^TK-03/.test(id)) epic = 'EP-03';
      else if (/^TK-04/.test(id)) epic = 'EP-04';
      else if (/^TK-05/.test(id)) epic = 'EP-05';
      else if (/^TK-06/.test(id)) epic = 'EP-06';
      else if (/^TK-07/.test(id)) epic = 'EP-07';
      else if (/^TK-08/.test(id)) epic = 'EP-08';
      else if (/^TK-09/.test(id)) epic = 'EP-09';
      if (!byEpic[epic]) {
        byEpic[epic] = { total: 0, DONE: 0, RUNNING: 0, REVIEW: 0, READY: 0, BLOCKED: 0, running_ids: [], review_ids: [] };
      }
      byEpic[epic].total++;
      if (byEpic[epic][st] !== undefined) byEpic[epic][st]++;
      if (st === 'RUNNING') byEpic[epic].running_ids.push(id);
      if (st === 'REVIEW') byEpic[epic].review_ids.push(id);
    }
  }

  var epics = epicOrder.map(function (k) {
    var e = byEpic[k];
    var tag = '待開工';
    var health = '灰';
    if (e.total === 0) {
      tag = '—';
      health = '灰';
    } else if (e.DONE === e.total) {
      tag = '完成';
      health = '綠';
    } else if (e.RUNNING > 0 || e.REVIEW > 0 || (e.READY > 0 && e.DONE > 0)) {
      tag = '進行中';
      health = '黃';
    }
    var bits = ['DONE ' + e.DONE + '/' + e.total];
    if (e.running_ids.length) bits.push(e.running_ids.join(',') + ' RUNNING');
    if (e.review_ids.length) bits.push(e.review_ids.join(',') + ' REVIEW');
    if (e.READY && e.DONE < e.total) bits.push('READY ' + e.READY);
    return { epic: k, tag: tag, health: health, text: bits.join('；'), done: e.DONE, total: e.total };
  });

  var focus = apiFocus_();
  var pct = counts.total ? Math.round((counts.DONE / counts.total) * 1000) / 10 : 0;
  return {
    ok: true,
    source: 'sheet',
    total: counts.total,
    done: counts.DONE,
    running: counts.RUNNING,
    review: counts.REVIEW,
    blocked: counts.BLOCKED,
    ready: counts.READY,
    pct: pct,
    pm_focus: focus.pm_focus || '',
    pm_phase: focus.pm_phase || '',
    ticket: focus.ticket || null,
    epics: epics,
    at: new Date().toISOString()
  };
}

/** 公開 OPEN 異常清單（免 token） */
function apiOpenActions_() {
  ensurePoSheet_();
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PO_Actions');
  var open = [];
  if (sh && sh.getLastRow() >= 2) {
    var data = sh.getDataRange().getValues();
    var map = {};
    data[0].forEach(function (n, i) { map[String(n)] = i; });
    for (var r = 1; r < data.length; r++) {
      var st = String(data[r][map.status] || '');
      var vis = String(data[r][map.web_visible] || 'TRUE').toUpperCase();
      if (st !== 'OPEN' || vis === 'FALSE') continue;
      open.push({
        id: data[r][map.action_id],
        role: data[r][map.source_role],
        type: data[r][map.type],
        priority: data[r][map.priority],
        title: data[r][map.title],
        detail: data[r][map.detail],
        tickets: data[r][map.related_tickets],
        created: data[r][map.created_at]
      });
    }
  }
  return { ok: true, open: open, at: new Date().toISOString() };
}

function assertChannelToken_(token) {
  var expected = String(getMeta_('channel_token') || '').trim();
  if (!expected) {
    expected = ensureChannelToken_();
  }
  if (!token || String(token).trim() !== expected) {
    return { ok: false, error: 'unauthorized: bad or missing token' };
  }
  return { ok: true };
}

/** 若無 channel_token 則產生並寫入 Meta；回傳 token */
function ensureChannelToken() {
  var t = ensureChannelToken_();
  try {
    SpreadsheetApp.getUi().alert('channel_token 已就緒（亦寫入 Meta）\n' + t);
  } catch (e) {}
  return t;
}

function ensureChannelToken_() {
  var existing = String(getMeta_('channel_token') || '').trim();
  if (existing) return existing;
  var t = Utilities.getUuid().replace(/-/g, '');
  setMeta_('channel_token', t);
  logEvent_('NOTE', '', 'ensureChannelToken_ created', '');
  return t;
}

function getMeta_(key) {
  ensureMeta_();
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_META);
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === key) return String(data[i][1] || '');
  }
  return '';
}

function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function buildExceptionsHtml_() {
  ensurePoSheet_();
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PO_Actions');
  var data = sh.getDataRange().getValues();
  var h = data[0];
  var map = {};
  h.forEach(function (n, i) { map[String(n)] = i; });

  var rows = [];
  for (var r = 1; r < data.length; r++) {
    var st = String(data[r][map.status] || '');
    var vis = String(data[r][map.web_visible] || 'TRUE').toUpperCase();
    if (st !== 'OPEN') continue;
    if (vis === 'FALSE') continue;
    rows.push({
      id: data[r][map.action_id],
      role: data[r][map.source_role],
      type: data[r][map.type],
      priority: data[r][map.priority],
      title: data[r][map.title],
      detail: data[r][map.detail],
      tickets: data[r][map.related_tickets],
      created: data[r][map.created_at]
    });
  }

  rows.sort(function (a, b) {
    var pa = a.priority === 'P0' ? 0 : a.priority === 'P1' ? 1 : 2;
    var pb = b.priority === 'P0' ? 0 : b.priority === 'P1' ? 1 : 2;
    return pa - pb;
  });

  var focus = '';
  var phase = '';
  try {
    focus = getMeta_('pm_focus');
    phase = getMeta_('pm_phase');
  } catch (err) {}

  var tr = rows.map(function (x) {
    return '<tr>' +
      '<td>' + esc_(x.priority) + '</td>' +
      '<td>' + esc_(x.type) + '</td>' +
      '<td>' + esc_(x.role) + '</td>' +
      '<td><strong>' + esc_(x.title) + '</strong><div class="d">' + esc_(x.detail) + '</div></td>' +
      '<td>' + esc_(x.tickets) + '</td>' +
      '<td>' + esc_(String(x.created)) + '</td>' +
      '</tr>';
  }).join('');

  if (!tr) {
    tr = '<tr><td colspan="6">目前沒有 OPEN 異常。排程下一輪或手動「立即排程一輪」。</td></tr>';
  }

  return '' +
    '<!DOCTYPE html><html lang="zh-Hant"><head><meta charset="utf-8"/>' +
    '<meta name="viewport" content="width=device-width,initial-scale=1"/>' +
    '<meta http-equiv="Cache-Control" content="no-cache"/>' +
    '<title>CG01 異常頁</title>' +
    '<style>' +
    'body{font-family:Segoe UI,Noto Sans TC,sans-serif;background:#0f1419;color:#e7ecf1;margin:0;padding:1.25rem}' +
    'h1{font-size:1.25rem;margin:0 0 .25rem}.sub{color:#8b9aab;font-size:.9rem;margin-bottom:1rem}' +
    'button{background:#1a222c;border:1px solid #2a3542;color:#4a8fd4;padding:.4rem .75rem;border-radius:6px;cursor:pointer}' +
    'table{width:100%;border-collapse:collapse;background:#1a222c;border-radius:8px;overflow:hidden;font-size:.9rem}' +
    'th,td{border-bottom:1px solid #2a3542;padding:.55rem .7rem;text-align:left;vertical-align:top}' +
    'th{color:#8b9aab}.d{color:#8b9aab;font-size:.8rem;margin-top:.25rem}' +
    '.focus{margin:0 0 1rem;padding:.75rem 1rem;background:#1a222c;border:1px solid #2a3542;border-radius:8px}' +
    '</style></head><body>' +
    '<h1>CG01 異常／需你處理</h1>' +
    '<p class="sub">來源：PO_Actions（OPEN）· 手動刷新本頁 · 處理後在試算表填 po_resolution 並改 DONE</p>' +
    '<p><button type="button" onclick="location.reload()">重新整理</button></p>' +
    '<div class="focus"><strong>PM 焦點票</strong>：' + esc_(focus || '（尚無）') +
    '　｜　<strong>階段</strong>：' + esc_(phase || '（尚無）') + '</div>' +
    '<table><thead><tr><th>優先</th><th>類型</th><th>來源</th><th>內容</th><th>相關票</th><th>建立</th></tr></thead>' +
    '<tbody>' + tr + '</tbody></table>' +
    '</body></html>';
}

function esc_(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/** 部署 Web App 後呼叫，把 URL 存 Meta */
function saveWebAppUrl() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.prompt('貼上 Web App 部署 URL');
  if (r.getSelectedButton() == ui.Button.OK) {
    setMeta_('webapp_url', r.getResponseText().trim());
    ensureChannelToken_();
    ui.alert('已寫入 Meta.webapp_url，並確保 channel_token');
  }
}


