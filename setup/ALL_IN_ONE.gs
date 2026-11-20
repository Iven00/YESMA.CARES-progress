/** YESMA.CARES CG01 — 合併腳本（一次性貼上用） */
/** 來源目錄 google-sheet-channel；函式勿重複定義 */

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
    .addSeparator()
    .addItem('5. 儲存異常頁 Web App URL', 'saveWebAppUrl')
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


// ===== FILE: PM.gs =====
/**
 * 程式 PM：排程推動、焦點、升級 PO_Actions
 */

var PO_HEADERS = [
  'action_id', 'created_at', 'source_role', 'type', 'priority',
  'title', 'detail', 'related_tickets', 'status', 'po_resolution',
  'resolved_at', 'web_visible'
];

function pmTick() {
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
    else if (st === 'BLOCKED') blocked.push(rowObj_(rows[r], idx));
  }

  // 焦點：READY 且依賴已滿足（簡化：depends 空或所列皆 DONE／封存）
  var focus = [];
  ready.forEach(function (t) {
    if (depsSatisfied_(t.depends, rows, idx)) focus.push(t.ticket_id);
  });

  // BLOCKED → 需 HANDLE／IMPEDIMENT（去重：同 ticket 已有 OPEN 則不重複開）
  blocked.forEach(function (t) {
    var reason = t.blocked_reason || t.notes || '未填 blocked_reason';
    upsertPoAction_({
      source_role: 'PM',
      type: looksLikeDecision_(reason) ? 'DECISION' : 'HANDLE',
      priority: 'P1',
      title: '票阻塞：' + t.ticket_id,
      detail: reason + '｜' + (t.objective || ''),
      related_tickets: t.ticket_id
    });
  });

  // 無 READY 且有 RUNNING 以外的空轉 → 提醒
  if (focus.length === 0 && ready.length === 0 && blocked.length > 0) {
    upsertPoAction_({
      source_role: 'PM',
      type: 'IMPEDIMENT',
      priority: 'P0',
      title: '無可開工票：全線受阻',
      detail: 'BLOCKED=' + blocked.length + '；請清 PO_Actions／依賴',
      related_tickets: blocked.map(function (b) { return b.ticket_id; }).join(',')
    });
  }

  var summary = {
    role: 'PM',
    ready: ready.length,
    focus: focus.slice(0, 5),
    running: running.length,
    blocked: blocked.length,
    done: done,
    sealed: sealed
  };
  logSchedule_('PM', 'OK', JSON.stringify(summary));
  setMeta_('last_pm_tick', new Date().toISOString());
  setMeta_('pm_focus', focus.slice(0, 5).join(','));
  Logger.log(JSON.stringify(summary));
  return summary;
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
    // 允許寫 EP-00 當粗依賴：若字串以 EP- 開頭，視為非票 id，暫不擋
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
    notes: String(row[idx.notes] || ''),
    updated_at: String(row[idx.updated_at] || '')
  };
}


// ===== FILE: SM.gs =====
/**
 * 程式 SM：阻塞衛生、逾時 RUNNING、缺證據
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

  for (var r = 1; r < rows.length; r++) {
    var t = rowObj_(rows[r], idx);
    if (!t.ticket_id) continue;

    if (t.status === 'RUNNING') {
      var updated = Date.parse(t.updated_at);
      if (!isNaN(updated) && (now - updated) > STALE_HOURS * 3600 * 1000) {
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
    }

    if (t.status === 'DONE') {
      var ev = String(rows[r][idx.evidence_url] || '').trim();
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
  SpreadsheetApp.getUi().alert('已安裝每 30 分鐘觸發：halfHourOrchestrate（含 30 分閒置催票規則）');
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
      // 無 updated_at 或超過 30 分未更新 → 閒置檢查
      if (age === null || age > IDLE_WINDOW_MS) {
        idleRunning.push(t.ticket_id);
      }
    }
  }

  // 規則：30 分鐘內已有完成 → 本輪不催
  if (recentDone.length > 0) {
    var skipDetail = {
      rule: 'recent_DONE_within_30m_skip_nudge',
      recentDone: recentDone,
      running: runningAll
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

  // 無近期完成：對可能閒置的 RUNNING 催票／檢查
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
      detail: 'RUNNING 超過 30 分鐘無 updated_at 更新，請確認是否仍在執行、未閒置｜' + obj,
      related_tickets: id
    });
  });

  // 有 RUNNING 但都在 30 分內有更新 → 不開催票，只記日誌
  if (runningAll.length > 0 && idleRunning.length === 0) {
    logSchedule_('IDLE', 'OK_ACTIVE', JSON.stringify({ running: runningAll }));
    setMeta_('last_idle_nudge', 'OK_ACTIVE@' + new Date().toISOString());
  } else if (idleRunning.length > 0) {
    logSchedule_('IDLE', 'NUDGE', JSON.stringify({ idleRunning: idleRunning }));
    setMeta_('last_idle_nudge', 'NUDGE@' + new Date().toISOString());
  } else {
    // 無 RUNNING、無近期 DONE → 催「推進下一張 READY」（給 PM 焦點，不重複洗版可開一則彙總）
    upsertPoAction_({
      source_role: 'PM',
      type: 'HANDLE',
      priority: 'P2',
      title: '30分鐘無完成且無 RUNNING：請推進下一張',
      detail: '請依 Meta.pm_focus／READY 票開工，並把 status 改 RUNNING／完成改 DONE',
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
    reason: idleRunning.length ? 'nudge_idle' : (runningAll.length ? 'active' : 'nudge_no_running')
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
 * 異常頁 Web App：部署為「執行身分＝自己／誰能存取＝你」
 * 部署後 URL 寫入 Meta.webapp_url；Pages 可轉址或嵌入
 */

function doGet(e) {
  return HtmlService.createHtmlOutput(buildExceptionsHtml_())
    .setTitle('CG01 異常／需 PO 處理')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** 回傳 JSON（給 Pages 或外部刷新）?format=json */
function doGetJson() {
  // 若要用 contentService，改 doGet 分支
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

  // P0 在前
  rows.sort(function (a, b) {
    var pa = a.priority === 'P0' ? 0 : a.priority === 'P1' ? 1 : 2;
    var pb = b.priority === 'P0' ? 0 : b.priority === 'P1' ? 1 : 2;
    return pa - pb;
  });

  var focus = '';
  try {
    var meta = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_META);
    var md = meta.getDataRange().getValues();
    for (var i = 1; i < md.length; i++) {
      if (String(md[i][0]) === 'pm_focus') focus = String(md[i][1] || '');
    }
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
    '<div class="focus"><strong>PM 焦點票</strong>：' + esc_(focus || '（尚無）') + '</div>' +
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
    ui.alert('已寫入 Meta.webapp_url');
  }
}


