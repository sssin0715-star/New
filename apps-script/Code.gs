/**
 * 신규입사자 면담 예약 — 서버 (Google Apps Script)
 *
 * 구글 스프레드시트에 붙여 쓰는 스크립트입니다. 예약은 "예약" 시트에,
 * 일정 설정은 "설정" 시트에 저장됩니다. 웹앱을 "모든 사용자"로 배포하면
 * 신규입사자는 로그인 없이 링크만으로 예약할 수 있습니다.
 *
 * 클라이언트에는 예약자 개인정보를 절대 내려보내지 않습니다.
 * 방문자에게는 "어떤 시간이 찼는지"만, 담당자에게는 관리자 코드를
 * 확인한 뒤에만 명단을 보냅니다.
 */

var SHEET_BOOKINGS = '예약';
var SHEET_CONFIG = '설정';

var BOOKING_HEADERS = ['예약ID', '날짜', '요일', '시간', '이름', '이메일', '연락처', '예약일시', '취소코드'];
var DOW = ['일', '월', '화', '수', '목', '금', '토'];

var CONFIG_DEFAULTS = [
  ['active', 'TRUE', '예약 페이지 공개 여부 (TRUE / FALSE)'],
  ['startDate', '2026-09-07', '예약 가능 시작일 (YYYY-MM-DD)'],
  ['endDate', '2026-09-11', '예약 가능 종료일 (YYYY-MM-DD)'],
  ['weekdays', '[1,4,5]', '반복 요일 (0=일 … 6=토)'],
  ['startTime', '09:30', '하루 시작 시각'],
  ['endTime', '17:30', '하루 종료 시각'],
  ['slotMinutes', '60', '면담 간격(분)'],
  ['breakStart', '11:30', '제외 시간대 시작 (없으면 비움)'],
  ['breakEnd', '13:30', '제외 시간대 종료 (없으면 비움)'],
  ['durationLabel', '30~40분', '화면에 표시할 소요 시간'],
  ['locationLabel', '메일로 안내', '화면에 표시할 면담 장소'],
  ['blocked', '[]', '개별로 마감한 시간 목록 (관리 화면에서 자동 관리)'],
  ['adminCode', '', '관리 화면 코드 (비우면 자동 생성)'],
  ['notifyEmail', '', '예약이 들어오면 알림을 받을 담당자 이메일 (비우면 안 보냄)'],
  ['confirmEmail', 'TRUE', '예약자에게 확인 메일 보내기 (TRUE / FALSE)']
];

/* ================= 웹앱 진입점 ================= */

/**
 * 두 가지 역할을 합니다.
 *  - action 파라미터가 없으면: 예약 페이지(Index.html)를 그대로 보여줍니다.
 *  - action 파라미터가 있으면: JSONP API로 응답합니다. GitHub Pages 등
 *    다른 곳에 올린 화면이 이 주소를 호출해 예약을 저장합니다.
 *
 * 브라우저에서 앱스스크립트로 보내는 fetch 는 리디렉션 때문에 CORS 가
 * 잘 깨집니다. 그래서 script 태그로 부르는 JSONP 를 씁니다.
 */
function doGet(e) {
  ensureSheets_();
  var p = (e && e.parameter) || {};

  if (!p.action) {
    return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('신규입사자 면담 예약')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  var result;
  try {
    result = dispatch_(p);
  } catch (err) {
    result = { ok: false, message: '처리 중 오류가 발생했어요.' };
  }

  var json = JSON.stringify(result);
  if (p.callback && /^[A-Za-z_$][A-Za-z0-9_$]*$/.test(p.callback)) {
    return ContentService.createTextOutput(p.callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function parseParamJson_(v, fallback) {
  if (!v) return fallback;
  try { return JSON.parse(v); } catch (e) { return fallback; }
}

function dispatch_(p) {
  switch (String(p.action)) {
    case 'state':
      return getPublicState();

    case 'myBooking':
      return { ok: true, booking: getMyBooking(p.slotId, p.cancelCode) };

    case 'book':
      return book({ slotId: p.slotId, name: p.name, email: p.email, phone: p.phone });

    case 'cancelMine':
      return cancelMyBooking(p.slotId, p.cancelCode);

    case 'adminState':
      return getAdminState(p.code);

    case 'adminSave':
      return adminSaveConfig(p.code, parseParamJson_(p.data, {}));

    case 'adminBlock':
      return adminToggleBlock(p.code, p.slotId, String(p.block) === 'true');

    case 'adminCancel':
      return adminCancelBooking(p.code, p.slotId);

    default:
      return { ok: false, message: '알 수 없는 요청이에요.' };
  }
}

/* ================= 시트 준비 ================= */

function ensureSheets_() {
  var ss = SpreadsheetApp.getActive();

  var cfg = ss.getSheetByName(SHEET_CONFIG);
  if (!cfg) {
    cfg = ss.insertSheet(SHEET_CONFIG);
    cfg.getRange(1, 1, 1, 3).setValues([['키', '값', '설명']]).setFontWeight('bold');
    cfg.getRange(2, 1, CONFIG_DEFAULTS.length, 3).setValues(CONFIG_DEFAULTS);
    cfg.setColumnWidth(1, 130);
    cfg.setColumnWidth(2, 200);
    cfg.setColumnWidth(3, 380);
    cfg.setFrozenRows(1);
  }

  var bk = ss.getSheetByName(SHEET_BOOKINGS);
  if (!bk) {
    bk = ss.insertSheet(SHEET_BOOKINGS);
    bk.getRange(1, 1, 1, BOOKING_HEADERS.length).setValues([BOOKING_HEADERS]).setFontWeight('bold');
    bk.setFrozenRows(1);
    bk.setColumnWidth(1, 140);
    bk.setColumnWidth(6, 200);
    bk.setColumnWidth(8, 160);
    bk.setColumnWidth(9, 260);
  }

  // 관리자 코드가 비어 있으면 한 번만 만들어 둡니다.
  if (!readConfigRaw_()['adminCode']) {
    writeConfigValue_('adminCode', Utilities.getUuid().replace(/-/g, '').slice(0, 8));
  }
  return ss;
}

/** 메뉴에서 한 번 눌러 초기화 + 관리자 코드 확인용. */
function 초기설정() {
  ensureSheets_();
  var code = readConfigRaw_()['adminCode'];
  SpreadsheetApp.getUi().alert(
    '준비가 끝났습니다.\n\n관리자 코드: ' + code +
    '\n\n예약 페이지에서 맨 아래 "담당자"를 누르고 이 코드를 입력하면 관리 화면이 열립니다.'
  );
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('면담 예약').addItem('초기 설정 / 관리자 코드 확인', '초기설정').addToUi();
}

/* ================= 설정 읽고 쓰기 ================= */

function readConfigRaw_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_CONFIG);
  var out = {};
  if (!sh) return out;
  var last = sh.getLastRow();
  if (last < 2) return out;
  var rows = sh.getRange(2, 1, last - 1, 2).getValues();
  for (var i = 0; i < rows.length; i++) {
    var key = String(rows[i][0]).trim();
    if (key) out[key] = rows[i][1];
  }
  return out;
}

function writeConfigValue_(key, value) {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_CONFIG);
  var last = sh.getLastRow();
  var rows = last >= 2 ? sh.getRange(2, 1, last - 1, 1).getValues() : [];
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === key) {
      sh.getRange(i + 2, 2).setValue(value);
      return;
    }
  }
  sh.appendRow([key, value, '']);
}

function asBool_(v) {
  var s = String(v).trim().toUpperCase();
  return s === 'TRUE' || s === 'Y' || s === '1' || s === '예';
}

/** 시트에 날짜/시간이 실제 Date 로 들어가도 문자열로 정규화합니다. */
function asDateStr_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, tz_(), 'yyyy-MM-dd');
  return String(v || '').trim();
}
function asTimeStr_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, tz_(), 'HH:mm');
  var s = String(v || '').trim();
  if (!s) return '';
  var p = s.split(':');
  return pad2_(Number(p[0])) + ':' + pad2_(Number(p[1] || 0));
}
function tz_() { return SpreadsheetApp.getActive().getSpreadsheetTimeZone() || 'Asia/Seoul'; }
function pad2_(n) { return (n < 10 ? '0' : '') + n; }

function parseJsonArray_(v, fallback) {
  try {
    var parsed = JSON.parse(String(v));
    return Object.prototype.toString.call(parsed) === '[object Array]' ? parsed : fallback;
  } catch (e) { return fallback; }
}

/** 화면에서 쓰는 형태의 설정. 비밀 항목은 빼고 돌려줍니다. */
function getConfig_() {
  var raw = readConfigRaw_();
  return {
    active: asBool_(raw.active),
    startDate: asDateStr_(raw.startDate),
    endDate: asDateStr_(raw.endDate),
    weekdays: parseJsonArray_(raw.weekdays, [1, 2, 3, 4, 5]).map(Number),
    startTime: asTimeStr_(raw.startTime) || '10:00',
    endTime: asTimeStr_(raw.endTime) || '17:00',
    slotMinutes: Number(raw.slotMinutes) || 45,
    breakStart: asTimeStr_(raw.breakStart),
    breakEnd: asTimeStr_(raw.breakEnd),
    durationLabel: String(raw.durationLabel || '30~40분'),
    locationLabel: String(raw.locationLabel || '메일로 안내'),
    blocked: parseJsonArray_(raw.blocked, []).map(String)
  };
}

/* ================= 시간 슬롯 계산 (화면과 같은 규칙) ================= */

function todayStr_() {
  return Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd');
}
function nowToken_() {
  return Utilities.formatDate(new Date(), tz_(), "yyyy-MM-dd'T'HH:mm");
}
function timeToMin_(t) {
  var p = String(t).split(':');
  return Number(p[0]) * 60 + Number(p[1]);
}
function minToTime_(m) { return pad2_(Math.floor(m / 60)) + ':' + pad2_(m % 60); }

function parseDate_(dateStr) {
  var p = String(dateStr).split('-');
  return new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
}
function fmtDate_(d) {
  return d.getFullYear() + '-' + pad2_(d.getMonth() + 1) + '-' + pad2_(d.getDate());
}

/** 설정이 만들어내는 모든 슬롯 id 집합. */
function allSlotIds_(config) {
  var set = {};
  if (!config.active || !config.startDate || !config.endDate) return set;

  var times = [];
  var hasBreak = !!(config.breakStart && config.breakEnd);
  var bs = hasBreak ? timeToMin_(config.breakStart) : 0;
  var be = hasBreak ? timeToMin_(config.breakEnd) : 0;
  var startMin = timeToMin_(config.startTime);
  var endMin = timeToMin_(config.endTime);
  for (var m = startMin; m + config.slotMinutes <= endMin; m += config.slotMinutes) {
    if (hasBreak && m < be && m + config.slotMinutes > bs) continue;
    times.push(minToTime_(m));
  }

  var cur = parseDate_(config.startDate);
  var end = parseDate_(config.endDate);
  var guard = 0;
  while (cur <= end && guard < 400) {
    if (config.weekdays.indexOf(cur.getDay()) !== -1) {
      var ds = fmtDate_(cur);
      for (var i = 0; i < times.length; i++) set[ds + 'T' + times[i]] = true;
    }
    cur.setDate(cur.getDate() + 1);
    guard++;
  }
  return set;
}

/* ================= 예약 시트 읽기 ================= */

function readBookings_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_BOOKINGS);
  var out = [];
  if (!sh) return out;
  var last = sh.getLastRow();
  if (last < 2) return out;
  var rows = sh.getRange(2, 1, last - 1, BOOKING_HEADERS.length).getValues();
  for (var i = 0; i < rows.length; i++) {
    var id = String(rows[i][0]).trim();
    if (!id) continue;
    out.push({
      row: i + 2,
      id: id,
      name: String(rows[i][4] || ''),
      email: String(rows[i][5] || ''),
      phone: String(rows[i][6] || ''),
      bookedAt: rows[i][7] instanceof Date
        ? Utilities.formatDate(rows[i][7], tz_(), 'yyyy-MM-dd HH:mm')
        : String(rows[i][7] || ''),
      cancelCode: String(rows[i][8] || '')
    });
  }
  return out;
}

/* ================= 방문자용 API ================= */

/** 개인정보 없이, 설정과 "찬 시간 목록"만 돌려줍니다. */
function getPublicState() {
  ensureSheets_();
  var config = getConfig_();
  var taken = readBookings_().map(function (b) { return b.id; });
  return { config: config, taken: taken, serverNow: nowToken_() };
}

/** 내 예약 확인 — 취소코드가 맞을 때만 내용을 돌려줍니다. */
function getMyBooking(slotId, cancelCode) {
  if (!slotId || !cancelCode) return null;
  var list = readBookings_();
  for (var i = 0; i < list.length; i++) {
    if (list[i].id === String(slotId) && list[i].cancelCode === String(cancelCode)) {
      return { id: list[i].id, name: list[i].name, email: list[i].email, bookedAt: list[i].bookedAt };
    }
  }
  return null;
}

function validEmail_(s) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(s)); }

/**
 * 예약하기. 같은 시간에 두 명이 동시에 눌러도 한 명만 확정되도록
 * 스크립트 락을 잡고 검사와 기록을 한 묶음으로 처리합니다.
 */
function book(payload) {
  payload = payload || {};
  var slotId = String(payload.slotId || '').trim();
  var name = String(payload.name || '').trim();
  var email = String(payload.email || '').trim();
  var phone = String(payload.phone || '').trim();

  if (!name) return { ok: false, reason: 'invalid', message: '이름을 입력해주세요.' };
  if (name.length > 20) return { ok: false, reason: 'invalid', message: '이름은 20자 이내로 입력해주세요.' };
  if (!validEmail_(email)) return { ok: false, reason: 'invalid', message: '이메일 주소를 확인해주세요.' };
  if (phone.replace(/[^0-9]/g, '').length < 9) return { ok: false, reason: 'invalid', message: '연락처를 확인해주세요.' };

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
  } catch (e) {
    return { ok: false, reason: 'busy', message: '지금 예약이 몰리고 있어요. 잠시 후 다시 시도해주세요.' };
  }

  try {
    var config = getConfig_();
    if (!config.active) return { ok: false, reason: 'closed', message: '지금은 예약을 받지 않고 있어요.' };

    var slots = allSlotIds_(config);
    if (!slots[slotId]) return { ok: false, reason: 'invalid', message: '선택할 수 없는 시간이에요. 화면을 새로고침해주세요.' };
    if (config.blocked.indexOf(slotId) !== -1) return { ok: false, reason: 'blocked', message: '담당자가 마감한 시간이에요.' };
    if (slotId < nowToken_()) return { ok: false, reason: 'past', message: '이미 지난 시간이에요.' };

    var list = readBookings_();
    for (var i = 0; i < list.length; i++) {
      if (list[i].id === slotId) {
        return { ok: false, reason: 'taken', message: '방금 다른 분이 예약했어요. 다른 시간을 선택해주세요.' };
      }
    }

    var cancelCode = Utilities.getUuid();
    var parts = slotId.split('T');
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_BOOKINGS);
    sh.appendRow([
      slotId,
      parts[0],
      DOW[parseDate_(parts[0]).getDay()],
      parts[1],
      name,
      email,
      phone,
      Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm'),
      cancelCode
    ]);
    SpreadsheetApp.flush();

    sendMails_(config, { id: slotId, name: name, email: email, phone: phone });
    return { ok: true, cancelCode: cancelCode };
  } finally {
    lock.releaseLock();
  }
}

/** 본인 취소 — 취소코드가 맞을 때만 지웁니다. */
function cancelMyBooking(slotId, cancelCode) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (e) {
    return { ok: false, message: '잠시 후 다시 시도해주세요.' };
  }
  try {
    var list = readBookings_();
    for (var i = 0; i < list.length; i++) {
      if (list[i].id === String(slotId) && list[i].cancelCode === String(cancelCode)) {
        SpreadsheetApp.getActive().getSheetByName(SHEET_BOOKINGS).deleteRow(list[i].row);
        return { ok: true };
      }
    }
    return { ok: false, message: '예약을 찾지 못했어요. 화면을 새로고침해주세요.' };
  } finally {
    lock.releaseLock();
  }
}

/* ================= 메일 ================= */

function fmtWhenKo_(slotId) {
  var parts = String(slotId).split('T');
  var d = parseDate_(parts[0]);
  return (d.getMonth() + 1) + '월 ' + d.getDate() + '일 (' + DOW[d.getDay()] + ') ' + parts[1];
}

function sendMails_(config, booking) {
  var raw = readConfigRaw_();
  var when = fmtWhenKo_(booking.id);

  if (asBool_(raw.confirmEmail)) {
    try {
      MailApp.sendEmail({
        to: booking.email,
        subject: '[면담 예약 확인] ' + when,
        htmlBody:
          '<p>' + booking.name + '님, 신규입사자 면담 예약이 확정되었습니다.</p>' +
          '<p><strong>일시</strong> ' + when + '<br>' +
          '<strong>소요 시간</strong> ' + config.durationLabel + '<br>' +
          '<strong>장소</strong> ' + config.locationLabel + '</p>' +
          '<p>변경이 필요하시면 예약 페이지에서 직접 취소하신 뒤 다시 예약해주세요.</p>'
      });
    } catch (e) { /* 메일 실패가 예약을 막지 않도록 */ }
  }

  var notify = String(raw.notifyEmail || '').trim();
  if (notify && validEmail_(notify)) {
    try {
      MailApp.sendEmail({
        to: notify,
        subject: '[면담 예약] ' + when + ' · ' + booking.name,
        htmlBody:
          '<p>새 면담 예약이 등록되었습니다.</p>' +
          '<p><strong>일시</strong> ' + when + '<br>' +
          '<strong>이름</strong> ' + booking.name + '<br>' +
          '<strong>이메일</strong> ' + booking.email + '<br>' +
          '<strong>연락처</strong> ' + booking.phone + '</p>' +
          '<p><a href="' + SpreadsheetApp.getActive().getUrl() + '">예약 시트 열기</a></p>'
      });
    } catch (e) { /* 무시 */ }
  }
}

/* ================= 관리자용 API ================= */

function checkAdmin_(code) {
  var expected = String(readConfigRaw_()['adminCode'] || '').trim();
  return !!expected && String(code || '').trim() === expected;
}

/** 코드 확인만. 화면 진입 게이트. */
function verifyAdmin(code) {
  return { ok: checkAdmin_(code) };
}

/** 관리 화면 전체 데이터 — 코드가 맞을 때만 개인정보를 포함합니다. */
function getAdminState(code) {
  if (!checkAdmin_(code)) return { ok: false };
  var raw = readConfigRaw_();
  return {
    ok: true,
    config: getConfig_(),
    bookings: readBookings_().map(function (b) {
      return { id: b.id, name: b.name, email: b.email, phone: b.phone, bookedAt: b.bookedAt };
    }),
    sheetUrl: SpreadsheetApp.getActive().getUrl(),
    notifyEmail: String(raw.notifyEmail || ''),
    confirmEmail: asBool_(raw.confirmEmail),
    serverNow: nowToken_()
  };
}

function adminSaveConfig(code, next) {
  if (!checkAdmin_(code)) return { ok: false, message: '코드가 맞지 않아요.' };
  next = next || {};

  if (!next.startDate || !next.endDate || next.startDate > next.endDate) {
    return { ok: false, message: '시작일과 종료일을 확인해주세요.' };
  }
  if (!next.weekdays || !next.weekdays.length) {
    return { ok: false, message: '요일을 하나 이상 선택해주세요.' };
  }
  if (timeToMin_(next.startTime) >= timeToMin_(next.endTime)) {
    return { ok: false, message: '시작 시각이 종료 시각보다 빨라야 해요.' };
  }
  if (next.breakStart && next.breakEnd && timeToMin_(next.breakStart) >= timeToMin_(next.breakEnd)) {
    return { ok: false, message: '제외 시간대를 확인해주세요.' };
  }

  writeConfigValue_('active', 'TRUE');
  writeConfigValue_('startDate', next.startDate);
  writeConfigValue_('endDate', next.endDate);
  writeConfigValue_('weekdays', JSON.stringify(next.weekdays.map(Number)));
  writeConfigValue_('startTime', next.startTime);
  writeConfigValue_('endTime', next.endTime);
  writeConfigValue_('slotMinutes', String(Number(next.slotMinutes) || 45));
  writeConfigValue_('breakStart', next.breakStart || '');
  writeConfigValue_('breakEnd', next.breakEnd || '');
  writeConfigValue_('durationLabel', next.durationLabel || '30~40분');
  writeConfigValue_('locationLabel', next.locationLabel || '메일로 안내');
  if (typeof next.notifyEmail === 'string') writeConfigValue_('notifyEmail', next.notifyEmail.trim());
  if (typeof next.confirmEmail === 'boolean') writeConfigValue_('confirmEmail', next.confirmEmail ? 'TRUE' : 'FALSE');

  return { ok: true, state: getAdminState(code) };
}

function adminToggleBlock(code, slotId, block) {
  if (!checkAdmin_(code)) return { ok: false, message: '코드가 맞지 않아요.' };
  var blocked = getConfig_().blocked;
  var idx = blocked.indexOf(String(slotId));
  if (block && idx === -1) blocked.push(String(slotId));
  if (!block && idx !== -1) blocked.splice(idx, 1);
  writeConfigValue_('blocked', JSON.stringify(blocked));
  return { ok: true, state: getAdminState(code) };
}

function adminCancelBooking(code, slotId) {
  if (!checkAdmin_(code)) return { ok: false, message: '코드가 맞지 않아요.' };
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (e) {
    return { ok: false, message: '잠시 후 다시 시도해주세요.' };
  }
  try {
    var list = readBookings_();
    for (var i = 0; i < list.length; i++) {
      if (list[i].id === String(slotId)) {
        SpreadsheetApp.getActive().getSheetByName(SHEET_BOOKINGS).deleteRow(list[i].row);
        return { ok: true, state: getAdminState(code) };
      }
    }
    return { ok: false, message: '예약을 찾지 못했어요.' };
  } finally {
    lock.releaseLock();
  }
}
