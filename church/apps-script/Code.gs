/**
 * 교회 관리 — Google Apps Script 백엔드
 *
 * 구글 스프레드시트를 데이터베이스로 쓰는 가벼운 교회 관리 도구입니다.
 * 이 스크립트는 화면을 그리지 않고 JSON 만 주고받습니다.
 * 화면은 깃허브 페이지에 올린 정적 HTML 한 장이 담당합니다.
 *
 * 조직: 목장 › 순 › 교인
 * 권한: 관리자(전체) / 목장(자기 목장) / 순장(자기 순)
 *
 * 설치할 때 편집기에서 setup() 을 한 번 실행하세요.
 * 자세한 순서는 같은 폴더의 README.md 에 있습니다.
 */

const TZ = 'Asia/Seoul';
const APP_TITLE = '교회 관리';
const TOKEN_TTL_HOURS = 12;      // 로그인 유지 시간
const LONG_ABSENCE_DAYS = 28;    // 이 기간 이상 안 보이면 "살펴볼 분"
const VISIT_GAP_DAYS = 180;      // 이 기간 이상 심방 기록이 없으면 안내

const ROLE_ADMIN = '관리자';
const ROLE_MOKJANG = '목장';
const ROLE_SOON = '순장';
const ROLES = [ROLE_ADMIN, ROLE_MOKJANG, ROLE_SOON];

const MEMBER_STATUS = ['재적', '장기결석', '이명', '별세'];
const ATTEND_STATUS = ['출석', '온라인', '결석'];
const VISIT_METHODS = ['심방', '전화', '문자', '만남', '기타'];

const SHEETS = {
  accounts: {
    name: '계정',
    headers: ['ID', '이름', '권한', '목장', '순', '비밀번호', '솔트', '상태', '최근접속', '메모']
  },
  members: {
    name: '교인',
    headers: ['ID', '이름', '성별', '생년월일', '음력', '휴대전화', '목장', '순', '직분',
              '등록일', '주소', '상태', '메모', '수정일시']
  },
  attendance: {
    name: '출결',
    headers: ['날짜', '교인ID', '이름', '목장', '순', '상태', '메모', '기록일시', '기록자']
  },
  prayers: {
    name: '기도제목',
    headers: ['ID', '교인ID', '이름', '목장', '순', '분류', '내용', '등록일', '상태',
              '응답일', '응답메모', '작성자', '수정일시']
  },
  visits: {
    name: '심방',
    headers: ['ID', '날짜', '교인ID', '이름', '목장', '순', '방식', '담당', '내용', '기록일시']
  }
};


/* ══════════════════════════ 설치 ══════════════════════════ */

/**
 * 최초 1회 실행합니다. 시트를 만들고 관리자 계정을 발급합니다.
 * 실행 뒤 [실행 로그]에 스프레드시트 주소와 초기 비밀번호가 찍힙니다.
 */
function setup() {
  const props = PropertiesService.getScriptProperties();
  let ss = null;

  const savedId = props.getProperty('SPREADSHEET_ID');
  if (savedId) {
    ss = SpreadsheetApp.openById(savedId);
  } else {
    ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) ss = SpreadsheetApp.create('교회 관리 대장');
    props.setProperty('SPREADSHEET_ID', ss.getId());
  }

  ensureSheets_(ss);

  if (!props.getProperty('SECRET')) {
    props.setProperty('SECRET', Utilities.getUuid() + Utilities.getUuid());
  }

  let notice;
  if (readAll_('accounts').length) {
    notice = '계정이 이미 있습니다. 비밀번호를 잊으셨으면 resetPassword("이름", "새비밀번호") 를 실행하세요.';
  } else {
    const pw = String(Math.floor(100000 + Math.random() * 900000));
    createAccount_({ name: '관리자', role: ROLE_ADMIN, mokjang: '', soon: '', password: pw });
    notice = '관리자 계정을 만들었습니다.\n  이름: 관리자\n  비밀번호: ' + pw +
             '\n  ← 꼭 적어두고, 앱에 들어가 바꾸십시오.';
  }

  const msg = [
    '설치가 끝났습니다.',
    '스프레드시트: ' + ss.getUrl(),
    '',
    notice,
    '',
    '이제 [배포 > 새 배포 > 웹 앱] 으로 배포하십시오.',
    '  실행 사용자: 나',
    '  액세스 권한: 모든 사용자',
    '배포 뒤 나오는 웹 앱 URL 을 앱 첫 화면에 붙여넣으면 됩니다.'
  ].join('\n');
  Logger.log(msg);
  return msg;
}

/** 비밀번호를 잊었을 때 편집기에서 실행합니다. resetPassword('홍길동', '새비밀번호') */
function resetPassword(name, password) {
  const acc = readAll_('accounts').filter(function (a) {
    return String(a['이름']).trim() === String(name).trim();
  })[0];
  if (!acc) throw new Error('그런 이름의 계정이 없습니다: ' + name);
  if (String(password || '').length < 4) throw new Error('비밀번호는 4자 이상이어야 합니다.');

  const salt = Utilities.getUuid();
  acc['솔트'] = salt;
  acc['비밀번호'] = hash_(salt + '|' + password);
  acc['상태'] = '사용';
  writeRow_('accounts', acc._row, acc);
  Logger.log(name + ' 님의 비밀번호를 바꿨습니다.');
  return '비밀번호를 바꿨습니다.';
}

/**
 * 비밀번호를 잊었을 때 쓰는 자리입니다.
 * 아래 두 줄에 이름과 새 비밀번호를 적고, 편집기에서 이 함수를 실행하세요.
 * 실행이 끝나면 적어 둔 비밀번호는 지워 두시는 편이 좋습니다.
 */
function 비밀번호_재설정() {
  const 이름 = '';          // 예: '이순장'
  const 새비밀번호 = '';     // 예: 'saemal2026'

  if (!이름 || !새비밀번호) {
    throw new Error('이 함수 안의 이름과 새비밀번호를 먼저 채워 주세요.');
  }
  return resetPassword(이름, 새비밀번호);
}

/** 편집기에서 관리자를 하나 더 만들 때 씁니다. addAdmin('김목사', '비밀번호') */
function addAdmin(name, password) {
  createAccount_({ name: name, role: ROLE_ADMIN, mokjang: '', soon: '', password: password });
  Logger.log(name + ' 관리자 계정을 만들었습니다.');
  return '만들었습니다.';
}

function ensureSheets_(ss) {
  Object.keys(SHEETS).forEach(function (key) {
    const def = SHEETS[key];
    let sh = ss.getSheetByName(def.name);
    if (!sh) sh = ss.insertSheet(def.name);

    const width = def.headers.length;
    const lastCol = sh.getLastColumn();
    const current = lastCol > 0
      ? sh.getRange(1, 1, 1, Math.max(lastCol, width)).getValues()[0]
      : [];
    const mismatch = def.headers.some(function (h, i) { return current[i] !== h; });
    if (mismatch) sh.getRange(1, 1, 1, width).setValues([def.headers]);

    sh.getRange(1, 1, 1, width).setFontWeight('bold').setBackground('#eef2f7');
    sh.setFrozenRows(1);
  });

  // 계정 시트의 비밀번호 칸은 눈에 띄지 않게 접어 둡니다.
  const acc = ss.getSheetByName(SHEETS.accounts.name);
  if (acc) {
    acc.hideColumns(SHEETS.accounts.headers.indexOf('비밀번호') + 1, 2);
  }

  const blank = ss.getSheetByName('시트1') || ss.getSheetByName('Sheet1');
  if (blank && ss.getSheets().length > 1 && blank.getLastRow() === 0) ss.deleteSheet(blank);
}


/* ══════════════════════════ JSON API ══════════════════════════
 *
 * 요청은 POST + text/plain 입니다. text/plain 으로 보내면 브라우저가
 * 사전 확인(preflight) 요청을 건너뛰어 한 번의 왕복으로 끝납니다.
 */

function doPost(e) {
  let params = {};
  try {
    params = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return respond_({ ok: false, error: '요청 형식이 올바르지 않습니다.' }, null);
  }
  return handle_(params, e);
}

function doGet(e) {
  const params = (e && e.parameter) || {};
  if (params.payload) {
    try { Object.assign(params, JSON.parse(params.payload)); } catch (err) {}
  }
  if (!params.action) {
    return respond_({ ok: true, data: { service: APP_TITLE, ready: isReady_() } }, e);
  }
  return handle_(params, e);
}

function handle_(params, e) {
  let result;
  try {
    result = { ok: true, data: route_(params) };
  } catch (err) {
    result = { ok: false, error: String((err && err.message) || err) };
  }
  return respond_(result, e);
}

function respond_(result, e) {
  const json = JSON.stringify(result);
  const callback = (e && e.parameter && e.parameter.callback) || '';
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function isReady_() {
  return !!PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
}

function route_(p) {
  const t = p.token;
  switch (String(p.action || '')) {
    /* 로그인 전에도 쓰는 것 */
    case 'who':            return apiWho();
    case 'login':          return apiLogin(p.name, p.password);

    /* 공통 */
    case 'bootstrap':      return apiBootstrap(t);
    case 'dashboard':      return apiDashboard(t);
    case 'changePassword': return apiChangePassword(t, p.current, p.next);

    case 'members':        return apiListMembers(t);
    case 'saveMember':     return apiSaveMember(t, p.member);
    case 'deleteMember':   return apiDeleteMember(t, p.id);

    case 'attendance':     return apiGetAttendance(t, p.date);
    case 'saveAttendance': return apiSaveAttendance(t, p.date, p.records);

    case 'prayers':        return apiListPrayers(t);
    case 'savePrayer':     return apiSavePrayer(t, p.prayer);
    case 'answerPrayer':   return apiAnswerPrayer(t, p.id, p.note);
    case 'deletePrayer':   return apiDeletePrayer(t, p.id);

    case 'visits':         return apiListVisits(t);
    case 'saveVisit':      return apiSaveVisit(t, p.visit);
    case 'deleteVisit':    return apiDeleteVisit(t, p.id);

    /* 관리자 전용 */
    case 'accounts':       return apiListAccounts(t);
    case 'saveAccount':    return apiSaveAccount(t, p.account);
    case 'deleteAccount':  return apiDeleteAccount(t, p.id);

    default:
      throw new Error('알 수 없는 요청입니다: ' + p.action);
  }
}


/* ══════════════════════════ 인증 ══════════════════════════ */

function hash_(text) {
  return Utilities.base64Encode(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8)
  );
}

function sign_(text) {
  const secret = PropertiesService.getScriptProperties().getProperty('SECRET') || 'no-secret';
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(text, secret));
}

/** 로그인 화면에 띄울 이름 목록. 개인정보는 담지 않습니다. */
function apiWho() {
  const list = readAll_('accounts')
    .filter(function (a) { return String(a['상태'] || '사용') !== '정지'; })
    .map(function (a) {
      return {
        id: String(a['ID']),
        name: String(a['이름']),
        role: String(a['권한'] || ROLE_SOON),
        label: String(a['이름']) + ' · ' +
               (String(a['권한']) === ROLE_ADMIN ? '관리자'
                 : [a['목장'], a['순']].filter(Boolean).join(' ') || String(a['권한']))
      };
    });
  list.sort(function (a, b) {
    const ra = ROLES.indexOf(a.role), rb = ROLES.indexOf(b.role);
    if (ra !== rb) return ra - rb;
    return a.name.localeCompare(b.name, 'ko');
  });
  return { ready: list.length > 0, accounts: list };
}

function apiLogin(nameOrId, password) {
  const cache = CacheService.getScriptCache();
  const key = 'fail_' + hash_(String(nameOrId)).slice(0, 12);
  const fails = Number(cache.get(key) || 0);
  if (fails >= 8) throw new Error('로그인 시도가 너무 많습니다. 10분 뒤에 다시 시도해 주세요.');

  const wanted = String(nameOrId || '').trim();
  const acc = readAll_('accounts').filter(function (a) {
    return String(a['ID']) === wanted || String(a['이름']).trim() === wanted;
  })[0];

  if (!acc || String(acc['상태'] || '사용') === '정지' ||
      hash_(String(acc['솔트']) + '|' + String(password || '')) !== String(acc['비밀번호'])) {
    cache.put(key, String(fails + 1), 600);
    throw new Error('이름이나 비밀번호가 맞지 않습니다.');
  }
  cache.remove(key);

  acc['최근접속'] = nowStamp_();
  writeRow_('accounts', acc._row, acc);

  const payload = Utilities.base64EncodeWebSafe(JSON.stringify({
    id: String(acc['ID']),
    exp: Date.now() + TOKEN_TTL_HOURS * 3600 * 1000
  }));
  return {
    token: payload + '.' + sign_(payload),
    me: accountToClient_(acc)
  };
}

/** 토큰을 확인하고 로그인한 사람의 정보를 돌려줍니다. */
function auth_(token) {
  const parts = String(token || '').split('.');
  if (parts.length !== 2 || sign_(parts[0]) !== parts[1]) throw new Error('로그인이 필요합니다.');

  let payload;
  try {
    payload = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString());
  } catch (err) {
    throw new Error('로그인이 필요합니다.');
  }
  if (!payload.exp || payload.exp < Date.now()) {
    throw new Error('로그인이 만료되었습니다. 다시 로그인해 주세요.');
  }

  const acc = readAll_('accounts').filter(function (a) {
    return String(a['ID']) === String(payload.id);
  })[0];
  if (!acc) throw new Error('로그인이 필요합니다.');
  if (String(acc['상태'] || '사용') === '정지') throw new Error('사용이 중지된 계정입니다.');

  return {
    _row: acc._row,
    id: String(acc['ID']),
    name: String(acc['이름']),
    role: String(acc['권한'] || ROLE_SOON),
    mokjang: String(acc['목장'] || '').trim(),
    soon: String(acc['순'] || '').trim()
  };
}

function requireAdmin_(user) {
  if (user.role !== ROLE_ADMIN) throw new Error('관리자만 할 수 있습니다.');
  return user;
}

function apiChangePassword(token, current, next) {
  const user = auth_(token);
  const acc = readAll_('accounts').filter(function (a) {
    return String(a['ID']) === user.id;
  })[0];
  if (hash_(String(acc['솔트']) + '|' + String(current || '')) !== String(acc['비밀번호'])) {
    throw new Error('지금 쓰는 비밀번호가 맞지 않습니다.');
  }
  if (String(next || '').length < 4) throw new Error('새 비밀번호는 4자 이상이어야 합니다.');

  const salt = Utilities.getUuid();
  acc['솔트'] = salt;
  acc['비밀번호'] = hash_(salt + '|' + next);
  writeRow_('accounts', acc._row, acc);
  return true;
}


/* ══════════════════════════ 권한 범위 ══════════════════════════ */

/**
 * 이 사람이 저 목장·순의 교인을 볼 수 있는지 판단합니다.
 *  관리자 — 전체
 *  목장   — 자기 목장 전체
 *  순장   — 자기 목장의 자기 순
 */
function inScope_(user, mokjang, soon) {
  if (user.role === ROLE_ADMIN) return true;
  if (!user.mokjang) return false;
  if (String(mokjang || '').trim() !== user.mokjang) return false;
  if (user.role === ROLE_SOON) {
    return !!user.soon && String(soon || '').trim() === user.soon;
  }
  return true;
}

function scopeLabel_(user) {
  if (user.role === ROLE_ADMIN) return '전체';
  if (user.role === ROLE_MOKJANG) return user.mokjang + ' 목장';
  return [user.mokjang, user.soon].filter(Boolean).join(' ') + ' 순';
}

function assertScope_(user, mokjang, soon) {
  if (!inScope_(user, mokjang, soon)) {
    throw new Error('맡으신 ' + scopeLabel_(user) + ' 밖의 자료는 다룰 수 없습니다.');
  }
}


/* ══════════════════════════ 시트 도우미 ══════════════════════════ */

function getSS_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  throw new Error('스프레드시트를 찾지 못했습니다. 편집기에서 setup() 을 실행하세요.');
}

function getSheet_(key) {
  const ss = getSS_();
  let sh = ss.getSheetByName(SHEETS[key].name);
  if (!sh) {
    ensureSheets_(ss);
    sh = ss.getSheetByName(SHEETS[key].name);
  }
  return sh;
}

/** 시트 전체를 [{헤더: 값}] 으로 읽습니다. _row 에 실제 행 번호가 붙습니다. */
function readAll_(key) {
  const sh = getSheet_(key);
  const lastRow = sh.getLastRow();
  const headers = SHEETS[key].headers;
  if (lastRow < 2) return [];

  const values = sh.getRange(2, 1, lastRow - 1, headers.length).getValues();
  const out = [];
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    if (row.every(function (c) { return c === '' || c === null; })) continue;
    const obj = { _row: i + 2 };
    headers.forEach(function (h, j) { obj[h] = row[j]; });
    out.push(obj);
  }
  return out;
}

function rowFrom_(key, obj) {
  return SHEETS[key].headers.map(function (h) {
    const v = obj[h];
    return (v === undefined || v === null) ? '' : v;
  });
}

function appendRow_(key, obj) {
  getSheet_(key).appendRow(rowFrom_(key, obj));
}

function writeRow_(key, rowNum, obj) {
  const headers = SHEETS[key].headers;
  getSheet_(key).getRange(rowNum, 1, 1, headers.length).setValues([rowFrom_(key, obj)]);
}

function findBy_(key, field, value) {
  return readAll_(key).filter(function (r) {
    return String(r[field]) === String(value);
  })[0] || null;
}

/**
 * 짧지만 겹치지 않는 ID 를 만듭니다.
 * 뒤쪽 8자리는 UUID 에서 뽑아 쓰고, 만약을 대비해 같은 시트에 이미 있으면 다시 뽑습니다.
 */
function newId_(prefix, sheetKey) {
  const stamp = Utilities.formatDate(new Date(), TZ, 'yyMMdd');
  for (let i = 0; i < 20; i++) {
    const id = prefix + '-' + stamp + '-' + Utilities.getUuid().replace(/-/g, '').slice(0, 8);
    if (!sheetKey || !findBy_(sheetKey, 'ID', id)) return id;
  }
  throw new Error('ID 를 만들지 못했습니다. 다시 시도해 주세요.');
}

function nowStamp_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'); }
function today_()    { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function pad2_(n)    { return ('0' + String(n)).slice(-2); }

/** 셀 값이 날짜든 문자열이든 'yyyy-MM-dd' 로 통일합니다. */
function toDateStr_(v) {
  if (v === '' || v === null || v === undefined) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  }
  const s = String(v).trim();
  let m = s.match(/^(\d{4})[-./](\d{1,2})[-./](\d{1,2})/);
  if (m) return m[1] + '-' + pad2_(m[2]) + '-' + pad2_(m[3]);
  m = s.match(/^(\d{1,2})[-./](\d{1,2})$/);        // 연도 없는 생일
  if (m) return '--' + pad2_(m[1]) + '-' + pad2_(m[2]);
  m = s.match(/^(\d{4})(\d{2})(\d{2})$/);          // 19800101
  if (m) return m[1] + '-' + m[2] + '-' + m[3];
  return s;
}

function monthDay_(dateStr) {
  const m = String(dateStr || '').match(/(\d{2})-(\d{2})$/);
  return m ? m[1] + '-' + m[2] : '';
}

function daysBetween_(fromStr, toStr) {
  if (!fromStr || !toStr) return null;
  const a = new Date(fromStr + 'T00:00:00Z').getTime();
  const b = new Date(toStr + 'T00:00:00Z').getTime();
  if (isNaN(a) || isNaN(b)) return null;
  return Math.round((b - a) / 86400000);
}

function addDays_(dateStr, n) {
  const d = new Date(dateStr + 'T00:00:00Z');
  d.setUTCDate(d.getUTCDate() + n);
  return Utilities.formatDate(d, 'UTC', 'yyyy-MM-dd');
}

function lastSunday_() {
  const t = today_();
  return addDays_(t, -new Date(t + 'T00:00:00Z').getUTCDay());
}

function truthy_(v) {
  const s = String(v).trim().toLowerCase();
  return s === 'true' || s === 'y' || s === 'yes' || s === '예' ||
         s === 'o' || s === '음력' || s === '1';
}

function uniqueSorted_(arr) {
  const seen = {}, out = [];
  arr.forEach(function (v) {
    const s = String(v || '').trim();
    if (!s || seen[s]) return;
    seen[s] = 1; out.push(s);
  });
  return out.sort(function (a, b) { return a.localeCompare(b, 'ko'); });
}


/* ══════════════════════════ 계정 ══════════════════════════ */

function accountToClient_(a) {
  return {
    id: String(a['ID']),
    name: String(a['이름']),
    role: String(a['권한'] || ROLE_SOON),
    mokjang: String(a['목장'] || ''),
    soon: String(a['순'] || ''),
    status: String(a['상태'] || '사용'),
    lastSeen: String(a['최근접속'] || ''),
    memo: String(a['메모'] || '')
  };
}

function createAccount_(spec) {
  const name = String(spec.name || '').trim();
  if (!name) throw new Error('이름을 입력해 주세요.');
  if (findBy_('accounts', '이름', name)) {
    throw new Error('같은 이름의 계정이 이미 있습니다: ' + name);
  }
  const password = String(spec.password || '');
  if (password.length < 4) throw new Error('비밀번호는 4자 이상이어야 합니다.');

  const role = ROLES.indexOf(spec.role) >= 0 ? spec.role : ROLE_SOON;
  if (role !== ROLE_ADMIN && !String(spec.mokjang || '').trim()) {
    throw new Error('목장을 지정해 주세요.');
  }
  if (role === ROLE_SOON && !String(spec.soon || '').trim()) {
    throw new Error('순을 지정해 주세요.');
  }

  const salt = Utilities.getUuid();
  const record = {
    'ID': newId_('A', 'accounts'),
    '이름': name,
    '권한': role,
    '목장': String(spec.mokjang || '').trim(),
    '순': String(spec.soon || '').trim(),
    '비밀번호': hash_(salt + '|' + password),
    '솔트': salt,
    '상태': '사용',
    '최근접속': '',
    '메모': String(spec.memo || '')
  };
  appendRow_('accounts', record);
  return record;
}

function apiListAccounts(token) {
  requireAdmin_(auth_(token));
  return readAll_('accounts').map(accountToClient_).sort(function (a, b) {
    const ra = ROLES.indexOf(a.role), rb = ROLES.indexOf(b.role);
    if (ra !== rb) return ra - rb;
    return (a.mokjang + a.soon + a.name).localeCompare(b.mokjang + b.soon + b.name, 'ko');
  });
}

function apiSaveAccount(token, spec) {
  const user = requireAdmin_(auth_(token));
  if (!spec) throw new Error('내용이 비어 있습니다.');

  if (!spec.id) return accountToClient_(createAccount_(spec));

  const acc = findBy_('accounts', 'ID', spec.id);
  if (!acc) throw new Error('그 계정을 찾지 못했습니다.');

  const role = ROLES.indexOf(spec.role) >= 0 ? spec.role : acc['권한'];
  if (role !== ROLE_ADMIN && !String(spec.mokjang || '').trim()) {
    throw new Error('목장을 지정해 주세요.');
  }
  if (role === ROLE_SOON && !String(spec.soon || '').trim()) {
    throw new Error('순을 지정해 주세요.');
  }
  // 마지막 관리자가 사라지지 않도록 막습니다.
  if (acc['권한'] === ROLE_ADMIN && (role !== ROLE_ADMIN || spec.status === '정지')) {
    const admins = readAll_('accounts').filter(function (a) {
      return a['권한'] === ROLE_ADMIN && String(a['상태'] || '사용') !== '정지';
    });
    if (admins.length <= 1) throw new Error('관리자가 한 분뿐입니다. 다른 관리자를 먼저 만들어 주세요.');
  }

  acc['이름'] = String(spec.name || acc['이름']).trim();
  acc['권한'] = role;
  acc['목장'] = String(spec.mokjang || '').trim();
  acc['순'] = String(spec.soon || '').trim();
  acc['상태'] = spec.status === '정지' ? '정지' : '사용';
  acc['메모'] = String(spec.memo || '');

  if (spec.password) {
    if (String(spec.password).length < 4) throw new Error('비밀번호는 4자 이상이어야 합니다.');
    const salt = Utilities.getUuid();
    acc['솔트'] = salt;
    acc['비밀번호'] = hash_(salt + '|' + spec.password);
  }
  writeRow_('accounts', acc._row, acc);
  return accountToClient_(acc);
}

function apiDeleteAccount(token, id) {
  const user = requireAdmin_(auth_(token));
  if (String(id) === user.id) throw new Error('자기 계정은 지울 수 없습니다.');
  const acc = findBy_('accounts', 'ID', id);
  if (!acc) throw new Error('그 계정을 찾지 못했습니다.');
  getSheet_('accounts').deleteRow(acc._row);
  return true;
}


/* ══════════════════════════ 교인 명단 ══════════════════════════ */

function memberToClient_(m) {
  return {
    id: String(m['ID'] || ''),
    name: String(m['이름'] || ''),
    gender: String(m['성별'] || ''),
    birth: toDateStr_(m['생년월일']),
    lunar: truthy_(m['음력']),
    phone: String(m['휴대전화'] || ''),
    mokjang: String(m['목장'] || ''),
    soon: String(m['순'] || ''),
    role: String(m['직분'] || ''),
    joined: toDateStr_(m['등록일']),
    address: String(m['주소'] || ''),
    status: String(m['상태'] || '재적') || '재적',
    memo: String(m['메모'] || '')
  };
}

/** 로그인한 사람이 볼 수 있는 교인만 (원본 행 정보 포함) */
function scopedMemberRows_(user) {
  return readAll_('members').filter(function (m) {
    return inScope_(user, m['목장'], m['순']);
  });
}

function apiListMembers(token) {
  const user = auth_(token);
  return scopedMemberRows_(user).map(memberToClient_).sort(function (a, b) {
    const k = (a.mokjang + a.soon).localeCompare(b.mokjang + b.soon, 'ko');
    return k !== 0 ? k : a.name.localeCompare(b.name, 'ko');
  });
}

function apiSaveMember(token, m) {
  const user = auth_(token);
  if (!m || !String(m.name || '').trim()) throw new Error('이름은 반드시 입력해야 합니다.');

  // 순장·목장은 자기 범위 안에서만 등록할 수 있습니다.
  const mokjang = user.role === ROLE_ADMIN ? String(m.mokjang || '').trim() : user.mokjang;
  const soon = user.role === ROLE_SOON ? user.soon : String(m.soon || '').trim();
  assertScope_(user, mokjang, soon);

  const existing = m.id ? findBy_('members', 'ID', m.id) : null;
  if (m.id && !existing) throw new Error('그 교인을 찾지 못했습니다.');
  if (existing) assertScope_(user, existing['목장'], existing['순']);

  const record = {
    'ID': m.id || newId_('M', 'members'),
    '이름': String(m.name).trim(),
    '성별': m.gender || '',
    '생년월일': toDateStr_(m.birth),
    '음력': m.lunar ? '음력' : '',
    '휴대전화': m.phone || '',
    '목장': mokjang,
    '순': soon,
    '직분': m.role || '',
    '등록일': toDateStr_(m.joined),
    '주소': m.address || '',
    '상태': MEMBER_STATUS.indexOf(m.status) >= 0 ? m.status : '재적',
    '메모': m.memo || '',
    '수정일시': nowStamp_() + ' ' + user.name
  };

  if (existing) writeRow_('members', existing._row, record);
  else appendRow_('members', record);

  syncMemberInfo_(record['ID'], record['이름'], mokjang, soon);
  return memberToClient_(findBy_('members', 'ID', record['ID']));
}

/** 교인의 이름·소속이 바뀌면 다른 시트의 표시용 값도 맞춰 둡니다. */
function syncMemberInfo_(id, name, mokjang, soon) {
  ['attendance', 'prayers', 'visits'].forEach(function (key) {
    const sh = getSheet_(key);
    if (sh.getLastRow() < 2) return;
    const headers = SHEETS[key].headers;
    const idCol = headers.indexOf('교인ID') + 1;
    if (idCol < 1) return;

    const n = sh.getLastRow() - 1;
    const ids = sh.getRange(2, idCol, n, 1).getValues();
    const cols = { '이름': name, '목장': mokjang, '순': soon };

    Object.keys(cols).forEach(function (h) {
      const col = headers.indexOf(h) + 1;
      if (col < 1) return;
      const vals = sh.getRange(2, col, n, 1).getValues();
      let changed = false;
      for (let i = 0; i < n; i++) {
        if (String(ids[i][0]) === String(id) && vals[i][0] !== cols[h]) {
          vals[i][0] = cols[h];
          changed = true;
        }
      }
      if (changed) sh.getRange(2, col, n, 1).setValues(vals);
    });
  });
}

function apiDeleteMember(token, id) {
  const user = auth_(token);
  const m = findBy_('members', 'ID', id);
  if (!m) throw new Error('그 교인을 찾지 못했습니다.');
  assertScope_(user, m['목장'], m['순']);
  getSheet_('members').deleteRow(m._row);
  return true;
}


/* ══════════════════════════ 출결 ══════════════════════════ */

function apiGetAttendance(token, dateStr) {
  const user = auth_(token);
  const date = toDateStr_(dateStr) || lastSunday_();
  const records = {};

  readAll_('attendance').forEach(function (r) {
    if (toDateStr_(r['날짜']) !== date) return;
    if (!inScope_(user, r['목장'], r['순'])) return;
    records[String(r['교인ID'])] = {
      status: String(r['상태'] || ''),
      memo: String(r['메모'] || '')
    };
  });
  return { date: date, records: records };
}

/**
 * 출결을 한 번에 저장합니다.
 * records: [{ id, status, memo }] — status 가 빈 값이면 그 기록을 지웁니다.
 */
function apiSaveAttendance(token, dateStr, records) {
  const user = auth_(token);
  const date = toDateStr_(dateStr);
  if (!date) throw new Error('날짜를 선택해 주세요.');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = getSheet_('attendance');
    const headers = SHEETS.attendance.headers;

    const memberById = {};
    scopedMemberRows_(user).forEach(function (m) { memberById[String(m['ID'])] = m; });

    const onDate = {};
    readAll_('attendance').forEach(function (r) {
      if (toDateStr_(r['날짜']) === date) onDate[String(r['교인ID'])] = r;
    });

    const stamp = nowStamp_();
    const toAppend = [];
    const toDelete = [];

    (records || []).forEach(function (rec) {
      const id = String(rec.id || '');
      const m = memberById[id];
      if (!m) return;                       // 범위 밖이면 조용히 건너뜁니다
      const status = String(rec.status || '').trim();
      if (status && ATTEND_STATUS.indexOf(status) < 0) return;

      const prev = onDate[id];
      if (!status) {
        if (prev) toDelete.push(prev._row);
        return;
      }
      const obj = {
        '날짜': date,
        '교인ID': id,
        '이름': String(m['이름']),
        '목장': String(m['목장'] || ''),
        '순': String(m['순'] || ''),
        '상태': status,
        '메모': rec.memo || '',
        '기록일시': stamp,
        '기록자': user.name
      };
      if (prev) writeRow_('attendance', prev._row, obj);
      else toAppend.push(rowFrom_('attendance', obj));
    });

    if (toAppend.length) {
      sh.getRange(sh.getLastRow() + 1, 1, toAppend.length, headers.length).setValues(toAppend);
    }
    // 행 번호가 밀리지 않도록 아래에서부터 지웁니다.
    toDelete.sort(function (a, b) { return b - a; })
            .forEach(function (r) { sh.deleteRow(r); });

    return apiGetAttendance(token, date);
  } finally {
    lock.releaseLock();
  }
}


/* ══════════════════════════ 기도제목 ══════════════════════════ */

function prayerToClient_(p) {
  return {
    id: String(p['ID'] || ''),
    memberId: String(p['교인ID'] || ''),
    name: String(p['이름'] || ''),
    mokjang: String(p['목장'] || ''),
    soon: String(p['순'] || ''),
    category: String(p['분류'] || ''),
    body: String(p['내용'] || ''),
    date: toDateStr_(p['등록일']),
    status: String(p['상태'] || '진행중') || '진행중',
    answeredAt: toDateStr_(p['응답일']),
    answerNote: String(p['응답메모'] || ''),
    author: String(p['작성자'] || '')
  };
}

/** 자기 범위의 기도제목 + 대상이 없는 공동 기도제목 */
function prayerVisible_(user, p) {
  if (!String(p['교인ID'] || '') && !String(p['목장'] || '')) return true;
  return inScope_(user, p['목장'], p['순']);
}

function apiListPrayers(token) {
  const user = auth_(token);
  return readAll_('prayers')
    .filter(function (p) { return prayerVisible_(user, p); })
    .map(prayerToClient_)
    .sort(function (a, b) { return String(b.date).localeCompare(String(a.date)); });
}

function apiSavePrayer(token, p) {
  const user = auth_(token);
  if (!p || !String(p.body || '').trim()) throw new Error('기도제목 내용을 입력해 주세요.');

  let name = '', mokjang = '', soon = '';
  if (p.memberId) {
    const m = findBy_('members', 'ID', p.memberId);
    if (!m) throw new Error('그 교인을 찾지 못했습니다.');
    assertScope_(user, m['목장'], m['순']);
    name = String(m['이름']);
    mokjang = String(m['목장'] || '');
    soon = String(m['순'] || '');
  } else if (user.role !== ROLE_ADMIN) {
    // 대상을 고르지 않으면 자기 범위의 공동 제목으로 둡니다.
    mokjang = user.mokjang;
    soon = user.role === ROLE_SOON ? user.soon : '';
  }

  const existing = p.id ? findBy_('prayers', 'ID', p.id) : null;
  if (p.id && !existing) throw new Error('그 기도제목을 찾지 못했습니다.');
  if (existing && !prayerVisible_(user, existing)) {
    throw new Error('맡으신 범위 밖의 기도제목입니다.');
  }

  const record = {
    'ID': p.id || newId_('P', 'prayers'),
    '교인ID': p.memberId || '',
    '이름': name,
    '목장': mokjang,
    '순': soon,
    '분류': p.category || '',
    '내용': String(p.body).trim(),
    '등록일': toDateStr_(p.date) || today_(),
    '상태': p.status === '응답됨' ? '응답됨' : '진행중',
    '응답일': toDateStr_(p.answeredAt),
    '응답메모': p.answerNote || '',
    '작성자': existing ? String(existing['작성자'] || user.name) : user.name,
    '수정일시': nowStamp_() + ' ' + user.name
  };

  if (existing) writeRow_('prayers', existing._row, record);
  else appendRow_('prayers', record);

  return prayerToClient_(findBy_('prayers', 'ID', record['ID']));
}

function apiAnswerPrayer(token, id, note) {
  const user = auth_(token);
  const row = findBy_('prayers', 'ID', id);
  if (!row) throw new Error('그 기도제목을 찾지 못했습니다.');
  if (!prayerVisible_(user, row)) throw new Error('맡으신 범위 밖의 기도제목입니다.');

  const answered = String(row['상태']) !== '응답됨';
  row['상태'] = answered ? '응답됨' : '진행중';
  row['응답일'] = answered ? today_() : '';
  row['응답메모'] = answered ? (note || String(row['응답메모'] || '')) : '';
  row['수정일시'] = nowStamp_() + ' ' + user.name;
  writeRow_('prayers', row._row, row);
  return prayerToClient_(row);
}

function apiDeletePrayer(token, id) {
  const user = auth_(token);
  const row = findBy_('prayers', 'ID', id);
  if (!row) throw new Error('그 기도제목을 찾지 못했습니다.');
  if (!prayerVisible_(user, row)) throw new Error('맡으신 범위 밖의 기도제목입니다.');
  getSheet_('prayers').deleteRow(row._row);
  return true;
}


/* ══════════════════════════ 심방 ══════════════════════════ */

function visitToClient_(v) {
  return {
    id: String(v['ID'] || ''),
    date: toDateStr_(v['날짜']),
    memberId: String(v['교인ID'] || ''),
    name: String(v['이름'] || ''),
    mokjang: String(v['목장'] || ''),
    soon: String(v['순'] || ''),
    method: String(v['방식'] || '심방'),
    by: String(v['담당'] || ''),
    body: String(v['내용'] || '')
  };
}

function apiListVisits(token) {
  const user = auth_(token);
  return readAll_('visits')
    .filter(function (v) { return inScope_(user, v['목장'], v['순']); })
    .map(visitToClient_)
    .sort(function (a, b) { return String(b.date).localeCompare(String(a.date)); });
}

function apiSaveVisit(token, v) {
  const user = auth_(token);
  if (!v || !v.memberId) throw new Error('심방한 분을 선택해 주세요.');

  const m = findBy_('members', 'ID', v.memberId);
  if (!m) throw new Error('그 교인을 찾지 못했습니다.');
  assertScope_(user, m['목장'], m['순']);

  const existing = v.id ? findBy_('visits', 'ID', v.id) : null;
  if (v.id && !existing) throw new Error('그 기록을 찾지 못했습니다.');
  if (existing) assertScope_(user, existing['목장'], existing['순']);

  const record = {
    'ID': v.id || newId_('V', 'visits'),
    '날짜': toDateStr_(v.date) || today_(),
    '교인ID': v.memberId,
    '이름': String(m['이름']),
    '목장': String(m['목장'] || ''),
    '순': String(m['순'] || ''),
    '방식': VISIT_METHODS.indexOf(v.method) >= 0 ? v.method : '심방',
    '담당': String(v.by || user.name),
    '내용': v.body || '',
    '기록일시': nowStamp_() + ' ' + user.name
  };

  if (existing) writeRow_('visits', existing._row, record);
  else appendRow_('visits', record);

  return visitToClient_(findBy_('visits', 'ID', record['ID']));
}

function apiDeleteVisit(token, id) {
  const user = auth_(token);
  const row = findBy_('visits', 'ID', id);
  if (!row) throw new Error('그 기록을 찾지 못했습니다.');
  assertScope_(user, row['목장'], row['순']);
  getSheet_('visits').deleteRow(row._row);
  return true;
}


/* ══════════════════════════ 첫 화면 자료 ══════════════════════════ */

function apiBootstrap(token) {
  const user = auth_(token);
  const memberRows = scopedMemberRows_(user);
  const members = memberRows.map(memberToClient_).sort(function (a, b) {
    const k = (a.mokjang + a.soon).localeCompare(b.mokjang + b.soon, 'ko');
    return k !== 0 ? k : a.name.localeCompare(b.name, 'ko');
  });

  const allMembers = user.role === ROLE_ADMIN ? memberRows : readAll_('members');

  return {
    today: today_(),
    lastSunday: lastSunday_(),
    me: {
      id: user.id, name: user.name, role: user.role,
      mokjang: user.mokjang, soon: user.soon, scope: scopeLabel_(user)
    },
    members: members,
    // 목장·순 선택지는 전체 시트에서 뽑아야 새 소속을 만들 수 있습니다.
    mokjangs: uniqueSorted_(allMembers.map(function (m) { return m['목장']; })),
    soons: uniqueSorted_(members.map(function (m) { return m.soon; })),
    prayers: apiListPrayers(token),
    visits: apiListVisits(token),
    dashboard: buildDashboard_(user, members),
    accounts: user.role === ROLE_ADMIN ? apiListAccounts(token) : [],
    options: {
      memberStatus: MEMBER_STATUS,
      attendStatus: ATTEND_STATUS,
      visitMethods: VISIT_METHODS,
      roles: ROLES
    }
  };
}

function apiDashboard(token) {
  const user = auth_(token);
  return buildDashboard_(user, scopedMemberRows_(user).map(memberToClient_));
}

function buildDashboard_(user, members) {
  const today = today_();
  const active = members.filter(function (m) {
    return m.status !== '이명' && m.status !== '별세';
  });
  const mine = {};
  active.forEach(function (m) { mine[m.id] = m; });

  /* ── 생일 ── */
  const upcoming = [], thisMonth = [];
  const curMonth = today.slice(5, 7);

  active.forEach(function (m) {
    const md = monthDay_(m.birth);
    if (!md) return;
    if (md.slice(0, 2) === curMonth) {
      thisMonth.push({ name: m.name, mokjang: m.mokjang, soon: m.soon, md: md, lunar: m.lunar });
    }
    const diff = daysUntilMonthDay_(today, md);
    if (diff !== null && diff <= 7) {
      upcoming.push({ name: m.name, mokjang: m.mokjang, soon: m.soon, md: md,
                      lunar: m.lunar, inDays: diff, phone: m.phone });
    }
  });
  upcoming.sort(function (a, b) { return a.inDays - b.inDays; });
  thisMonth.sort(function (a, b) { return a.md.localeCompare(b.md); });

  /* ── 출결 ── */
  const lastSeen = {}, byDate = {};
  readAll_('attendance').forEach(function (r) {
    const id = String(r['교인ID']);
    if (!mine[id]) return;                       // 내 범위만 셉니다
    const d = toDateStr_(r['날짜']);
    if (!d) return;
    const st = String(r['상태'] || '');
    if (!byDate[d]) byDate[d] = { date: d, '출석': 0, '온라인': 0, '결석': 0 };
    if (byDate[d][st] !== undefined) byDate[d][st]++;
    if (st === '출석' || st === '온라인') {
      if (!lastSeen[id] || lastSeen[id] < d) lastSeen[id] = d;
    }
  });

  const dates = Object.keys(byDate).sort();
  const recentDate = dates.length ? dates[dates.length - 1] : '';
  const trend = dates.slice(-8).map(function (d) { return byDate[d]; });

  /* ── 오래 못 뵌 분 ── */
  const away = [];
  active.forEach(function (m) {
    const seen = lastSeen[m.id] || '';
    const gap = seen ? daysBetween_(seen, today) : null;
    if (gap === null) {
      if (dates.length >= 4) {
        away.push({ name: m.name, mokjang: m.mokjang, soon: m.soon,
                    phone: m.phone, lastSeen: '', days: null });
      }
    } else if (gap >= LONG_ABSENCE_DAYS) {
      away.push({ name: m.name, mokjang: m.mokjang, soon: m.soon,
                  phone: m.phone, lastSeen: seen, days: gap });
    }
  });
  away.sort(function (a, b) {
    if (a.days === null) return 1;
    if (b.days === null) return -1;
    return b.days - a.days;
  });

  /* ── 기도제목 ── */
  const openPrayers = readAll_('prayers')
    .filter(function (p) {
      return String(p['상태']) !== '응답됨' && prayerVisible_(user, p);
    })
    .map(prayerToClient_)
    .sort(function (a, b) { return String(b.date).localeCompare(String(a.date)); });

  /* ── 심방 ── */
  const visits = readAll_('visits')
    .filter(function (v) { return inScope_(user, v['목장'], v['순']); })
    .map(visitToClient_)
    .sort(function (a, b) { return String(b.date).localeCompare(String(a.date)); });

  const lastVisit = {};
  visits.forEach(function (v) {
    if (!v.memberId || !v.date) return;
    if (!lastVisit[v.memberId] || lastVisit[v.memberId] < v.date) lastVisit[v.memberId] = v.date;
  });

  const needVisit = active.map(function (m) {
    const last = lastVisit[m.id] || '';
    return {
      name: m.name, mokjang: m.mokjang, soon: m.soon, phone: m.phone,
      lastVisit: last, days: last ? daysBetween_(last, today) : null
    };
  }).filter(function (x) {
    return x.days === null || x.days >= VISIT_GAP_DAYS;
  }).sort(function (a, b) {
    if (a.days === null) return -1;
    if (b.days === null) return 1;
    return b.days - a.days;
  }).slice(0, 20);

  return {
    today: today,
    scope: scopeLabel_(user),
    counts: {
      total: members.length,
      active: active.filter(function (m) { return m.status === '재적'; }).length,
      away: away.length,
      openPrayers: openPrayers.length
    },
    birthdaysUpcoming: upcoming,
    birthdaysThisMonth: thisMonth,
    recentService: recentDate ? byDate[recentDate] : null,
    trend: trend,
    awayMembers: away.slice(0, 20),
    openPrayers: openPrayers.slice(0, 6),
    recentVisits: visits.slice(0, 6),
    needVisit: needVisit
  };
}

/** 오늘부터 'MM-dd' 생일까지 남은 날수 (연말·연초를 넘겨도 계산됩니다) */
function daysUntilMonthDay_(todayStr, md) {
  if (!md) return null;
  const year = Number(todayStr.slice(0, 4));
  const tries = [year + '-' + md, (year + 1) + '-' + md];
  for (let i = 0; i < tries.length; i++) {
    const diff = daysBetween_(todayStr, tries[i]);
    if (diff !== null && diff >= 0) return diff;
  }
  return null;
}
