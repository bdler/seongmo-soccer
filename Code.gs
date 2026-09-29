/**
 * 성모 사커 (Seongmo Soccer) - Google Apps Script 서버 코드
 *
 * - doGet(): 웹 앱으로 게임 화면(index.html)을 제공합니다.
 * - saveMatchResult(): 경기 기록을 구글 시트에 저장합니다. (점수는 서버에서 계산)
 * - getLeaderboard(): 점수 순위 상위 N개 기록을 반환합니다.
 * - getStats(): 특정 닉네임의 누적 전적을 반환합니다.
 *
 * 이름이 밑줄(_)로 끝나는 함수는 내부 도우미 함수이며 클라이언트(google.script.run)에서 호출할 수 없습니다.
 */

var SHEET_NAME_ = '경기기록';
var SS_PROP_KEY_ = 'SEONGMO_SOCCER_SPREADSHEET_ID';
var SS_TITLE_ = '성모 사커 경기 기록';
var TZ_ = 'Asia/Seoul';
var HEADERS_ = ['일시', '닉네임', '내 팀', '상대 팀', '득점', '실점', '결과', '난이도', '점수'];

// index.html 의 TEAMS 목록과 반드시 같아야 합니다.
var TEAM_NAMES_ = [
  '성모 FC', '한강 유나이티드', '푸른별 시티', '백두 타이거즈',
  '은하 로버스', '청솔 워리어스', '바다 갈매기', '불꽃 레인저스'
];
var DIFFICULTY_MULT_ = { '쉬움': 1, '보통': 1.5, '어려움': 2 };
var RESULTS_ = ['승', '무', '패'];
var NICK_MAX_ = 12;
// 경기 시간(분) 화이트리스트와 현실적인 득점 상한 (내 팀 득점은 분당 최대 GOALS_PER_MIN_ 골, 실점은 GOALS_MAX_ 까지)
var MINUTES_ = [2, 4, 6];
var GOALS_PER_MIN_ = 4;
var GOALS_MAX_ = 24;
// 저장 요청 제한 (CacheService): 전체 분당 최대 저장 수, 같은 닉네임 연속 저장 간격(초)
var RATE_GLOBAL_PER_MIN_ = 30;
var RATE_NICK_COOLDOWN_SEC_ = 15;

/**
 * 웹 앱 진입점.
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('성모 사커')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no, viewport-fit=cover')
    // 태블릿/휴대폰에서 '홈 화면에 추가'로 열 때 주소창 없이 전체 화면처럼 보이도록
    .addMetaTag('mobile-web-app-capable', 'yes')
    .addMetaTag('apple-mobile-web-app-capable', 'yes')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 경기 결과 저장.
 * @param {{nickname:string, myTeam:string, oppTeam:string, goalsFor:number, goalsAgainst:number, difficulty:string, minutes:number}} record
 * @return {{ok:boolean, message?:string, score?:number, result?:string}}
 */
function saveMatchResult(record) {
  var clean;
  try {
    clean = sanitizeRecord_(record);
  } catch (err) {
    return { ok: false, message: String((err && err.message) || err) };
  }

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return { ok: false, message: '서버가 혼잡합니다. 잠시 후 다시 시도해 주세요.' };
  }
  try {
    var limited = checkRateLimit_(clean.nickname);
    if (limited) return { ok: false, message: limited };
    var sheet = getOrCreateSheet_();
    // 글자 칸은 모두 textCell_ 로 감싸 시트가 숫자/날짜/불리언/수식으로 바꾸지 않게 합니다.
    sheet.appendRow([
      new Date(),
      textCell_(clean.nickname),
      textCell_(clean.myTeam),
      textCell_(clean.oppTeam),
      clean.goalsFor,
      clean.goalsAgainst,
      textCell_(clean.result),
      textCell_(clean.difficulty),
      clean.score
    ]);
    SpreadsheetApp.flush();
    recordRateLimit_(clean.nickname);
    return { ok: true, score: clean.score, result: clean.result };
  } catch (err) {
    console.error('saveMatchResult 실패: ' + ((err && err.stack) || err));
    return { ok: false, message: '기록 저장 중 오류가 발생했습니다.' };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 점수 순위 상위 기록.
 * @param {number} limit 1~100 (기본 20)
 * @return {Array<Object>} JSON 직렬화 가능한 기록 배열 (Date 는 문자열로 변환)
 */
function getLeaderboard(limit) {
  // 범위를 벗어나면 1~100 으로 맞추고, 숫자가 아니면 기본값 20
  var n = Math.floor(Number(limit));
  n = isFinite(n) ? Math.min(100, Math.max(1, n)) : 20;
  var rows;
  try {
    rows = readAllRows_();
  } catch (err) {
    console.error('getLeaderboard 실패: ' + ((err && err.stack) || err));
    throw new Error('랭킹을 불러오지 못했습니다.');
  }
  rows.sort(function (a, b) {
    return (b.score - a.score) || (b.ts - a.ts);
  });
  return rows.slice(0, n).map(function (r) {
    return {
      date: r.date,
      nickname: r.nickname,
      myTeam: r.myTeam,
      oppTeam: r.oppTeam,
      goalsFor: r.goalsFor,
      goalsAgainst: r.goalsAgainst,
      result: r.result,
      difficulty: r.difficulty,
      score: r.score
    };
  });
}

/**
 * 닉네임별 누적 전적.
 * @param {string} nickname
 */
function getStats(nickname) {
  var nick = cleanText_(nickname, NICK_MAX_);
  var out = { nickname: nick, games: 0, wins: 0, draws: 0, losses: 0, goalsFor: 0, goalsAgainst: 0, bestScore: 0 };
  if (!nick) return out;
  var rows;
  try {
    rows = readAllRows_();
  } catch (err) {
    console.error('getStats 실패: ' + ((err && err.stack) || err));
    throw new Error('전적을 불러오지 못했습니다.');
  }
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (r.nickname !== nick) continue;
    out.games++;
    if (r.result === '승') out.wins++;
    else if (r.result === '무') out.draws++;
    else if (r.result === '패') out.losses++;
    out.goalsFor += r.goalsFor;
    out.goalsAgainst += r.goalsAgainst;
    if (r.score > out.bestScore) out.bestScore = r.score;
  }
  return out;
}

/* ------------------------------------------------------------------ */
/* 내부 도우미 함수                                                     */
/* ------------------------------------------------------------------ */

/** 점수 계산 (index.html 의 computeScore 와 동일한 공식) */
function computeScore_(gf, ga, difficulty) {
  var base = gf > ga ? 300 : (gf === ga ? 100 : 0);
  var raw = base + gf * 50 - ga * 20 + Math.max(0, gf - ga) * 30;
  var mult = DIFFICULTY_MULT_[difficulty] || 1;
  return Math.max(0, Math.round(raw * mult));
}

/** 입력 검증/정리. 잘못된 값이면 Error 를 던집니다. */
function sanitizeRecord_(r) {
  if (!r || typeof r !== 'object') throw new Error('잘못된 기록 형식입니다.');
  var nickname = cleanText_(r.nickname, NICK_MAX_) || '익명';
  var myTeam = cleanText_(r.myTeam, 30);
  var oppTeam = cleanText_(r.oppTeam, 30);
  if (TEAM_NAMES_.indexOf(myTeam) < 0 || TEAM_NAMES_.indexOf(oppTeam) < 0) {
    throw new Error('알 수 없는 팀입니다.');
  }
  if (myTeam === oppTeam) throw new Error('같은 팀끼리는 경기할 수 없습니다.');
  var minutes = toInt_(r.minutes, 1, 60);
  if (minutes === null || MINUTES_.indexOf(minutes) < 0) throw new Error('경기 시간 값이 올바르지 않습니다.');
  // 득점은 점수를 올리므로 경기 시간에 비례한 현실적인 상한을 둡니다.
  // 실점은 많을수록 점수가 내려가므로(조작할 이유가 없음) 실제 경기 결과가 거부되지 않도록 전체 상한만 적용합니다.
  var goalCap = Math.min(GOALS_MAX_, minutes * GOALS_PER_MIN_);
  var gf = toInt_(r.goalsFor, 0, goalCap);
  var ga = toInt_(r.goalsAgainst, 0, GOALS_MAX_);
  if (gf === null || ga === null) throw new Error('득점/실점 값이 올바르지 않습니다.');
  var difficulty = cleanText_(r.difficulty, 5);
  if (!Object.prototype.hasOwnProperty.call(DIFFICULTY_MULT_, difficulty)) {
    throw new Error('알 수 없는 난이도입니다.');
  }
  // 결과는 클라이언트 값을 믿지 않고 득실로 다시 계산합니다.
  var result = gf > ga ? '승' : (gf === ga ? '무' : '패');
  if (RESULTS_.indexOf(result) < 0) throw new Error('결과 값이 올바르지 않습니다.');
  return {
    nickname: nickname,
    myTeam: myTeam,
    oppTeam: oppTeam,
    goalsFor: gf,
    goalsAgainst: ga,
    difficulty: difficulty,
    minutes: minutes,
    result: result,
    score: computeScore_(gf, ga, difficulty)
  };
}

/** 제어 문자 제거, 공백 정리, 길이 제한 */
function cleanText_(v, max) {
  if (v === null || v === undefined) return '';
  var s = String(v)
    .replace(/[\u0000-\u001F\u007F-\u009F\u200B-\u200F\u2028\u2029\uFEFF]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
  if (s.length > max) s = s.substring(0, max).trim();
  return s;
}

/**
 * 글자 칸을 항상 '텍스트'로 저장합니다.
 * 앞에 ' (따옴표 접두사)를 붙이면 시트는 값을 그대로 문자열로 보관하고, 읽을 때 ' 는 값에 포함되지 않습니다.
 * - = + - @ 로 시작하는 수식 주입 방지
 * - '007' → 7, '50%' → 0.5, 'TRUE' → true, '2024-01-01' → 날짜 처럼 자동 변환되는 것 방지
 */
function textCell_(s) {
  return "'" + String(s);
}

/**
 * 엄격한 정수 검증. 진짜 정수(number) 또는 숫자로만 된 짧은 문자열('3')만 허용합니다.
 * 2.9, true, [5], '0x1e', '1e1' 같은 값은 거부(null)합니다.
 */
function toInt_(v, min, max) {
  var n;
  if (typeof v === 'number') n = v;
  else if (typeof v === 'string' && /^\d{1,3}$/.test(v)) n = parseInt(v, 10);
  else return null;
  if (!isFinite(n) || Math.floor(n) !== n) return null;
  if (n < min || n > max) return null;
  return n;
}

/**
 * 저장 요청 제한 확인 (반드시 스크립트 잠금 안에서 호출). 제한에 걸리면 안내 문구, 아니면 null.
 * - 전체: 1분에 RATE_GLOBAL_PER_MIN_ 건
 * - 같은 닉네임: RATE_NICK_COOLDOWN_SEC_ 초에 1건
 */
function checkRateLimit_(nickname) {
  try {
    var cache = CacheService.getScriptCache();
    if (!cache) return null;
    var n = Number(cache.get(rateGlobalKey_())) || 0;
    if (n >= RATE_GLOBAL_PER_MIN_) return '지금은 저장 요청이 너무 많습니다. 1분 뒤에 다시 시도해 주세요.';
    if (cache.get(rateNickKey_(nickname))) return '같은 닉네임으로 방금 저장했습니다. 잠시 후 다시 시도해 주세요.';
  } catch (e) {
    console.warn('요청 제한 확인 실패(무시): ' + e);
  }
  return null;
}

/** 저장 성공 후 요청 제한 카운터 갱신 */
function recordRateLimit_(nickname) {
  try {
    var cache = CacheService.getScriptCache();
    var key = rateGlobalKey_();
    var n = Number(cache.get(key)) || 0;
    cache.put(key, String(n + 1), 120);
    cache.put(rateNickKey_(nickname), '1', RATE_NICK_COOLDOWN_SEC_);
  } catch (e) {
    console.warn('요청 제한 기록 실패(무시): ' + e);
  }
}

function rateGlobalKey_() {
  return 'rate_g_' + Math.floor(Date.now() / 60000);
}

function rateNickKey_(nickname) {
  return 'rate_n_' + encodeURIComponent(String(nickname));
}

/** 시트 셀 값을 문자열로 (시트가 날짜/숫자로 자동 변환한 값도 안전하게 처리) */
function cellText_(v) {
  if (v === null || v === undefined) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return isNaN(v.getTime()) ? '' : Utilities.formatDate(v, TZ_, 'yyyy-MM-dd');
  }
  return String(v);
}

/**
 * 스프레드시트 찾기. 컨테이너 바인딩 시트를 우선 사용합니다.
 * - 바인딩 시트도, 저장된 ID 도 없으면 null (이때만 새로 만들어도 됩니다)
 * - ID 가 저장되어 있는데 열리지 않으면 한 번 재시도 후 오류를 던집니다.
 *   (일시적인 "Service timed out" 등으로 새 시트를 만들어 기존 기록과 연결이 끊기는 것을 막기 위함.
 *    저장된 ID 는 절대 자동으로 덮어쓰지 않습니다.)
 */
function findSpreadsheet_() {
  var ss = null;
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {
    ss = null;
  }
  if (ss) return ss;
  var id = PropertiesService.getScriptProperties().getProperty(SS_PROP_KEY_);
  if (!id) return null;
  var lastErr = null;
  for (var attempt = 0; attempt < 2; attempt++) {
    try {
      return SpreadsheetApp.openById(id);
    } catch (e) {
      lastErr = e;
      console.warn('저장된 스프레드시트를 열 수 없습니다 (시도 ' + (attempt + 1) + '/2): ' + e);
      if (attempt === 0) Utilities.sleep(800);
    }
  }
  throw new Error('기록 스프레드시트(ID ' + id + ')를 열 수 없습니다: ' + ((lastErr && lastErr.message) || lastErr) +
    ' / 시트를 삭제했다면 스크립트 속성 ' + SS_PROP_KEY_ + ' 를 지운 뒤 다시 시도하세요.');
}

/**
 * 기록 시트를 가져오거나 새로 만듭니다. (반드시 스크립트 잠금 안에서 호출)
 * - 스크립트가 시트에 바인딩되어 있으면 그 시트를 사용
 * - 독립형 스크립트면 스프레드시트를 한 번 만들고 ID 를 스크립트 속성에 저장
 */
function getOrCreateSheet_() {
  var ss = findSpreadsheet_();
  var created = false;
  if (!ss) {
    ss = SpreadsheetApp.create(SS_TITLE_);
    PropertiesService.getScriptProperties().setProperty(SS_PROP_KEY_, ss.getId());
    created = true;
  }
  var sheet = ss.getSheetByName(SHEET_NAME_);
  if (!sheet) {
    if (created && ss.getSheets().length === 1 && ss.getSheets()[0].getLastRow() === 0) {
      sheet = ss.getSheets()[0].setName(SHEET_NAME_);
    } else {
      sheet = ss.insertSheet(SHEET_NAME_);
    }
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS_);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS_.length).setFontWeight('bold').setBackground('#1e3a8a').setFontColor('#ffffff');
    sheet.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm');
  }
  return sheet;
}

/** 기존 기록 시트 (없으면 null, 새로 만들지 않음) */
function findSheet_() {
  var ss = findSpreadsheet_();
  if (!ss) return null;
  return ss.getSheetByName(SHEET_NAME_);
}

/** 모든 기록을 JSON 친화적인 객체 배열로 읽기 */
function readAllRows_() {
  var sheet = findSheet_();
  if (!sheet) return [];
  var last = sheet.getLastRow();
  if (last < 2) return [];
  var values = sheet.getRange(2, 1, last - 1, HEADERS_.length).getValues();
  var rows = [];
  for (var i = 0; i < values.length; i++) {
    var v = values[i];
    var score = Number(v[8]);
    if (v[1] === '' || !isFinite(score)) continue;
    var d = v[0];
    var isDate = Object.prototype.toString.call(d) === '[object Date]' && !isNaN(d.getTime());
    rows.push({
      ts: isDate ? d.getTime() : 0,
      date: isDate ? Utilities.formatDate(d, TZ_, 'yyyy-MM-dd HH:mm') : String(d),
      nickname: cellText_(v[1]),
      myTeam: cellText_(v[2]),
      oppTeam: cellText_(v[3]),
      goalsFor: Number(v[4]) || 0,
      goalsAgainst: Number(v[5]) || 0,
      result: cellText_(v[6]),
      difficulty: cellText_(v[7]),
      score: score
    });
  }
  return rows;
}
