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
var GOALS_MAX_ = 30;

/**
 * 웹 앱 진입점.
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('성모 사커')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 경기 결과 저장.
 * @param {{nickname:string, myTeam:string, oppTeam:string, goalsFor:number, goalsAgainst:number, difficulty:string}} record
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
    var sheet = getOrCreateSheet_();
    sheet.appendRow([
      new Date(),
      safeCell_(clean.nickname),
      safeCell_(clean.myTeam),
      safeCell_(clean.oppTeam),
      clean.goalsFor,
      clean.goalsAgainst,
      clean.result,
      clean.difficulty,
      clean.score
    ]);
    SpreadsheetApp.flush();
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
  var n = toInt_(limit, 1, 100);
  if (n === null) n = 20;
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
  var gf = toInt_(r.goalsFor, 0, GOALS_MAX_);
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

/** 스프레드시트 수식 주입 방지: = + - @ 로 시작하면 앞에 ' 를 붙입니다. */
function safeCell_(s) {
  s = String(s);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

/** 정수 변환 + 범위 확인. 실패 시 null */
function toInt_(v, min, max) {
  if (v === null || v === undefined || v === '') return null;
  var n = Number(v);
  if (!isFinite(n)) return null;
  n = Math.floor(n);
  if (n < min || n > max) return null;
  return n;
}

/** 스프레드시트 찾기 (없으면 null). 컨테이너 바인딩 시트를 우선 사용합니다. */
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
  try {
    return SpreadsheetApp.openById(id);
  } catch (e) {
    console.warn('저장된 스프레드시트를 열 수 없습니다: ' + e);
    return null;
  }
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
      nickname: String(v[1]),
      myTeam: String(v[2]),
      oppTeam: String(v[3]),
      goalsFor: Number(v[4]) || 0,
      goalsAgainst: Number(v[5]) || 0,
      result: String(v[6]),
      difficulty: String(v[7]),
      score: score
    });
  }
  return rows;
}
