/**
 * 기념품 신청 — Google Apps Script 웹앱
 *
 * 스프레드시트 한 개로 전부 처리합니다.
 *   시트1      참가자 명단 (참가일자 / 협력사명 / 직위 / 참가자명)  ← 직접 관리
 *   신청내역    접수 결과                                        ← 자동 생성
 *
 * 설치
 *  1. 아래 스프레드시트를 연다
 *     https://docs.google.com/spreadsheets/d/1P9bgchqFiJMl0OYXGDLy7_rWGNPz4Iss19ZUYvIYB24/edit
 *  2. 확장 프로그램 > Apps Script 에 이 파일 내용을 붙여넣고 저장
 *  3. 배포 > 새 배포 > 유형: 웹 앱
 *     - 실행 계정: 나
 *     - 액세스 권한: 모든 사용자          <- 이걸 안 바꾸면 전부 실패한다
 *  4. 나오는 /exec URL 을 index.html 의 API 값에 붙여넣는다
 *
 * 코드를 고친 뒤에는 "배포 관리 > 수정 > 버전: 새 버전"으로 다시 배포해야 반영됩니다.
 * 저장만으로는 URL 에 적용되지 않습니다.
 */

var SS_ID      = '1P9bgchqFiJMl0OYXGDLy7_rWGNPz4Iss19ZUYvIYB24';
var LIST_SHEET = '시트1';      // 참가자 명단
var DATA_SHEET = '신청내역';    // 접수 결과 (없으면 자동 생성)
var TZ         = 'Asia/Seoul';

var HEADERS = [
  '신청일시', '최종수정', '참가일자', '협력사명', '직위', '참가자명',
  '기념품', '옵션', '수령인', '연락처', '우편번호', '주소', '상세주소', '요청사항', '기념품ID'
];

// 시트 열 순서 <-> 전송 데이터 키 대응
var FIELDS = [
  'created', 'updated', 'event_date', 'company', 'position', 'attendee',
  'gift_name', 'option', 'recipient', 'phone', 'postcode', 'address1', 'address2', 'memo', 'gift_id'
];

var COL_COMPANY = 4;   // '협력사명'  — 이 열로 중복을 찾는다
var COL_PHONE   = 10;  // '연락처'    — 텍스트 서식 고정
var COL_POST    = 11;  // '우편번호'  — 앞자리 0 이 사라지지 않도록


// ---------------------------------------------------------------- 공통
function ss_() {
  return SpreadsheetApp.openById(SS_ID);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function now_() {
  return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
}

var WEEKDAY = ['일', '월', '화', '수', '목', '금', '토'];

/**
 * 참가일자 셀을 '9/22(화)' 형태로 만든다.
 * 시트에 날짜 값으로 들어가 있으면 Date 객체가 넘어오므로 그대로 쓰면
 * 'Tue Sep 22 2026 00:00:00 GMT+0900' 이 되어 버린다.
 * 텍스트로 적혀 있으면 손대지 않고 그대로 돌려준다.
 */
function fmtDate_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) {
    var ymd = Utilities.formatDate(v, TZ, 'yyyy/M/d').split('/');
    var d = new Date(Number(ymd[0]), Number(ymd[1]) - 1, Number(ymd[2]));
    return ymd[1] + '/' + ymd[2] + '(' + WEEKDAY[d.getDay()] + ')';
  }
  return String(v == null ? '' : v).trim();
}

function dataSheet_() {
  var ss = ss_();
  var sh = ss.getSheetByName(DATA_SHEET);
  if (!sh) {
    sh = ss.insertSheet(DATA_SHEET);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  }
  return sh;
}


// ---------------------------------------------------------------- 명단
/**
 * 시트1 을 읽어 [{date, company, position, name}] 로 돌려준다.
 * 협력사명이 아이디이므로 중복은 첫 줄만 쓴다.
 * 헤더 이름으로 열을 찾으므로 열 순서가 바뀌어도 동작한다.
 */
function readList_() {
  var sh = ss_().getSheetByName(LIST_SHEET);
  if (!sh) throw new Error('명단 시트("' + LIST_SHEET + '")를 찾을 수 없습니다.');

  var last = sh.getLastRow();
  if (last < 2) return [];

  var vals = sh.getRange(1, 1, last, sh.getLastColumn()).getValues();
  var head = vals[0].map(function (h) { return String(h).trim(); });
  var iDate = head.indexOf('참가일자');
  var iComp = head.indexOf('협력사명');
  var iPos  = head.indexOf('직위');
  var iName = head.indexOf('참가자명');
  if (iComp < 0) throw new Error('명단 시트에 "협력사명" 열이 없습니다.');

  var out = [], seen = {};
  for (var r = 1; r < vals.length; r++) {
    var company = String(vals[r][iComp] || '').trim();
    if (!company || seen[company]) continue;
    seen[company] = true;
    out.push({
      date:     iDate < 0 ? '' : fmtDate_(vals[r][iDate]),
      company:  company,
      position: iPos  < 0 ? '' : String(vals[r][iPos]  || '').trim(),
      name:     iName < 0 ? '' : String(vals[r][iName] || '').trim()
    });
  }
  return out;
}

function findRow_(sh, company) {
  var last = sh.getLastRow();
  if (last < 2) return 0;
  var vals = sh.getRange(2, COL_COMPANY, last - 1, 1).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === String(company).trim()) return i + 2;
  }
  return 0;
}

function submittedCompanies_() {
  var sh = dataSheet_();
  var last = sh.getLastRow();
  if (last < 2) return [];
  var out = [];
  sh.getRange(2, COL_COMPANY, last - 1, 1).getValues().forEach(function (r) {
    if (r[0]) out.push(String(r[0]).trim());
  });
  return out;
}

// 휴대폰 번호를 010-0000-0000 형태로
function normPhone_(raw) {
  var d = String(raw || '').replace(/\D/g, '');
  if (d.length === 11) return d.slice(0, 3) + '-' + d.slice(3, 7) + '-' + d.slice(7);
  if (d.length === 10) return d.slice(0, 3) + '-' + d.slice(3, 6) + '-' + d.slice(6);
  return String(raw || '').trim();
}


// ---------------------------------------------------------------- 조회
/**
 * ?action=init
 *   첫 화면용. 일자 목록과 (일자, 협력사, 신청여부) 만 내려보낸다.
 *   직위·참가자명은 담지 않는다 — 회사를 고른 사람에게만 pick 으로 준다.
 *
 * ?action=pick&company=...
 *   고른 회사의 직위·참가자명과, 이미 신청했다면 그 내용
 */
function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) || 'init';

    if (action === 'init') {
      var people = readList_();
      var done = {};
      submittedCompanies_().forEach(function (c) { done[c] = true; });

      var dates = [];
      var companies = people.map(function (p) {
        if (dates.indexOf(p.date) < 0) dates.push(p.date);
        return { date: p.date, company: p.company, submitted: !!done[p.company] };
      });
      return json_({ ok: true, dates: dates, companies: companies });
    }

    if (action === 'pick') {
      var company = String(e.parameter.company || '').trim();
      var hit = null;
      readList_().forEach(function (p) { if (p.company === company) hit = p; });
      if (!hit) return json_({ ok: false, error: '명단에 없는 협력사입니다.' });

      var res = { ok: true, position: hit.position, name: hit.name, found: false };

      var sh = dataSheet_();
      var row = findRow_(sh, company);
      if (row) {
        var vals = sh.getRange(row, 1, 1, HEADERS.length).getValues()[0];
        var data = {};
        FIELDS.forEach(function (k, i) { data[k] = String(vals[i] == null ? '' : vals[i]); });
        res.found = true;
        res.data = data;
      }
      return json_(res);
    }

    return json_({ ok: false, error: 'unknown action' });

  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}


// ---------------------------------------------------------------- 접수
/** 협력사명이 아이디. 이미 있으면 그 줄을 덮어쓴다 */
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);

    var d = JSON.parse(e.postData.contents);

    var company   = String(d.company || '').trim();
    var recipient = String(d.recipient || '').trim();
    var phone     = normPhone_(d.phone);
    var address1  = String(d.address1 || '').trim();

    // 명단에 있는 회사인지 서버에서도 확인한다
    var hit = null;
    readList_().forEach(function (p) { if (p.company === company) hit = p; });
    if (!hit) return json_({ ok: false, error: '명단에 없는 협력사입니다. 담당자에게 문의해 주세요.' });

    if (!d.gift_id) return json_({ ok: false, error: '기념품을 선택해 주세요.' });
    if (recipient.length < 2) return json_({ ok: false, error: '수령인 이름을 입력해 주세요.' });
    if (!/^01[016789]-\d{3,4}-\d{4}$/.test(phone)) {
      return json_({ ok: false, error: '연락처를 휴대폰 번호 형식으로 입력해 주세요.' });
    }
    if (address1.length < 5) return json_({ ok: false, error: '주소를 입력해 주세요.' });

    var sh = dataSheet_();
    var row = findRow_(sh, company);
    var existed = !!row;
    var ts = now_();

    var values = [
      existed ? sh.getRange(row, 1).getValue() : ts,   // 최초 신청일시는 유지
      ts,
      hit.date,
      company,
      hit.position,
      hit.name,
      String(d.gift_name || d.gift_id),
      String(d.option || '').trim(),
      recipient,
      phone,
      String(d.postcode || '').trim(),
      address1,
      String(d.address2 || '').trim(),
      String(d.memo || '').trim(),
      String(d.gift_id)
    ];

    if (!row) row = Math.max(sh.getLastRow() + 1, 2);

    // 값을 넣기 전에 텍스트 서식을 잡아야 우편번호 앞자리 0 이 살아남는다
    sh.getRange(row, COL_PHONE, 1, 2).setNumberFormat('@');
    sh.getRange(row, 1, 1, values.length).setValues([values]);

    return json_({ ok: true, updated: existed });

  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}


// ---------------------------------------------------------------- 점검용
/** 편집기에서 이 함수를 실행하면 명단이 제대로 읽히는지 로그로 확인됩니다 */
function 명단확인() {
  var people = readList_();
  Logger.log('총 ' + people.length + '개 협력사');
  Logger.log(people.slice(0, 5));
}
