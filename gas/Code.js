/**
 * 일정콕 — 업무 일정 관리 (구글 시트 + 구글 캘린더)
 * 화면(index.html)은 GitHub Pages에 있고, 이 스크립트는 웹 앱(API)으로 배포한다.
 *
 * 처음 한 번: 편집기에서 setup() 실행 → 권한 허용 → 실행 로그의 '연결 링크'를 휴대폰·PC에서 한 번 연다.
 * 시트를 직접 고쳤으면: 시트 메뉴 [일정콕] → [캘린더와 다시 맞추기].
 */
const SHEET = '일정';
const CAL_NAME = '업무일정(일정콕)';
const TZ = 'Asia/Seoul';
const PAGE = 'https://j333333766-hub.github.io/iljeong-kok/';
const HEAD = ['id', '날짜', '끝날짜', '시작', '끝', '종일', '제목', '분류', '장소', '메모', '상태', '캘린더ID', '등록', '수정'];
const C = Object.fromEntries(HEAD.map((h, i) => [h, i]));
const CATS = ['수업', '회의', '마감', '행사', '출장', '기타'];

// ---------------- 처음 설정 ----------------
function setup() {
  const p = PropertiesService.getScriptProperties();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  p.setProperty('SS_ID', ss.getId());
  let sh = ss.getSheetByName(SHEET);
  if (!sh) { sh = ss.getSheets()[0]; sh.setName(SHEET); }
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold').setBackground('#e8f2fb');
    sh.setFrozenRows(1);
    sh.getRange('A:F').setNumberFormat('@');           // 날짜·시각을 글자로 두어 시트가 멋대로 바꾸지 않게
    sh.getRange('K:N').setNumberFormat('@');
    sh.setColumnWidth(C['제목'] + 1, 240); sh.setColumnWidth(C['메모'] + 1, 240);
    sh.getRange(2, C['분류'] + 1, 999, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(CATS, true).setAllowInvalid(true).build());
    sh.getRange(2, C['상태'] + 1, 999, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['예정', '완료'], true).setAllowInvalid(true).build());
  }
  let calId = p.getProperty('CAL_ID');
  if (!calId || !CalendarApp.getCalendarById(calId)) {
    const found = CalendarApp.getCalendarsByName(CAL_NAME)[0];
    const cal = found || CalendarApp.createCalendar(CAL_NAME, { color: CalendarApp.Color.BLUE, timeZone: TZ });
    p.setProperty('CAL_ID', cal.getId());
  }
  if (!p.getProperty('KEY')) p.setProperty('KEY', Utilities.getUuid().replace(/-/g, '').slice(0, 10));
  Logger.log('✅ 준비 끝. 시트: %s', ss.getUrl());
  Logger.log('캘린더: %s', CAL_NAME);
  Logger.log('다음: [배포] → [새 배포] → 웹 앱(실행: 나, 액세스: 모든 사용자) 뒤 linkInfo()를 실행하세요.');
}

/** 배포한 뒤 실행: 화면과 이 스크립트를 잇는 연결 링크를 로그에 적는다(링크에 접속 키가 들어 있으니 남에게 보내지 말 것). */
function linkInfo() {
  const url = ScriptApp.getService().getUrl();
  if (!url) { Logger.log('아직 웹 앱으로 배포되지 않았습니다.'); return; }
  const key = PropertiesService.getScriptProperties().getProperty('KEY');
  const code = Utilities.base64EncodeWebSafe(JSON.stringify({ u: url, k: key }));
  Logger.log('연결 링크(한 번만 열면 그 기기에 저장됨):\n%s#connect=%s', PAGE, code);
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('일정콕')
    .addItem('캘린더와 다시 맞추기', 'syncAll')
    .addItem('화면 열기 안내', 'showLink')
    .addToUi();
}
function showLink() { SpreadsheetApp.getUi().alert('일정콕 화면', PAGE + '\n\n처음 기기라면 편집기에서 linkInfo()를 실행해 나온 연결 링크로 여세요.', SpreadsheetApp.getUi().ButtonSet.OK); }

// ---------------- 웹 앱(API) ----------------
function doGet(e) { return handle_(e.parameter || {}); }
function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: '요청 형식이 잘못되었습니다.' }); }
  return handle_(body);
}
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function handle_(q) {
  try {
    const key = PropertiesService.getScriptProperties().getProperty('KEY');
    if (!key || q.key !== key) return out_({ ok: false, error: '접속 키가 맞지 않습니다.' });
    switch (q.action) {
      case 'ping': return out_({ ok: true, cal: CAL_NAME, cats: CATS });
      case 'list': return out_({ ok: true, items: list_(q.from, q.to) });
      case 'add': return out_({ ok: true, item: save_(q.item, true) });
      case 'update': return out_({ ok: true, item: save_(q.item, false) });
      case 'done': return out_({ ok: true, item: setDone_(q.id, !!q.done) });
      case 'remove': remove_(q.id); return out_({ ok: true });
      case 'brief': return out_({ ok: true, text: brief_(q.date) });
      default: return out_({ ok: false, error: '알 수 없는 요청: ' + q.action });
    }
  } catch (err) {
    return out_({ ok: false, error: String(err && err.message || err) });
  }
}

// ---------------- 시트 ----------------
function sheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('SS_ID');
  return SpreadsheetApp.openById(id).getSheetByName(SHEET);
}
function rows_() {
  const sh = sheet_(), n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, HEAD.length).getDisplayValues().map((r, i) => ({ r, row: i + 2 }));
}
function toItem_(r) {
  return { id: r[C.id], date: r[C['날짜']], endDate: r[C['끝날짜']], start: r[C['시작']], end: r[C['끝']], allDay: r[C['종일']] === 'Y',
    title: r[C['제목']], cat: r[C['분류']], place: r[C['장소']], memo: r[C['메모']], status: r[C['상태']] || '예정', eventId: r[C['캘린더ID']] };
}
function list_(from, to) {
  return rows_().map(x => toItem_(x.r)).filter(it => it.id && it.date &&
    (!to || it.date <= to) && (!from || (it.endDate || it.date) >= from))
    .sort((a, b) => (a.date + (a.allDay ? '00:00' : a.start)).localeCompare(b.date + (b.allDay ? '00:00' : b.start)));
}
const now_ = () => Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd HH:mm");
function clean_(it) {
  const s = v => String(v == null ? '' : v).trim();
  const o = { title: s(it.title).slice(0, 120), date: s(it.date), endDate: s(it.endDate), start: s(it.start), end: s(it.end), allDay: !!it.allDay,
    cat: CATS.includes(s(it.cat)) ? s(it.cat) : '기타', place: s(it.place).slice(0, 120), memo: s(it.memo).slice(0, 1000) };
  if (!o.title) throw new Error('제목을 적어 주세요.');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(o.date)) throw new Error('날짜 형식이 잘못되었습니다.');
  if (o.endDate && (!/^\d{4}-\d{2}-\d{2}$/.test(o.endDate) || o.endDate < o.date)) o.endDate = '';
  if (o.endDate === o.date) o.endDate = '';
  if (!o.allDay) {
    if (!/^\d{2}:\d{2}$/.test(o.start)) o.allDay = true;
    else if (!/^\d{2}:\d{2}$/.test(o.end) || (!o.endDate && o.end <= o.start)) o.end = addMin_(o.start, 60);
  }
  if (o.allDay) { o.start = ''; o.end = ''; }
  return o;
}
function addMin_(hm, m) { const [h, mi] = hm.split(':').map(Number); const t = Math.min(h * 60 + mi + m, 23 * 60 + 59); return ('0' + Math.floor(t / 60)).slice(-2) + ':' + ('0' + t % 60).slice(-2); }

function save_(raw, isNew) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const it = clean_(raw), sh = sheet_();
    let row = 0, old = null;
    if (!isNew) { const f = rows_().find(x => x.r[C.id] === raw.id); if (!f) throw new Error('일정을 찾지 못했습니다.'); row = f.row; old = toItem_(f.r); }
    it.id = isNew ? 'S' + Utilities.formatDate(new Date(), TZ, 'yyMMddHHmmss') + Math.floor(Math.random() * 90 + 10) : raw.id;
    it.status = old ? old.status : '예정';
    it.eventId = syncEvent_(it, old && old.eventId);
    const vals = [it.id, it.date, it.endDate, it.start, it.end, it.allDay ? 'Y' : '', it.title, it.cat, it.place, it.memo, it.status, it.eventId, old ? (sh.getRange(row, C['등록'] + 1).getDisplayValue()) : now_(), now_()];
    if (isNew) sh.appendRow(vals); else sh.getRange(row, 1, 1, HEAD.length).setValues([vals]);
    return it;
  } finally { lock.releaseLock(); }
}
function setDone_(id, done) {
  const f = rows_().find(x => x.r[C.id] === id); if (!f) throw new Error('일정을 찾지 못했습니다.');
  const it = toItem_(f.r); it.status = done ? '완료' : '예정';
  const sh = sheet_(); sh.getRange(f.row, C['상태'] + 1).setValue(it.status); sh.getRange(f.row, C['수정'] + 1).setValue(now_());
  const ev = event_(it.eventId); if (ev) ev.setTitle(calTitle_(it));
  return it;
}
function remove_(id) {
  const f = rows_().find(x => x.r[C.id] === id); if (!f) return;
  const ev = event_(f.r[C['캘린더ID']]); if (ev) ev.deleteEvent();
  sheet_().deleteRow(f.row);
}

// ---------------- 캘린더 ----------------
function cal_() { return CalendarApp.getCalendarById(PropertiesService.getScriptProperties().getProperty('CAL_ID')); }
function event_(id) { if (!id) return null; try { return cal_().getEventById(id); } catch (e) { return null; } }
function d_(ymd, hm) { const [y, m, d] = ymd.split('-').map(Number); const [h, mi] = (hm || '00:00').split(':').map(Number); return new Date(y, m - 1, d, h, mi); }
function calTitle_(it) { return (it.status === '완료' ? '✅ ' : '') + (it.cat && it.cat !== '기타' ? '[' + it.cat + '] ' : '') + it.title; }
/** 시트의 일정 하나를 캘린더에 맞춘다(있으면 지우고 다시 만든다). 새 캘린더 일정 ID를 돌려준다. */
function syncEvent_(it, oldId) {
  const old = event_(oldId); if (old) old.deleteEvent();
  const cal = cal_(), opt = { location: it.place || '', description: (it.memo ? it.memo + '\n\n' : '') + '— 일정콕에서 등록' };
  let ev;
  if (it.allDay) {
    const end = d_(it.endDate || it.date); end.setDate(end.getDate() + 1);   // 종일 일정의 끝날은 '다음 날'로 넘겨야 한다
    ev = (it.endDate ? cal.createAllDayEvent(calTitle_(it), d_(it.date), end, opt) : cal.createAllDayEvent(calTitle_(it), d_(it.date), opt));
  } else {
    ev = cal.createEvent(calTitle_(it), d_(it.date, it.start), d_(it.endDate || it.date, it.end), opt);
  }
  if (it.cat === '마감') ev.setColor(CalendarApp.EventColor.RED);
  return ev.getId();
}
/** 시트를 직접 고쳤을 때: 시트 내용대로 캘린더를 다시 만든다(id가 비어 있으면 새로 붙인다). */
function syncAll() {
  const sh = sheet_(); let n = 0;
  rows_().forEach(({ r, row }) => {
    if (!r[C['제목']] || !r[C['날짜']]) return;
    let it;
    try { it = clean_(toItem_(r)); } catch (e) { return; }
    it.id = r[C.id] || 'S' + Utilities.formatDate(new Date(), TZ, 'yyMMddHHmmss') + Math.floor(Math.random() * 90 + 10);
    it.status = r[C['상태']] || '예정';
    it.eventId = syncEvent_(it, r[C['캘린더ID']]);
    sh.getRange(row, 1, 1, HEAD.length).setValues([[it.id, it.date, it.endDate, it.start, it.end, it.allDay ? 'Y' : '', it.title, it.cat, it.place, it.memo, it.status, it.eventId, r[C['등록']] || now_(), now_()]]);
    n++;
  });
  try { SpreadsheetApp.getActive().toast(n + '건을 캘린더와 맞췄습니다.', '일정콕'); } catch (e) { Logger.log(n + '건'); }
}

// ---------------- 브리핑(화면의 '오늘 브리핑'과 같은 글) ----------------
function brief_(date) {
  const today = date || Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  const t = new Date(d_(today)); t.setDate(t.getDate() + 7);
  const week = Utilities.formatDate(t, TZ, 'yyyy-MM-dd');
  const all = list_(today, week);
  const on = it => it.date <= today && (it.endDate || it.date) >= today;
  const todays = all.filter(on), dues = all.filter(it => it.cat === '마감' && !on(it) && it.status !== '완료');
  const line = it => '· ' + (it.allDay ? '종일' : it.start + '~' + it.end) + ' ' + it.title + (it.place ? ' (' + it.place + ')' : '') + (it.status === '완료' ? ' ✅' : '');
  const dd = it => { const n = Math.round((d_(it.date) - d_(today)) / 864e5); return '· D-' + n + ' ' + it.title + ' (' + it.date.slice(5).replace('-', '/') + ')'; };
  return ['📅 ' + today + ' 오늘 일정 ' + todays.length + '건', ...(todays.length ? todays.map(line) : ['· 없음']),
    '', '⏰ 다가오는 마감(7일)', ...(dues.length ? dues.map(dd) : ['· 없음'])].join('\n');
}
