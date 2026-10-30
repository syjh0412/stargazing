/**
 * 고천중 가을밤 천체관측 — 신청 페이지 (Google Apps Script 하나로 페이지 + 접수)
 * ------------------------------------------------------------------
 * 설치 (5분)
 *  1) 새 구글 시트 → 확장 프로그램 → Apps Script
 *  2) Code.gs 내용을 지우고 이 파일 전체 붙여넣기
 *  3) 왼쪽 ＋ → HTML → 이름을 Index 로 → Index.html 전체 붙여넣기 → 저장
 *  4) 배포 → 새 배포 → 유형: 웹 앱 (실행: 나 / 액세스: 모든 사용자) → 배포 → 권한 허용
 *  5) 나온 웹 앱 주소(…/exec)를 가정통신문·QR로 안내
 * 학생에게 보이는 페이지는 GitHub Pages(https://syjh0412.github.io/stargazing/)에 있고,
 *   이 스크립트는 접수 창구(doPost) 역할을 합니다. 이 주소를 직접 열어도 같은 페이지가 나옵니다.
 * 신청은 시트의 '신청' 탭에 접수 순서대로 쌓이고, '상태' 칸에 확정 / 대기가 자동으로 적힙니다.
 * 취소가 생기면 그 줄의 상태를 '취소'로 바꾸세요 → 남은 자리가 다시 열립니다.
 *   (대기자를 올릴 때는 그 학생의 상태를 '확정'으로 바꾸고 연락하세요)
 */

var EVENT = {
  OPEN_AT: '2026-11-07T09:00:00+09:00',   // 신청 시작 (이 시각 전에는 버튼이 잠김)
  CLOSE_AT: '2026-11-10T18:00:00+09:00',  // 신청 마감
  CAP: 25,                                // 정원
  WAIT: 5,                                // 대기 인원
  GRADES: ['1', '2']                      // 신청 가능 학년
};
var SHEET = '신청';
var HEAD = ['접수번호','접수시각','상태','학년','반','번호','이름','신청자','보호자 연락처','귀가 방법','개인정보 동의'];

function doGet(){
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('가을밤 천체관측 신청 | 고천중학교')
    .addMetaTag('viewport','width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* GitHub Pages(syjh0412.github.io/stargazing)에서 부르는 창구 — {fn, arg}를 받아 JSON으로 답함 */
function doPost(e){
  var out;
  try{
    var p = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if(p.fn==='status') out = status();
    else if(p.fn==='apply') out = apply(p.arg||{});
    else if(p.fn==='check') out = check(p.arg||{});
    else out = {error:'알 수 없는 요청'};
  }catch(err){ out = {error:String(err)}; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function sh(){
  var ss = SpreadsheetApp.getActiveSpreadsheet(), s = ss.getSheetByName(SHEET);
  if(!s){ s = ss.insertSheet(SHEET); s.appendRow(HEAD); s.setFrozenRows(1); s.getRange(1,1,1,HEAD.length).setFontWeight('bold').setBackground('#1f2a55').setFontColor('#ffffff'); }
  return s;
}
function rows_(){ var s = sh(), n = s.getLastRow(); return n<2 ? [] : s.getRange(2,1,n-1,HEAD.length).getValues(); }
function counts_(r){
  var c = {conf:0, wait:0};
  r.forEach(function(x){ var st=String(x[2]); if(st==='확정') c.conf++; else if(st==='대기') c.wait++; });
  return c;
}

/* 페이지가 처음 열릴 때 / 30초마다: 남은 자리와 신청 가능 여부 */
function status(){
  var now = new Date(), open = new Date(EVENT.OPEN_AT), close = new Date(EVENT.CLOSE_AT), c = counts_(rows_());
  return { now: now.getTime(), openAt: open.getTime(), closeAt: close.getTime(), cap: EVENT.CAP, wait: EVENT.WAIT,
           conf: c.conf, waiting: c.wait, full: c.conf>=EVENT.CAP && c.wait>=EVENT.WAIT };
}

function clean_(v, n){ return String(v==null?'':v).replace(/[<>]/g,'').trim().slice(0, n||40); }

/* 신청 — 잠금(LockService)을 걸어 동시에 눌러도 순서가 꼬이지 않게 */
function apply(f){
  var lock = LockService.getScriptLock();
  if(!lock.tryLock(20000)) return {ok:false, msg:'신청이 몰리고 있어요. 잠시 뒤 다시 눌러 주세요.'};
  try{
    var now = new Date();
    if(now < new Date(EVENT.OPEN_AT)) return {ok:false, msg:'아직 신청 시간이 아니에요. 11월 7일(토) 오전 9시에 열립니다.'};
    if(now > new Date(EVENT.CLOSE_AT)) return {ok:false, msg:'신청이 마감되었어요.'};
    var g = clean_(f.grade,2), b = clean_(f.cls,3), num = clean_(f.num,3), name = clean_(f.name,20);
    var who = clean_(f.who,10), phone = clean_(f.phone,20), home = clean_(f.home,20);
    if(EVENT.GRADES.indexOf(g)<0) return {ok:false, msg:'1·2학년만 신청할 수 있어요.'};
    if(!/^\d{1,2}$/.test(b) || !/^\d{1,2}$/.test(num)) return {ok:false, msg:'반과 번호를 숫자로 적어 주세요.'};
    if(name.length<2) return {ok:false, msg:'이름을 정확히 적어 주세요.'};
    if(!/^[0-9\-\s]{9,14}$/.test(phone)) return {ok:false, msg:'보호자 연락처를 확인해 주세요. (예: 010-1234-5678)'};
    if(f.agree!==true) return {ok:false, msg:'개인정보 수집·이용에 동의해야 신청할 수 있어요.'};
    if(home!=='보호자 동행' && home!=='단독 귀가(보호자 동의)') return {ok:false, msg:'귀가 방법을 골라 주세요.'};

    var r = rows_();
    for(var i=0;i<r.length;i++){
      if(String(r[i][3])===g && String(r[i][4])===b && String(r[i][5])===num && String(r[i][2])!=='취소')
        return {ok:false, dup:true, msg:'이미 신청되어 있어요. (접수번호 '+r[i][0]+', '+r[i][2]+') 아래 \'신청 확인\'에서 볼 수 있어요.'};
    }
    var c = counts_(r), st;
    if(c.conf < EVENT.CAP) st = '확정';
    else if(c.wait < EVENT.WAIT) st = '대기';
    else return {ok:false, msg:'정원과 대기 인원이 모두 찼어요. 관심 가져 주셔서 고맙습니다!'};
    var no = r.length + 1;
    sh().appendRow([no, now, st, g, b, num, name, who, "'"+phone, home, '동의']);
    return {ok:true, no:no, status:st, rank: st==='확정' ? c.conf+1 : c.wait+1, name:name};
  } finally { lock.releaseLock(); }
}

/* 신청 확인 — 학년·반·번호·이름이 모두 맞아야 보여 줌 */
function check(f){
  var g = clean_(f.grade,2), b = clean_(f.cls,3), num = clean_(f.num,3), name = clean_(f.name,20), r = rows_();
  for(var i=r.length-1;i>=0;i--){
    if(String(r[i][3])===g && String(r[i][4])===b && String(r[i][5])===num && String(r[i][6])===name)
      return {ok:true, no:r[i][0], status:String(r[i][2]), name:name};
  }
  return {ok:false, msg:'신청 기록을 찾지 못했어요. 학년·반·번호·이름을 다시 확인해 주세요.'};
}
