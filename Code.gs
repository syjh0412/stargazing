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
  CAP: 30,                                // 정원 (11/9: 처음 대기 5명까지 모두 확정해 30명)
  WAIT: 5,                                // 대기 인원
  GRADES: ['1'],                          // 정식 신청 가능 학년 — 2학년 추가 모집 때 ['1', '2'] 로 바꾸고 '새 버전' 배포
  PRE_GRADES: ['1', '2'],                 // 예비신청은 1·2학년 모두 (추가 모집 인원 가늠용)
  REFLECT_AT: '2026-11-23T20:30:00+09:00', // 관측 일지(소감) 쓰기가 열리는 시각
  STAFF_PIN: '1234',                      // 선생님 탭 비밀번호
  NOTIFY: true                            // 불참 알림이 오면 이 스크립트 주인(선생님) 메일로 알림
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
    else if(p.fn==='pre') out = preApply(p.arg||{});
    else if(p.fn==='me') out = me(p.arg||{});
    else if(p.fn==='staff') out = staff(p.arg||{});
    else if(p.fn==='absent') out = absent(p.arg||{});
    else if(p.fn==='savePre') out = savePre(p.arg||{});
    else if(p.fn==='saveRefl') out = saveRefl(p.arg||{});
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
           conf: c.conf, waiting: c.wait, full: (c.conf>=EVENT.CAP || c.wait>0) && c.wait>=EVENT.WAIT, pre: preCount_(), grades: EVENT.GRADES, reflectAt: new Date(EVENT.REFLECT_AT).getTime() };
}

/* ───────── 예비신청: 정식 신청(OPEN_AT) 전까지 인원 파악용. 선착순 순서와는 무관 ───────── */
var PRE_SHEET = '예비신청';
var PRE_HEAD = ['신청시각','학년','반','번호','이름'];
function preSh_(){
  var ss = SpreadsheetApp.getActiveSpreadsheet(), s = ss.getSheetByName(PRE_SHEET);
  if(!s){ s = ss.insertSheet(PRE_SHEET); s.appendRow(PRE_HEAD); s.setFrozenRows(1); s.getRange(1,1,1,PRE_HEAD.length).setFontWeight('bold').setBackground('#3b2a6e').setFontColor('#ffffff'); }
  return s;
}
function preRows_(){ var s = preSh_(), n = s.getLastRow(); return n<2 ? [] : s.getRange(2,1,n-1,PRE_HEAD.length).getValues(); }
function preCount_(){ return preRows_().length; }
function preApply(f){
  var lock = LockService.getScriptLock();
  if(!lock.tryLock(20000)) return {ok:false, msg:'잠시 뒤 다시 눌러 주세요.'};
  try{
    if(new Date() >= new Date(EVENT.OPEN_AT)) return {ok:false, msg:'예비신청은 끝났어요. 이제 정식 신청을 해 주세요.'};
    var g = clean_(f.grade,2), b = clean_(f.cls,3), num = clean_(f.num,3), name = clean_(f.name,20);
    if(EVENT.PRE_GRADES.indexOf(g)<0) return {ok:false, msg:'1·2학년만 예비 신청할 수 있어요.'};
    if(!/^\d{1,2}$/.test(b) || !/^\d{1,2}$/.test(num)) return {ok:false, msg:'반과 번호를 숫자로 적어 주세요.'};
    if(name.length<2) return {ok:false, msg:'이름을 정확히 적어 주세요.'};
    var r = preRows_();
    for(var i=0;i<r.length;i++){
      if(String(r[i][1])===g && String(r[i][2])===b && String(r[i][3])===num)
        return {ok:true, again:true, count:r.length};
    }
    preSh_().appendRow([new Date(), g, b, num, name]);
    return {ok:true, count:r.length+1};
  } finally { lock.releaseLock(); }
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
    if(EVENT.GRADES.indexOf(g)<0) return {ok:false, msg: EVENT.GRADES.length===1 ? '지금은 1학년만 신청할 수 있어요. 자리가 남으면 2학년 추가 모집을 따로 안내할게요.' : '1·2학년만 신청할 수 있어요.'};
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
    /* 대기자가 한 명이라도 있으면 빈자리가 생겨도 새 신청은 대기로 — 빈자리는 선생님이 대기 순서대로 연락 */
    if(c.conf < EVENT.CAP && c.wait === 0) st = '확정';
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


/* ═════════ 참가자 공간 (모집이 끝난 뒤) ═════════
   참가자 = '신청' 탭에서 상태가 '확정' 또는 '대기'인 학생 (대기자도 모두 참가)
   학생이 불참을 알리면 상태가 '불참 알림'으로 바뀌고 '불참알림' 탭에 기록, 선생님 메일로 알림 */
var ABS_SHEET = '불참알림', ABS_HEAD = ['알린시각','학년','반','번호','이름','사유','처리'];
var PREACT_SHEET = '사전활동', PREACT_HEAD = ['저장시각','학년','반','번호','이름','안전약속','가장 보고 싶은 것','이유','예측1 달 모양','예측2 토성 고리','예측3 크레이터','강사님께 질문'];
var REFL_SHEET = '관측일지', REFL_HEAD = ['저장시각','학년','반','번호','이름','본 것','관측 기록(어떻게 보였나)','예측과 실제','배운 점','느낀 점','실천할 점(관측 계획)','해시태그','별점'];

function tab_(name, head, color){
  var ss = SpreadsheetApp.getActiveSpreadsheet(), t = ss.getSheetByName(name);
  if(!t){ t = ss.insertSheet(name); t.appendRow(head); t.setFrozenRows(1); }
  /* 항목이 바뀌었으면 머리글만 새로 맞춤 */
  var cur = t.getRange(1,1,1,head.length).getValues()[0];
  if(cur.join('|') !== head.join('|')) t.getRange(1,1,1,head.length).setValues([head]);
  t.getRange(1,1,1,head.length).setFontWeight('bold').setBackground(color||'#1f2a55').setFontColor('#ffffff');
  return t;
}
function who_(f){ return {g:clean_(f.grade,2), b:clean_(f.cls,3), n:clean_(f.num,3), name:clean_(f.name,20)}; }
function findP_(w){
  var r = rows_();
  for(var i=r.length-1;i>=0;i--){
    if(String(r[i][3])===w.g && String(r[i][4])===w.b && String(r[i][5])===w.n && String(r[i][6])===w.name){
      var st = String(r[i][2]);
      if(st==='확정' || st==='대기' || st==='불참 알림') return {row:i+2, status:st, no:r[i][0]};
    }
  }
  return null;
}
function rowOf_(t, head, w){
  var n = t.getLastRow(); if(n<2) return null;
  var v = t.getRange(2,1,n-1,head.length).getValues();
  for(var i=0;i<v.length;i++) if(String(v[i][1])===w.g && String(v[i][2])===w.b && String(v[i][3])===w.n) return {row:i+2, vals:v[i]};
  return null;
}
function upsert_(t, head, w, vals){
  var line = [new Date(), w.g, w.b, w.n, w.name].concat(vals), hit = rowOf_(t, head, w);
  if(hit) t.getRange(hit.row,1,1,line.length).setValues([line]); else t.appendRow(line);
}
function me(f){
  var w = who_(f), p = findP_(w);
  if(!p) return {ok:false, msg:'참가자 명단에서 찾지 못했어요. 학년·반·번호·이름을 신청할 때와 똑같이 적어 주세요.'};
  var cc = checkCode_(w, f.code, true);
  if(cc && cc!=='new') return {ok:false, msg:cc, badCode:true};
  var pre = rowOf_(tab_(PREACT_SHEET, PREACT_HEAD, '#3b2a6e'), PREACT_HEAD, w);
  var rf = rowOf_(tab_(REFL_SHEET, REFL_HEAD, '#0f5c4d'), REFL_HEAD, w);
  var wrank = 0;
  if(p.status==='대기'){ var r = rows_(); for(var i=0;i<r.length && i+2<=p.row;i++) if(String(r[i][2])==='대기') wrank++; }
  return {ok:true, firstCode: cc==='new', name:w.name, no:p.no, absent: p.status==='불참 알림', status: p.status, waitRank: wrank,
          pre: pre ? pre.vals.slice(5) : null, refl: rf ? rf.vals.slice(5) : null,
          reflectAt: new Date(EVENT.REFLECT_AT).getTime(), now: Date.now()};
}
function absent(f){
  var lock = LockService.getScriptLock(); if(!lock.tryLock(20000)) return {ok:false, msg:'잠시 뒤 다시 눌러 주세요.'};
  try{
    var w = who_(f), p = findP_(w), why = clean_(f.reason, 200);
    if(!p) return {ok:false, msg:'참가자 명단에서 찾지 못했어요.'};
    var cc = checkCode_(w, f.code, false); if(cc) return {ok:false, msg:cc};
    if(why.length<2) return {ok:false, msg:'못 오는 까닭을 짧게 적어 주세요.'};
    sh().getRange(p.row, 3).setValue('불참 알림');
    tab_(ABS_SHEET, ABS_HEAD, '#7a2c47').appendRow([new Date(), w.g, w.b, w.n, w.name, why, '']);
    if(EVENT.NOTIFY){
      try{ MailApp.sendEmail(Session.getEffectiveUser().getEmail(), '[천체관측] 불참 알림 – '+w.g+'학년 '+w.b+'반 '+w.name,
        w.g+'학년 '+w.b+'반 '+w.n+'번 '+w.name+' 학생이 11/23 천체관측에 참여하지 못한다고 알렸습니다.\n사유: '+why+'\n\n구글 시트의 \'불참알림\' 탭과 \'신청\' 탭(상태: 불참 알림)을 확인해 주세요.'); }catch(e){}
    }
    return {ok:true};
  } finally { lock.releaseLock(); }
}
function savePre(f){
  var w = who_(f); if(!findP_(w)) return {ok:false, msg:'참가자 명단에서 찾지 못했어요.'};
  var cc = checkCode_(w, f.code, false); if(cc) return {ok:false, msg:cc};
  var a = f.a || {};
  upsert_(tab_(PREACT_SHEET, PREACT_HEAD, '#3b2a6e'), PREACT_HEAD, w,
    [a.safety?'동의':'', clean_(a.want,20), clean_(a.why,200), clean_(a.q1,20), clean_(a.q2,20), clean_(a.q3,20), clean_(a.ask,200)]);
  return {ok:true};
}
function saveRefl(f){
  var w = who_(f); if(!findP_(w)) return {ok:false, msg:'참가자 명단에서 찾지 못했어요.'};
  var cc = checkCode_(w, f.code, false); if(cc) return {ok:false, msg:cc};
  if(new Date() < new Date(EVENT.REFLECT_AT)) return {ok:false, msg:'관측 일지는 관측이 끝난 뒤(11/23 20:30)부터 쓸 수 있어요.'};
  var a = f.a || {};
  upsert_(tab_(REFL_SHEET, REFL_HEAD, '#0f5c4d'), REFL_HEAD, w,
    [clean_(a.seen,80), clean_(a.desc,500), clean_(a.diff,400), clean_(a.learn,600), clean_(a.feel,600), clean_(a.plan,400), clean_(a.tags,80), clean_(a.stars,2)]);
  return {ok:true};
}


/* ═════════ 선생님 탭: 비밀번호가 맞으면 명단·현황을 보여 줌 (보호자 연락처는 보내지 않음) ═════════ */
function staff(f){
  if(String(f.pin||'') !== String(EVENT.STAFF_PIN)) return {ok:false, msg:'비밀번호가 맞지 않아요.'};
  var r = rows_(), people = [];
  var ct = tab_(CODE_SHEET, CODE_HEAD, '#5b3a9e');
  r.forEach(function(x){ var st=String(x[2]); if(st==='확정'||st==='대기'||st==='불참 알림'){ var pw={g:String(x[3]),b:String(x[4]),n:String(x[5])}; people.push({no:x[0], st:st, g:pw.g, b:pw.b, n:pw.n, name:String(x[6]), home:String(x[9]), code: !!rowOf_(ct, CODE_HEAD, pw)}); } });
  people.sort(function(a,b){ return (Number(a.g)-Number(b.g)) || (Number(a.b)-Number(b.b)) || (Number(a.n)-Number(b.n)); });
  var pre = tab_(PREACT_SHEET, PREACT_HEAD, '#3b2a6e'), pv = pre.getLastRow()>1 ? pre.getRange(2,1,pre.getLastRow()-1,PREACT_HEAD.length).getValues() : [];
  var asks = pv.filter(function(x){ return String(x[11]).trim(); }).map(function(x){ return {who:x[1]+'-'+x[2]+' '+x[4], q:String(x[11])}; });
  var wants = {}; pv.forEach(function(x){ var w=String(x[6]); if(w) wants[w]=(wants[w]||0)+1; });
  var rf = tab_(REFL_SHEET, REFL_HEAD, '#0f5c4d'), rcount = Math.max(0, rf.getLastRow()-1);
  var ab = tab_(ABS_SHEET, ABS_HEAD, '#7a2c47'), av = ab.getLastRow()>1 ? ab.getRange(2,1,ab.getLastRow()-1,ABS_HEAD.length).getValues() : [];
  return {ok:true, people:people, preCount:pv.length, safeCount:pv.filter(function(x){return x[5]==='동의';}).length, wants:wants, asks:asks, reflCount:rcount,
          absents: av.map(function(x){ return {who:x[1]+'-'+x[2]+' '+x[4], why:String(x[5]), done:String(x[6])}; })};
}


/* ═════════ 별빛 암호: 참가자 공간에 들어갈 때 쓰는 학생별 숫자 4자리 ═════════
   처음 들어올 때 정한 암호가 '별빛암호' 탭에 (그대로가 아니라 바꾼 값으로) 저장됨.
   학생이 암호를 잊으면 그 학생 줄을 지우면 됨 → 다음에 들어올 때 새로 정함 */
var CODE_SHEET = '별빛암호', CODE_HEAD = ['정한 시각','학년','반','번호','이름','암호(변환값)'];
function hash_(w, code){
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'gocheon-star|'+w.g+'-'+w.b+'-'+w.n+'|'+code, Utilities.Charset.UTF_8);
  return raw.map(function(b){ return ('0'+(b&255).toString(16)).slice(-2); }).join('');
}
/* 맞으면 '' / 처음이면 새로 저장 후 'new' / 틀리면 오류 문구 */
function checkCode_(w, code, allowSet){
  code = String(code||'').trim();
  if(!/^\d{4}$/.test(code)) return '별빛 암호는 숫자 4자리예요.';
  var t = tab_(CODE_SHEET, CODE_HEAD, '#5b3a9e'), hit = rowOf_(t, CODE_HEAD, w);
  if(hit){ return String(hit.vals[5])===hash_(w, code) ? '' : '별빛 암호가 맞지 않아요. 잊었다면 담당 선생님께 말해 주세요.'; }
  if(!allowSet) return '먼저 참가자 공간에 들어가 별빛 암호를 정해 주세요.';
  t.appendRow([new Date(), w.g, w.b, w.n, w.name, hash_(w, code)]);
  return 'new';
}
