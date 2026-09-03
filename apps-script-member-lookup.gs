const ACCESS_TOKEN = '여기에_관리프로그램과_같은_토큰을_입력하세요';
const FOLDER_NAME = '부천세무사회_회비관리_자동저장';
const LOOKUP_FILE_NAME = '회원조회자료.json';
const MAX_FAILURES = 3;
const LOCK_PREFIX = 'LOOKUP_LOCK_';

function doPost(e) {
  try {
    const data = JSON.parse((e.postData && e.postData.contents) || '{}');
    if (!data.token || data.token !== ACCESS_TOKEN) return jsonResponse({ ok: false, error: '인증에 실패했습니다.' });
    if (data.action === 'listLookupLocks') return jsonResponse({ ok: true, locks: listLookupLocks() });
    if (data.action === 'unlockLookupLock') return jsonResponse({ ok: unlockLookupLock(data.lockId) });
    if (data.action === 'unlockUnknownLocks') return jsonResponse({ ok: true, removed: unlockUnknownLocks() });
    if (data.action === 'unlockAllLookupLocks') return jsonResponse({ ok: true, removed: unlockAllLookupLocks() });
    if (!data.xlsxBase64 || !data.lookupData) return jsonResponse({ ok: false, error: '저장할 자료가 없습니다.' });

    const folder = getOrCreateFolder();
    const filename = safeFilename(data.filename || '회비미납현황.xlsx');
    const bytes = Utilities.base64Decode(data.xlsxBase64);
    const blob = Utilities.newBlob(bytes, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', filename);
    folder.createFile(blob);
    saveTextFile(folder, LOOKUP_FILE_NAME, JSON.stringify(data.lookupData));

    return jsonResponse({ ok: true, filename: filename, savedAt: new Date().toISOString() });
  } catch (error) {
    return jsonResponse({ ok: false, error: '저장 처리 중 오류가 발생했습니다.' });
  }
}

function doGet() {
  return HtmlService.createHtmlOutput(memberLookupHtmlV2())
    .setTitle('부천지역세무사회 회비 납입내역 조회')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function lookupMember(regNo, phoneLast4, clientId) {
  const clientKey = lookupLockKey(clientId);
  const currentLock = readLookupLock(clientKey);
  if (currentLock && currentLock.locked) return lockedLookupResponse(currentLock);
  const reg = digits(regNo);
  const last4 = digits(phoneLast4);
  if (reg.length < 4 || last4.length !== 4) return recordLookupFailure(clientKey, reg, null);

  const data = readLookupData();
  if (!data || !Array.isArray(data.members)) return { ok: false, noData: true };
  const memberByReg = data.members.find(function (item) { return digits(item.regNo) === reg; });
  const member = memberByReg && digits(memberByReg.phoneLast4) === last4 ? memberByReg : null;
  if (!member) return recordLookupFailure(clientKey, reg, memberByReg || null);
  deleteLookupLock(clientKey);
  return {
    ok: true,
    name: String(member.name || ''),
    generatedAt: String(data.generatedAt || ''),
    bankLastTransactionAt: String(data.bankLastTransactionAt || ''),
    totalArrears: Number(member.totalArrears || 0),
    years: (member.years || []).map(function (item) {
      return {
        year: item.year,
        label: String(item.label || ''),
        required: Number(item.required || 0),
        allocated: Number(item.allocated || 0),
        remaining: Number(item.remaining || 0),
        status: String(item.status || ''),
        payments: (item.payments || []).map(function (payment) {
          return {
            date: String(payment.date || ''),
            time: String(payment.time || ''),
            fullAmount: Number(payment.fullAmount || 0),
            appliedAmount: Number(payment.appliedAmount || 0)
          };
        })
      };
    }),
    payments: (member.payments || []).map(function (payment) {
      return { date: payment.date, time: payment.time, amount: Number(payment.amount || 0) };
    })
  };
}

function getLookupLockStatus(clientId) {
  const lock = readLookupLock(lookupLockKey(clientId));
  return lock && lock.locked ? lockedLookupResponse(lock) : { ok: true, locked: false };
}

function recordLookupFailure(key, reg, member) {
  Utilities.sleep(350);
  const previous = readLookupLock(key) || {};
  const failures = Number(previous.failures || 0) + 1;
  const now = new Date().toISOString();
  const record = {
    lockId: key.replace(LOCK_PREFIX, ''), failures: failures, locked: failures >= MAX_FAILURES,
    regNo: reg, memberName: member ? String(member.name || '') : '',
    updatedAt: now, lockedAt: failures >= MAX_FAILURES ? now : ''
  };
  writeLookupLock(key, record);
  if (record.locked) return lockedLookupResponse(record);
  return { ok: false, mismatch: true, failures: failures, remainingAttempts: MAX_FAILURES - failures };
}

function lockedLookupResponse(lock) {
  return { ok: false, locked: true, failures: Number(lock.failures || MAX_FAILURES) };
}

function lookupLockKey(clientId) {
  return LOCK_PREFIX + hashText(String(clientId || 'missing-client'));
}

function readLookupLock(key) {
  try {
    const value = PropertiesService.getScriptProperties().getProperty(key);
    return value ? JSON.parse(value) : null;
  } catch (error) { return null; }
}

function writeLookupLock(key, value) {
  PropertiesService.getScriptProperties().setProperty(key, JSON.stringify(value));
}

function deleteLookupLock(key) {
  PropertiesService.getScriptProperties().deleteProperty(key);
}

function listLookupLocks() {
  const properties = PropertiesService.getScriptProperties().getProperties();
  return Object.keys(properties).filter(function (key) { return key.indexOf(LOCK_PREFIX) === 0; }).map(function (key) {
    try { return JSON.parse(properties[key]); } catch (error) { return null; }
  }).filter(function (item) { return item && item.locked; }).map(function (item) {
    return {
      lockId: String(item.lockId || ''), memberName: String(item.memberName || ''),
      maskedRegNo: maskLookupRegNo(item.regNo), failures: Number(item.failures || 0),
      lockedAt: String(item.lockedAt || item.updatedAt || ''), knownMember: !!item.memberName
    };
  }).sort(function (a, b) { return b.lockedAt.localeCompare(a.lockedAt); });
}

function unlockLookupLock(lockId) {
  const id = String(lockId || '').replace(/[^a-f0-9]/gi, '').slice(0, 32);
  if (!id) return false;
  PropertiesService.getScriptProperties().deleteProperty(LOCK_PREFIX + id);
  return true;
}

function unlockUnknownLocks() {
  return deleteLookupLocks(function (item) { return !item.memberName; });
}

function unlockAllLookupLocks() {
  return deleteLookupLocks(function () { return true; });
}

function deleteLookupLocks(predicate) {
  const store = PropertiesService.getScriptProperties();
  const properties = store.getProperties();
  const keys = [];
  Object.keys(properties).forEach(function (key) {
    if (key.indexOf(LOCK_PREFIX) !== 0) return;
    let item = null;
    try { item = JSON.parse(properties[key]); } catch (error) {}
    if (item && item.locked && predicate(item)) keys.push(key);
  });
  keys.forEach(function (key) { store.deleteProperty(key); });
  return keys.length;
}

function maskLookupRegNo(value) {
  const reg = digits(value);
  if (!reg) return '확인 불가';
  if (reg.length <= 2) return reg;
  return reg.slice(0, 2) + '••' + reg.slice(-2);
}

function readLookupData() {
  const folder = getOrCreateFolder();
  const files = folder.getFilesByName(LOOKUP_FILE_NAME);
  if (!files.hasNext()) return null;
  return JSON.parse(files.next().getBlob().getDataAsString('UTF-8'));
}

function getOrCreateFolder() {
  const folders = DriveApp.getFoldersByName(FOLDER_NAME);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(FOLDER_NAME);
}

function saveTextFile(folder, name, content) {
  const files = folder.getFilesByName(name);
  if (files.hasNext()) {
    files.next().setContent(content);
  } else {
    folder.createFile(name, content, MimeType.PLAIN_TEXT);
  }
}

function safeFilename(value) {
  return String(value).replace(/[\\/:*?"<>|]/g, '_');
}

function digits(value) {
  return String(value == null ? '' : value).replace(/[^0-9]/g, '');
}

function hashText(value) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, value, Utilities.Charset.UTF_8);
  return bytes.map(function (byte) { return ('0' + ((byte + 256) % 256).toString(16)).slice(-2); }).join('').slice(0, 32);
}

function jsonResponse(value) {
  return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON);
}

function memberLookupHtml() {
  return `<!doctype html>
<html lang="ko"><head><meta charset="utf-8"><link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/pretendard@1.3.9/dist/web/static/pretendard.css"><style>
*{box-sizing:border-box}body{margin:0;background:#edf2f7;color:#172840;font-family:Pretendard,-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif}.wrap{width:min(100%,440px);margin:0 auto;padding:22px 10px}.shell{overflow:hidden;border:1px solid rgba(31,61,96,.12);border-radius:28px;background:linear-gradient(180deg,#f9fbfd 0,#eef3f8 100%);box-shadow:0 18px 50px rgba(28,52,83,.13),0 3px 10px rgba(28,52,83,.06)}.header{display:flex;align-items:center;gap:12px;padding:26px;color:#fff;background:linear-gradient(135deg,#102f57,#173f72)}.mark{display:grid;place-items:center;width:48px;height:48px;flex:0 0 48px;border:1px solid rgba(255,255,255,.22);border-radius:14px;background:rgba(255,255,255,.14);font-size:16px;font-weight:600}.org{margin:0;color:#fff;font-size:18px;font-weight:600;letter-spacing:-.02em}.service{margin:4px 0 0;color:#dce8f6;font-size:14px}.intro{padding:34px 26px 26px}.intro h1{margin:0 0 14px;font-size:30px;font-weight:600;letter-spacing:-.04em}.intro p{margin:0;color:#56677d;font-size:17px;line-height:1.7;word-break:keep-all}.lookup-panel,.summary,.year-card{margin:0 16px 16px;border:1px solid #d8e1eb;border-radius:22px;background:#fff;box-shadow:0 10px 28px rgba(34,59,89,.08)}.lookup-panel{padding:28px 22px 24px}.lookup-panel h2{margin:0 0 24px;font-size:21px;font-weight:600}.field{margin-bottom:21px}label{display:block;margin:0 0 10px;color:#283b53;font-size:17px;font-weight:600}input{width:100%;height:64px;border:1px solid #cbd6e2;border-radius:14px;padding:0 18px;background:#f8fafc;color:#172840;font:inherit;font-size:18px}input::placeholder{color:#8997a8}.primary,.secondary{width:100%;height:64px;border-radius:14px;font:inherit;font-size:18px;font-weight:600;cursor:pointer}.primary{border:0;background:linear-gradient(180deg,#173f72,#102f57);color:#fff;box-shadow:0 8px 18px rgba(20,58,104,.22)}.primary:disabled{opacity:.55}.secondary{border:1px solid #c8d4e1;background:#f9fbfd;color:#173f72}.msg{display:none;margin:0 16px 16px;padding:15px 16px;border-radius:14px;font-size:15px;line-height:1.55;word-break:keep-all}.err{display:block;background:#fff1f2;color:#a33a36}.privacy{display:flex;gap:10px;align-items:flex-start;margin:18px 16px 0;padding:17px 18px;border-radius:14px;background:#e9f1fa;color:#465c75;font-size:14px;line-height:1.55}.privacy span:last-child{word-break:keep-all}.foot{padding:0 24px 28px;text-align:center}.contact{margin:20px 0 0;color:#173f72;font-size:15px;font-weight:600;word-break:keep-all}.result{display:none}.greeting{padding:30px 25px 22px}.greeting h1{margin:0 0 12px;font-size:27px;font-weight:600;letter-spacing:-.04em}.greeting p{margin:0;color:#52657c;font-size:16px;line-height:1.75;word-break:keep-all}.summary{padding:25px 22px}.summary-label{margin:0;color:#50647c;font-size:16px;font-weight:600}.summary-amount{margin:8px 0 12px;color:#a33a36;font-size:36px;font-weight:600;letter-spacing:-.04em}.summary-amount.paid{color:#18704c}.summary-note{margin:0;color:#697b8e;font-size:15px;line-height:1.6;word-break:keep-all}.years{padding:4px 16px 8px}.years h2{margin:0 4px 16px;font-size:21px;font-weight:600}.year-card{margin-left:0;margin-right:0;padding:21px 20px;border-radius:18px;box-shadow:0 6px 18px rgba(37,62,91,.06)}.year-top{display:flex;align-items:center;justify-content:space-between;gap:12px}.year-title{margin:0;font-size:19px;font-weight:600}.status{flex:0 0 auto;padding:6px 10px;border-radius:999px;font-size:14px;font-weight:600}.unpaid{color:#973b38;background:#f9e9e7}.partial{color:#765213;background:#f9f0d8}.paid-status{color:#18704c;background:#e7f4ed}.remaining{margin:16px 0 0;font-size:18px}.remaining strong{color:#a33a36}.details{display:grid;grid-template-columns:1fr auto;gap:10px 14px;margin-top:17px;padding-top:16px;border-top:1px solid #e1e7ee}.details span{color:#617287;font-size:15px}.details .value{color:#2b3f56;text-align:right}.basis{margin:3px 0 18px;color:#718093;font-size:14px;line-height:1.6;word-break:keep-all}.result-foot{padding:12px 24px 27px;text-align:center}.empty{margin:0 16px 18px;padding:24px;border-radius:18px;background:#fff;color:#617287;text-align:center;line-height:1.65;word-break:keep-all}@media(max-width:380px){.wrap{padding-inline:4px}.header,.intro{padding-left:20px;padding-right:20px}.lookup-panel,.summary{margin-left:12px;margin-right:12px}.years{padding-left:12px;padding-right:12px}}
</style></head><body><div class="wrap"><main class="shell"><header class="header"><div class="mark">부천</div><div><p class="org">부천지역세무사회</p><p class="service">회비 납부내역 조회 서비스</p></div></header><section id="lookup" class="lookup"><div class="intro"><h1>회비 납부내역 조회</h1><p>등록번호와 휴대전화 번호 뒤 4자리를 입력해 주세요.</p></div><form id="form" class="lookup-panel"><h2>조회정보를 입력해 주세요</h2><div class="field"><label for="reg">등록번호</label><input id="reg" inputmode="numeric" autocomplete="off" placeholder="등록번호 입력" required></div><div class="field"><label for="phone">휴대전화 번호 뒤 4자리</label><input id="phone" inputmode="numeric" maxlength="4" autocomplete="off" placeholder="숫자 4자리 입력" required></div><button id="submit" class="primary" type="submit">납부내역 조회하기</button></form><div id="msg" class="msg" role="alert"></div><div class="foot"><div class="privacy"><span>✓</span><span>입력하신 정보는 회비 납부내역 조회에만 사용됩니다.</span></div><p class="contact">문의 · 재무간사 김선영 세무사</p></div></section><section id="result" class="result"><div class="greeting"><h1 id="member-name"></h1><p>납부해 주신 회비는 지역세무사회 활동에 귀하게 잘 쓰이고 있습니다. 지역회 활동에 적극적으로 동참해 주셔서 감사합니다.</p></div><section class="summary"><p class="summary-label">현재 총 미납액</p><p id="total" class="summary-amount"></p><p id="summary-note" class="summary-note"></p></section><section class="years"><h2 id="year-heading">연도별 미납 현황</h2><div id="year-list"></div></section><div class="result-foot"><p id="basis" class="basis"></p><button id="again" class="secondary" type="button">다시 조회하기</button><p class="contact">문의 · 재무간사 김선영 세무사</p></div></section></main></div><script>
const form=document.getElementById('form'),button=document.getElementById('submit'),msg=document.getElementById('msg'),lookup=document.getElementById('lookup'),result=document.getElementById('result'),phone=document.getElementById('phone');
function won(value){return Number(value||0).toLocaleString('ko-KR')+'원'}
function esc(value){return String(value==null?'':value).replace(/[&<>"']/g,function(ch){return({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'})[ch]})}
function basisText(value){if(!value)return '통장 거래내역 기준 · 확인 가능한 마지막 거래 시각 없음';var text=String(value).replace('T',' ').replace('Z','').slice(0,19);var match=text.match(/^(\d{4})-(\d{2})-(\d{2})\s+(\d{2}):(\d{2})/);if(!match)return '통장 거래내역 기준 · '+text+'까지';var hour=Number(match[4]),minute=Number(match[5]),ampm=hour<12?'오전':'오후',displayHour=hour%12||12;return '통장 거래내역 기준 · '+Number(match[1])+'. '+Number(match[2])+'. '+Number(match[3])+'. '+ampm+' '+displayHour+'시'+(minute?' '+minute+'분':'')+'까지'}
function paymentDetails(item){var payments=item.payments||[];if(!payments.length)return'';return payments.map(function(payment){return '<span>입금일</span><span class="value">'+esc(payment.date)+(payment.time?' '+esc(payment.time):'')+'</span><span>해당 연도 충당액</span><span class="value">'+won(payment.appliedAmount)+'</span>'}).join('')}
function yearCard(item){var status=item.status==='paid'?'완납':item.status==='partial'?'일부납':'미납';var statusClass=item.status==='paid'?'paid-status':item.status;var headline=item.status==='paid'?'납부 완료':item.status==='partial'?'남은 미납액':'미납액';var amount=item.status==='paid'?won(item.required):won(item.remaining);var details='<div class="details"><span>'+esc(item.label)+' 회비</span><span class="value">'+won(item.required)+'</span>';if(item.status!=='unpaid')details+='<span>납부해 주신 금액</span><span class="value">'+won(item.allocated)+'</span>';details+=paymentDetails(item)+'</div>';return '<article class="year-card"><div class="year-top"><p class="year-title">'+esc(item.label)+' 회비</p><span class="status '+statusClass+'">'+status+'</span></div><p class="remaining">'+headline+' <strong>'+amount+'</strong></p>'+details+'</article>'}
function renderResult(data){var years=(data.years||[]).filter(function(item){return Number(item.required||0)>0}),arrears=Number(data.totalArrears||0);document.getElementById('member-name').textContent=data.name+' 세무사님';var total=document.getElementById('total'),note=document.getElementById('summary-note'),heading=document.getElementById('year-heading'),list=document.getElementById('year-list');total.textContent=won(arrears);total.className='summary-amount'+(arrears===0?' paid':'');if(arrears>0){var unpaidYears=years.filter(function(item){return Number(item.remaining||0)>0});note.textContent='현재 확인되는 연도별 미납 회비 내역입니다. 총 '+unpaidYears.length+'개 연도입니다.';heading.textContent='연도별 미납 현황';var paidYears=years.filter(function(item){return Number(item.remaining||0)===0});list.innerHTML=unpaidYears.map(yearCard).join('')+(paidYears.length?'<h2 style="margin-top:28px">완납 연도</h2>'+paidYears.map(yearCard).join(''):'')}else if(years.length){note.textContent='현재 확인되는 미납 회비는 없습니다.';heading.textContent='연도별 회비 납입내역';list.innerHTML=years.map(yearCard).join('')}else{note.textContent='현재 확인할 수 있는 연도별 회비 내역이 없습니다.';heading.textContent='회비 내역';list.innerHTML='<div class="empty">자세한 사항은 부천지역세무사회로 문의해 주세요.</div>'}document.getElementById('basis').textContent=basisText(data.bankLastTransactionAt);lookup.style.display='none';result.style.display='block';window.scrollTo(0,0)}
phone.addEventListener('input',function(){phone.value=phone.value.replace(/\D/g,'').slice(0,4)});
form.addEventListener('submit',function(e){e.preventDefault();button.disabled=true;button.textContent='조회 중…';msg.className='msg';google.script.run.withSuccessHandler(function(data){button.disabled=false;button.textContent='납부내역 조회하기';if(!data||!data.ok){msg.textContent=(data&&data.error)||'조회하지 못했습니다.';msg.className='msg err';return}renderResult(data)}).withFailureHandler(function(){button.disabled=false;button.textContent='납부내역 조회하기';msg.textContent='조회 중 오류가 발생했습니다. 잠시 후 다시 시도해 주세요.';msg.className='msg err'}).lookupMember(document.getElementById('reg').value,phone.value)});
document.getElementById('again').addEventListener('click',function(){result.style.display='none';lookup.style.display='block';msg.className='msg';form.reset();window.scrollTo(0,0)});
</script></body></html>`;
}

function memberLookupHtmlV2() {
  return `<!doctype html><html lang="ko"><head><meta charset="utf-8"><link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/pretendard@1.3.9/dist/web/static/pretendard.css"><style>
*{box-sizing:border-box}body{margin:0;background:#f4f6f8;color:#17243b;font-family:Pretendard,-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif}.wrap{width:min(100%,500px);margin:auto;padding:32px 18px}.brand{display:flex;align-items:center;gap:11px;margin:0 4px 20px;color:#21334f;font-weight:700}.mark{display:grid;place-items:center;width:38px;height:38px;flex:0 0 38px;border-radius:12px;background:#173f72;color:#fff;font-size:13px}.panel,.summary,.years{background:#fff;border:1px solid #e2e7ed;border-radius:22px;box-shadow:0 10px 30px rgba(25,42,68,.08)}.panel{padding:34px 32px 30px}.title{margin:0;font-size:28px;line-height:1.3;letter-spacing:-.045em}.intro{margin:14px 0 28px;color:#516078;font-size:17px;line-height:1.7;word-break:keep-all}.field{margin-top:20px}label{display:block;margin-bottom:9px;color:#253652;font-size:16px;font-weight:650}input{width:100%;height:58px;border:1px solid #cbd4df;border-radius:13px;background:#fff;color:#17243b;padding:0 17px;font:inherit;font-size:17px}input:focus{outline:none;border-color:#285b96;box-shadow:0 0 0 3px rgba(40,91,150,.13)}button{font-family:inherit}.primary,.secondary{width:100%;min-height:58px;border-radius:14px;font-size:17px;font-weight:700;cursor:pointer}.primary{margin-top:28px;border:0;background:#173f72;color:#fff}.primary:disabled{opacity:.58}.secondary{margin-top:20px;border:1px solid #cbd4df;background:#fff;color:#31445f}.privacy,.basis{display:flex;align-items:flex-start;gap:9px;margin:20px 1px 0;color:#69768a;font-size:14px;line-height:1.6;word-break:keep-all}.contact{margin:20px 0 0;text-align:center;color:#657286;font-size:14px}.contact strong{color:#31445f}.notice{display:none;margin-top:18px;padding:15px 16px;border-radius:13px;background:#fff4db;color:#765213;font-size:14px;line-height:1.65;word-break:keep-all}.view{display:none}.summary{padding:31px 30px 28px}.person{margin:0;color:#4f5f75;font-size:16px;line-height:1.5}.person strong{color:#253652}.total-label{margin-top:23px;color:#526178;font-size:15px}.total{margin:5px 0 0;color:#b23a36;font-size:35px;font-weight:750;letter-spacing:-.045em}.total.paid{color:#246746;font-size:29px}.help{margin:15px 0 0;color:#68768a;font-size:14px;line-height:1.65;word-break:keep-all}.years{margin-top:16px;padding:28px 26px 10px}.years h2{margin:0 0 9px;font-size:21px}.caption{margin:0 0 16px;color:#6b7789;font-size:14px;line-height:1.55}.year{padding:19px 0 20px;border-top:1px solid #edf0f3}.year-head{display:grid;grid-template-columns:62px minmax(0,1fr) auto;gap:13px;align-items:center}.year-name{color:#263852;font-size:16px;font-weight:700}.year-detail{color:#647287;font-size:14px;line-height:1.5}.year-detail strong{display:block;color:#31445f;font-size:15px}.state{min-width:62px;padding:8px 10px;border-radius:999px;text-align:center;font-size:13px;font-weight:700}.state-paid{background:#eaf5ef;color:#28704b}.state-partial{background:#fff4db;color:#8a6111}.state-unpaid{background:#fbeceb;color:#a13d39}.payments{margin:14px 0 0 75px;padding:13px 15px;border-radius:12px;background:#f6f8fa}.payments-title{margin:0 0 8px;color:#31445f;font-size:13px;font-weight:700}.payment{display:grid;grid-template-columns:minmax(0,1fr) auto;gap:9px;color:#607087;font-size:13px;line-height:1.55}.payment+.payment{margin-top:5px}.payment strong{color:#263852;white-space:nowrap}.payment-empty{margin:13px 0 0 75px;color:#7a8595;font-size:13px}.state-panel{padding:38px 32px 30px;text-align:center}.state-icon{display:grid;place-items:center;width:56px;height:56px;margin:0 auto 19px;border-radius:18px;background:#edf2f7;color:#536881;font-size:25px}.state-icon.lock{background:#fbeceb;color:#a13d39}.state-title{margin:0;font-size:25px;line-height:1.4;word-break:keep-all}.state-text{margin:14px auto 0;max-width:360px;color:#5f6d81;font-size:16px;line-height:1.75;word-break:keep-all}.state-note{margin:20px 0 0;padding:15px 16px;border-radius:13px;background:#f6f8fa;color:#667489;font-size:14px;line-height:1.65;text-align:left;word-break:keep-all}@media(max-width:520px){.wrap{padding:24px 12px}.brand{margin-bottom:16px}.panel,.summary{padding:28px 21px 25px;border-radius:19px}.years{padding:25px 19px 8px;border-radius:19px}.title{font-size:25px}.intro{font-size:16px}.year-head{grid-template-columns:55px minmax(0,1fr) auto;gap:9px}.state{min-width:57px;padding-inline:8px}.payments,.payment-empty{margin-left:0}.state-panel{padding:32px 21px 25px}}
</style></head><body><div class="wrap"><div class="brand"><div class="mark">부천</div><span>부천지역세무사회</span></div><main>
<section id="lookup" class="panel"><h1 class="title">회비 납부내역 조회</h1><p class="intro">등록번호와 휴대전화 번호 뒤 4자리를 입력해 주세요.</p><form id="form"><div class="field"><label for="reg">등록번호</label><input id="reg" inputmode="numeric" autocomplete="off" placeholder="등록번호 입력" required></div><div class="field"><label for="phone">휴대전화 번호 뒤 4자리</label><input id="phone" inputmode="numeric" maxlength="4" autocomplete="off" placeholder="숫자 4자리" required></div><button id="submit" class="primary" type="submit">납부내역 조회</button></form><div id="notice" class="notice" role="alert"></div><div class="privacy"><span>✓</span><span>입력하신 정보는 회비 납부내역 조회에만 사용됩니다.</span></div></section>
<section id="result" class="view"><div class="summary"><p class="person"><strong id="member-name"></strong>의 회비 납부내역입니다.</p><div id="total-label" class="total-label"></div><div id="total" class="total"></div><p id="result-help" class="help"></p></div><div class="years"><h2>연도별 납부현황</h2><p class="caption">최근 연도부터 표시되며, 입금액은 오래된 연도 회비부터 충당됩니다.</p><div id="year-list"></div></div><div id="basis" class="basis"></div><button id="again" class="secondary" type="button">처음 화면으로</button></section>
<section id="no-data" class="view panel state-panel"><div class="state-icon">⌕</div><h1 class="state-title">현재 조회할 수 있는<br>납부내역이 없습니다</h1><p class="state-text">회비 납부자료가 준비된 후 다시 조회해 주세요.</p><div class="state-note">최근에 회비를 입금하신 경우 통장 거래내역이 반영된 후 확인하실 수 있습니다.</div><button class="primary reset" type="button">다시 조회</button></section>
<section id="locked" class="view panel state-panel"><div class="state-icon lock">🔒</div><h1 class="state-title">조회가 잠겼습니다</h1><p class="state-text">입력정보가 3회 연속 일치하지 않아 더 이상 입력할 수 없습니다.</p><div class="state-note">확인이 필요하시면 재무간사 김선영 세무사에게 문의해 주세요.</div></section>
<p class="contact">문의&nbsp; <strong>재무간사 김선영 세무사</strong></p></main></div><script>
var form=document.getElementById('form'),submit=document.getElementById('submit'),notice=document.getElementById('notice'),phone=document.getElementById('phone'),CLIENT_KEY='bucheon-fee-lookup-client-v1';
function clientId(){var value=localStorage.getItem(CLIENT_KEY);if(!value){value=String(Date.now())+'-'+Math.random().toString(36).slice(2)+Math.random().toString(36).slice(2);localStorage.setItem(CLIENT_KEY,value)}return value}
function show(id){['lookup','result','no-data','locked'].forEach(function(key){document.getElementById(key).style.display=key===id?'block':'none'});window.scrollTo(0,0)}
function won(value){return Number(value||0).toLocaleString('ko-KR')+'원'}
function esc(value){return String(value==null?'':value).replace(/[&<>"']/g,function(ch){return({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'})[ch]})}
function basisText(value){if(!value)return'통장 거래내역 기준 · 확인 가능한 마지막 거래 시각 없음';var text=String(value).replace('T',' ').replace('Z','').slice(0,19),m=text.match(/^(\\d{4})-(\\d{2})-(\\d{2})\\s+(\\d{2}):(\\d{2})/);if(!m)return'통장 거래내역 기준 · '+text+'까지';var h=Number(m[4]),ampm=h<12?'오전':'오후';return'통장 거래내역 기준 · '+Number(m[1])+'. '+Number(m[2])+'. '+Number(m[3])+'. '+ampm+' '+(h%12||12)+'시'+(Number(m[5])?' '+Number(m[5])+'분':'')+'까지'}
function paymentHtml(p){var when=esc(p.date)+(p.time?' '+esc(p.time):''),full=Number(p.fullAmount||0),applied=Number(p.appliedAmount||0);return'<div class="payment"><span>'+when+' 입금 '+won(full)+(full!==applied?' 중':'')+'</span><strong>'+won(applied)+' 충당</strong></div>'}
function yearHtml(item){var status=item.status==='paid'?'완납':item.status==='partial'?'일부납':'미납',statusClass='state-'+(item.status==='paid'?'paid':item.status),main=item.status==='paid'?'충당 '+won(item.allocated):item.status==='partial'?'충당 '+won(item.allocated):'남은 금액 '+won(item.remaining),sub=item.status==='partial'?'남은 금액 '+won(item.remaining):'회비 '+won(item.required),payments=(item.payments||[]).map(paymentHtml).join(''),paymentBlock=payments?'<div class="payments"><p class="payments-title">입금·충당 내역</p>'+payments+'</div>':'<p class="payment-empty">입금·충당 내역이 없습니다.</p>',year=item.year==='prior'?'이월':String(item.year);return'<div class="year"><div class="year-head"><div class="year-name">'+esc(year)+'</div><div class="year-detail"><strong>'+main+'</strong>'+sub+'</div><div class="state '+statusClass+'">'+status+'</div></div>'+paymentBlock+'</div>'}
function renderResult(data){var years=(data.years||[]).filter(function(item){return Number(item.required||0)>0}).sort(function(a,b){if(a.year==='prior')return 1;if(b.year==='prior')return-1;return Number(b.year)-Number(a.year)}),arrears=Number(data.totalArrears||0);document.getElementById('member-name').textContent=data.name+' 세무사님';var total=document.getElementById('total');if(arrears>0){document.getElementById('total-label').textContent='총 미납액';total.textContent=won(arrears);total.className='total';document.getElementById('result-help').textContent='아래에서 입금일과 입금액, 각 연도 회비에 충당된 금액을 확인하실 수 있습니다.'}else{document.getElementById('total-label').textContent='';total.textContent='미납액이 없습니다';total.className='total paid';document.getElementById('result-help').textContent='연도별 입금일과 입금액, 각 연도 회비에 충당된 금액은 아래에서 확인하실 수 있습니다.'}document.getElementById('year-list').innerHTML=years.map(yearHtml).join('');document.getElementById('basis').textContent=basisText(data.bankLastTransactionAt);show('result')}
function showMismatch(data){var left=Number(data.remainingAttempts||0);notice.innerHTML='입력하신 정보와 일치하는 회원정보를 찾지 못했습니다.<br>등록번호와 휴대전화 번호 뒤 4자리를 다시 확인해 주세요.<br><strong>입력정보가 3회 연속 일치하지 않으면 조회가 잠깁니다.'+(left?' 남은 입력 횟수 '+left+'회.':'')+'</strong>';notice.style.display='block'}
function handleResponse(data){submit.disabled=false;submit.textContent='납부내역 조회';if(data&&data.locked){show('locked');return}if(data&&data.noData){show('no-data');return}if(!data||!data.ok){showMismatch(data||{});return}notice.style.display='none';renderResult(data)}
phone.addEventListener('input',function(){phone.value=phone.value.replace(/\\D/g,'').slice(0,4)});form.addEventListener('submit',function(e){e.preventDefault();submit.disabled=true;submit.textContent='조회 중…';google.script.run.withSuccessHandler(handleResponse).withFailureHandler(function(){submit.disabled=false;submit.textContent='납부내역 조회';notice.textContent='잠시 후 다시 조회해 주세요.';notice.style.display='block'}).lookupMember(document.getElementById('reg').value,phone.value,clientId())});document.getElementById('again').addEventListener('click',function(){form.reset();notice.style.display='none';show('lookup')});Array.prototype.forEach.call(document.querySelectorAll('.reset'),function(button){button.addEventListener('click',function(){form.reset();notice.style.display='none';show('lookup')})});google.script.run.withSuccessHandler(function(data){if(data&&data.locked)show('locked')}).getLookupLockStatus(clientId());
</script></body></html>`;
}
