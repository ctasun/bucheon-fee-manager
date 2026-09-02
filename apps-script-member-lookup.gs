const ACCESS_TOKEN = '여기에_관리프로그램과_같은_토큰을_입력하세요';
const FOLDER_NAME = '부천세무사회_회비관리_자동저장';
const LOOKUP_FILE_NAME = '회원조회자료.json';
const MAX_FAILURES = 5;
const LOCK_SECONDS = 600;

function doPost(e) {
  try {
    const data = JSON.parse((e.postData && e.postData.contents) || '{}');
    if (!data.token || data.token !== ACCESS_TOKEN) return jsonResponse({ ok: false, error: '인증에 실패했습니다.' });
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
  return HtmlService.createHtmlOutput(memberLookupHtml())
    .setTitle('부천지역세무사회 회비 납입내역 조회')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function lookupMember(regNo, phoneLast4) {
  const reg = digits(regNo);
  const last4 = digits(phoneLast4);
  if (reg.length < 4 || last4.length !== 4) return lookupFailure('입력정보를 확인해 주세요.');

  const cache = CacheService.getScriptCache();
  const failureKey = 'fail_' + hashText(reg);
  const failures = Number(cache.get(failureKey) || 0);
  if (failures >= MAX_FAILURES) return { ok: false, error: '조회 시도가 많습니다. 10분 후 다시 시도해 주세요.' };

  const data = readLookupData();
  if (!data || !Array.isArray(data.members)) return { ok: false, error: '조회자료가 아직 준비되지 않았습니다.' };
  const member = data.members.find(function (item) {
    return digits(item.regNo) === reg && digits(item.phoneLast4) === last4;
  });
  if (!member) {
    cache.put(failureKey, String(failures + 1), LOCK_SECONDS);
    return lookupFailure('등록번호 또는 휴대전화 번호를 확인해 주세요.');
  }
  cache.remove(failureKey);
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

function lookupFailure(message) {
  Utilities.sleep(350);
  return { ok: false, error: message };
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
