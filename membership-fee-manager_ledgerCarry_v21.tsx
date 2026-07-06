import React, { useState, useEffect, useMemo, useRef } from 'react';
import * as XLSX from 'xlsx';
import { Users, Settings, Upload, Grid3x3, Bell, Download, Plus, Trash2, Edit2, X, Check, Copy, FileSpreadsheet, Save } from 'lucide-react';

const SK = { M: 'mfm-members-v2', F: 'mfm-fees-v2', T: 'mfm-tx-v2', C: 'mfm-config-v2', U: 'mfm-upload-v2' };

const DEFAULT_FEE = (y) => (y >= 2020 && y <= 2022) ? 100000 : 200000;
const FEE_START = 2011;
const LEDGER_CUTOFF_YEAR = 2020; // 옛 장부는 이 연도까지만 신뢰(이후는 통장 기준)

const fmt = (n) => new Intl.NumberFormat('ko-KR').format(Math.round(n || 0));
const today = () => new Date().toISOString().split('T')[0];
const yearOf = (d) => d ? new Date(d).getFullYear() : new Date().getFullYear();
const monthOf = (d) => d ? new Date(d).getMonth() + 1 : 1;
const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
const fmtDateTime = (iso) => {
  if (!iso) return '-';
  const d = new Date(iso);
  if (isNaN(d.getTime())) return '-';
  const p = (n) => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${p(d.getMonth()+1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}`;
};

function buildDefaultFees(from, to) {
  const arr = [];
  for (let y = from; y <= to; y++) arr.push({ year: y, amount: DEFAULT_FEE(y) });
  return arr;
}

function isWithdrawn(m) {
  if (!m.leaveDate) return false;
  const d = new Date(m.leaveDate);
  return !isNaN(d.getTime()) && d <= new Date();
}

function feeForYear(year, fees) {
  const f = (fees || []).find(x => x.year === year);
  return f ? f.amount : DEFAULT_FEE(year);
}

function requiredFor(member, year, fees) {
  const base = feeForYear(year, fees);
  if (base === 0) return 0;
  if (year <= LEDGER_CUTOFF_YEAR) return 0; // 2020년까지는 '이월 미납'(priorArrears)으로 통합 — 연도별 이중계산 방지
  if (year > new Date().getFullYear()) return 0; // 아직 도래 안 한 다음 연도는 미납으로 잡지 않음
  const jy = yearOf(member.joinDate);
  const jm = monthOf(member.joinDate);
  if (member.joinDate && year < jy) return 0;
  if (member.leaveDate && year > yearOf(member.leaveDate)) return 0;
  if (member.joinDate && year === jy && jm >= 7) return Math.round(base / 2);
  return base;
}

// Allocate payments oldest-first. Includes prior(2010이전) bucket.
function allocate(member, memberTx, fees, upToYear) {
  const jy = yearOf(member.joinDate);
  const startYear = Math.max(jy, FEE_START);
  const endYear = member.leaveDate ? Math.min(upToYear, yearOf(member.leaveDate) - 1) : upToYear;
  const years = [];
  const prior = Number(member.priorArrears || 0);
  if (prior > 0) years.push({ year: 'prior', label: '2020까지이월', required: prior, allocated: 0, payments: [], isPrior: true });
  for (let y = startYear; y <= endYear; y++) {
    years.push({ year: y, required: requiredFor(member, y, fees), allocated: 0, payments: [] });
  }
  const sorted = [...memberTx].sort((a, b) => new Date(a.date) - new Date(b.date));
  for (const tx of sorted) {
    let rem = tx.amount;
    for (const yr of years) {
      if (rem <= 0) break;
      if (yr.allocated >= yr.required) continue;
      const give = Math.min(rem, yr.required - yr.allocated);
      yr.allocated += give;
      yr.payments.push({ txId: tx.id, date: tx.date, time: tx.time, amount: give, fullAmount: tx.amount, depositorName: tx.depositorName, description: tx.description });
      rem -= give;
    }
    if (rem > 0 && years.length) {
      const last = years[years.length - 1];
      last.payments.push({ txId: tx.id, date: tx.date, time: tx.time, amount: rem, fullAmount: tx.amount, depositorName: tx.depositorName, description: tx.description, overflow: true });
      last.allocated += rem;
    }
  }
  return years;
}

function cellStatus(yr) {
  if (yr.required === 0) return 'none';
  if (yr.allocated <= 0) return 'unpaid';
  if (yr.allocated < yr.required) return 'partial';
  return 'paid';
}

// ---------- parsing helpers ----------
function parseDate(v) {
  if (!v && v !== 0) return null;
  if (v instanceof Date) return v.toISOString().split('T')[0];
  if (typeof v === 'number') {
    const d = XLSX.SSF.parse_date_code(v);
    if (d) return `${d.y}-${String(d.m).padStart(2,'0')}-${String(d.d).padStart(2,'0')}`;
  }
  const s = String(v).trim();
  const m = s.match(/(\d{4})[.\-/년]\s*(\d{1,2})[.\-/월]\s*(\d{1,2})/);
  if (m) return `${m[1]}-${m[2].padStart(2,'0')}-${m[3].padStart(2,'0')}`;
  const m2 = s.match(/^(\d{4})(\d{2})(\d{2})/);
  if (m2) return `${m2[1]}-${m2[2]}-${m2[3]}`;
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d.toISOString().split('T')[0];
}
function parseDate8(v) {
  if (!v && v !== 0) return '';
  const s = String(v).trim().replace(/[^\d]/g, '');
  if (/^\d{8}$/.test(s)) return `${s.slice(0,4)}-${s.slice(4,6)}-${s.slice(6,8)}`;
  return parseDate(v) || '';
}
function parseDateTime(v) {
  if (!v && v !== 0) return { date: null, time: '' };
  if (v instanceof Date) return { date: v.toISOString().split('T')[0], time: v.toTimeString().slice(0,8) };
  const s = String(v).trim();
  const dt = s.match(/(\d{4})[.\-/](\d{1,2})[.\-/](\d{1,2})[\sT]+(\d{1,2}):(\d{2})(?::(\d{2}))?/);
  if (dt) return { date: `${dt[1]}-${dt[2].padStart(2,'0')}-${dt[3].padStart(2,'0')}`, time: `${dt[4].padStart(2,'0')}:${dt[5]}:${(dt[6]||'00').padStart(2,'0')}` };
  return { date: parseDate(v), time: '' };
}
function parseAmount(v) {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return v;
  const n = parseFloat(String(v).replace(/[,\s원]/g, ''));
  return isNaN(n) ? 0 : n;
}

const DATE_KEYS = ['거래일시','거래일자','거래일','날짜','일자','일시','이체일','입금일'];
const AMT_KEYS = ['입금액','입금','받은금액','맡기신금액','입금금액'];
const NAME_KEYS = ['보낸분','받는분','보낸분/받는분','입금자','의뢰인','입금자명','보낸이','기재내용'];
const DESC_KEYS = ['적요','내용','메모','비고','거래기록사항','송금메모'];

function cleanName(raw) {
  if (!raw) return '';
  return String(raw).trim().replace(/[(（[].*?[)）\]]/g, '').trim();
}
function extractNameCore(raw) {
  if (!raw) return '';
  let s = cleanName(raw);
  const noise = ['세무사','세무회계','회계법인','세무','회계','회비','정기회비','부천지역세무사회','부천지역','지역세무사회','세무사회','부천회','부천','년분','년도'];
  for (const w of noise) s = s.split(w).join(' ');
  s = s.replace(/[0-9]/g, ' ').replace(/[^\uAC00-\uD7A3\s]/g, ' ').replace(/\s+/g, ' ').trim();
  return s;
}
function extractRegNos(raw) {
  const m = String(raw || '').match(/\d{4,6}/g);
  return m || [];
}
function detectColumns(headers) {
  const find = (keys) => {
    for (let i = 0; i < headers.length; i++) {
      const h = String(headers[i] || '').trim();
      if (keys.some(k => h.includes(k))) return i;
    }
    return -1;
  };
  return { date: find(DATE_KEYS), amount: find(AMT_KEYS), name: find(NAME_KEYS), desc: find(DESC_KEYS) };
}

function smartMatch(rawName, members, aliasMap, regIndex, nameIndex) {
  const raw = String(rawName || '').trim();
  if (!raw) return null;
  const exact = aliasMap.get(raw) || aliasMap.get(cleanName(raw));
  if (exact) return { memberId: exact };
  for (const rn of extractRegNos(raw)) {
    const hit = regIndex.get(rn);
    if (hit) return { memberId: hit };
  }
  const core = extractNameCore(raw);
  if (core && core.length >= 2 && core.length <= 4 && !core.includes(' ')) {
    const ids = nameIndex.get(core);
    if (ids && ids.length === 1) return { memberId: ids[0] };
  }
  if (core && core.length >= 2) {
    const cand = [];
    for (const m of members) {
      if (m.name && m.name.length >= 2 && (core === m.name || core.startsWith(m.name) || raw.includes(m.name))) cand.push(m.id);
    }
    const u = [...new Set(cand)];
    if (u.length === 1) return { memberId: u[0] };
  }
  return null;
}
function suggestCandidates(rawName, members) {
  const raw = String(rawName || '').trim();
  if (!raw) return [];
  const byReg = [];
  for (const rn of extractRegNos(raw)) {
    for (const m of members) if (String(m.regNo||'').trim() === rn) byReg.push(m);
  }
  if (byReg.length) return [...new Map(byReg.map(m => [m.id, m])).values()];
  const core = extractNameCore(raw);
  const cand = [];
  if (core && core.length >= 2) {
    for (const m of members) {
      if (!m.name) continue;
      if (core === m.name || core.replace(/\s/g,'') === m.name || raw.includes(m.name) || core.startsWith(m.name) || m.name.startsWith(core)) cand.push(m);
    }
  }
  return [...new Map(cand.map(m => [m.id, m])).values()];
}

// ============================================================
export default function App() {
  const [tab, setTab] = useState('status');
  const [members, setMembers] = useState([]);
  const [fees, setFees] = useState([]);
  const [tx, setTx] = useState([]);
  const [config, setConfig] = useState({ groupName: '부천지역세무사회', cellMode: 'unpaid', relayUrl: '', relayToken: '', msgTemplate: '안녕하세요, {이름}님.\n{단체명}입니다.\n{연도}년 회비 {금액}원이 미납 상태이오니 확인 부탁드립니다.\n감사합니다.' });
  const [uploadMeta, setUploadMeta] = useState(null);
  const [loaded, setLoaded] = useState(false);
  const [modal, setModal] = useState(null);
  const [editing, setEditing] = useState(null);
  const [resetConfirm, setResetConfirm] = useState(false);
  const [lastSync, setLastSync] = useState(null);
  const [syncMsg, setSyncMsg] = useState('');
  const syncTimer = useRef(null);

  useEffect(() => {
    (async () => {
      const ld = async (k, fb) => { try { const r = await window.storage.get(k); return r ? JSON.parse(r.value) : fb; } catch { return fb; } };
      setMembers(await ld(SK.M, []));
      const lf = await ld(SK.F, []);
      setFees(lf && lf.length ? lf : buildDefaultFees(FEE_START, new Date().getFullYear() + 1));
      setTx(await ld(SK.T, []));
      setUploadMeta(await ld(SK.U, null));
      const c = await ld(SK.C, null);
      if (c) setConfig(c);
      setLoaded(true);
    })();
  }, []);

  const sv = (k, v) => { try { window.storage.set(k, JSON.stringify(v)); } catch {} };
  useEffect(() => { if (loaded) sv(SK.M, members); }, [members, loaded]);
  useEffect(() => { if (loaded) sv(SK.F, fees); }, [fees, loaded]);
  useEffect(() => { if (loaded) sv(SK.T, tx); }, [tx, loaded]);
  useEffect(() => { if (loaded) sv(SK.C, config); }, [config, loaded]);
  useEffect(() => { if (loaded && uploadMeta) sv(SK.U, uploadMeta); }, [uploadMeta, loaded]);
  useEffect(() => {
    if (!loaded) return;
    if (!(config.relayUrl||'').trim()) return;
    if (members.length === 0) return;
    if (syncTimer.current) clearTimeout(syncTimer.current);
    syncTimer.current = setTimeout(() => { pushToDrive(false); }, 3000);
    return () => { if (syncTimer.current) clearTimeout(syncTimer.current); };
  }, [members, fees, tx, loaded]);

  const aliasMap = useMemo(() => {
    const map = new Map();
    const put = (k, id) => { const s = String(k||'').trim(); if (s && !map.has(s)) map.set(s, id); };
    for (const m of members) {
      put(m.depositorName, m.id); put(cleanName(m.depositorName), m.id); put(m.name, m.id);
      for (const a of (m.aliases || [])) { put(a, m.id); put(cleanName(a), m.id); }
    }
    return map;
  }, [members]);
  const regIndex = useMemo(() => {
    const map = new Map();
    for (const m of members) { const r = String(m.regNo||'').trim(); if (r) map.set(r, m.id); }
    return map;
  }, [members]);
  const nameIndex = useMemo(() => {
    const map = new Map();
    for (const m of members) { const n = String(m.name||'').trim(); if (!n) continue; if (!map.has(n)) map.set(n, []); map.get(n).push(m.id); }
    return map;
  }, [members]);

  useEffect(() => {
    if (!loaded) return;
    setTx(prev => prev.map(t => {
      if (t.memberId || t.manualAssign) return t;
      const r = smartMatch(t.depositorName, members, aliasMap, regIndex, nameIndex);
      return r && r.memberId ? { ...t, memberId: r.memberId } : t;
    }));
    // eslint-disable-next-line
  }, [aliasMap, regIndex, nameIndex, loaded]);

  const yearRange = useMemo(() => {
    const now = new Date().getFullYear();
    const max = Math.max(now, ...fees.map(f => f.year), FEE_START);
    const arr = [];
    for (let y = LEDGER_CUTOFF_YEAR + 1; y <= max; y++) arr.push(y);
    return arr;
  }, [fees]);

  const allocations = useMemo(() => {
    const upTo = yearRange.length ? yearRange[yearRange.length - 1] : new Date().getFullYear();
    const map = new Map();
    for (const m of members) map.set(m.id, allocate(m, tx.filter(t => t.memberId === m.id), fees, upTo));
    return map;
  }, [members, tx, fees, yearRange]);

  const currentMembers = useMemo(() => members.filter(m => !isWithdrawn(m)), [members]);
  const withdrawnMembers = useMemo(() => members.filter(m => isWithdrawn(m)), [members]);

  // handlers
  const addMember = (d) => setMembers(p => [...p, { id: uid(), aliases: [], priorArrears: 0, ...d }]);
  const updateMember = (id, d) => setMembers(p => p.map(m => m.id === id ? { ...m, ...d } : m));
  const deleteMember = (id) => { setMembers(p => p.filter(m => m.id !== id)); setTx(p => p.map(t => t.memberId === id ? { ...t, memberId: null, manualAssign: false } : t)); };
  const setFeeForYear = (year, amount) => setFees(p => { const e = p.find(f => f.year === year); return e ? p.map(f => f.year === year ? { ...f, amount } : f) : [...p, { year, amount }].sort((a,b)=>a.year-b.year); });
  const deleteFee = (year) => setFees(p => p.filter(f => f.year !== year));
  const assignTx = (txId, memberId, saveAlias) => {
    const t = tx.find(x => x.id === txId); if (!t) return;
    setTx(p => p.map(x => x.id === txId ? { ...x, memberId, manualAssign: true } : x));
    if (saveAlias && t.depositorName) {
      setMembers(p => p.map(m => {
        if (m.id !== memberId) return m;
        const al = m.aliases || [];
        if (al.includes(t.depositorName) || m.depositorName === t.depositorName) return m;
        return { ...m, aliases: [...al, t.depositorName] };
      }));
    }
  };
  const unassignTx = (txId) => setTx(p => p.map(x => x.id === txId ? { ...x, memberId: null, manualAssign: true } : x));
  const deleteTx = (txId) => setTx(p => p.filter(x => x.id !== txId));

  const importLedger = (rows) => {
    // 장부는 등록번호로 기존 현회원에게만 부착(비현회원 제외).
    // 2020년까지 순미납(2010이전 + 2011~2020 미납, 선납 음수 포함)을 '이월 미납'(priorArrears)으로 저장.
    // 순액이 음수(선납/과오납)면 2021년 이후로 이월되는 크레딧(입금)으로 처리.
    const byReg = new Map();
    for (const m of members) { const rg = String(m.regNo||'').trim(); if (rg) byReg.set(rg, m); }
    const nt = []; const priorMap = new Map();
    let attached = 0, skipped = 0;
    for (const r of rows) {
      const reg = String(r.regNo||'').trim();
      const mem = reg ? byReg.get(reg) : null;
      if (!mem) { skipped++; continue; }
      attached++;
      let net = Number(r.priorArrears||0);
      for (const [ys, uv] of Object.entries(r.years)) { if (Number(ys) <= LEDGER_CUTOFF_YEAR) net += Number(uv||0); }
      if (net > 0) priorMap.set(mem.id, net);
      else if (net < 0) nt.push({ id: uid(), date: `${LEDGER_CUTOFF_YEAR}-12-31`, time: '', amount: -net, depositorName: mem.name, description: '기존장부 이관(선납/과오납 크레딧)', memberId: mem.id, manualAssign: true, carryOver: true });
    }
    setMembers(p => p.map(m => priorMap.has(m.id) ? { ...m, priorArrears: priorMap.get(m.id) } : m));
    setTx(p => [...p, ...nt]);
    return { members: attached, skipped, tx: nt.length };
  };

  const importRoster = (rows, mode) => {
    let added = 0, updated = 0;
    setMembers(prev => {
      let base = mode === 'replace' ? [] : [...prev];
      const byReg = new Map(), byName = new Map();
      base.forEach((m, i) => { if (m.regNo) byReg.set(String(m.regNo).trim(), i); if (m.name) byName.set(m.name.trim(), i); });
      for (const r of rows) {
        const reg = String(r.regNo||'').trim();
        const idx = reg ? (byReg.has(reg) ? byReg.get(reg) : -1) : (byName.has(r.name) ? byName.get(r.name) : -1);
        const noteVal = r.reason ? `최신명단·${r.reason}` : '최신명단';
        if (idx >= 0) {
          const ex = base[idx];
          base[idx] = { ...ex, regNo: reg || ex.regNo, name: r.name || ex.name, society: r.society || ex.society, joinDate: r.joinDate || ex.joinDate, leaveDate: r.leaveDate, depositorName: ex.depositorName || r.name, note: r.reason ? noteVal : ex.note, review: !!r.review };
          updated++;
        } else {
          base.push({ id: uid(), aliases: [], regNo: reg, name: r.name, depositorName: r.name, society: r.society||'', phone: '', joinDate: r.joinDate||'', leaveDate: r.leaveDate||'', priorArrears: 0, note: noteVal, review: !!r.review });
          if (reg) byReg.set(reg, base.length - 1);
          if (r.name) byName.set(r.name, base.length - 1);
          added++;
        }
      }
      return base;
    });
    return { added, updated };
  };

  const doFullReset = async () => {
    setMembers([]); setTx([]); setUploadMeta(null);
    setFees(buildDefaultFees(FEE_START, new Date().getFullYear() + 1));
    try { await window.storage.delete(SK.M); await window.storage.delete(SK.T); await window.storage.delete(SK.U); await window.storage.delete(SK.F); } catch {}
    setResetConfirm(false);
  };

  const buildWorkbook = () => {
    const wb = XLSX.utils.book_new();
    const anyPrior = members.some(m => Number(m.priorArrears||0) > 0);
    const mrows = [['등록번호','성명','입금자명','별칭','지역세무사회명','연락처','개업/전입일','폐업/전출일','2020까지이월미납','비고']];
    for (const m of members) mrows.push([m.regNo||'', m.name, m.depositorName, (m.aliases||[]).join(', '), m.society||'', m.phone||'', m.joinDate||'', m.leaveDate||'', Number(m.priorArrears||0)||'', m.note||'']);
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(mrows), '회원명부');
    const frows = [['연도','회비']];
    for (const f of [...fees].sort((a,b)=>a.year-b.year)) frows.push([f.year, f.amount]);
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(frows), '연도별회비');
    const head = ['연번','등록번호','성명','지역세무사회명','개업/전입일','폐업/전출일','상태', ...(anyPrior?['2020까지이월미납']:[]), ...yearRange.map(y=>`${y}년`), '총계'];
    const grid = [head];
    const sm = [...members].sort((a,b)=>(a.name||'').localeCompare(b.name||'','ko'));
    sm.forEach((m, idx) => {
      const alloc = allocations.get(m.id) || [];
      const wd = isWithdrawn(m);
      const row = [idx+1, m.regNo||'', m.name, m.society||'', m.joinDate||'', m.leaveDate||'', wd?'탈회':'현회원'];
      let sum = 0; const pa = Number(m.priorArrears||0);
      const pyr = alloc.find(a => a.isPrior); const prem = pyr ? Math.max(0, pyr.required - pyr.allocated) : 0;
      if (anyPrior) { row.push(pa>0?(prem>0?prem:'-'):'-'); if (!wd) sum += prem; }
      const jy = yearOf(m.joinDate);
      for (const y of yearRange) {
        const bj = m.joinDate && y < jy, al = m.leaveDate && y >= yearOf(m.leaveDate);
        const yr = alloc.find(a => a.year === y);
        if (bj) { row.push(''); continue; }
        if (al || !yr || yr.required === 0) { row.push('-'); continue; }
        const sf = yr.required - yr.allocated;
        if (sf <= 0) row.push('-'); else { row.push(sf); if (!wd) sum += sf; }
      }
      row.push(wd ? '(탈회·참고)' : sum);
      grid.push(row);
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(grid), '미납액현황');
    const trows = [['날짜','시각','입금자명','금액','적요','매칭회원']];
    for (const t of [...tx].sort((a,b)=>new Date(a.date)-new Date(b.date))) { const mb = members.find(m=>m.id===t.memberId); trows.push([t.date, t.time||'', t.depositorName, t.amount, t.description||'', mb?mb.name:'(미매칭)']); }
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(trows), '거래내역');
    return wb;
  };
  const exportAll = () => { XLSX.writeFile(buildWorkbook(), `회비관리_${today()}.xlsx`); };
  const pushToDrive = async (manual) => {
    const url = (config.relayUrl||'').trim();
    if (!url) { if (manual) setSyncMsg('중계 URL이 비어 있습니다. 회비설정 탭에서 입력하세요.'); return; }
    if (members.length === 0) { if (manual) setSyncMsg('회원이 0명이라 저장을 건너뜁니다(안전장치).'); return; }
    try {
      const b64 = XLSX.write(buildWorkbook(), { bookType: 'xlsx', type: 'base64' });
      setSyncMsg('저장 중…');
      const r = await fetch(url, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify({ token: (config.relayToken||'').trim(), xlsxBase64: b64, filename: `회비미납현황_${today()}.xlsx` }) });
      const j = await r.json();
      if (j && j.ok) { setLastSync(new Date().toISOString()); setSyncMsg('드라이브 저장 완료'); }
      else { setSyncMsg('저장 실패: ' + (j && j.error ? j.error : '알 수 없음')); }
    } catch (e) { setSyncMsg('저장 실패(네트워크/응답): ' + String(e)); }
  };

  const exportBackup = () => {
    const obj = { __bhm_backup: 1, version: 18, savedAt: new Date().toISOString(), members, fees, tx, config, uploadMeta };
    const blob = new Blob([JSON.stringify(obj)], { type: 'application/json' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `회비관리_백업_${today()}.json`;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(a.href), 2000);
  };
  const importBackup = (file) => {
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const o = JSON.parse(String(reader.result));
        if (!o || !o.__bhm_backup) { alert('이 파일은 회비관리 백업 파일이 아닙니다.'); return; }
        if (!window.confirm('현재 데이터를 백업 파일 내용으로 교체합니다. 계속할까요?')) return;
        setMembers(Array.isArray(o.members) ? o.members : []);
        setFees(Array.isArray(o.fees) && o.fees.length ? o.fees : buildDefaultFees(FEE_START, new Date().getFullYear() + 1));
        setTx(Array.isArray(o.tx) ? o.tx : []);
        if (o.config) setConfig(o.config);
        setUploadMeta(o.uploadMeta || null);
        alert('백업을 불러왔습니다.');
      } catch (e) { alert('백업 파일을 읽지 못했습니다: ' + e); }
    };
    reader.readAsText(file);
  };
  const tabs = [
    { id: 'status', name: '납입현황', icon: Grid3x3 },
    { id: 'members', name: '회원관리', icon: Users },
    { id: 'import', name: '기존장부 가져오기', icon: FileSpreadsheet },
    { id: 'fees', name: '회비설정', icon: Settings },
    { id: 'upload', name: '거래내역 업로드', icon: Upload },
    { id: 'overdue', name: '미납자 관리', icon: Bell },
    { id: 'withdrawn', name: '탈회회원', icon: Users },
  ];

  if (!loaded) return <div className="p-8 text-center text-gray-500">불러오는 중...</div>;
  const unmatchedN = tx.filter(x => !x.memberId).length;

  return (
    <div className="min-h-screen bg-gray-50 text-gray-900">
      <div className="bg-white border-b border-gray-200">
        <div className="max-w-7xl mx-auto px-4 py-3 flex items-center justify-between flex-wrap gap-2">
          <div className="flex items-center gap-3">
            <FileSpreadsheet className="w-6 h-6 text-indigo-600" />
            <div>
              <input className="font-bold text-lg bg-transparent border-b border-transparent hover:border-gray-300 focus:border-indigo-400 outline-none" value={config.groupName} onChange={e => setConfig({ ...config, groupName: e.target.value })} />
              <div className="text-xs text-gray-500">회비 관리 프로그램</div>
            </div>
          </div>
          <div className="flex items-center gap-2">
            <button onClick={exportAll} className="flex items-center gap-2 px-4 py-2 bg-indigo-600 text-white rounded-lg hover:bg-indigo-700 text-sm font-medium"><Download className="w-4 h-4" /> 엑셀 전체 다운로드</button>
            <button onClick={() => setResetConfirm(true)} className="flex items-center gap-2 px-3 py-2 bg-rose-50 text-rose-600 border border-rose-200 rounded-lg hover:bg-rose-100 text-sm font-medium"><Trash2 className="w-4 h-4" /> 완전 초기화</button>
          </div>
        </div>
        <div className="max-w-7xl mx-auto px-4 flex gap-1 overflow-x-auto">
          {tabs.map(t => {
            const Ic = t.icon; const active = tab === t.id;
            return (
              <button key={t.id} onClick={() => setTab(t.id)} className={`flex items-center gap-2 px-4 py-2.5 text-sm font-medium border-b-2 whitespace-nowrap ${active ? 'border-indigo-600 text-indigo-600' : 'border-transparent text-gray-600 hover:text-gray-900'}`}>
                <Ic className="w-4 h-4" /> {t.name}
                {t.id === 'upload' && unmatchedN > 0 && <span className="ml-1 px-1.5 py-0.5 bg-amber-100 text-amber-800 text-xs rounded-full">{unmatchedN}</span>}
              </button>
            );
          })}
        </div>
      </div>

      <div className="max-w-7xl mx-auto p-4">
        {tab === 'status' && <StatusGrid members={currentMembers} yearRange={yearRange} allocations={allocations} fees={fees} cellMode={config.cellMode} setCellMode={(m) => setConfig({ ...config, cellMode: m })} onCellClick={(m, yr) => setModal({ type: 'cell', member: m, yearData: yr })} />}
        {tab === 'members' && <MembersTab members={currentMembers} onAdd={addMember} onUpdate={updateMember} onDelete={deleteMember} editing={editing} setEditing={setEditing} />}
        {tab === 'import' && <ImportTab onImportLedger={importLedger} onImportRoster={importRoster} memberCount={members.length} onReset={() => setResetConfirm(true)} />}
        {tab === 'fees' && <><SyncCard config={config} setConfig={setConfig} lastSync={lastSync} syncMsg={syncMsg} onTest={() => pushToDrive(true)} /><BackupCard onExport={exportBackup} onImport={importBackup} memberCount={members.length} txCount={tx.length} /><FeesTab fees={fees} setFeeForYear={setFeeForYear} deleteFee={deleteFee} /></>}
        {tab === 'upload' && <UploadTab tx={tx} setTx={setTx} members={members} aliasMap={aliasMap} regIndex={regIndex} nameIndex={nameIndex} uploadMeta={uploadMeta} setUploadMeta={setUploadMeta} onAssign={assignTx} onDeleteTx={deleteTx} />}
        {tab === 'overdue' && <OverdueTab members={members} allocations={allocations} config={config} setConfig={setConfig} />}
        {tab === 'withdrawn' && <WithdrawnTab members={withdrawnMembers} allocations={allocations} onUpdate={updateMember} onDelete={deleteMember} />}
      </div>

      {modal?.type === 'cell' && <CellModal member={modal.member} yearData={modal.yearData} onClose={() => setModal(null)} onUnassign={(id) => { unassignTx(id); setModal(null); }} />}

      {resetConfirm && (
        <Modal onClose={() => setResetConfirm(false)} title="완전 초기화">
          <p className="text-sm text-gray-700 mb-2">모든 데이터를 삭제하고 처음부터 시작합니다.</p>
          <p className="text-sm text-gray-600 mb-4">삭제 대상: 회원 <b>{members.length}명</b> · 거래내역 <b>{tx.length}건</b></p>
          <div className="flex gap-2 justify-end">
            <button onClick={() => setResetConfirm(false)} className="px-4 py-2 text-gray-600 text-sm">취소</button>
            <button onClick={doFullReset} className="px-4 py-2 bg-rose-600 text-white rounded-md text-sm font-medium">전부 삭제</button>
          </div>
        </Modal>
      )}
    </div>
  );
}

function BackupCard({ onExport, onImport, memberCount, txCount }) {
  const fileRef = React.useRef(null);
  return (
    <div className="bg-white rounded-lg border border-gray-200 p-4 mb-4">
      <h3 className="font-semibold mb-1">전체 백업 · 인수인계</h3>
      <p className="text-xs text-gray-500 mb-3">회원·회비·통장·설정 전체를 파일 1개로 저장합니다. 다음 담당자는 이 파일을 "백업 불러오기"로 불러오면 그대로 이어받습니다. (현재 회원 {memberCount}명 · 거래 {txCount}건)</p>
      <div className="flex flex-wrap gap-2">
        <button onClick={onExport} className="flex items-center gap-2 px-4 py-2 bg-emerald-600 text-white rounded-lg hover:bg-emerald-700 text-sm font-medium"><Download className="w-4 h-4" /> 전체 백업 내려받기</button>
        <button onClick={() => fileRef.current && fileRef.current.click()} className="flex items-center gap-2 px-4 py-2 bg-white text-gray-700 border border-gray-300 rounded-lg hover:bg-gray-50 text-sm font-medium"><Upload className="w-4 h-4" /> 백업 불러오기</button>
        <input ref={fileRef} type="file" accept=".json,application/json" className="hidden" onChange={e => { const f = e.target.files && e.target.files[0]; if (f) onImport(f); e.target.value=''; }} />
      </div>
      <p className="mt-2 text-xs text-gray-400">불러오기는 현재 데이터를 백업 내용으로 교체합니다(불러오기 전 확인창이 뜹니다).</p>
    </div>
  );
}

function SyncCard({ config, setConfig, lastSync, syncMsg, onTest }) {
  const fmtT = (iso) => { if (!iso) return '없음'; const d = new Date(iso); const p = n => String(n).padStart(2,'0'); return `${d.getFullYear()}-${p(d.getMonth()+1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}`; };
  const on = !!(config.relayUrl||'').trim();
  return (
    <div className="bg-white rounded-lg border border-gray-200 p-4 mb-4">
      <div className="flex items-center justify-between mb-3 flex-wrap gap-2">
        <h3 className="font-semibold">구글 드라이브 자동저장 <span className={`ml-1 text-xs px-2 py-0.5 rounded-full ${on?'bg-emerald-100 text-emerald-700':'bg-gray-100 text-gray-500'}`}>{on?'켜짐':'꺼짐'}</span></h3>
        <button onClick={onTest} className="px-3 py-1.5 bg-indigo-600 text-white rounded-md text-sm hover:bg-indigo-700"><Save className="w-4 h-4 inline mr-1" />지금 저장</button>
      </div>
      <div className="space-y-2">
        <div><label className="text-xs text-gray-600 block mb-1">중계 URL (…/exec)</label><input value={config.relayUrl||''} onChange={e=>setConfig({...config, relayUrl:e.target.value})} placeholder="https://script.google.com/macros/s/.../exec" className="w-full border border-gray-300 rounded px-2 py-1.5 text-sm" /></div>
        <div><label className="text-xs text-gray-600 block mb-1">토큰</label><input value={config.relayToken||''} onChange={e=>setConfig({...config, relayToken:e.target.value})} placeholder="예: bcstax2026x9k3m2" className="w-full border border-gray-300 rounded px-2 py-1.5 text-sm" /></div>
      </div>
      <div className="mt-3 text-xs text-gray-600">마지막 자동저장: <b>{fmtT(lastSync)}</b>{syncMsg && <span className="ml-2 text-indigo-600">· {syncMsg}</span>}</div>
      <div className="mt-1 text-xs text-gray-400">회원 0명일 때는 자동저장을 건너뜁니다(빈 데이터 덮어쓰기 방지). 데이터가 바뀌면 몇 초 뒤 자동 저장됩니다.</div>
    </div>
  );
}

// ============ Status Grid ============
function StatusGrid({ members, yearRange, allocations, fees, cellMode, setCellMode, onCellClick }) {
  const mode = cellMode || 'unpaid';
  const anyPrior = members.some(m => Number(m.priorArrears || 0) > 0);
  const stats = useMemo(() => {
    let req = 0, col = 0, unpaid = 0, wUnpaid = 0, wCount = 0;
    for (const m of members) {
      const alloc = allocations.get(m.id) || [];
      const wd = isWithdrawn(m);
      if (wd) wCount++;
      for (const yr of alloc) {
        if (yr.required <= 0) continue;
        const sf = yr.required - Math.min(yr.allocated, yr.required);
        if (wd) { wUnpaid += sf; continue; }
        req += yr.required; col += Math.min(yr.allocated, yr.required); unpaid += sf;
      }
    }
    return { req, col, unpaid, wUnpaid, wCount };
  }, [members, allocations]);

  if (!members.length) return <EmptyState icon={Users} title="회원이 없습니다" desc="기존장부 가져오기 또는 회원관리 탭에서 회원을 먼저 등록해 주세요." />;
  const sorted = [...members].sort((a, b) => (a.name || '').localeCompare(b.name || '', 'ko'));

  return (
    <div>
      <div className="grid grid-cols-2 md:grid-cols-3 gap-3 mb-4">
        <StatCard label="현 회원 회비 의무" value={`${fmt(stats.req)}원`} />
        <StatCard label="납입 완료" value={`${fmt(stats.col)}원`} color="text-emerald-600" />
        <StatCard label="현재 미납액 (현 회원)" value={`${fmt(stats.unpaid)}원`} color="text-rose-600" />
      </div>
      <div className="flex flex-wrap gap-3 mb-3 text-xs items-center">
        <Legend color="bg-emerald-100 text-emerald-800 border-emerald-300" label="완납" />
        <Legend color="bg-amber-100 text-amber-800 border-amber-300" label="일부납" />
        <Legend color="bg-rose-100 text-rose-800 border-rose-300" label="미납" />
        <Legend color="bg-gray-100 text-gray-400 border-gray-300" label="의무전/전출/탈회" />
        <div className="ml-auto flex items-center gap-2">
          <span className="text-gray-500">표시:</span>
          <div className="inline-flex rounded-md border border-gray-300 overflow-hidden">
            <button onClick={() => setCellMode('unpaid')} className={`px-2 py-1 ${mode === 'unpaid' ? 'bg-indigo-600 text-white' : 'bg-white text-gray-600'}`}>미납액</button>
            <button onClick={() => setCellMode('status')} className={`px-2 py-1 ${mode === 'status' ? 'bg-indigo-600 text-white' : 'bg-white text-gray-600'}`}>납입상태</button>
          </div>
        </div>
      </div>
      <div className="text-xs text-gray-500 mb-2">셀을 클릭하면 입금 상세내역을 볼 수 있습니다</div>
      <div className="bg-white rounded-lg border border-gray-200 overflow-auto max-h-[75vh]">
        <table className="w-full text-sm border-collapse">
          <thead>
            <tr className="bg-gray-50 border-b border-gray-200">
              <th className="sticky left-0 top-0 z-30 bg-gray-50 px-3 py-2 text-left font-semibold border-r border-gray-200">회원</th>
              <th className="sticky top-0 z-20 bg-gray-50 px-2 py-2 text-left font-semibold border-r border-gray-200 text-gray-600">입회일</th>
              {anyPrior && <th className="sticky top-0 z-20 px-2 py-2 text-center font-semibold border-r border-gray-200 min-w-[90px] bg-purple-50">2020까지<div className="text-xs font-normal text-gray-500">이월미납</div></th>}
              {yearRange.map(y => <th key={y} className="sticky top-0 z-20 bg-gray-50 px-2 py-2 text-center font-semibold border-r border-gray-200 min-w-[105px]">{y}년<div className="text-xs font-normal text-gray-500">{fmt(feeForYear(y, fees))}원</div></th>)}
              <th className="sticky top-0 z-20 bg-gray-50 px-2 py-2 text-center font-semibold min-w-[100px]">미납합계</th>
            </tr>
          </thead>
          <tbody>
            {sorted.map(m => {
              const alloc = allocations.get(m.id) || [];
              const jy = yearOf(m.joinDate);
              const wd = isWithdrawn(m);
              const unpaidSum = alloc.reduce((s, yr) => s + (yr.required > yr.allocated ? yr.required - yr.allocated : 0), 0);
              const pa = Number(m.priorArrears || 0);
              return (
                <tr key={m.id} className={`border-b border-gray-100 ${wd ? 'bg-gray-100 text-gray-400' : 'hover:bg-gray-50/50'}`}>
                  <td className={`sticky left-0 px-3 py-2 font-medium border-r border-gray-200 ${wd ? 'bg-gray-100 text-gray-400' : 'bg-white'}`}>
                    {m.name}{wd && <span className="block text-[10px] text-gray-400 font-normal">탈회 {m.leaveDate}</span>}
                  </td>
                  <td className={`px-2 py-2 text-xs border-r border-gray-200 ${wd ? 'text-gray-400' : 'text-gray-500'}`}>{m.joinDate || '-'}</td>
                  {anyPrior && (() => { const pyr = alloc.find(a => a.isPrior); const prem = pyr ? Math.max(0, pyr.required - pyr.allocated) : 0; return <td className={`p-0 border-r border-gray-100 text-center font-medium ${pa <= 0 ? 'bg-purple-50/40 text-gray-300' : (wd ? 'bg-gray-100 text-gray-400' : 'bg-purple-50 text-purple-800')}`}>{pa > 0 ? <button onClick={() => onCellClick(m, pyr)} className="w-full h-full px-2 py-2 hover:bg-purple-100">{prem > 0 ? fmt(prem) : <span className="text-emerald-600 font-normal">-</span>}</button> : '—'}</td>; })()}
                  {yearRange.map(y => {
                    const yr = alloc.find(a => a.year === y);
                    const bj = m.joinDate && y < jy;
                    const al = m.leaveDate && y >= yearOf(m.leaveDate);
                    if (bj || al) return <td key={y} className="px-2 py-2 border-r border-gray-100 text-center bg-gray-100 text-gray-300" title={bj ? '입회 전' : (y === yearOf(m.leaveDate) ? '전출 연도' : '전출 후')}></td>;
                    if (!yr || yr.required === 0) return <td key={y} className="px-2 py-2 border-r border-gray-100 text-center text-gray-400">—</td>;
                    const st = cellStatus(yr);
                    const sf = yr.required - yr.allocated;
                    const half = y === jy && m.joinDate && monthOf(m.joinDate) >= 7;
                    if (wd) return (
                      <td key={y} className="border-r border-gray-100 p-0">
                        <button onClick={() => onCellClick(m, yr)} className="w-full h-full px-2 py-2 text-center bg-gray-100 text-gray-400">{st === 'paid' ? '-' : fmt(sf)}</button>
                      </td>
                    );
                    const cls = st === 'paid' ? 'bg-emerald-50 text-emerald-800 hover:bg-emerald-100' : st === 'partial' ? 'bg-amber-50 text-amber-800 hover:bg-amber-100' : 'bg-rose-50 text-rose-800 hover:bg-rose-100';
                    return (
                      <td key={y} className="border-r border-gray-100 p-0">
                        <button onClick={() => onCellClick(m, yr)} className={`w-full h-full px-2 py-2 text-center font-medium ${cls}`} title={`의무 ${fmt(yr.required)}원${half ? ' (하반기 절반)' : ''} / 납입 ${fmt(yr.allocated)}원`}>
                          {mode === 'unpaid'
                            ? (st === 'paid' ? <span className="text-emerald-600 font-normal">-</span> : <>{fmt(sf)}{half && <span className="block text-[10px] text-gray-400 font-normal">½</span>}</>)
                            : (<>{st === 'paid' && <>완납<br /><span className="text-xs font-normal">{fmt(yr.allocated)}</span></>}{st === 'partial' && <>일부<br /><span className="text-xs font-normal">{fmt(yr.allocated)}/{fmt(yr.required)}</span></>}{st === 'unpaid' && <>미납<br /><span className="text-xs font-normal">{fmt(yr.required)}</span></>}</>)}
                        </button>
                      </td>
                    );
                  })}
                  <td className={`px-2 py-2 text-center font-semibold ${wd ? 'text-gray-400' : unpaidSum > 0 ? 'text-rose-600' : 'text-gray-400'}`}>{unpaidSum > 0 ? fmt(unpaidSum) : '—'}{wd && unpaidSum > 0 && <span className="block text-[10px] font-normal">참고용</span>}</td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ============ Members Tab ============
function MembersTab({ members, onAdd, onUpdate, onDelete, editing, setEditing }) {
  const [form, setForm] = useState({ regNo: '', name: '', depositorName: '', society: '부천지역세무사회', phone: '', joinDate: today(), leaveDate: '', note: '' });
  const [showAdd, setShowAdd] = useState(false);
  const [search, setSearch] = useState('');
  const [statusFilter, setStatusFilter] = useState('all');
  const reviewCount = members.filter(m => m.review).length;
  const filtered = members.filter(m => {
    if (statusFilter === 'review' && !m.review) return false;
    if (!search) return true;
    return m.name.includes(search) || (m.depositorName||'').includes(search) || (m.phone||'').includes(search) || (m.regNo||'').includes(search) || (m.note||'').includes(search);
  });
  const submit = () => {
    if (!form.name.trim()) return alert('성명을 입력하세요');
    if (!form.joinDate) return alert('개업/전입일을 입력하세요');
    onAdd({ ...form, depositorName: form.depositorName.trim() || form.name.trim() });
    setForm({ regNo: '', name: '', depositorName: '', society: '부천지역세무사회', phone: '', joinDate: today(), leaveDate: '', note: '' });
    setShowAdd(false);
  };
  return (
    <div className="bg-white rounded-lg border border-gray-200 p-4">
      <div className="flex items-center justify-between mb-3 flex-wrap gap-2">
        <div className="flex items-center gap-2 flex-wrap">
          <input placeholder="이름/입금자명/등록번호/사유 검색" value={search} onChange={e=>setSearch(e.target.value)} className="px-3 py-1.5 border border-gray-300 rounded-md text-sm" />
          <span className="text-sm text-gray-500">현회원 {members.length}명</span>
          <div className="inline-flex rounded-md border border-gray-300 overflow-hidden text-xs">
            <button onClick={()=>setStatusFilter('all')} className={`px-3 py-1.5 ${statusFilter==='all'?'bg-indigo-600 text-white':'bg-white text-gray-600'}`}>전체</button>
            <button onClick={()=>setStatusFilter('review')} className={`px-3 py-1.5 border-l border-gray-300 ${statusFilter==='review'?'bg-amber-500 text-white':'bg-white text-amber-600'}`}>확인필요 {reviewCount}</button>
          </div>
        </div>
        <button onClick={() => setShowAdd(true)} className="flex items-center gap-1 px-3 py-1.5 bg-indigo-600 text-white rounded-md text-sm hover:bg-indigo-700"><Plus className="w-4 h-4" /> 회원 추가</button>
      </div>
      {showAdd && (
        <div className="border border-indigo-200 bg-indigo-50/50 rounded-lg p-3 mb-3">
          <div className="grid grid-cols-1 md:grid-cols-3 gap-2 mb-2">
            <Field label="등록번호" value={form.regNo} onChange={v => setForm({...form, regNo: v})} />
            <Field label="성명 *" value={form.name} onChange={v => setForm({...form, name: v})} />
            <Field label="입금자명 (미입력 시 성명)" value={form.depositorName} onChange={v => setForm({...form, depositorName: v})} />
            <Field label="지역세무사회명" value={form.society} onChange={v => setForm({...form, society: v})} />
            <Field label="연락처" value={form.phone} onChange={v => setForm({...form, phone: v})} />
            <Field label="개업/전입일 *" type="date" value={form.joinDate} onChange={v => setForm({...form, joinDate: v})} />
            <Field label="폐업/전출일" type="date" value={form.leaveDate} onChange={v => setForm({...form, leaveDate: v})} />
            <div className="md:col-span-2"><Field label="비고" value={form.note} onChange={v => setForm({...form, note: v})} /></div>
          </div>
          <div className="text-xs text-gray-500 mb-2">💡 하반기(7~12월) 입회는 첫 해 절반 자동 적용 · 전출일 이후는 회비 의무 없음</div>
          <div className="flex gap-2 justify-end">
            <button onClick={() => setShowAdd(false)} className="px-3 py-1.5 text-sm text-gray-600">취소</button>
            <button onClick={submit} className="px-3 py-1.5 bg-indigo-600 text-white rounded-md text-sm">저장</button>
          </div>
        </div>
      )}
      <div className="overflow-auto max-h-[70vh] border border-gray-100 rounded">
        <table className="w-full text-sm">
          <thead><tr className="border-b border-gray-200 text-left text-gray-600 bg-gray-50 sticky top-0 z-10">
            <th className="px-2 py-2">등록번호</th><th className="px-2 py-2">성명</th><th className="px-2 py-2">입금자명/별칭</th><th className="px-2 py-2">개업/전입</th><th className="px-2 py-2">전출</th><th className="px-2 py-2">상태</th><th className="px-2 py-2"></th>
          </tr></thead>
          <tbody>
            {filtered.length === 0 && <tr><td colSpan={7} className="text-center text-gray-400 py-8">회원이 없습니다</td></tr>}
            {filtered.map(m => editing === m.id ? (
              <MemberEditRow key={m.id} m={m} onSave={(d) => { onUpdate(m.id, d); setEditing(null); }} onCancel={() => setEditing(null)} />
            ) : (
              <tr key={m.id} className={`border-b border-gray-100 hover:bg-gray-50 ${isWithdrawn(m) ? 'text-gray-400' : ''}`}>
                <td className="px-2 py-2 text-xs text-gray-500">{m.regNo || '-'}</td>
                <td className="px-2 py-2 font-medium">{m.name}</td>
                <td className="px-2 py-2 text-xs">{m.depositorName}{(m.aliases||[]).length ? <span className="text-gray-400"> / {(m.aliases||[]).join(', ')}</span> : ''}</td>
                <td className="px-2 py-2 text-xs">{m.joinDate || '-'}{m.joinDate && <span className="text-gray-400">{monthOf(m.joinDate) >= 7 ? ' (하)' : ' (상)'}</span>}</td>
                <td className="px-2 py-2 text-xs text-rose-500">{m.leaveDate || '-'}</td>
                <td className="px-2 py-2 text-xs">{isWithdrawn(m) ? <span className="text-gray-500">탈회</span> : (m.review ? <span className="text-amber-600 font-medium">확인필요<span className="block text-[10px] text-gray-400 font-normal">{(m.note||'').replace('최신명단·','')}</span></span> : <span className="text-emerald-600">현회원</span>)}</td>
                <td className="px-2 py-2 text-right whitespace-nowrap">
                  <button onClick={() => setEditing(m.id)} className="p-1 text-gray-500 hover:text-indigo-600"><Edit2 className="w-4 h-4" /></button>
                  <button onClick={() => onDelete(m.id)} className="p-1 text-gray-500 hover:text-rose-600"><Trash2 className="w-4 h-4" /></button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function MemberEditRow({ m, onSave, onCancel }) {
  const [f, setF] = useState({ ...m, aliases: (m.aliases || []).join(', ') });
  const save = () => onSave({ ...f, aliases: String(f.aliases||'').split(',').map(s=>s.trim()).filter(Boolean), priorArrears: Number(f.priorArrears||0), review: false });
  return (
    <tr className="bg-indigo-50/30 border-b border-gray-200">
      <td className="px-1 py-1"><input className="border border-gray-300 rounded px-1 py-1 w-16 text-sm" value={f.regNo||''} onChange={e=>setF({...f, regNo:e.target.value})} /></td>
      <td className="px-1 py-1"><input className="border border-gray-300 rounded px-1 py-1 w-full text-sm" value={f.name} onChange={e=>setF({...f, name:e.target.value})} /></td>
      <td className="px-1 py-1"><input className="border border-gray-300 rounded px-1 py-1 w-full text-sm mb-1" placeholder="입금자명" value={f.depositorName} onChange={e=>setF({...f, depositorName:e.target.value})} /><input className="border border-gray-300 rounded px-1 py-1 w-full text-sm" placeholder="별칭,별칭2" value={f.aliases} onChange={e=>setF({...f, aliases:e.target.value})} /></td>
      <td className="px-1 py-1"><input type="date" className="border border-gray-300 rounded px-1 py-1 text-sm" value={f.joinDate||''} onChange={e=>setF({...f, joinDate:e.target.value})} /></td>
      <td className="px-1 py-1"><input type="date" className="border border-gray-300 rounded px-1 py-1 text-sm" value={f.leaveDate||''} onChange={e=>setF({...f, leaveDate:e.target.value})} /></td>
      <td className="px-1 py-1"><input type="number" className="border border-gray-300 rounded px-1 py-1 w-20 text-sm" placeholder="이월미납" value={f.priorArrears||''} onChange={e=>setF({...f, priorArrears:e.target.value})} title="2020까지 이월 미납" /></td>
      <td className="px-1 py-1 text-right whitespace-nowrap"><button onClick={save} className="p-1 text-emerald-600"><Check className="w-4 h-4" /></button><button onClick={onCancel} className="p-1 text-gray-500"><X className="w-4 h-4" /></button></td>
    </tr>
  );
}

// ============ Import Tab ============
function ImportTab({ onImportLedger, onImportRoster, memberCount, onReset }) {
  const [ledgerPrev, setLedgerPrev] = useState(null);
  const [rosterPrev, setRosterPrev] = useState(null);
  const [result, setResult] = useState(null);
  const ledgerRef = useRef(), rosterRef = useRef();

  const parseLedger = async (file) => {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { cellDates: true });
    const data = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: '' });
    let hi = -1, bestY = 0;
    for (let i = 0; i < Math.min(data.length, 15); i++) {
      const row = (data[i]||[]).map(c => String(c));
      const nameish = row.some(c => c.includes('성명') || c.includes('이전') || c.includes('이후') || c.includes('미납'));
      if (!nameish) continue;
      const yc = row.filter(c => /(20\d\d)/.test(c)).length;
      if (yc > bestY) { bestY = yc; hi = i; }
    }
    if (hi < 0) { alert('헤더(성명/연도)를 찾을 수 없습니다.'); return; }
    const hrow = (data[hi]||[]).map(c => String(c).trim());
    const above = (data[hi-1]||[]).map(c => String(c).trim());
    const headers = hrow.map((h, i) => h || above[i] || '');
    const col = (keys) => headers.findIndex(h => keys.some(k => h.includes(k)));
    const nameCol = col(['성명','이름']), regCol = col(['등록번호']), socCol = headers.findIndex((h,idx) => idx !== nameCol && !h.includes('성명') && !h.includes('이름') && ['지역회','지역세무사회','세무사회'].some(k => h.includes(k))), joinCol = col(['개업','전입']), leaveCol = col(['폐업','전출']);
    const priorCols = []; headers.forEach((h, i) => { if (h.includes('이전') || h.includes('이후')) priorCols.push(i); });
    const yearCols = []; headers.forEach((h, i) => { if (priorCols.includes(i)) return; const m = h.match(/(20\d\d)/); if (m) yearCols.push({ year: Number(m[1]), idx: i }); });
    if (nameCol < 0) { alert('성명 컬럼을 찾을 수 없습니다.'); return; }
    const rows = [];
    for (let i = hi + 1; i < data.length; i++) {
      const row = data[i] || [];
      const name = String(row[nameCol] || '').trim();
      if (!name || name.includes('계')) continue;
      let prior = 0;
      for (const pc of priorCols) { const raw = String(row[pc] ?? '').trim(); if (raw === '' || raw === '-' || raw.includes('신')) continue; const n = parseAmount(raw); if (n > 0) prior += n; }
      const years = {};
      for (const yc of yearCols) { if (yc.year > LEDGER_CUTOFF_YEAR) continue; const raw = String(row[yc.idx] ?? '').trim(); if (raw === '' || raw === '-' || raw.includes('신')) continue; const n = parseAmount(raw); if (n !== 0 && n % 50000 === 0) years[yc.year] = n; } // 음수(선납/과오납)도 저장
      rows.push({ name, regNo: regCol>=0?String(row[regCol]||'').trim():'', society: socCol>=0?String(row[socCol]||'').trim():'', joinDate: joinCol>=0?parseDate8(row[joinCol]):'', leaveDate: leaveCol>=0?parseDate8(row[leaveCol]):'', priorArrears: prior, years });
    }
    setLedgerPrev({ rows, yearCols: yearCols.map(y=>y.year), hasPrior: priorCols.length>0, fileName: file.name });
    setResult(null);
  };

  const parseRoster = async (file) => {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { cellDates: true });
    const data = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: '' });
    let hi = -1;
    for (let i = 0; i < Math.min(data.length, 15); i++) { const row = (data[i]||[]).map(c=>String(c)); if (row.some(c=>c.includes('성명')||c.includes('이름')) && row.some(c=>c.includes('개업')||c.includes('전입')||c.includes('등록번호'))) { hi = i; break; } }
    if (hi < 0) hi = 0;
    const headers = (data[hi]||[]).map(c=>String(c).trim());
    const col = (keys) => headers.findIndex(h => keys.some(k => h.includes(k)));
    const nameCol = col(['성명','이름']), regCol = col(['등록번호']), socCol = headers.findIndex((h,idx) => idx !== nameCol && !h.includes('성명') && !h.includes('이름') && ['지역회','지역세무사회','세무사회'].some(k => h.includes(k))), joinCol = col(['개업','전입']), leaveCol = col(['폐업','전출']);
    if (nameCol < 0) { alert('성명 컬럼을 찾을 수 없습니다.'); return; }

    // 전출 사유 판정: 날짜면 그 날짜, 글자면 탈회 사유
    const classifyLeave = (rawVal) => {
      const s = String(rawVal ?? '').trim();
      if (!s) return { withdrawn: false, leaveDate: '', reason: '', review: false };
      // 현회원 유지: 개업·입회·전입·재등록·주소변경
      if (s.includes('개업') || s.includes('입회') || s.includes('전입') || s.includes('재등록') || s.includes('주소')) return { withdrawn: false, leaveDate: '', reason: '', review: false };
      const d = parseDate8(s);
      if (d) return { withdrawn: true, leaveDate: d, reason: '', review: false }; // 날짜 → 전출일
      // 명시적 전출 사유
      if (s.includes('휴업')) return { withdrawn: true, leaveDate: '', reason: '휴업', review: false };
      if (s.includes('사망')) return { withdrawn: true, leaveDate: '', reason: '사망', review: false };
      if (s.includes('폐업')) return { withdrawn: true, leaveDate: '', reason: '폐업', review: false };
      if (s.includes('부천') || s.includes('-->') || s.includes('전출') || s.includes('타')) return { withdrawn: true, leaveDate: '', reason: '타지역 전출', review: false };
      // 인식 안 되는 사유 → 확인필요(현회원 유지, 사용자 검토 대상)
      return { withdrawn: false, leaveDate: '', reason: `확인필요: ${s}`, review: true };
    };

    const rows = [];
    let inWithdrawnSection = false; // <전출명단> 이후 true
    const seenRegs = new Set(); // 같은 등록번호가 두 번 나오면 먼저 나온 현회원 구간 행을 채택
    for (let i = hi + 1; i < data.length; i++) {
      const row = data[i] || [];
      const rowText = row.map(c => String(c)).join(' ');
      // 구분 행 감지: "전출명단"
      if (rowText.includes('전출명단') || rowText.includes('전출 명단')) { inWithdrawnSection = true; continue; }
      const name = String(row[nameCol] || '').trim();
      if (!name || name.includes('계')) continue;

      const regC = regCol >= 0 ? String(row[regCol]||'').trim() : '';
      if (regC && seenRegs.has(regC)) continue; // 이미 나온 등록번호(현회원 구간 우선) → 전출명단 중복행 건너뜀
      if (regC) seenRegs.add(regC);
      const rawLeave = leaveCol >= 0 ? row[leaveCol] : '';
      const cls0 = classifyLeave(rawLeave);
      // 위치 기준: <전출명단> 아래=탈회, 위=현회원(개업·전입일 있는 209명). F열 값은 참고만.
      const cls = inWithdrawnSection
        ? { withdrawn: true, leaveDate: cls0.leaveDate || '', reason: (cls0.reason && !cls0.review) ? cls0.reason : '전출명단', review: false }
        : { withdrawn: false, leaveDate: '', reason: '', review: false };
      const joinDate = joinCol >= 0 ? parseDate8(row[joinCol]) : '';
      // 탈회인데 전출일(날짜)이 없으면(휴업 등 글자 사유): 입회일 이후로 두어 논리 오류 방지.
      // 입회일이 있으면 그 해 말일, 없으면 파일 기준(현재)로.
      let leaveDate = cls.leaveDate;
      if (cls.withdrawn && !leaveDate) {
        if (joinDate) {
          const jyr = new Date(joinDate).getFullYear();
          leaveDate = `${jyr}-12-31`; // 입회한 해에 휴업 → 그 해까지는 의무, 이후 회색
        } else {
          leaveDate = today();
        }
      }

      rows.push({
        name,
        regNo: regCol>=0?String(row[regCol]||'').trim():'',
        society: socCol>=0?String(row[socCol]||'').trim():'',
        joinDate,
        leaveDate: cls.withdrawn ? leaveDate : '',
        withdrawn: cls.withdrawn,
        reason: cls.reason,
        review: cls.review,
      });
    }
    const activeN = rows.filter(r => !r.withdrawn && !r.review).length;
    const wN = rows.filter(r => r.withdrawn).length;
    const reviewN = rows.filter(r => r.review).length;
    setRosterPrev({ rows, fileName: file.name, activeN, wN, reviewN });
    setResult(null);
  };

  const handle = (fn) => async (e) => { const f = e.target.files[0]; if (!f) return; try { await fn(f); } catch (err) { alert('파일 읽기 실패: ' + err.message + '\n(은행/명단 파일을 엑셀에서 xlsx로 다시 저장해 보세요)'); } e.target.value = ''; };

  return (
    <div className="space-y-3">
      <div className="bg-indigo-50 border border-indigo-200 rounded-lg p-4">
        <h3 className="font-semibold text-indigo-900 mb-2">가져오기 안내</h3>
        <div className="text-sm text-gray-700 space-y-1">
          <div><b>평소(한두 달마다):</b> 아래 <b>최신 회원명단</b>만 반영하면 됩니다. 신규 추가·전출 탈회가 자동 처리됩니다.</div>
          <div><b>최초 1회 설정:</b> ①회원명단 반영 → ②옛 미납장부 → ③거래내역 업로드 탭에서 통장 파일.</div>
        </div>
        <div className="text-xs text-gray-500 mt-2">※ 회비(2011년~20만원, 2020~2022년 10만원)는 자동 설정됩니다. 2021년 이후는 통장 기준으로 계산됩니다.</div>
      </div>

      {result && <div className="p-3 bg-emerald-50 border border-emerald-200 rounded-md text-sm text-emerald-800">{result.type === 'roster' ? `✅ 회원명단 반영 — 신규 ${result.added}명, 갱신 ${result.updated}명` : `✅ 미납장부 반영 — 현회원 ${result.members}명에 부착, 이관 입금 ${result.tx}건 (현회원 아님 ${result.skipped||0}명 제외)`}</div>}

      {/* Roster */}
      <div className="bg-white rounded-lg border border-gray-200 p-4">
        <div className="flex items-center gap-2 mb-2"><span className="w-6 h-6 rounded-full bg-indigo-600 text-white text-sm flex items-center justify-center font-bold">①</span><h3 className="font-semibold">최신 회원명단 가져오기 <span className="ml-1 text-xs px-2 py-0.5 rounded-full bg-emerald-100 text-emerald-700">평소 사용</span></h3></div>
        <p className="text-sm text-gray-600 mb-3">미납액 없는 순수 명단. 신규 회원을 추가하고 기존 회원 정보를 갱신합니다. 8자리 날짜(19820203) 자동 변환.</p>
        <input ref={rosterRef} type="file" accept=".xlsx,.xls,.csv" onChange={handle(parseRoster)} className="hidden" />
        <button onClick={() => rosterRef.current?.click()} className="px-4 py-2 bg-indigo-600 text-white rounded-md text-sm hover:bg-indigo-700"><Upload className="w-4 h-4 inline mr-1" /> 회원명단 파일 선택</button>
        {rosterPrev && (
          <div className="mt-3 border border-indigo-200 rounded-lg p-3">
            <div className="flex items-center justify-between mb-2">
              <div className="font-medium text-sm">미리보기 — {rosterPrev.fileName} (현회원 {rosterPrev.activeN}명 · 탈회 {rosterPrev.wN}명 · 확인필요 {rosterPrev.reviewN}명)</div>
              <div className="flex gap-2">
                <button onClick={() => setRosterPrev(null)} className="px-3 py-1.5 text-gray-600 text-sm">취소</button>
                <button onClick={() => { const r = onImportRoster(rosterPrev.rows, 'merge'); setResult({ type:'roster', ...r }); setRosterPrev(null); }} className="px-3 py-1.5 bg-emerald-600 text-white rounded-md text-sm">이 명단으로 반영</button>
                
              </div>
            </div>
            <div className="mb-2 p-2 bg-blue-50 border border-blue-200 rounded text-xs text-blue-800"><b>"전출명단"</b> 아래 회원만 <b>탈회(회색)</b>로 처리됩니다. 전출명단 위(개업·전입일이 있는 회원)는 F열 내용과 무관하게 모두 <b>현회원</b>으로 관리됩니다.</div>
            <div className="overflow-x-auto max-h-80 border border-gray-200 rounded">
              <table className="w-full text-xs">
                <thead className="bg-gray-50 sticky top-0"><tr><th className="px-2 py-1 text-left border-r">등록번호</th><th className="px-2 py-1 text-left border-r">성명</th><th className="px-2 py-1 text-left border-r">개업/전입일</th><th className="px-2 py-1 text-left border-r">상태</th><th className="px-2 py-1 text-left border-r">사유</th></tr></thead>
                <tbody>{rosterPrev.rows.slice(0,120).map((r,i) => <tr key={i} className={`border-t border-gray-100 ${r.withdrawn ? 'bg-gray-50 text-gray-400' : (r.review ? 'bg-amber-50' : '')}`}><td className="px-2 py-1 border-r text-gray-500">{r.regNo}</td><td className="px-2 py-1 border-r font-medium">{r.name}</td><td className="px-2 py-1 border-r">{r.joinDate}{r.joinDate && (monthOf(r.joinDate)>=7?' (하)':' (상)')}</td><td className="px-2 py-1 border-r">{r.withdrawn ? <span className="text-rose-500">탈회</span> : (r.review ? <span className="text-amber-600 font-medium">확인필요</span> : <span className="text-emerald-600">현회원</span>)}</td><td className="px-2 py-1 border-r text-gray-500">{r.reason || (r.leaveDate && r.withdrawn ? r.leaveDate : '')}</td></tr>)}</tbody>
              </table>
              {rosterPrev.rows.length > 120 && <div className="text-center text-xs text-gray-500 py-2">... 외 {rosterPrev.rows.length-120}명</div>}
            </div>
          </div>
        )}
      </div>

      {/* Ledger */}
      <div className="bg-white rounded-lg border border-gray-200 p-4">
        <div className="flex items-center gap-2 mb-2"><span className="w-6 h-6 rounded-full bg-amber-600 text-white text-sm flex items-center justify-center font-bold">②</span><h3 className="font-semibold">옛 미납장부 가져오기 <span className="ml-1 text-xs px-2 py-0.5 rounded-full bg-gray-100 text-gray-500">최초 1회만</span></h3></div>
        <p className="text-sm text-gray-600 mb-3">2010이전 미수금액 + 연도별 미납액이 있는 장부입니다. 2021년 이후는 통장 기준이라 자동 제외됩니다.</p>
        <input ref={ledgerRef} type="file" accept=".xlsx,.xls,.csv" onChange={handle(parseLedger)} className="hidden" />
        <button onClick={() => ledgerRef.current?.click()} className="px-4 py-2 bg-amber-600 text-white rounded-md text-sm hover:bg-amber-700"><Upload className="w-4 h-4 inline mr-1" /> 미납장부 파일 선택</button>
        {ledgerPrev && (
          <div className="mt-3 border border-amber-200 rounded-lg p-3">
            <div className="flex items-center justify-between mb-2">
              <div className="font-medium text-sm">미리보기 — {ledgerPrev.fileName} ({ledgerPrev.rows.length}명)</div>
              <div className="flex gap-2">
                <button onClick={() => setLedgerPrev(null)} className="px-3 py-1.5 text-gray-600 text-sm">취소</button>
                <button onClick={() => { const r = onImportLedger(ledgerPrev.rows); setResult({ type:'ledger', ...r }); setLedgerPrev(null); }} className="px-3 py-1.5 bg-emerald-600 text-white rounded-md text-sm">가져오기</button>
              </div>
            </div>
            <div className="mb-2 p-2 bg-blue-50 border border-blue-200 rounded text-xs text-blue-800">2010이전 미수 + <b>{LEDGER_CUTOFF_YEAR}년까지</b>만 반영. 2021년 이후는 통장으로 계산(이중계산 방지).</div>
            <div className="overflow-x-auto max-h-80 border border-gray-200 rounded">
              <table className="w-full text-xs">
                <thead className="bg-gray-50 sticky top-0"><tr><th className="px-2 py-1 text-left border-r">등록번호</th><th className="px-2 py-1 text-left border-r">성명</th><th className="px-2 py-1 text-left border-r">개업일</th>{ledgerPrev.hasPrior && <th className="px-2 py-1 text-right border-r bg-purple-50">2010이전</th>}{ledgerPrev.yearCols.map(y => <th key={y} className={`px-2 py-1 text-right border-r ${y > LEDGER_CUTOFF_YEAR ? 'bg-gray-100 text-gray-400 line-through' : ''}`}>{y}</th>)}</tr></thead>
                <tbody>{ledgerPrev.rows.slice(0,100).map((r,i) => <tr key={i} className="border-t border-gray-100"><td className="px-2 py-1 border-r text-gray-500">{r.regNo}</td><td className="px-2 py-1 border-r font-medium">{r.name}</td><td className="px-2 py-1 border-r">{r.joinDate}</td>{ledgerPrev.hasPrior && <td className="px-2 py-1 border-r text-right bg-purple-50/50 text-purple-700">{r.priorArrears?fmt(r.priorArrears):<span className="text-gray-300">-</span>}</td>}{ledgerPrev.yearCols.map(y => <td key={y} className={`px-2 py-1 border-r text-right ${y > LEDGER_CUTOFF_YEAR ? 'bg-gray-50 text-gray-300' : ''}`}>{y > LEDGER_CUTOFF_YEAR ? '제외' : (r.years[y]?fmt(r.years[y]):<span className="text-gray-300">-</span>)}</td>)}</tr>)}</tbody>
              </table>
              {ledgerPrev.rows.length > 100 && <div className="text-center text-xs text-gray-500 py-2">... 외 {ledgerPrev.rows.length-100}명</div>}
            </div>
          </div>
        )}
      </div>

      <div className="bg-white rounded-lg border border-gray-200 p-4 flex items-center justify-between">
        <div className="text-sm text-gray-600">현재 등록된 회원: <b>{memberCount}명</b></div>
        {memberCount > 0 && <button onClick={onReset} className="px-3 py-1.5 text-rose-600 text-sm hover:bg-rose-50 rounded-md border border-rose-200"><Trash2 className="w-4 h-4 inline mr-1" /> 전체 초기화</button>}
      </div>
    </div>
  );
}

// ============ Fees Tab ============
function FeesTab({ fees, setFeeForYear, deleteFee }) {
  const [ny, setNy] = useState(new Date().getFullYear());
  const [na, setNa] = useState(200000);
  const [edits, setEdits] = useState({});
  const sorted = [...fees].sort((a, b) => a.year - b.year);
  return (
    <div className="bg-white rounded-lg border border-gray-200 p-4 max-w-2xl">
      <h3 className="font-semibold mb-3">연도별 회비</h3>
      <div className="mb-4 p-3 bg-emerald-50 border border-emerald-200 rounded-lg text-sm text-emerald-800">회비 규칙 자동 적용: <b>2011년부터 20만원</b>, <b>2020~2022년만 10만원</b>. 하반기 입회 첫 해는 자동 절반. 특정 연도만 예외로 바꾸려면 아래에서 수정하세요.</div>
      <div className="text-sm font-medium mb-2">개별 연도 수정 (필요 시)</div>
      <div className="flex gap-2 mb-4 items-end">
        <div><label className="text-xs text-gray-600">연도</label><input type="number" value={ny} onChange={e=>setNy(e.target.value)} className="block border border-gray-300 rounded px-2 py-1.5 w-24 text-sm" /></div>
        <div><label className="text-xs text-gray-600">회비(원)</label><input type="number" value={na} onChange={e=>setNa(e.target.value)} className="block border border-gray-300 rounded px-2 py-1.5 w-36 text-sm" /></div>
        <button onClick={() => { if (ny && na) setFeeForYear(Number(ny), Number(na)); }} className="px-3 py-1.5 bg-indigo-600 text-white rounded-md text-sm">추가/수정</button>
      </div>
      <table className="w-full text-sm">
        <thead><tr className="border-b border-gray-200 text-left text-gray-600"><th className="px-2 py-2">연도</th><th className="px-2 py-2">회비</th><th className="px-2 py-2"></th></tr></thead>
        <tbody>{sorted.map(f => (
          <tr key={f.year} className="border-b border-gray-100">
            <td className="px-2 py-2 font-medium">{f.year}년</td>
            <td className="px-2 py-2"><div className="flex items-center gap-2"><input type="number" defaultValue={f.amount} onChange={e=>setEdits({...edits, [f.year]: e.target.value})} className="border border-gray-300 rounded px-2 py-1 w-32 text-sm" /><span className="text-gray-500">원</span>{edits[f.year] !== undefined && Number(edits[f.year]) !== f.amount && <button onClick={() => { setFeeForYear(f.year, Number(edits[f.year])); const e2 = {...edits}; delete e2[f.year]; setEdits(e2); }} className="text-emerald-600 text-xs px-2 py-1 hover:bg-emerald-50 rounded">저장</button>}</div></td>
            <td className="px-2 py-2 text-right"><button onClick={() => deleteFee(f.year)} className="text-gray-500 hover:text-rose-600"><Trash2 className="w-4 h-4" /></button></td>
          </tr>
        ))}</tbody>
      </table>
    </div>
  );
}

// ============ Upload Tab ============
function UploadTab({ tx, setTx, members, aliasMap, regIndex, nameIndex, uploadMeta, setUploadMeta, onAssign, onDeleteTx }) {
  const [preview, setPreview] = useState(null);
  const [stats, setStats] = useState(null);
  const [confirmClear, setConfirmClear] = useState(false);
  const fileRef = useRef();

  const handleFile = async (e) => {
    const file = e.target.files[0]; if (!file) return;
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { cellDates: true });
      const data = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: '' });
      let hi = -1, best = 0;
      for (let i = 0; i < Math.min(data.length, 25); i++) { const m = detectColumns(data[i]||[]); const sc = (m.date>=0?1:0)+(m.amount>=0?1:0)+(m.name>=0?1:0)+(m.desc>=0?1:0); if (sc > best) { best = sc; hi = i; } }
      if (hi < 0) hi = 0;
      const headers = data[hi] || [];
      const rows = data.slice(hi + 1).filter(r => r && r.some(c => c !== '' && c != null));
      setPreview({ rows, headers, mapping: detectColumns(headers), fileName: file.name, headerIdx: hi });
    } catch (err) { alert('파일 읽기 실패: ' + err.message + '\n(은행 파일을 엑셀에서 xlsx로 다시 저장해 보세요)'); }
    e.target.value = '';
  };

  const confirmImport = () => {
    const { rows, mapping } = preview;
    if (mapping.date < 0 || mapping.amount < 0) return alert('날짜/입금액 컬럼을 선택하세요');
    const nt = []; let dup = 0, skip = 0, matched = 0;
    const keys = new Set(tx.map(t => t.dedupKey || `${t.date}|${t.time||''}|${t.amount}|${t.depositorName}`));
    for (const row of rows) {
      const { date: d, time } = parseDateTime(row[mapping.date]);
      const a = parseAmount(row[mapping.amount]);
      const rawName = String(row[mapping.name] || '').trim();
      const desc = mapping.desc >= 0 ? String(row[mapping.desc] || '').trim() : '';
      if (!d || a <= 0) { skip++; continue; }
      const key = `${d}|${time}|${a}|${rawName}|${desc}`;
      if (keys.has(key)) { dup++; continue; }
      keys.add(key);
      const r = smartMatch(rawName, members, aliasMap, regIndex, nameIndex);
      const mid = r && r.memberId ? r.memberId : null;
      if (mid) matched++;
      nt.push({ id: uid(), date: d, time, amount: a, depositorName: rawName, description: desc, memberId: mid, manualAssign: false, dedupKey: key });
    }
    setTx(prev => [...prev, ...nt]);
    setStats({ added: nt.length, dup, skip, matched, unmatched: nt.length - matched });
    const all = [...tx, ...nt].map(t => `${t.date} ${t.time||'00:00:00'}`).filter(Boolean).sort();
    setUploadMeta({ uploadedAt: new Date().toISOString(), fileName: preview.fileName, minDate: all[0]||'', maxDate: all[all.length-1]||'', count: tx.length + nt.length });
    setPreview(null);
  };

  const clearAll = () => { setTx([]); setStats(null); setConfirmClear(false); try { window.storage.set(SK.T, JSON.stringify([])); } catch {} };
  const setMap = (k, i) => setPreview({ ...preview, mapping: { ...preview.mapping, [k]: i } });

  return (
    <div>
      <div className="bg-white rounded-lg border border-gray-200 p-4 mb-3">
        <h3 className="font-semibold mb-2">거래내역 업로드</h3>
        <p className="text-sm text-gray-600 mb-3">은행에서 받은 통장 엑셀을 업로드하세요. 입금만 추출하고, 초 단위 중복을 자동 제거하며, 입금자명에서 이름을 추출해 회원과 매칭합니다.</p>
        <div className="flex items-center gap-2 flex-wrap">
          <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" onChange={handleFile} className="hidden" />
          <button onClick={() => fileRef.current?.click()} className="px-4 py-2 bg-indigo-600 text-white rounded-md text-sm hover:bg-indigo-700"><Upload className="w-4 h-4 inline mr-1" /> 파일 선택</button>
          <span className="text-sm text-gray-500">현재 등록된 거래: <b>{tx.length}건</b></span>
          {tx.length > 0 && <button onClick={() => setConfirmClear(true)} className="ml-auto px-3 py-1.5 text-rose-600 text-sm hover:bg-rose-50 rounded-md"><Trash2 className="w-4 h-4 inline mr-1" /> 전체 삭제</button>}
        </div>
        {confirmClear && <div className="mt-3 p-3 bg-rose-50 border border-rose-200 rounded-md"><p className="text-sm text-rose-800 mb-2">거래내역 {tx.length}건을 모두 삭제할까요? (회원·회비 유지)</p><div className="flex gap-2"><button onClick={clearAll} className="px-3 py-1.5 bg-rose-600 text-white rounded-md text-sm">삭제</button><button onClick={() => setConfirmClear(false)} className="px-3 py-1.5 text-gray-600 text-sm">취소</button></div></div>}
        {stats && <div className="mt-3 p-3 bg-emerald-50 border border-emerald-200 rounded-md text-sm"><b className="text-emerald-800">업로드 완료</b><div className="mt-1 text-gray-700">추가 {stats.added} / 자동매칭 {stats.matched} / 미매칭 {stats.unmatched} / 중복제외 {stats.dup} / 무효행 {stats.skip}</div></div>}
        {uploadMeta && <div className="mt-3 p-3 bg-slate-50 border border-slate-200 rounded-md text-sm"><div className="flex flex-wrap gap-x-6 gap-y-1 text-gray-700"><div>📅 <b>최근 업로드</b>: {fmtDateTime(uploadMeta.uploadedAt)}</div><div>📄 {uploadMeta.fileName || '-'}</div><div>🧾 자료 기간: <b>{uploadMeta.minDate?.slice(0,16)}</b> ~ <b>{uploadMeta.maxDate?.slice(0,16)}</b></div><div>총 {uploadMeta.count}건</div></div><div className="text-xs text-gray-500 mt-1">반영된 통장 내역의 마지막 시점: <b>{uploadMeta.maxDate?.slice(0,10)}</b>. 다음 업로드 전까지 이 상태로 조회됩니다.</div></div>}
      </div>

      {preview && (
        <div className="bg-white rounded-lg border border-indigo-200 p-4">
          <div className="flex items-center justify-between mb-3">
            <h3 className="font-semibold">미리보기 — {preview.fileName}</h3>
            <div className="flex gap-2"><button onClick={() => setPreview(null)} className="px-3 py-1.5 text-gray-600 text-sm">취소</button><button onClick={confirmImport} className="px-3 py-1.5 bg-emerald-600 text-white rounded-md text-sm"><Save className="w-4 h-4 inline mr-1" /> 가져오기 ({preview.rows.length}건)</button></div>
          </div>
          <div className="grid grid-cols-2 md:grid-cols-4 gap-2 mb-3 p-3 bg-gray-50 rounded-md">
            <MapSel label="날짜 *" headers={preview.headers} value={preview.mapping.date} onChange={v => setMap('date', v)} />
            <MapSel label="입금액 *" headers={preview.headers} value={preview.mapping.amount} onChange={v => setMap('amount', v)} />
            <MapSel label="입금자명" headers={preview.headers} value={preview.mapping.name} onChange={v => setMap('name', v)} />
            <MapSel label="적요" headers={preview.headers} value={preview.mapping.desc} onChange={v => setMap('desc', v)} />
          </div>
          <div className="text-xs text-gray-500 mb-2">입금액 0 초과 행만 가져옵니다(출금 제외) · 초 단위로 중복 방지 · 괄호/세무사 등 부가정보는 매칭 시 자동 제거{preview.headerIdx > 0 && ` · 상단 ${preview.headerIdx}개 행(계좌정보) 건너뜀`}</div>
          <div className="overflow-x-auto max-h-80 border border-gray-200 rounded">
            <table className="w-full text-xs"><thead className="bg-gray-50 sticky top-0"><tr>{preview.headers.map((h,i)=><th key={i} className="px-2 py-1 text-left border-r">{String(h)}</th>)}</tr></thead>
            <tbody>{preview.rows.slice(0,50).map((row,i)=><tr key={i} className="border-t border-gray-100">{preview.headers.map((_,j)=><td key={j} className="px-2 py-1 border-r">{String(row[j] != null ? row[j] : '')}</td>)}</tr>)}</tbody></table>
            {preview.rows.length > 50 && <div className="text-center text-xs text-gray-500 py-2">... 외 {preview.rows.length-50}건</div>}
          </div>
        </div>
      )}

      {tx.length > 0 && !preview && <TxList tx={tx} members={members} onAssign={onAssign} onDeleteTx={onDeleteTx} />}
    </div>
  );
}

function TxList({ tx, members, onAssign, onDeleteTx }) {
  const [q, setQ] = useState('');
  const [filter, setFilter] = useState('all');
  const [assignTarget, setAssignTarget] = useState(null);
  const [search, setSearch] = useState('');
  const [saveAlias, setSaveAlias] = useState(true);
  const mName = (id) => members.find(m => m.id === id)?.name || '';
  const isEarly = (t) => { const m = members.find(mm => mm.id === t.memberId); return !!(m && m.joinDate && t.date && yearOf(t.date) < yearOf(m.joinDate)); };
  const filtered = useMemo(() => {
    let list = [...tx];
    if (filter === 'matched') list = list.filter(t => t.memberId);
    else if (filter === 'unmatched') list = list.filter(t => !t.memberId);
    else if (filter === 'early') list = list.filter(t => isEarly(t));
    if (q.trim()) { const kw = q.trim().toLowerCase(); list = list.filter(t => (t.depositorName||'').toLowerCase().includes(kw) || (t.description||'').toLowerCase().includes(kw) || (t.date||'').includes(kw) || String(t.amount).includes(kw.replace(/,/g,'')) || mName(t.memberId).toLowerCase().includes(kw)); }
    return list.sort((a,b)=>new Date(b.date)-new Date(a.date));
  }, [tx, q, filter, members]);
  const matchedN = tx.filter(t => t.memberId).length;
  const earlyN = tx.filter(isEarly).length;
  const sum = filtered.reduce((s,t)=>s+t.amount,0);
  const memberList = members.filter(m => !search || m.name.includes(search) || (m.depositorName||'').includes(search) || (m.regNo||'').includes(search));
  return (
    <div className="bg-white rounded-lg border border-gray-200 p-4 mt-3">
      <div className="flex items-center justify-between mb-3 flex-wrap gap-2">
        <h3 className="font-semibold">전체 거래내역 <span className="text-gray-400 font-normal">({filtered.length}/{tx.length}건)</span></h3>
        <div className="flex items-center gap-2 flex-wrap">
          <div className="inline-flex rounded-md border border-gray-300 overflow-hidden text-xs">
            <button onClick={()=>setFilter('all')} className={`px-3 py-1.5 ${filter==='all'?'bg-indigo-600 text-white':'bg-white text-gray-600'}`}>전체 {tx.length}</button>
            <button onClick={()=>setFilter('matched')} className={`px-3 py-1.5 border-l border-gray-300 ${filter==='matched'?'bg-emerald-600 text-white':'bg-white text-emerald-700'}`}>매칭 {matchedN}</button>
            <button onClick={()=>setFilter('unmatched')} className={`px-3 py-1.5 border-l border-gray-300 ${filter==='unmatched'?'bg-amber-500 text-white':'bg-white text-amber-600'}`}>미매칭 {tx.length-matchedN}</button>
            <button onClick={()=>setFilter('early')} className={`px-3 py-1.5 border-l border-gray-300 ${filter==='early'?'bg-rose-500 text-white':'bg-white text-rose-600'}`}>입금일빠름 {earlyN}</button>
          </div>
          <input value={q} onChange={e=>setQ(e.target.value)} placeholder="입금자·회원·적요·날짜·금액 검색" className="px-3 py-1.5 border border-gray-300 rounded-md text-sm w-56" />
        </div>
      </div>
      <div className="text-xs text-gray-500 mb-2">표시 {filtered.length}건 합계: <b>{fmt(sum)}원</b>{earlyN > 0 && <span className="text-rose-600"> · ⚠ 입금일이 개업/전입일보다 빠른 거래 {earlyN}건(개업/전입일 확인 필요)</span>} · 미매칭 행에서 바로 연결 가능</div>
      <div className="overflow-x-auto max-h-[30rem]">
        <table className="w-full text-sm">
          <thead className="bg-gray-50 sticky top-0"><tr><th className="px-2 py-2 text-left">날짜</th><th className="px-2 py-2 text-left">입금자명</th><th className="px-2 py-2 text-right">금액</th><th className="px-2 py-2 text-left">적요</th><th className="px-2 py-2 text-left">매칭/연결</th></tr></thead>
          <tbody>
            {filtered.length === 0 && <tr><td colSpan={5} className="text-center text-gray-400 py-8">결과 없음</td></tr>}
            {filtered.map(t => {
              const mn = mName(t.memberId);
              const cands = t.memberId ? [] : suggestCandidates(t.depositorName, members);
              return (
                <tr key={t.id} className={`border-b border-gray-100 hover:bg-gray-50 ${!t.memberId ? 'bg-amber-50/20' : (isEarly(t) ? 'bg-rose-50/40' : '')}`}>
                  <td className="px-2 py-1.5 whitespace-nowrap">{t.date}<span className="text-gray-400 text-xs ml-1">{t.time||''}</span></td>
                  <td className="px-2 py-1.5">{t.depositorName}</td>
                  <td className="px-2 py-1.5 text-right font-mono">{fmt(t.amount)}</td>
                  <td className="px-2 py-1.5 text-xs text-gray-500">{t.description}</td>
                  <td className="px-2 py-1.5">
                    {mn ? <span className="text-emerald-700 font-medium">{mn}{isEarly(t) && <span className="ml-1 px-1 py-0.5 bg-rose-100 text-rose-700 rounded text-[10px]" title="입금일이 개업/전입일보다 빠릅니다. 개업/전입일을 확인하세요.">⚠확인</span>}</span> : (
                      <div className="flex flex-wrap items-center gap-1">
                        {cands.slice(0,3).map(m => <button key={m.id} onClick={() => onAssign(t.id, m.id, true)} className="px-2 py-1 bg-emerald-600 text-white rounded text-xs hover:bg-emerald-700">{m.name}{m.regNo?`(${m.regNo})`:''} 연결</button>)}
                        <button onClick={() => { setAssignTarget(t); setSearch(''); setSaveAlias(true); }} className="px-2 py-1 border border-gray-300 rounded text-xs text-gray-600 hover:bg-gray-50">직접 선택</button>
                        <button onClick={() => onDeleteTx(t.id)} className="p-1 text-gray-400 hover:text-rose-600"><Trash2 className="w-3.5 h-3.5" /></button>
                      </div>
                    )}
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
      {assignTarget && (
        <Modal onClose={() => setAssignTarget(null)} title="회원 연결">
          <div className="text-sm text-gray-600 mb-3"><b>{assignTarget.date}</b> · <b>{assignTarget.depositorName}</b> · <b>{fmt(assignTarget.amount)}원</b></div>
          <input placeholder="회원 검색(이름/등록번호)" value={search} onChange={e=>setSearch(e.target.value)} className="w-full border border-gray-300 rounded px-3 py-2 text-sm mb-2" />
          <div className="max-h-60 overflow-y-auto border border-gray-200 rounded mb-3">
            {memberList.map(m => <button key={m.id} onClick={() => { onAssign(assignTarget.id, m.id, saveAlias); setAssignTarget(null); }} className="block w-full text-left px-3 py-2 text-sm border-b border-gray-100 last:border-b-0 hover:bg-indigo-50"><div className="font-medium">{m.name} <span className="text-xs text-gray-400">{m.regNo}</span></div><div className="text-xs text-gray-500">입금자명: {m.depositorName}{m.aliases?.length ? ` / 별칭: ${m.aliases.join(', ')}` : ''}</div></button>)}
            {memberList.length === 0 && <div className="p-4 text-center text-gray-400 text-sm">일치 회원 없음</div>}
          </div>
          <label className="flex items-center gap-2 text-sm"><input type="checkbox" checked={saveAlias} onChange={e=>setSaveAlias(e.target.checked)} />"<b>{assignTarget.depositorName}</b>"을 별칭으로 저장</label>
        </Modal>
      )}
    </div>
  );
}

function MapSel({ label, headers, value, onChange }) {
  return <div><label className="text-xs text-gray-600 block mb-1">{label}</label><select value={value} onChange={e=>onChange(Number(e.target.value))} className="border border-gray-300 rounded px-2 py-1.5 w-full text-sm"><option value={-1}>선택 안 함</option>{headers.map((h,i)=><option key={i} value={i}>{String(h)}</option>)}</select></div>;
}

// ============ Overdue Tab ============
function OverdueTab({ members, allocations, config, setConfig }) {
  const [threshold, setThreshold] = useState(1);
  const [includePartial, setIncludePartial] = useState(true);
  const [showMsg, setShowMsg] = useState(false);
  const [copyState, setCopyState] = useState({});
  const overdue = useMemo(() => {
    const list = [];
    for (const m of members) {
      if (isWithdrawn(m)) continue;
      const alloc = allocations.get(m.id) || [];
      const unpaid = alloc.filter(yr => yr.required > 0 && (includePartial ? yr.allocated < yr.required : yr.allocated <= 0));
      if (unpaid.length < threshold) continue;
      const sum = unpaid.reduce((s, yr) => s + (yr.required - yr.allocated), 0);
      list.push({ member: m, unpaidYears: unpaid, totalUnpaid: sum });
    }
    return list.sort((a, b) => b.totalUnpaid - a.totalUnpaid);
  }, [members, allocations, threshold, includePartial]);
  const renderMsg = (it) => config.msgTemplate.replace(/\{이름\}/g, it.member.name).replace(/\{단체명\}/g, config.groupName).replace(/\{연도\}/g, it.unpaidYears.map(y => y.isPrior ? '2020까지이월' : y.year).join(', ')).replace(/\{금액\}/g, fmt(it.totalUnpaid));
  const copy = (text, key) => { navigator.clipboard.writeText(text).then(() => { setCopyState(s => ({ ...s, [key]: true })); setTimeout(() => setCopyState(s => ({ ...s, [key]: false })), 1500); }); };
  return (
    <div className="space-y-4">
      <div className="bg-white rounded-lg border border-gray-200 p-4">
        <h3 className="font-semibold mb-3">미납자 필터 <span className="text-xs text-gray-400 font-normal">(현 회원만 · 탈회원 제외)</span></h3>
        <div className="flex flex-wrap gap-4 items-center">
          <div><label className="text-xs text-gray-600 block mb-1">미납 기준</label><select value={threshold} onChange={e=>setThreshold(Number(e.target.value))} className="border border-gray-300 rounded px-2 py-1.5 text-sm">{[1,2,3,4,5].map(n=><option key={n} value={n}>{n}년 이상 미납</option>)}</select></div>
          <label className="flex items-center gap-2 text-sm"><input type="checkbox" checked={includePartial} onChange={e=>setIncludePartial(e.target.checked)} />일부납도 미납 포함</label>
          <div className="ml-auto text-sm text-gray-600">대상: <b className="text-rose-600">{overdue.length}명</b> / 미납총액: <b className="text-rose-600">{fmt(overdue.reduce((s,o)=>s+o.totalUnpaid,0))}원</b></div>
        </div>
      </div>
      <div className="bg-white rounded-lg border border-gray-200 p-4">
        <div className="flex items-center justify-between mb-3 flex-wrap gap-2"><h3 className="font-semibold">미납자 명단</h3><button onClick={() => setShowMsg(!showMsg)} className="px-3 py-1.5 bg-amber-600 text-white rounded-md text-sm hover:bg-amber-700"><Bell className="w-4 h-4 inline mr-1" /> 메시지 {showMsg ? '닫기' : '생성'}</button></div>
        {overdue.length === 0 ? <div className="text-center text-gray-400 py-8">조건에 해당하는 미납자가 없습니다 🎉</div> : (
          <table className="w-full text-sm"><thead><tr className="border-b border-gray-200 text-left text-gray-600"><th className="px-2 py-2">등록번호</th><th className="px-2 py-2">성명</th><th className="px-2 py-2">연락처</th><th className="px-2 py-2">미납 연도</th><th className="px-2 py-2 text-right">미납액</th></tr></thead>
          <tbody>{overdue.map(it => <tr key={it.member.id} className="border-b border-gray-100"><td className="px-2 py-2 text-xs text-gray-500">{it.member.regNo||'-'}</td><td className="px-2 py-2 font-medium">{it.member.name}</td><td className="px-2 py-2">{it.member.phone||'-'}</td><td className="px-2 py-2 text-xs">{it.unpaidYears.map(yr => <span key={String(yr.year)} className={`inline-block px-1.5 py-0.5 rounded mr-1 mb-1 ${yr.allocated>0?'bg-amber-100 text-amber-800':'bg-rose-100 text-rose-800'}`}>{yr.isPrior?'2020까지이월':yr.year+'년'} ({fmt(yr.required-yr.allocated)})</span>)}</td><td className="px-2 py-2 text-right font-mono font-semibold text-rose-600">{fmt(it.totalUnpaid)}</td></tr>)}</tbody></table>
        )}
      </div>
      {showMsg && overdue.length > 0 && (
        <div className="bg-white rounded-lg border border-amber-300 p-4">
          <h3 className="font-semibold mb-3">💬 알림 메시지</h3>
          <div className="mb-3"><label className="text-xs text-gray-600 block mb-1">템플릿 (변수: {`{이름} {단체명} {연도} {금액}`})</label><textarea value={config.msgTemplate} onChange={e=>setConfig({ ...config, msgTemplate: e.target.value })} className="w-full border border-gray-300 rounded px-3 py-2 text-sm font-mono" rows={5} /></div>
          <div className="flex gap-2 mb-3">
            <button onClick={() => copy(overdue.map(it => `[${it.member.name}${it.member.phone?' '+it.member.phone:''}]\n${renderMsg(it)}`).join('\n\n---\n\n'), '__all__')} className="px-3 py-1.5 bg-indigo-600 text-white rounded-md text-sm"><Copy className="w-4 h-4 inline mr-1" /> 전체 메시지 복사{copyState['__all__'] && <span className="ml-1 text-xs">✓</span>}</button>
            <button onClick={() => copy(overdue.map(it => it.member.phone).filter(Boolean).join(', '), '__ph__')} className="px-3 py-1.5 border border-gray-300 rounded-md text-sm"><Copy className="w-4 h-4 inline mr-1" /> 전화번호 복사{copyState['__ph__'] && <span className="ml-1 text-xs text-emerald-600">✓</span>}</button>
          </div>
          <div className="space-y-2 max-h-96 overflow-y-auto">{overdue.map(it => <div key={it.member.id} className="border border-gray-200 rounded p-3"><div className="flex justify-between items-start mb-2"><div className="text-sm"><b>{it.member.name}</b>{it.member.phone && <span className="text-gray-500 ml-2">{it.member.phone}</span>}</div><button onClick={() => copy(renderMsg(it), it.member.id)} className="text-xs px-2 py-1 bg-gray-100 hover:bg-gray-200 rounded">{copyState[it.member.id] ? '✓ 복사됨' : '복사'}</button></div><pre className="text-sm whitespace-pre-wrap bg-gray-50 p-2 rounded font-sans">{renderMsg(it)}</pre></div>)}</div>
        </div>
      )}
    </div>
  );
}

// ============ Withdrawn Tab ============
function WithdrawnTab({ members, allocations, onUpdate, onDelete }) {
  const [search, setSearch] = useState('');
  const list = members.filter(m => !search || (m.name||'').includes(search) || (m.depositorName||'').includes(search) || (m.regNo||'').includes(search) || (m.note||'').includes(search));
  const sorted = [...list].sort((a, b) => String(b.leaveDate||'').localeCompare(String(a.leaveDate||'')));
  return (
    <div className="bg-white rounded-lg border border-gray-200 p-4">
      <div className="flex items-center justify-between mb-2 flex-wrap gap-2">
        <div className="flex items-center gap-2">
          <input placeholder="이름/등록번호/사유 검색" value={search} onChange={e=>setSearch(e.target.value)} className="px-3 py-1.5 border border-gray-300 rounded-md text-sm" />
          <span className="text-sm text-gray-500">탈회 {list.length}명</span>
        </div>
      </div>
      <div className="text-xs text-gray-400 mb-3">전출·폐업·휴업·사망 등으로 제외된 회원입니다. 납입현황·회원관리·미납자 관리(현회원 209명 기준)에서는 집계되지 않습니다. 오분류라면 "복귀"로 현회원으로 되돌릴 수 있습니다.</div>
      <div className="overflow-auto max-h-[70vh] border border-gray-100 rounded">
        <table className="w-full text-sm">
          <thead><tr className="border-b border-gray-200 text-left text-gray-600 bg-gray-50 sticky top-0 z-10">
            <th className="px-2 py-2">등록번호</th><th className="px-2 py-2">성명</th><th className="px-2 py-2">개업/전입</th><th className="px-2 py-2">전출일</th><th className="px-2 py-2">사유</th><th className="px-2 py-2 text-right">참고 미납</th><th className="px-2 py-2"></th>
          </tr></thead>
          <tbody>
            {sorted.length === 0 && <tr><td colSpan={7} className="text-center text-gray-400 py-8">탈회 회원이 없습니다</td></tr>}
            {sorted.map(m => {
              const alloc = allocations.get(m.id) || [];
              const unpaid = alloc.reduce((sm, yr) => sm + (yr.required > yr.allocated ? yr.required - yr.allocated : 0), 0);
              return (
                <tr key={m.id} className="border-b border-gray-100 hover:bg-gray-50 text-gray-500">
                  <td className="px-2 py-2 text-xs">{m.regNo || '-'}</td>
                  <td className="px-2 py-2 font-medium text-gray-700">{m.name}</td>
                  <td className="px-2 py-2 text-xs">{m.joinDate || '-'}</td>
                  <td className="px-2 py-2 text-xs text-rose-500">{m.leaveDate || '-'}</td>
                  <td className="px-2 py-2 text-xs">{m.note || '-'}</td>
                  <td className="px-2 py-2 text-right text-xs">{unpaid > 0 ? fmt(unpaid) : '\u2014'}</td>
                  <td className="px-2 py-2 text-right whitespace-nowrap">
                    <button onClick={() => onUpdate(m.id, { leaveDate: '' })} className="px-2 py-1 text-xs text-emerald-600 hover:bg-emerald-50 rounded" title="현회원으로 복귀">복귀</button>
                    <button onClick={() => onDelete(m.id)} className="p-1 text-gray-400 hover:text-rose-600"><Trash2 className="w-4 h-4" /></button>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ============ Cell Modal ============
function CellModal({ member, yearData, onClose, onUnassign }) {
  const st = cellStatus(yearData);
  const jy = yearOf(member.joinDate);
  const isPrior = yearData.isPrior;
  const half = !isPrior && yearData.year === jy && member.joinDate && monthOf(member.joinDate) >= 7;
  return (
    <Modal onClose={onClose} title={`${member.name} · ${isPrior ? '2020년까지 이월 미납' : yearData.year + '년 회비'} 상세`}>
      {isPrior && <div className="mb-3 text-xs text-purple-700 bg-purple-50 p-2 rounded">ℹ️ 2020년까지의 이월 미납(선납 반영)입니다. 총 미납액에 포함됩니다.</div>}
      {half && <div className="mb-3 text-xs text-indigo-700 bg-indigo-50 p-2 rounded">ℹ️ 하반기({monthOf(member.joinDate)}월) 입회 첫 해로 의무 회비가 절반 적용되었습니다.</div>}
      {!isPrior && yearData.payments.some(p => p.date && yearOf(p.date) < jy) && <div className="mb-3 text-xs text-rose-700 bg-rose-50 p-2 rounded">⚠ 개업/전입일({member.joinDate})보다 이른 연도의 입금이 있습니다. 엑셀 입력 오류일 수 있으니 개업/전입일을 확인하세요.</div>}
      <div className="mb-3 grid grid-cols-3 gap-2 text-sm">
        <div className="bg-gray-50 rounded p-2"><div className="text-xs text-gray-500">의무 금액</div><div className="font-semibold">{fmt(yearData.required)}원</div></div>
        <div className="bg-gray-50 rounded p-2"><div className="text-xs text-gray-500">납입 금액</div><div className={`font-semibold ${st==='paid'?'text-emerald-600':st==='partial'?'text-amber-600':'text-rose-600'}`}>{fmt(yearData.allocated)}원</div></div>
        <div className="bg-gray-50 rounded p-2"><div className="text-xs text-gray-500">상태</div><div className="font-semibold">{st==='paid'?'완납':st==='partial'?'일부납':'미납'}</div></div>
      </div>
      <h4 className="font-semibold text-sm mb-2">입금 내역</h4>
      {yearData.payments.length === 0 ? <div className="text-center text-gray-400 py-6 border border-dashed border-gray-200 rounded">충당된 입금이 없습니다</div> : (
        <div className="border border-gray-200 rounded overflow-hidden"><table className="w-full text-sm"><thead className="bg-gray-50"><tr className="text-left text-gray-600 text-xs"><th className="px-2 py-1.5">입금일시</th><th className="px-2 py-1.5">입금자명</th><th className="px-2 py-1.5 text-right">충당액</th><th className="px-2 py-1.5 text-right">실입금</th><th className="px-2 py-1.5"></th></tr></thead>
        <tbody>{yearData.payments.map((p,i) => <tr key={i} className={`border-t border-gray-100 ${!isPrior && p.date && yearOf(p.date) < jy ? 'bg-rose-50/40' : ''}`}><td className="px-2 py-1.5 whitespace-nowrap">{p.date}<span className="text-gray-400 text-xs ml-1">{p.time||''}</span>{!isPrior && p.date && yearOf(p.date) < jy && <span className="ml-1 px-1 py-0.5 bg-rose-100 text-rose-700 rounded text-[10px]" title="입금일이 개업/전입일보다 빠릅니다. 개업/전입일을 확인하세요.">⚠확인</span>}</td><td className="px-2 py-1.5">{p.depositorName}</td><td className="px-2 py-1.5 text-right font-mono">{fmt(p.amount)}{p.overflow && <span className="ml-1 text-xs text-indigo-500">+</span>}</td><td className="px-2 py-1.5 text-right font-mono text-xs text-gray-500">{fmt(p.fullAmount)}</td><td className="px-2 py-1.5 text-right">{!p.carryOver && <button onClick={() => onUnassign(p.txId)} className="text-xs text-gray-400 hover:text-rose-600">해제</button>}</td></tr>)}</tbody></table></div>
      )}
    </Modal>
  );
}

// ============ Utilities ============
function Modal({ children, onClose, title }) {
  return (
    <div className="fixed inset-0 bg-black/40 z-50 flex items-center justify-center p-4" onClick={onClose}>
      <div className="bg-white rounded-lg max-w-2xl w-full max-h-[90vh] overflow-y-auto" onClick={e => e.stopPropagation()}>
        <div className="flex items-center justify-between border-b border-gray-200 px-4 py-3"><h3 className="font-semibold">{title}</h3><button onClick={onClose} className="text-gray-500 hover:text-gray-900"><X className="w-5 h-5" /></button></div>
        <div className="p-4">{children}</div>
      </div>
    </div>
  );
}
function Field({ label, value, onChange, type = 'text' }) {
  return <div><label className="text-xs text-gray-600 block mb-1">{label}</label><input type={type} value={value} onChange={e=>onChange(e.target.value)} className="w-full border border-gray-300 rounded px-2 py-1.5 text-sm" /></div>;
}
function StatCard({ label, value, color, small }) {
  return <div className="bg-white rounded-lg border border-gray-200 p-3"><div className="text-xs text-gray-500">{label}</div><div className={`font-bold ${small ? 'text-sm mt-1' : 'text-lg mt-0.5'} ${color || ''}`}>{value}</div></div>;
}
function Legend({ color, label }) {
  return <div className="flex items-center gap-1.5"><div className={`w-4 h-4 rounded border ${color}`}></div><span className="text-gray-700">{label}</span></div>;
}
function EmptyState({ icon: Ic, title, desc }) {
  return <div className="bg-white rounded-lg border border-dashed border-gray-300 p-12 text-center"><Ic className="w-12 h-12 text-gray-300 mx-auto mb-3" /><h3 className="font-semibold text-gray-700 mb-1">{title}</h3><p className="text-sm text-gray-500">{desc}</p></div>;
}
