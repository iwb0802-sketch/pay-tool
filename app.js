const $ = (s) => document.querySelector(s);
const nf = new Intl.NumberFormat('ko-KR');
let rawRows = [];
let detailRows = [];
let summaryRows = [];
let warningRows = [];
const TAX_RATE = 0.033;
const TAX_KEY='bns_tax_keywords';

const fileInput = $('#fileInput');
const dropZone = $('#dropZone');

['dragenter','dragover'].forEach(evt => dropZone.addEventListener(evt, e => {e.preventDefault(); dropZone.classList.add('drag')}));
['dragleave','drop'].forEach(evt => dropZone.addEventListener(evt, e => {e.preventDefault(); dropZone.classList.remove('drag')}));
dropZone.addEventListener('drop', e => { const f=e.dataTransfer.files?.[0]; if(f) loadFile(f); });
fileInput.addEventListener('change', e => { const f=e.target.files?.[0]; if(f) loadFile(f); });
$('#runBtn').addEventListener('click', runSettlement);
$('#resetBtn').addEventListener('click', () => { $('#keyword').value=''; render(detailRows, summaryRows, warningRows); });
const taxKeywordsEl = $('#taxKeywords');
if(taxKeywordsEl){ taxKeywordsEl.value = localStorage.getItem(TAX_KEY) || '세금x'; taxKeywordsEl.addEventListener('input', ()=>{ localStorage.setItem(TAX_KEY, taxKeywordsEl.value); if(rawRows.length) runSettlement(); }); }
$('#downloadSummary').addEventListener('click', () => downloadExcel('연주자별_정산요약.xls', summaryRows.map(r => ({연주자:r.performer, 건수:r.count, 총페이:r.total, '3.3%공제후':r.net, 세금제외건수:r.taxExemptCount}))));
$('#downloadDetail').addEventListener('click', () => downloadExcel('연주자별_정산상세.xls', detailRows.map(toKoreanDetail)));

async function loadFile(file){
  const buf = await file.arrayBuffer();
  const name = file.name.toLowerCase();
  let rows = [];
  try{
    const text = new TextDecoder('utf-8').decode(buf);
    if(text.trim().startsWith('<') || text.includes('<table')) rows = parseHtmlTable(text);
  }catch(e){}
  if(!rows.length && window.XLSX){
    const wb = XLSX.read(buf, {type:'array', cellDates:false});
    const sheet = wb.Sheets[wb.SheetNames[0]];
    rows = XLSX.utils.sheet_to_json(sheet, {defval:''});
  }
  if(!rows.length){ alert('엑셀 내용을 읽지 못했습니다. 관리자 페이지에서 받은 .xls 또는 .xlsx 파일을 다시 올려주세요.'); return; }
  rawRows = rows.map(cleanRow);
  $('#fileInfo').textContent = `${file.name} / ${rawRows.length}개 행사 불러옴`;
  setDateDefaults(rawRows);
  runSettlement();
}

function parseHtmlTable(html){
  const doc = new DOMParser().parseFromString(html, 'text/html');
  const table = doc.querySelector('table');
  if(!table) return [];
  const headers = [...table.querySelectorAll('thead th')].map(th => th.textContent.trim());
  const trs = [...table.querySelectorAll('tbody tr')];
  return trs.map(tr => {
    const cells = [...tr.children].map(td => td.textContent.replace(/\s+/g,' ').trim());
    const obj = {};
    headers.forEach((h,i) => obj[h] = cells[i] ?? '');
    return obj;
  }).filter(r => Object.values(r).some(Boolean));
}

function cleanRow(row){
  const out = {};
  for(const [k,v] of Object.entries(row)) out[String(k).trim()] = typeof v === 'string' ? v.trim() : v;
  return out;
}

function setDateDefaults(rows){
  const dates = rows.map(r => normalizeDate(r['행사날짜'])).filter(Boolean).sort();
  if(dates.length){ $('#startDate').value = dates[0]; $('#endDate').value = dates[dates.length-1]; }
}

function normalizeDate(v){
  if(!v) return '';
  if(v instanceof Date && !isNaN(v)) return v.toISOString().slice(0,10);
  const s = String(v).trim();
  const m = s.match(/(20\d{2})[-./년\s]*(\d{1,2})[-./월\s]*(\d{1,2})/);
  if(!m) return '';
  return `${m[1]}-${m[2].padStart(2,'0')}-${m[3].padStart(2,'0')}`;
}

function money(v){
  if(v == null || v === '') return 0;
  const n = Number(String(v).replace(/[^0-9.-]/g,''));
  return Number.isFinite(n) ? n : 0;
}

function runSettlement(){
  if(!rawRows.length){ alert('먼저 엑셀 파일을 올려주세요.'); return; }
  const start = $('#startDate').value || '0000-00-00';
  const end = $('#endDate').value || '9999-99-99';
  const kw = $('#keyword').value.trim().toLowerCase();

  detailRows = [];
  warningRows = [];
  const seenEvents = new Set();

  rawRows.forEach((r, idx) => {
    const date = normalizeDate(r['행사날짜']);
    if(!date || date < start || date > end) return;
    const base = `${date} ${r['시간']||''} ${r['장소/층수']||''} ${r['연주편성']||''} ${r['발주처']||''} ${r['발주담당자']||''}`.toLowerCase();
    let rowPaySum = 0;
    let hasAny = false;
    for(let i=1;i<=12;i++){
      const performerRaw = String(r[`악기구성${i}`] || '').replace(/\s+/g,' ').trim();
      const performer = cleanPerformerName(performerRaw);
      const taxExempt = isTaxExemptName(performerRaw);
      const pay = money(r[`악기페이${i}`]);
      const netPay = taxExempt ? pay : taxAdjusted(pay);
      if(!performer && !pay) continue;
      hasAny = true;
      rowPaySum += pay;
      const searchable = `${base} ${performerRaw} ${performer}`.toLowerCase();
      if(kw && !searchable.includes(kw)) continue;
      detailRows.push({
        eventKey: `${date}-${idx}`,
        date, time:r['시간']||'', place:r['장소/층수']||'', order:r['발주처']||'', formation:r['연주편성']||'', performer, performerRaw, pay, netPay, taxExempt,
        manager:r['발주담당자']||''
      });
      seenEvents.add(`${date}-${idx}`);
    }
    const declared = money(r['연주자페이합계']);
    if(hasAny && declared && Math.abs(declared-rowPaySum) > 1){
      warningRows.push({date, place:r['장소/층수']||'', declared, actual:rowPaySum, diff:rowPaySum-declared});
    }
  });

  const map = new Map();
  detailRows.forEach(d => {
    if(!d.performer) return;
    const cur = map.get(d.performer) || {performer:d.performer, count:0, total:0, net:0, taxExemptCount:0};
    cur.count += 1;
    cur.total += d.pay;
    cur.net += d.netPay;
    if(d.taxExempt) cur.taxExemptCount += 1;
    map.set(d.performer, cur);
  });
  summaryRows = [...map.values()].sort((a,b) => b.total-a.total || b.count-a.count || a.performer.localeCompare(b.performer,'ko'));
  render(detailRows, summaryRows, warningRows, seenEvents.size, start, end);
}

function render(details, summaries, warnings, eventCount=null, start='', end=''){
  $('#stats').hidden = $('#results').hidden = $('#detailPanel').hidden = false;
  $('#eventCount').textContent = nf.format(eventCount ?? new Set(details.map(d=>d.eventKey)).size);
  $('#lineCount').textContent = nf.format(details.length);
  $('#totalPay').textContent = nf.format(details.reduce((s,d)=>s+d.pay,0))+'원';
  $('#netPay').textContent = nf.format(details.reduce((s,d)=>s+d.netPay,0))+'원';
  $('#performerCount').textContent = nf.format(summaries.length)+'명';
  $('#rangeText').textContent = start && end ? `${start} ~ ${end}` : '';

  const stbody = $('#summaryTable tbody'); stbody.innerHTML = '';
  summaries.forEach(r => {
    const tr = document.createElement('tr'); tr.className='summary-row';
    tr.innerHTML = `<td><strong>${esc(r.performer)}</strong></td><td class="num">${nf.format(r.count)}</td><td class="num"><strong>${nf.format(r.total)}원</strong></td><td class="num"><strong>${nf.format(r.net)}원</strong></td><td class="num">${r.taxExemptCount ? nf.format(r.taxExemptCount)+'건' : '-'}</td>`;
    tr.addEventListener('click', () => renderDetails(details.filter(d => d.performer === r.performer)));
    stbody.appendChild(tr);
  });
  renderDetails(details);
  const wp = $('#warningPanel'); wp.hidden = !warnings.length;
  const wbody = $('#warningTable tbody'); wbody.innerHTML = '';
  warnings.forEach(w => {
    const tr=document.createElement('tr');
    tr.innerHTML = `<td>${esc(w.date)}</td><td>${esc(w.place)}</td><td class="num">${nf.format(w.declared)}원</td><td class="num">${nf.format(w.actual)}원</td><td class="num bad">${nf.format(w.diff)}원</td>`;
    wbody.appendChild(tr);
  });
}

function renderDetails(details){
  const body = $('#detailTable tbody'); body.innerHTML = '';
  details.sort((a,b)=> a.date.localeCompare(b.date) || a.time.localeCompare(b.time) || a.performer.localeCompare(b.performer,'ko'))
    .forEach(d => {
      const tr=document.createElement('tr');
      tr.innerHTML = `<td>${esc(d.date)}</td><td>${esc(d.time)}</td><td>${esc(d.place)}</td><td>${esc(d.formation)}</td><td><strong>${esc(d.performer)}</strong></td><td class="num">${nf.format(d.pay)}원</td><td class="num"><strong>${nf.format(d.netPay)}원</strong></td><td>${d.taxExempt ? '세금 제외' : '3.3% 공제'}</td><td>${esc(d.order)}</td><td>${esc(d.manager)}</td>`;
      body.appendChild(tr);
    });
}

function toKoreanDetail(d){ return {날짜:d.date, 시간:d.time, 장소:d.place, 연주편성:d.formation, 연주자:d.performer, 원페이:d.pay, '3.3%공제후':d.netPay, 세금처리:d.taxExempt ? '세금 제외' : '3.3% 공제', 발주처:d.order, 발주담당:d.manager}; }
function isTaxExemptName(name){ const raw=String(name||'').toLowerCase(); const list=(taxKeywordsEl?.value||'세금x').split(/[\n,]+/).map(v=>v.trim().toLowerCase()).filter(Boolean); return list.some(k=>raw.includes(k)); }
function cleanPerformerName(name){ return String(name || '').replace(/[\s\(\[\{]*세금\s*x[\s\)\]\}]*/ig, ' ').replace(/\s+/g,' ').trim(); }
function taxAdjusted(amount){ return Math.round(amount * (1 - TAX_RATE)); }
function esc(s){ return String(s ?? '').replace(/[&<>'"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[c])); }

function downloadExcel(filename, rows){
  if(!rows.length){ alert('다운로드할 데이터가 없습니다.'); return; }
  const headers = Object.keys(rows[0]);
  const html = `<html><head><meta charset="utf-8"></head><body><table border="1"><thead><tr>${headers.map(h=>`<th>${esc(h)}</th>`).join('')}</tr></thead><tbody>${rows.map(r=>`<tr>${headers.map(h=>`<td>${esc(r[h])}</td>`).join('')}</tr>`).join('')}</tbody></table></body></html>`;
  const blob = new Blob([html], {type:'application/vnd.ms-excel;charset=utf-8'});
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = filename; a.click(); URL.revokeObjectURL(a.href);
}
