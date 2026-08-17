/** Freshy Badminton - Google Apps Script API
 *  1) วางไฟล์นี้ใน Apps Script ของ Google Sheet
 *  2) ตั้งค่า SHEET_ID และ API_KEY ให้ตรงกับเว็บ
 */
const CONFIG = {
  SHEET_ID: '1hSVQ9e_V5ZelrBaHYZ3ap5Ha0D2CAPEhxvSd9_NvLEg',
  SCHEDULE_SHEET: 'ตารางการแข่งขัน',
  RESULTS_SHEET: 'ผลการแข่งขัน',
  SUMMARY_SHEET: 'สรุปผลการแข่งขัน',
  API_KEY: ''
};

const HEADERS = ['id','dateISO','scheduledTime','court','eventType','status','winner','teamA','teamB','sets','totalA','totalB','setWinsA','setWinsB','updatedAt','updatedBy','version','auditLog'];

function doGet(e) {
  try {
    checkKey_(e.parameter.key);
    // เปิด URL ตรงๆ ก็ให้โหลดข้อมูลได้เลย ส่วนเว็บจะส่ง action=list มาอยู่แล้ว
    if (e.parameter.action && e.parameter.action !== 'list') throw new Error('action ไม่ถูกต้อง');
    return json_({ ok: true, matches: readMatches_() });
  } catch (err) { return json_({ ok: false, error: err.message }); }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    checkKey_(body.key);
    if (body.action !== 'save' || !body.match || !body.match.id) throw new Error('ข้อมูลบันทึกไม่ครบ');
    const match = saveMatch_(body.match);
    // การสรุปผลไม่ควรทำให้การบันทึกคะแนนล้มเหลว หากชีตสรุปมีเซลล์ merge
    try { updateSummary_(); } catch (summaryError) { console.error(summaryError); }
    return json_({ ok: true, match: match });
  } catch (err) { return json_({ ok: false, error: err.message }); }
}

function setupSheet() {
  const sheet = getResultsSheet_();
  if (sheet.getLastRow() === 0) sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  else if (sheet.getRange(1, 1).getValue() !== 'id') sheet.insertRowBefore(1).getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
}

function readMatches_() {
  const schedule = getSheetByName_(CONFIG.SCHEDULE_SHEET).getDataRange().getDisplayValues();
  const results = readResultRows_();
  let eventType = 'MD', matches = [];
  for (let i = 5; i < schedule.length; i++) {
    const r = schedule[i];
    if (r[0] && !/^\d+$/.test(String(r[0]).trim())) {
      const heading = String(r[0]).trim();
      eventType = heading.includes('ผสม') ? 'XD' : (heading.includes('หญิง') ? 'WD' : 'MD');
    }
    if (!/^\d+$/.test(String(r[0]).trim())) continue;
    const court = String(r[6] || 'สนาม 1').trim();
    const id = court + '-คู่ที่-' + String(r[0]).trim();
    const saved = results[id] || {};
    matches.push({ id:id, dateISO:normalizeDate_(r[1]), scheduledTime:String(r[5] || '').replace(' น.',''), court:court, eventType:eventType, status:saved.status || 'scheduled', winner:saved.winner || null, teamA:team_(r[2]), teamB:team_(r[4]), sets:saved.sets || emptySets_(), totalA:Number(saved.totalA||0), totalB:Number(saved.totalB||0), setWinsA:Number(saved.setWinsA||0), setWinsB:Number(saved.setWinsB||0), updatedAt:saved.updatedAt||'', updatedBy:saved.updatedBy||'', version:Number(saved.version||0), auditLog:saved.auditLog||[] });
  }
  return matches;
}

function saveMatch_(match) {
  const sheet = getResultsSheet_();
  const rows = sheet.getDataRange().getDisplayValues();
  const rowIndex = rows.findIndex((r, i) => i > 0 && matchId_(r[6], r[0]) === String(match.id));
  const old = rowIndex > 0 ? readResult_(rows[rowIndex]) : null;
  if (old && Number(match.version) !== Number(old.version)) throw new Error('ข้อมูลนี้ถูกแก้ไขโดยกรรมการคนอื่นแล้ว');
  match.version = (old ? Number(old.version) : 0) + 1;
  match.updatedAt = new Date().toISOString();
  if (rowIndex <= 0) throw new Error('หาแถวของแมตช์ในแท็บผลการแข่งขันไม่พบ');
  sheet.getRange(rowIndex + 1, 8, 1, 2).setValues([[formatScore_(match), match.winner || '']]);
  return match;
}

function getSheetByName_(name) { return SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName(name) || SpreadsheetApp.openById(CONFIG.SHEET_ID).insertSheet(name); }
function getResultsSheet_() { return getSheetByName_(CONFIG.RESULTS_SHEET); }
function readResultRows_() { const rows=getResultsSheet_().getDataRange().getDisplayValues(); const out={}; rows.slice(1).forEach(r=>{ if (/^\d+$/.test(String(r[0]).trim())) out[matchId_(r[6],r[0])] = readResult_(r); }); return out; }
function matchId_(court,pair) { return String(court||'สนาม 1').trim() + '-คู่ที่-' + String(pair||'').trim(); }
function readResult_(r) { const score=String(r[7]||''); const nums=score.match(/\d+/g)||[]; const a1=Number(nums[0]||0), b1=Number(nums[1]||0), a2=Number(nums[2]||0), b2=Number(nums[3]||0); const sets={set1:{scoreA:a1,scoreB:b1},set2:{scoreA:a2,scoreB:b2}}; return {status:score?'finished':'scheduled',winner:r[8]||null,sets:sets,totalA:a1+a2,totalB:b1+b2,setWinsA:(a1>b1?1:0)+(a2>b2?1:0),setWinsB:(b1>a1?1:0)+(b2>a2?1:0),version:score?1:0,auditLog:[]}; }
function formatScore_(m) { return [m.sets.set1.scoreA,m.sets.set1.scoreB,m.sets.set2.scoreA,m.sets.set2.scoreB].join('-'); }
function updateSummary_() {
  const sheet = getSheetByName_(CONFIG.SUMMARY_SHEET);
  const matches = readMatches_().filter(m => m.status === 'finished' || m.status === 'walkover');
  const names = ['รังสีเทคนิค','แพทย์แผนไทย','นวัตกรรมและฉุกเฉินการแพทย์','วทบ.เวชและปวส.เวช'];
  const stats = {}; ['MD','WD','XD'].forEach(type => { stats[type]={}; names.forEach(n => stats[type][n]={matches:0,points:0}); });
  matches.forEach(m => { const type=m.eventType||'MD', a=m.teamA.name, b=m.teamB.name; if(!stats[type]) return; if(stats[type][a]){stats[type][a].matches++;stats[type][a].points+=m.totalA||0;} if(stats[type][b]){stats[type][b].matches++;stats[type][b].points+=m.totalB||0;} });
  const rows = sheet.getDataRange().getDisplayValues();
  let currentType='MD';
  rows.forEach((r,i) => {
    const label=String(r[0]||'').trim();
    if(label.includes('หญิง')) currentType='WD'; else if(label.includes('ผสม')) currentType='XD'; else if(label.includes('ชาย')) currentType='MD';
    if(stats[currentType][label]) sheet.getRange(i+1,2,1,2).setValues([[stats[currentType][label].matches,stats[currentType][label].points]]);
  });
}
function emptySets_() { return {set1:{scoreA:0,scoreB:0},set2:{scoreA:0,scoreB:0}}; }
function team_(name) { const n=String(name||'').trim(); const map={'รังสีเทคนิค':'d2','แพทย์แผนไทย':'d1','นวัตกรรมและฉุกเฉินการแพทย์':'d4','นวัตกรรมสื่อสารและฉุกเฉินการแพทย์':'d4','วทบ.เวชและปวส.เวช':'d3','วทบ.เวชและปวส.เวชระเบียน':'d3'}; return {id:n,name:n,departmentId:map[n]||'d3'}; }
function normalizeDate_(value) { const s=String(value||''); if (/^\d{4}\/\d{1,2}\/\d{1,2}$/.test(s)) return s.replaceAll('/','-'); return s; }
function checkKey_(key) { if (CONFIG.API_KEY && key !== CONFIG.API_KEY) throw new Error('API key ไม่ถูกต้อง'); }
function rowToMatch_(r) { return { id:r[0], dateISO:r[1], scheduledTime:r[2], court:r[3], eventType:r[4], status:r[5], winner:r[6] || null, teamA:JSON.parse(r[7]), teamB:JSON.parse(r[8]), sets:JSON.parse(r[9]), totalA:Number(r[10]||0), totalB:Number(r[11]||0), setWinsA:Number(r[12]||0), setWinsB:Number(r[13]||0), updatedAt:r[14], updatedBy:r[15], version:Number(r[16]||1), auditLog:JSON.parse(r[17] || '[]') }; }
function matchToRow_(m) { return [m.id,m.dateISO,m.scheduledTime,m.court,m.eventType,m.status,m.winner||'',JSON.stringify(m.teamA),JSON.stringify(m.teamB),JSON.stringify(m.sets),m.totalA||0,m.totalB||0,m.setWinsA||0,m.setWinsB||0,m.updatedAt||'',m.updatedBy||'',m.version||1,JSON.stringify(m.auditLog||[])]; }
function json_(data) { return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON); }
