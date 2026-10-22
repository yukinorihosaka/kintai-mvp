const SHEETS = {
  employees: '従業員', shifts: 'シフト', punches: '打刻', settings: '設定'
};

const HEADERS = {
  employees: ['社員番号', '氏名', 'PINハッシュ', '権限', '有効'],
  shifts: ['日付', '社員番号', '開始予定', '終了予定', '休憩予定分', '勤務日区分', '備考', '更新日時'],
  punches: ['記録ID', '記録日時', '社員番号', '種別', '対象日時', '備考'],
  settings: ['項目', '値']
};

function onOpen() {
  SpreadsheetApp.getUi().createMenu('勤怠管理')
    .addItem('初期設定', 'setupSystem')
    .addItem('WebアプリURLを表示', 'showWebAppUrl')
    .addToUi();
}

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('勤怠管理').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function setupSystem() {
  const ss = SpreadsheetApp.getActive();
  Object.keys(SHEETS).forEach(key => {
    let sh = ss.getSheetByName(SHEETS[key]);
    if (!sh) sh = ss.insertSheet(SHEETS[key]);
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, HEADERS[key].length).setValues([HEADERS[key]])
        .setFontWeight('bold').setBackground('#dbeafe');
      sh.setFrozenRows(1);
    }
  });
  const settings = ss.getSheetByName(SHEETS.settings);
  if (settings.getLastRow() === 1) {
    settings.getRange(2, 1, 5, 2).setValues([
      ['会社名', '勤怠管理'], ['締め日', '末日'], ['週開始曜日', '月'],
      ['法定休日曜日', '日'], ['深夜開始/終了', '22:00-05:00']
    ]);
  }
  const employees = ss.getSheetByName(SHEETS.employees);
  if (employees.getLastRow() === 1) {
    employees.appendRow(['ADMIN', '管理者', hashPin_('ADMIN', 'change-me-now'), '管理者', true]);
  }
  ss.toast(
    '初期ログイン：社員番号 ADMIN / PIN change-me-now',
    '初期設定が完了しました',
    10
  );
  return '初期設定が完了しました。';
}

function showWebAppUrl() {
  SpreadsheetApp.getUi().alert(ScriptApp.getService().getUrl() || '先に「デプロイ」→「新しいデプロイ」→「ウェブアプリ」を実行してください。');
}

function login(employeeId, pin) {
  const emp = authenticate_(employeeId, pin);
  return { employeeId: emp.id, name: emp.name, role: emp.role };
}

function changePin(employeeId, currentPin, newPin) {
  authenticate_(employeeId, currentPin);
  if (!newPin || String(newPin).length < 6) throw new Error('新しいPINは6文字以上にしてください。');
  const sh = sheet_(SHEETS.employees);
  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(employeeId)) {
      sh.getRange(i + 1, 3).setValue(hashPin_(employeeId, newPin));
      return true;
    }
  }
  throw new Error('従業員が見つかりません。');
}

function addEmployee(adminId, adminPin, employee) {
  requireAdmin_(adminId, adminPin);
  if (!employee.id || !employee.name || !employee.pin || String(employee.pin).length < 6) {
    throw new Error('社員番号、氏名、6文字以上のPINを入力してください。');
  }
  const sh = sheet_(SHEETS.employees);
  if (findEmployee_(employee.id)) throw new Error('同じ社員番号が登録されています。');
  sh.appendRow([employee.id, employee.name, hashPin_(employee.id, employee.pin), employee.role === '管理者' ? '管理者' : '従業員', true]);
  return true;
}

function punch(employeeId, pin, type) {
  authenticate_(employeeId, pin);
  if (!['出勤', '退勤', '休憩開始', '休憩終了'].includes(type)) throw new Error('打刻種別が不正です。');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const now = new Date();
    const today = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const recent = getPunchRows_(employeeId, today, today);
    if (recent.length && recent[recent.length - 1].type === type) throw new Error('同じ打刻が連続しています。');
    sheet_(SHEETS.punches).appendRow([Utilities.getUuid(), now, employeeId, type, now, '本人打刻']);
    return { message: type + 'を記録しました', time: formatDateTime_(now) };
  } finally { lock.releaseLock(); }
}

function getDashboard(employeeId, pin) {
  const emp = authenticate_(employeeId, pin);
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return {
    employee: { id: emp.id, name: emp.name, role: emp.role },
    shift: getShiftRows_(employeeId, today, today)[0] || null,
    punches: getPunchRows_(employeeId, today, today),
    serverTime: formatDateTime_(new Date())
  };
}

function saveShift(adminId, adminPin, shift) {
  requireAdmin_(adminId, adminPin);
  if (!findEmployee_(shift.employeeId)) throw new Error('社員番号が存在しません。');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(shift.date)) throw new Error('日付が不正です。');
  if (!/^\d{2}:\d{2}$/.test(shift.start) || !/^\d{2}:\d{2}$/.test(shift.end)) throw new Error('時刻が不正です。');
  const sh = sheet_(SHEETS.shifts);
  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (dateKey_(values[i][0]) === shift.date && String(values[i][1]) === String(shift.employeeId)) {
      sh.getRange(i + 1, 1, 1, 8).setValues([[shift.date, shift.employeeId, shift.start, shift.end, Number(shift.breakMinutes || 0), shift.dayType || '所定労働日', shift.note || '', new Date()]]);
      return true;
    }
  }
  sh.appendRow([shift.date, shift.employeeId, shift.start, shift.end, Number(shift.breakMinutes || 0), shift.dayType || '所定労働日', shift.note || '', new Date()]);
  return true;
}

function getMonthlySummary(employeeId, pin, month, targetEmployeeId) {
  const actor = authenticate_(employeeId, pin);
  const target = actor.role === '管理者' && targetEmployeeId ? targetEmployeeId : actor.id;
  const range = monthRange_(month);
  const days = buildDaily_(target, range.start, range.end);
  return { month: month, employee: publicEmployee_(findEmployee_(target)), days: days, totals: sumDays_(days) };
}

function exportFreeeCsv(adminId, adminPin, month) {
  requireAdmin_(adminId, adminPin);
  const range = monthRange_(month);
  const headers = ['従業員番号','氏名','所定労働時間（分）','法定内残業時間（分）','時間外労働時間（分）','所定休日労働時間（分）','深夜労働時間（分）','法定休日労働時間（分）','総労働時間（分）','総労働日数','所定労働出勤日数','所定休日出勤日数','法定休日出勤日数','遅刻時間（分）','早退時間（分）','欠勤日数','遅刻日数','早退日数','有休取得日数','集計開始日','集計終了日'];
  const rows = activeEmployees_().map(emp => {
    const t = sumDays_(buildDaily_(emp.id, range.start, range.end));
    return [emp.id, emp.name, t.scheduledMinutes, 0, t.overtimeMinutes, t.nonStatutoryHolidayMinutes, t.nightMinutes, t.statutoryHolidayMinutes, t.workMinutes, t.workDays, t.regularWorkDays, t.nonStatutoryHolidayDays, t.statutoryHolidayDays, t.lateMinutes, t.earlyMinutes, 0, t.lateDays, t.earlyDays, 0, range.start.replace(/-/g,'/'), range.end.replace(/-/g,'/')];
  });
  const csv = [headers].concat(rows).map(r => r.map(csvCell_).join(',')).join('\r\n');
  return { filename: 'freee_勤怠_' + month + '.csv', content: '\uFEFF' + csv };
}

function buildDaily_(employeeId, start, end) {
  const shifts = getShiftRows_(employeeId, start, end);
  const punches = getPunchRows_(employeeId, start, end);
  const shiftMap = {}; shifts.forEach(s => shiftMap[s.date] = s);
  const punchMap = {};
  punches.forEach(p => {
    if (!punchMap[p.date]) punchMap[p.date] = [];
    punchMap[p.date].push(p);
  });
  const dates = Array.from(new Set(Object.keys(shiftMap).concat(Object.keys(punchMap)))).sort();
  return dates.map(date => calculateDay_(date, shiftMap[date], punchMap[date] || []));
}

function calculateDay_(date, shift, punches) {
  const first = punches.find(p => p.type === '出勤');
  const last = punches.slice().reverse().find(p => p.type === '退勤');
  let work = 0, night = 0, status = '';
  if (first && last && last.epoch > first.epoch) {
    let breaks = 0, breakStart = null;
    punches.forEach(p => {
      if (p.type === '休憩開始') breakStart = p.epoch;
      if (p.type === '休憩終了' && breakStart && p.epoch > breakStart) { breaks += p.epoch - breakStart; breakStart = null; }
    });
    work = Math.max(0, Math.round((last.epoch - first.epoch - breaks) / 60000));
    night = overlapNightMinutes_(first.epoch, last.epoch);
  } else if (punches.length) status = '打刻漏れ';
  const planned = shift ? plannedMinutes_(shift) : 0;
  const late = shift && first ? Math.max(0, Math.round((first.epoch - localEpoch_(date, shift.start)) / 60000)) : 0;
  const early = shift && last ? Math.max(0, Math.round((localEpoch_(date, shift.end, shift.end <= shift.start) - last.epoch) / 60000)) : 0;
  return { date, shift: shift || null, clockIn: first ? first.time : '', clockOut: last ? last.time : '', workMinutes: work, plannedMinutes: planned, scheduledMinutes: Math.min(work, planned), overtimeMinutes: Math.max(0, work - planned), nightMinutes: night, lateMinutes: late, earlyMinutes: early, dayType: shift ? shift.dayType : 'シフトなし', status };
}

function sumDays_(days) {
  const t = { workMinutes:0, scheduledMinutes:0, overtimeMinutes:0, nightMinutes:0, lateMinutes:0, earlyMinutes:0, workDays:0, regularWorkDays:0, nonStatutoryHolidayMinutes:0, statutoryHolidayMinutes:0, nonStatutoryHolidayDays:0, statutoryHolidayDays:0, lateDays:0, earlyDays:0 };
  days.forEach(d => {
    t.workMinutes += d.workMinutes; t.scheduledMinutes += d.scheduledMinutes; t.overtimeMinutes += d.overtimeMinutes;
    t.nightMinutes += d.nightMinutes; t.lateMinutes += d.lateMinutes; t.earlyMinutes += d.earlyMinutes;
    if (d.workMinutes > 0) { t.workDays++; if (d.dayType === '法定休日') { t.statutoryHolidayDays++; t.statutoryHolidayMinutes += d.workMinutes; } else if (d.dayType === '所定休日') { t.nonStatutoryHolidayDays++; t.nonStatutoryHolidayMinutes += d.workMinutes; } else t.regularWorkDays++; }
    if (d.lateMinutes > 0) t.lateDays++; if (d.earlyMinutes > 0) t.earlyDays++;
  }); return t;
}

function getShiftRows_(employeeId, start, end) {
  return dataRows_(SHEETS.shifts).filter(r => String(r[1]) === String(employeeId) && dateKey_(r[0]) >= start && dateKey_(r[0]) <= end).map(r => ({ date:dateKey_(r[0]), employeeId:String(r[1]), start:timeKey_(r[2]), end:timeKey_(r[3]), breakMinutes:Number(r[4]||0), dayType:String(r[5]||'所定労働日'), note:String(r[6]||'') }));
}

function getPunchRows_(employeeId, start, end) {
  return dataRows_(SHEETS.punches).filter(r => String(r[2]) === String(employeeId) && dateKey_(r[4]) >= start && dateKey_(r[4]) <= end).map(r => ({ id:String(r[0]), date:dateKey_(r[4]), type:String(r[3]), time:formatDateTime_(new Date(r[4])), epoch:new Date(r[4]).getTime() })).sort((a,b) => a.epoch-b.epoch);
}

function authenticate_(id, pin) {
  const emp = findEmployee_(id);
  if (!emp || !emp.active || emp.hash !== hashPin_(id, pin)) throw new Error('社員番号またはPINが違います。');
  return emp;
}
function requireAdmin_(id, pin) { const e = authenticate_(id, pin); if (e.role !== '管理者') throw new Error('管理者権限が必要です。'); return e; }
function findEmployee_(id) { return allEmployees_().find(e => String(e.id) === String(id)); }
function allEmployees_() { return dataRows_(SHEETS.employees).map(r => ({id:String(r[0]), name:String(r[1]), hash:String(r[2]), role:String(r[3]), active:r[4] === true || String(r[4]).toUpperCase()==='TRUE'})); }
function activeEmployees_() { return allEmployees_().filter(e => e.active); }
function publicEmployee_(e) { return e ? {id:e.id, name:e.name, role:e.role} : null; }
function hashPin_(id, pin) { return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(id) + '|' + String(pin), Utilities.Charset.UTF_8)); }
function sheet_(name) { const sh = SpreadsheetApp.getActive().getSheetByName(name); if (!sh) throw new Error('先に初期設定を実行してください。'); return sh; }
function dataRows_(name) { const v=sheet_(name).getDataRange().getValues(); return v.length > 1 ? v.slice(1) : []; }
function dateKey_(v) { return Utilities.formatDate(v instanceof Date ? v : new Date(String(v).replace(/-/g,'/')), Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function timeKey_(v) { return v instanceof Date ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'HH:mm') : String(v).slice(0,5); }
function formatDateTime_(d) { return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'); }
function monthRange_(m) { if (!/^\d{4}-\d{2}$/.test(m)) throw new Error('対象月が不正です。'); const a=m.split('-').map(Number), last=new Date(a[0],a[1],0).getDate(); return {start:m+'-01',end:m+'-'+String(last).padStart(2,'0')}; }
function localEpoch_(date,time,nextDay) { const p=date.split('-').map(Number), q=time.split(':').map(Number); return new Date(p[0],p[1]-1,p[2]+(nextDay?1:0),q[0],q[1]).getTime(); }
function plannedMinutes_(s) { const a=localEpoch_(s.date,s.start), b=localEpoch_(s.date,s.end,s.end<=s.start); return Math.max(0,Math.round((b-a)/60000)-s.breakMinutes); }
function overlapNightMinutes_(start,end) { let total=0, cursor=new Date(start); cursor.setHours(0,0,0,0); cursor.setDate(cursor.getDate()-1); while(cursor.getTime()<end){ const n1=new Date(cursor);n1.setHours(22);const n2=new Date(cursor);n2.setDate(n2.getDate()+1);n2.setHours(5);total+=Math.max(0,Math.min(end,n2.getTime())-Math.max(start,n1.getTime()));cursor.setDate(cursor.getDate()+1);} return Math.round(total/60000); }
function csvCell_(v) { const s=String(v==null?'':v); return /[",\r\n]/.test(s)?'"'+s.replace(/"/g,'""')+'"':s; }
