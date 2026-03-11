// ══════════════════════════════════════════════════════
// AZEM — مزامنة Firestore مع Google Sheets
// الصقه في: Extensions → Apps Script
// ══════════════════════════════════════════════════════

// ⚙️ إعدادات — عدّل هذه القيم فقط
const FIREBASE_PROJECT_ID = 'azem-ad49b';
const FIREBASE_API_KEY    = 'AIzaSyDnKud5gR8a_Fyq8cNdzgHMNQw4GMuX0-Q';
const SHEET_NAME          = 'المستخدمون';

// ══════════════════════════════════════════════════════
// الدالة الرئيسية — تُشغَّل يدوياً أو تلقائياً
// ══════════════════════════════════════════════════════
function syncUsersToSheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName(SHEET_NAME);

  // أنشئ الورقة إن لم تكن موجودة
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  // جلب البيانات من Firestore REST API
  const url = `https://firestore.googleapis.com/v1/projects/${FIREBASE_PROJECT_ID}/databases/(default)/documents/users?key=${FIREBASE_API_KEY}&pageSize=500`;

  const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  const data     = JSON.parse(response.getContentText());

  if (!data.documents) {
    SpreadsheetApp.getUi().alert('لا توجد بيانات أو خطأ في الاتصال:\n' + response.getContentText());
    return;
  }

  // ── رؤوس الجدول ──────────────────────────────────
  const headers = [
    'الاسم',
    'الإيميل',
    'الصورة (رابط)',
    'المدينة',
    'الدولة',
    'المنطقة',
    'IP',
    'المنطقة الزمنية',
    'الوزن (كغ)',
    'الطول (سم)',
    'الهدف',
    'أيام البرنامج',
    'اليوم الحالي',
    '🔥 سلسلة',
    'أيام مكتملة',
    'آخر تسجيل دخول',
    'آخر مزامنة',
    'المعرف (UID)',
  ];

  // ── مسح الورقة وإعادة الكتابة ────────────────────
  sheet.clearContents();

  // تنسيق رأس الجدول
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#1a1c2a');
  headerRange.setFontColor('#d4a843');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  sheet.setFrozenRows(1);

  // ── معالجة البيانات ───────────────────────────────
  const rows = [];

  data.documents.forEach(docObj => {
    const fields  = docObj.fields || {};
    const profile = getNestedMap(fields, 'profile');
    const state   = getNestedMap(fields, 'state');
    const geo     = getNestedMap(profile, 'geo');
    const user    = getNestedMap(state, 'user');

    const displayName = getStr(profile, 'displayName') || getStr(user, 'name')         || '—';
    const email       = getStr(profile, 'email')       || getStr(user, 'email')        || '—';
    const photoURL    = getStr(profile, 'photoURL')    || '';
    const city        = getStr(geo,     'city')        || '—';
    const country     = getStr(geo,     'country')     || '—';
    const region      = getStr(geo,     'region')      || '—';
    const ip          = getStr(geo,     'ip')          || '—';
    const timezone    = getStr(geo,     'timezone')    || '—';
    const weight      = getNum(user,    'weight')      || '—';
    const height      = getNum(user,    'height')      || '—';
    const goal        = translateGoal(getStr(user, 'goal'));
    const programDays = getNum(user,    'programDays') || '—';
    const currentDay  = getNum(state,   'currentDay')  || '—';
    const streak      = getNum(state,   'streak')      || 0;
    const completedDays = getCompletedDays(state);
    const lastLogin   = formatTs(getTs(profile, 'lastLogin'));
    const lastSync    = formatTs(getTs(state,   '_syncedAt'));

    // استخراج UID من مسار الوثيقة
    const uid = docObj.name ? docObj.name.split('/').pop() : '—';

    rows.push([
      displayName, email, photoURL,
      city, country, region, ip, timezone,
      weight, height, goal,
      programDays, currentDay, streak, completedDays,
      lastLogin, lastSync, uid
    ]);
  });

  // كتابة البيانات
  if (rows.length > 0) {
    const dataRange = sheet.getRange(2, 1, rows.length, headers.length);
    dataRange.setValues(rows);
    dataRange.setHorizontalAlignment('center');

    // تلوين الصفوف بالتناوب
    for (let i = 0; i < rows.length; i++) {
      const bg = i % 2 === 0 ? '#0e1018' : '#12141f';
      sheet.getRange(i + 2, 1, 1, headers.length).setBackground(bg);
      sheet.getRange(i + 2, 1, 1, headers.length).setFontColor('#e8e0d0');
    }
  }

  // تحجيم الأعمدة تلقائياً
  sheet.autoResizeColumns(1, headers.length);

  // ── ورقة الإحصاء ──────────────────────────────────
  updateStatsSheet(ss, rows);

  // رسالة نجاح
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  SpreadsheetApp.getUi().alert(`✅ تمت المزامنة بنجاح!\n${rows.length} مستخدم · ${now}`);
}

// ══════════════════════════════════════════════════════
// ورقة الإحصاء السريع
// ══════════════════════════════════════════════════════
function updateStatsSheet(ss, rows) {
  const STATS_SHEET = 'إحصاء سريع';
  let statsSheet = ss.getSheetByName(STATS_SHEET);
  if (!statsSheet) statsSheet = ss.insertSheet(STATS_SHEET);
  statsSheet.clearContents();

  const total     = rows.length;
  const withGeo   = rows.filter(r => r[3] !== '—').length;
  const weights   = rows.map(r => parseFloat(r[8])).filter(w => !isNaN(w));
  const avgWeight = weights.length ? (weights.reduce((a,b) => a+b, 0) / weights.length).toFixed(1) : '—';
  const goals     = {};
  rows.forEach(r => { goals[r[10]] = (goals[r[10]] || 0) + 1; });
  const countries = {};
  rows.forEach(r => { if (r[4] !== '—') countries[r[4]] = (countries[r[4]] || 0) + 1; });

  const stats = [
    ['📊 إحصاء AZEM', ''],
    ['', ''],
    ['إجمالي المستخدمين', total],
    ['مستخدمون بموقع جغرافي', withGeo],
    ['متوسط الوزن (كغ)', avgWeight],
    ['', ''],
    ['── الأهداف ──', ''],
    ...Object.entries(goals).map(([k,v]) => [k, v]),
    ['', ''],
    ['── الدول ──', ''],
    ...Object.entries(countries).sort((a,b) => b[1]-a[1]).slice(0,10).map(([k,v]) => [k, v]),
  ];

  statsSheet.getRange(1, 1, stats.length, 2).setValues(stats);
  statsSheet.getRange(1, 1, 1, 2).setFontWeight('bold').setFontColor('#d4a843').setBackground('#1a1c2a');
  statsSheet.autoResizeColumns(1, 2);
}

// ══════════════════════════════════════════════════════
// إعداد التشغيل التلقائي كل ساعة
// ══════════════════════════════════════════════════════
function setupTrigger() {
  // احذف المحفزات القديمة
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  // أضف محفز جديد كل ساعة
  ScriptApp.newTrigger('syncUsersToSheet')
    .timeBased()
    .everyHours(1)
    .create();
  SpreadsheetApp.getUi().alert('✅ تم إعداد التحديث التلقائي كل ساعة!');
}

// ══════════════════════════════════════════════════════
// إضافة قائمة مخصصة في Sheets
// ══════════════════════════════════════════════════════
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚡ AZEM')
    .addItem('🔄 مزامنة الآن', 'syncUsersToSheet')
    .addSeparator()
    .addItem('⏰ تفعيل التحديث التلقائي (كل ساعة)', 'setupTrigger')
    .addToUi();
}

// ══════════════════════════════════════════════════════
// دوال مساعدة لاستخراج البيانات من Firestore format
// ══════════════════════════════════════════════════════
function getNestedMap(fields, key) {
  if (!fields || !fields[key]) return {};
  const val = fields[key];
  return val.mapValue ? (val.mapValue.fields || {}) : {};
}

function getStr(fields, key) {
  if (!fields || !fields[key]) return '';
  const val = fields[key];
  return val.stringValue || '';
}

function getNum(fields, key) {
  if (!fields || !fields[key]) return null;
  const val = fields[key];
  return val.integerValue || val.doubleValue || null;
}

function getTs(fields, key) {
  if (!fields || !fields[key]) return null;
  const val = fields[key];
  return val.integerValue || val.timestampValue || null;
}

function formatTs(ts) {
  if (!ts) return '—';
  try {
    const d = new Date(typeof ts === 'number' ? ts : ts);
    return Utilities.formatDate(d, 'Asia/Riyadh', 'yyyy-MM-dd HH:mm');
  } catch(e) { return '—'; }
}

function getCompletedDays(state) {
  if (!state || !state.completedDays) return 0;
  const val = state.completedDays;
  if (val.arrayValue && val.arrayValue.values) return val.arrayValue.values.length;
  return 0;
}

function translateGoal(g) {
  const map = { burn:'حرق دهون', muscle:'بناء عضلات', fitness:'لياقة', health:'صحة عامة' };
  return map[g] || g || '—';
}
