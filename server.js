// =========================================================================
// 📦 1. استدعاء المكتبات وإعداد التطبيق
// =========================================================================
const express = require('express');
const path = require('path');
const fs = require('fs');
const xlsx = require('xlsx');
const bodyParser = require('body-parser');
const bcrypt = require('bcryptjs');
const multer = require('multer');
const session = require('express-session');
const { Pool } = require('pg');

const app = express();
const PORT = process.env.PORT || 3000;

// 🏠 توجيه الزائر عند فتح الرابط الرئيسي إلى صفحة تسجيل الدخول مباشرة
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'login.html'));
});

// تعطيل الفهرس التلقائي index.html لضمان عدم تجاوز التوجيه
app.use(express.static(__dirname, { index: false }));
app.use(bodyParser.json());
app.use(express.urlencoded({ extended: true }));

// 🎯 إعداد الجلسات (Session)
app.use(session({
  secret: 'school_management_secret_key_2026',
  resave: false,
  saveUninitialized: false,
  cookie: { 
    maxAge: 24 * 60 * 60 * 1000,
    httpOnly: true 
  }
}));

// منع التخزين المؤقت (Cache)
app.use((req, res, next) => {
  res.set('Cache-Control', 'no-store, no-cache, must-revalidate, private');
  next();
});

// =========================================================================
// 🗄️ 2. التعامل مع قاعدة البيانات PostgreSQL والتخزين الدائم
// =========================================================================

console.log("🔍 حالة متغير DATABASE_URL:", process.env.DATABASE_URL ? "موجود ومقروء ✅" : "غير موجود ❌");

const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: process.env.DATABASE_URL ? { rejectUnauthorized: false } : false,
  connectionTimeoutMillis: 5000
});

async function initDb() {
  if (!process.env.DATABASE_URL) {
    console.error("🚨 خطأ: لم يتم العثور على DATABASE_URL بداخل بيئة Render!");
    return;
  }

  try {
    console.log("⏳ جاري الاتصال بقاعدة بيانات PostgreSQL...");
    const client = await pool.connect();
    
    await client.query(`
      CREATE TABLE IF NOT EXISTS users (
        id SERIAL PRIMARY KEY,
        username VARCHAR(255) UNIQUE,
        password TEXT,
        full_name TEXT,
        region TEXT,
        school_name TEXT,
        school_excel_file TEXT,
        excel_data BYTEA
      );
    `);

    await client.query(`
      ALTER TABLE users ADD COLUMN IF NOT EXISTS excel_data BYTEA;
    `);

    // 🟢 إضافة جدول المساءلات إلى PostgreSQL
    await client.query(`
      CREATE TABLE IF NOT EXISTS accountability (
        id SERIAL PRIMARY KEY,
        username VARCHAR(255),
        teacher_name TEXT NOT NULL,
        violation_type TEXT NOT NULL,
        hijri_date TEXT NOT NULL,
        delay_minutes INTEGER DEFAULT 0,
        details TEXT,
        status TEXT DEFAULT 'بانتظار الرد',
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
      );
    `);
    
    // 🟢 جدول الإشراف الأسبوعي (صف لكل معلم في كل مهمة في كل يوم)
    await client.query(`
      CREATE TABLE IF NOT EXISTS monitoring_schedule (
        id SERIAL PRIMARY KEY,
        username VARCHAR(255) NOT NULL,
        day_key VARCHAR(3) NOT NULL CHECK (day_key IN ('sun','mon','tue','wed','thu')),
        task TEXT NOT NULL,
        teacher_name TEXT NOT NULL,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        UNIQUE (username, day_key, task, teacher_name)
      );
    `);
    await client.query(`CREATE INDEX IF NOT EXISTS idx_monitoring_username ON monitoring_schedule (username);`);

    // ✅ سجل أداء الإشراف الفعلي (لكل معلم في كل تاريخ): done = أدّى ، missed = لم يؤدِّ
    await client.query(`
      CREATE TABLE IF NOT EXISTS monitoring_attendance (
        id SERIAL PRIMARY KEY,
        username VARCHAR(255) NOT NULL,
        att_date DATE NOT NULL,
        day_key VARCHAR(3) NOT NULL,
        task TEXT NOT NULL,
        teacher_name TEXT NOT NULL,
        status VARCHAR(10) NOT NULL CHECK (status IN ('done','missed')),
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        UNIQUE (username, att_date, task, teacher_name)
      );
    `);
    await client.query(`CREATE INDEX IF NOT EXISTS idx_mon_att_user_date ON monitoring_attendance (username, att_date);`);

    // علامة لكل مدرسة: هل تم نقل جدول الإشراف القديم من الإكسل إلى القاعدة؟
    await client.query(`ALTER TABLE users ADD COLUMN IF NOT EXISTS monitoring_migrated BOOLEAN DEFAULT FALSE;`);

    // 🗄️ بيانات المدرسة الأساسية (معلمون/فصول/معلومات) + الانتظار + الغياب
    await client.query(`
      CREATE TABLE IF NOT EXISTS school_teachers (
        id SERIAL PRIMARY KEY,
        username VARCHAR(255) NOT NULL,
        national_id TEXT,
        name TEXT NOT NULL,
        phone TEXT,
        sort_order INTEGER DEFAULT 0
      );
    `);
    await client.query(`CREATE INDEX IF NOT EXISTS idx_school_teachers_user ON school_teachers (username);`);
    await client.query(`
      CREATE TABLE IF NOT EXISTS school_classes (
        id SERIAL PRIMARY KEY,
        username VARCHAR(255) NOT NULL,
        name TEXT NOT NULL,
        sort_order INTEGER DEFAULT 0
      );
    `);
    await client.query(`CREATE INDEX IF NOT EXISTS idx_school_classes_user ON school_classes (username);`);
    await client.query(`
      CREATE TABLE IF NOT EXISTS school_info (
        username VARCHAR(255) PRIMARY KEY,
        education_dept TEXT DEFAULT '',
        school_name TEXT DEFAULT '',
        principal TEXT DEFAULT '',
        deputy TEXT DEFAULT ''
      );
    `);
    await client.query(`
      CREATE TABLE IF NOT EXISTS waiting_slots (
        id SERIAL PRIMARY KEY,
        username VARCHAR(255) NOT NULL,
        day_name TEXT NOT NULL,
        row_no INTEGER NOT NULL,
        period INTEGER NOT NULL,
        teacher_id INTEGER REFERENCES school_teachers(id) ON DELETE CASCADE,
        teacher_label TEXT DEFAULT '',
        UNIQUE (username, day_name, row_no, period)
      );
    `);
    await client.query(`
      CREATE TABLE IF NOT EXISTS absence_records (
        id SERIAL PRIMARY KEY,
        username VARCHAR(255) NOT NULL,
        abs_date TEXT NOT NULL,
        day_name TEXT NOT NULL,
        teacher_id INTEGER REFERENCES school_teachers(id) ON DELETE SET NULL,
        teacher_name TEXT NOT NULL,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
      );
    `);
    await client.query(`CREATE INDEX IF NOT EXISTS idx_absence_user_date ON absence_records (username, abs_date);`);
    await client.query(`
      CREATE TABLE IF NOT EXISTS absence_slots (
        id SERIAL PRIMARY KEY,
        absence_id INTEGER NOT NULL REFERENCES absence_records(id) ON DELETE CASCADE,
        period INTEGER NOT NULL,
        class_name TEXT DEFAULT '',
        sub_id INTEGER REFERENCES school_teachers(id) ON DELETE SET NULL,
        sub_name TEXT DEFAULT '',
        UNIQUE (absence_id, period)
      );
    `);
    // علامة لكل مدرسة: هل نُقل الانتظار والغياب والمعلمون من الإكسل إلى القاعدة؟
    await client.query(`ALTER TABLE users ADD COLUMN IF NOT EXISTS waiting_migrated BOOLEAN DEFAULT FALSE;`);

    // 💬 جدول حفظ جلسات واتساب (ربط دائم لكل مدرسة)
    await client.query(`
      CREATE TABLE IF NOT EXISTS wa_auth (
        username VARCHAR(255) NOT NULL,
        key TEXT NOT NULL,
        value TEXT,
        PRIMARY KEY (username, key)
      );
    `);

    client.release();
    console.log("✅ تم الاتصال بقاعدة بيانات PostgreSQL وبناء الجداول وتأمين التخزين الدائم بنجاح!");
  } catch (err) {
    console.error("❌ خطأ في الاتصال بقاعدة البيانات:", err.message);
  }
}

initDb();

// 🔄 دالة حفظ ملف الإكسل بداخل قاعدة البيانات
async function syncExcelToDb(username, filePath) {
  try {
    if (fs.existsSync(filePath)) {
      const fileBuffer = fs.readFileSync(filePath);
      await pool.query('UPDATE users SET excel_data = $1 WHERE username = $2', [fileBuffer, String(username).trim()]);
      console.log(`💾 تم مزامنة وحفظ ملف الإكسل للمستخدم (${username}) في قاعدة البيانات الدائمة.`);
      return true;
    }
    return false;
  } catch (e) {
    console.error(`❌ فشل مزامنة ملف الإكسل لقاعدة البيانات: ${e.message}`);
    return false;
  }
}

// ✅ حفظ موثوق: لا نردّ "تم الحفظ" إلا بعد التأكد أن الملف وصل فعلاً لقاعدة البيانات الدائمة
async function syncOrThrow(username, filePath) {
  if (!username) throw new Error('انتهت الجلسة، يرجى تسجيل الدخول من جديد');
  const ok = await syncExcelToDb(username, filePath);
  if (!ok) throw new Error('تعذر حفظ التعديل في قاعدة البيانات الدائمة، حاول مرة أخرى');
}

// 🔐 مسار تسجيل الدخول للمدارس
app.post('/login', async (req, res) => {
  const { username, password } = req.body;

  try {
    const cleanUsername = String(username).trim();
    const result = await pool.query('SELECT * FROM users WHERE username = $1', [cleanUsername]);

    if (result.rows.length > 0) {
      const user = result.rows[0];

      let isPasswordValid = false;
      if (user.password.startsWith('$2a$') || user.password.startsWith('$2b$')) {
        isPasswordValid = bcrypt.compareSync(password, user.password);
      } else {
        isPasswordValid = (password === user.password);
      }

      if (isPasswordValid) {
        req.session.username = user.username;

        return res.json({
          success: true,
          message: 'تم تسجيل الدخول بنجاح!',
          username: user.username,
          excelFile: user.school_excel_file || '',
          schoolName: user.school_name || '',
          fullName: user.full_name || ''
        });
      }
    }

    return res.status(401).json({
      success: false,
      message: 'اسم المستخدم أو كلمة المرور غير صحيحة.'
    });

  } catch (error) {
    console.error('خطأ في تسجيل الدخول:', error);
    res.status(500).json({ success: false, message: 'حدث خطأ في السيرفر أثناء تسجيل الدخول' });
  }
});

// 🚪 مسار تسجيل الخروج
app.post('/logout', (req, res) => {
  req.session.destroy((err) => {
    if (err) return res.status(500).json({ success: false, message: 'تعذر تسجيل الخروج' });
    res.clearCookie('connect.sid');
    return res.json({ success: true, message: 'تم تسجيل الخروج بنجاح' });
  });
});

// =========================================================================
// ⚙️ إعدادات أسماء الملف والورقات ودوال إنشاء ملف نظيف
// =========================================================================
const EXCEL_FILE = path.join(__dirname, 'waiting_data.xlsx');

const TEACHERS_SHEET = 'teacher_name';
const CLASSES_SHEET = 'classes';
const WAITING_SHEET = ' Waiting_table';
const REPORT_SHEET = 'report';
const MAIN_INFO_SHEET = 'maininfo';
const MONITORING_SHEET = 'moni';
const ROTATION_SHEET = 'rotation';

// 🧹 دالة إنشاء ملف إكسل مخصص ونظيف تماماً بدون أي بيانات معلمين مسبقة
function createCleanSchoolWorkbook(schoolName = '', region = '', fullName = '') {
  const wb = xlsx.utils.book_new();

  // 1. sheet: teacher_name (العناوين فقط)
  const teachersSheet = xlsx.utils.aoa_to_sheet([["الهوية", "اسم المعلم"]]);
  xlsx.utils.book_append_sheet(wb, teachersSheet, TEACHERS_SHEET);

  // 2. sheet: classes (العناوين فقط)
  const classesSheet = xlsx.utils.aoa_to_sheet([["الفصل"]]);
  xlsx.utils.book_append_sheet(wb, classesSheet, CLASSES_SHEET);

  // 3. sheet: Waiting_table (جدول حصص فارغ)
  const waitingSheet = xlsx.utils.aoa_to_sheet([
    ["اليوم", "1", "2", "3", "4", "5", "6", "7"]
  ]);
  xlsx.utils.book_append_sheet(wb, waitingSheet, WAITING_SHEET);

  // 4. sheet: report
  const reportHeader = ["التاريخ", "اليوم", "المعلم الغائب", "فصل 1", "بديل 1", "فصل 2", "بديل 2", "فصل 3", "بديل 3", "فصل 4", "بديل 4", "فصل 5", "بديل 5", "فصل 6", "بديل 6", "فصل 7", "بديل 7"];
  const reportSheet = xlsx.utils.aoa_to_sheet([reportHeader]);
  xlsx.utils.book_append_sheet(wb, reportSheet, REPORT_SHEET);

  // 5. sheet: maininfo (معلومات المدرسة المسجلة فقط)
  const mainInfoSheet = xlsx.utils.aoa_to_sheet([
    ["المدرسة", schoolName],
    ["القطاع", region],
    ["المدير", fullName]
  ]);
  xlsx.utils.book_append_sheet(wb, mainInfoSheet, MAIN_INFO_SHEET);

  // 6. sheet: moni
  const moniHeader = ["المهمة", "الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس"];
  const moniSheet = xlsx.utils.aoa_to_sheet([moniHeader]);
  xlsx.utils.book_append_sheet(wb, moniSheet, MONITORING_SHEET);

  // 7. sheet: rotation
  const rotationHeader = ["الأسبوع", "اليوم", "التاريخ", "المناوب"];
  const rotationSheet = xlsx.utils.aoa_to_sheet([rotationHeader]);
  xlsx.utils.book_append_sheet(wb, rotationSheet, ROTATION_SHEET);

  return wb;
}

// 🛠️ دالة تحديد وإعادة بناء ملف المدرسة (نسخة محدثة وآمنة)
async function getUserExcelPath(req) {
  const username = req.session?.username;
  
  if (!username) {
    throw new Error("UNAUTHORIZED: يرجى تسجيل الدخول أولاً للوصول لبيانات المدرسة.");
  }

  const cleanUsername = String(username).trim();
  const uploadsDir = path.join(__dirname, 'uploads');
  if (!fs.existsSync(uploadsDir)) {
    fs.mkdirSync(uploadsDir, { recursive: true });
  }

  const userFilePath = path.join(uploadsDir, `data_${cleanUsername}.xlsx`);

  try {
    const dbResult = await pool.query('SELECT excel_data FROM users WHERE username = $1', [cleanUsername]);
    
    if (dbResult.rows.length > 0 && dbResult.rows[0].excel_data) {
      if (fs.existsSync(userFilePath)) {
        fs.unlinkSync(userFilePath);
      }
      fs.writeFileSync(userFilePath, dbResult.rows[0].excel_data);
      console.log(`⚡ تم تحديث الملف للمستخدم (${cleanUsername}) من PostgreSQL`);
      return userFilePath;
    } else {
      if (fs.existsSync(userFilePath)) {
        fs.unlinkSync(userFilePath);
      }
      console.log(`🧹 إنشاء ملف جديد نظيف تماماً للمستخدم (${cleanUsername})`);
      const cleanWb = createCleanSchoolWorkbook();
      xlsx.writeFile(cleanWb, userFilePath);
      await syncExcelToDb(cleanUsername, userFilePath);
      return userFilePath;
    }
  } catch (err) {
    console.error('❌ خطأ في معالجة ملف الإكسل:', err.message);
    throw err;
  }
}

function getSheet(workbook, sheetIdentifier) {
  if (typeof sheetIdentifier === 'number') {
    return workbook.Sheets[workbook.SheetNames[sheetIdentifier]];
  }
  return workbook.Sheets[sheetIdentifier];
}

function normalizeArabicText(str) {
  if (!str) return '';
  return str.toString().trim().replace(/[أإآا]/g, 'ا').replace(/ة/g, 'ه').replace(/\s+/g, '');
}

const CANONICAL_WAITING_DAYS = ['الأحد', 'الإثنين', 'الثلاثاء', 'الأربعاء', 'الخميس'];

function canonicalizeDayName(day) {
  const normalizedInput = normalizeArabicText(day);
  const match = CANONICAL_WAITING_DAYS.find(d => normalizeArabicText(d) === normalizedInput);
  return match || day;
}

// =========================================================================
// 📊 مسارات جلب وحفظ البيانات الأساسية
// =========================================================================

// 📥 مسار تنزيل نموذج الإكسل المباشر من ملفات السيرفر
app.get('/download-template', (req, res) => {
  const templatePath = path.join(__dirname, 'templates', 'waiting_data.xlsx');

  // التحقق من وجود الملف في المسار قبل تنزيله
  if (fs.existsSync(templatePath)) {
    return res.download(templatePath, 'waiting_data.xlsx', (err) => {
      if (err) {
        console.error("خطأ أثناء إرسال الملف:", err);
        if (!res.headersSent) {
          return res.status(500).send("حدث خطأ أثناء تنزيل الملف.");
        }
      }
    });
  } else {
    console.error("❌ ملف النموذج غير موجود في المسار:", templatePath);
    return res.status(404).send("عذراً، لم يتم العثور على ملف النموذج في السيرفر.");
  }
});

app.get('/monitoring', (req, res) => {
  res.sendFile(path.join(__dirname, 'monitoring.html'));
});

// =========================================================================
// 🗄️ بيانات المدرسة الأساسية + الانتظار + الغياب — كلها في PostgreSQL
// (المعلمون، الفصول، معلومات المدرسة، جدول الانتظار، سجلات الغياب والبدلاء)
// لا يوجد أي اعتماد على ملف الإكسل أثناء الطلبات. الإكسل يُقرأ فقط:
//   1) مرة واحدة تلقائياً لنقل البيانات القديمة، 2) عند رفع ملف جديد من المدير/الأدمن.
// =========================================================================

function normName(s) {
  return String(s ?? '')
    .replace(/[ً-ْـ]/g, '')
    .replace(/[أإآ]/g, 'ا').replace(/ى/g, 'ي').replace(/ة/g, 'ه')
    .replace(/\s+/g, ' ').trim();
}

function toWesternDigits(s) {
  return String(s ?? '')
    .replace(/[٠-٩]/g, d => '٠١٢٣٤٥٦٧٨٩'.indexOf(d))
    .replace(/[۰-۹]/g, d => '۰۱۲۳۴۵۶۷۸۹'.indexOf(d));
}

// ---------------------------- التواريخ ----------------------------------
// المعتمد في المنصة: تاريخ هجري بصيغة YYYY/MM/DD. أي تاريخ ميلادي قديم يُحوَّل لهجري.
const _hijriFmt = new Intl.DateTimeFormat('en-US-u-ca-islamic-umalqura', {
  year: 'numeric', month: '2-digit', day: '2-digit', timeZone: 'UTC'
});

function gregorianToHijriStr(y, m, d) {
  const dt = new Date(Date.UTC(y, m - 1, d));
  if (isNaN(dt.getTime()) || dt.getUTCMonth() !== m - 1) return null;
  const p = {};
  _hijriFmt.formatToParts(dt).forEach(x => { p[x.type] = x.value; });
  return `${String(p.year).padStart(4, '0')}/${String(p.month).padStart(2, '0')}/${String(p.day).padStart(2, '0')}`;
}

function ymdToHijri(y, m, d) {
  if (m < 1 || m > 12 || d < 1 || d > 31) return null;
  if (y >= 1300 && y <= 1600) return `${y}/${String(m).padStart(2, '0')}/${String(d).padStart(2, '0')}`;
  if (y >= 1900 && y <= 2200) return gregorianToHijriStr(y, m, d);
  return null;
}

function parseDateStrict(v) {
  if (v === null || v === undefined || v === '') return null;
  if (v instanceof Date) return gregorianToHijriStr(v.getUTCFullYear(), v.getUTCMonth() + 1, v.getUTCDate());
  if (typeof v === 'number') {
    if (v > 20000 && v < 80000) { // رقم تسلسلي من إكسل
      const dt = new Date(Math.round((v - 25569) * 86400000));
      return gregorianToHijriStr(dt.getUTCFullYear(), dt.getUTCMonth() + 1, dt.getUTCDate());
    }
    return null;
  }
  const s = toWesternDigits(v).replace(/[‎‏]/g, '').trim();
  let m = s.match(/^(\d{4})\s*[\/\-.]\s*(\d{1,2})\s*[\/\-.]\s*(\d{1,2})$/);
  if (m) return ymdToHijri(+m[1], +m[2], +m[3]);
  m = s.match(/^(\d{1,2})\s*[\/\-.]\s*(\d{1,2})\s*[\/\-.]\s*(\d{4})$/);
  if (m) return ymdToHijri(+m[3], +m[2], +m[1]);
  return null;
}

function normalizeDateLoose(v) {
  return parseDateStrict(v) || toWesternDigits(v).replace(/[‎‏]/g, '').trim().replace(/-/g, '/');
}

// ---------------------------- أدوات قاعدة البيانات -------------------------
async function withTx(fn) {
  const client = await pool.connect();
  try {
    await client.query('BEGIN');
    const out = await fn(client);
    await client.query('COMMIT');
    return out;
  } catch (e) {
    try { await client.query('ROLLBACK'); } catch (_) {}
    throw e;
  } finally {
    client.release();
  }
}

function sendWaitingError(res, e, prefix) {
  const status = e.status || 500;
  const msg = (prefix ? prefix + ' ' : '') + e.message;
  if (status >= 500) console.error('❌', msg);
  return res.status(status).json({ success: false, error: msg, message: msg });
}

// ---------------------------- الاستيراد من الإكسل -------------------------
function findSheetName(workbook, wanted) {
  const w = String(wanted).trim().toLowerCase();
  return workbook.SheetNames.find(n => String(n).trim().toLowerCase() === w) || null;
}

function cellStr(v) {
  if (v === null || v === undefined) return '';
  return String(v).trim();
}

function parseTeachersSheet(workbook) {
  const name = findSheetName(workbook, TEACHERS_SHEET);
  if (!name) return null;
  const rows = xlsx.utils.sheet_to_json(workbook.Sheets[name], { header: 1 });
  const header = (rows[0] || []).map(h => String(h ?? ''));
  let phoneCol = header.findIndex(h => /جوال|هاتف|موبايل|تواصل|phone|mobile/i.test(h));
  if (phoneCol < 0) phoneCol = 2;
  const out = [];
  rows.slice(1).forEach(r => {
    if (!r) return;
    const nm = cellStr(r[1]);
    if (!nm) return;
    out.push({ national_id: cellStr(r[0]), name: nm, phone: cellStr(r[phoneCol]) });
  });
  return out;
}

function parseClassesSheet(workbook) {
  const name = findSheetName(workbook, CLASSES_SHEET);
  if (!name) return null;
  const rows = xlsx.utils.sheet_to_json(workbook.Sheets[name], { header: 1 });
  const seen = new Set();
  const out = [];
  rows.slice(1).forEach(r => {
    const c = cellStr(r && r[0]);
    if (c && !seen.has(c)) { seen.add(c); out.push(c); }
  });
  return out;
}

function parseInfoSheet(workbook) {
  const name = findSheetName(workbook, MAIN_INFO_SHEET);
  if (!name) return null;
  const rows = xlsx.utils.sheet_to_json(workbook.Sheets[name], { header: 1 });
  const v = i => cellStr(rows[i] && rows[i][1]);
  // القالب النظيف القديم: (المدرسة، القطاع، المدير) بدل (الإدارة، المدرسة، المدير، الوكيل)
  if (cellStr(rows[0] && rows[0][0]) === 'المدرسة') {
    return { education_dept: v(1), school_name: v(0), principal: v(2), deputy: v(3) };
  }
  return { education_dept: v(0), school_name: v(1), principal: v(2), deputy: v(3) };
}

function parseWaitingSheet(workbook) {
  const name = findSheetName(workbook, WAITING_SHEET);
  if (!name) return [];
  const data = xlsx.utils.sheet_to_json(workbook.Sheets[name], { header: 1 });
  const out = []; // { day, row_no, period, value }
  CANONICAL_WAITING_DAYS.forEach(day => {
    const idx = data.findIndex(r => r && normalizeArabicText(r[0]) === normalizeArabicText(day));
    if (idx === -1) return;
    for (let r = 0; r < 4; r++) {
      const row = data[idx + 1 + r] || [];
      for (let c = 0; c < 7; c++) {
        const val = cellStr(row[c]);
        if (val) out.push({ day, row_no: r, period: c + 1, value: val });
      }
    }
  });
  return out;
}

function parseReportSheet(workbook) {
  const name = findSheetName(workbook, REPORT_SHEET);
  if (!name) return [];
  const data = xlsx.utils.sheet_to_json(workbook.Sheets[name], { header: 1 });
  const out = [];
  data.forEach(row => {
    if (!row) return;
    const date = parseDateStrict(row[0]);
    const teacher = cellStr(row[2]);
    if (!date || !teacher) return; // يتجاوز صفوف العناوين والصفوف الفارغة
    const dayRaw = cellStr(row[1]);
    const slots = [];
    for (let p = 1; p <= 7; p++) {
      const cls = cellStr(row[3 + (p - 1) * 2]);
      const sub = cellStr(row[4 + (p - 1) * 2]);
      if (cls || sub) slots.push({ period: p, class_name: cls, sub_name: sub });
    }
    out.push({ date, day: canonicalizeDayName(dayRaw), teacher, slots });
  });
  return out;
}

async function loadTeacherMaps(client, username) {
  const r = await client.query(
    'SELECT id, national_id, name, phone FROM school_teachers WHERE username = $1 ORDER BY sort_order, id', [username]);
  const byName = new Map();
  r.rows.forEach(t => { const k = normName(t.name); if (k && !byName.has(k)) byName.set(k, t); });
  return { rows: r.rows, byName };
}

// يحدّث المعلمين/الفصول/معلومات المدرسة من ملف إكسل دون المساس بالانتظار والغياب.
// المعلم يُطابَق بالسجل المدني (ثم بالاسم)، فتعديل اسمه ينعكس تلقائياً على الجداول والتقارير.
async function applyBaseFromWorkbook(client, username, workbook) {
  const teachers = parseTeachersSheet(workbook);
  if (teachers) {
    const existing = (await client.query(
      'SELECT id, national_id, name FROM school_teachers WHERE username = $1', [username])).rows;
    const used = new Set();
    const byNid = new Map();
    const byNm = new Map();
    existing.forEach(t => {
      if (t.national_id) byNid.set(String(t.national_id), t);
      const k = normName(t.name);
      if (k && !byNm.has(k)) byNm.set(k, t);
    });
    let order = 0;
    for (const t of teachers) {
      order += 1;
      let match = t.national_id ? byNid.get(t.national_id) : null;
      if (match && used.has(match.id)) match = null;
      if (!match) {
        const m2 = byNm.get(normName(t.name));
        if (m2 && !used.has(m2.id)) match = m2;
      }
      if (match) {
        used.add(match.id);
        await client.query(
          'UPDATE school_teachers SET national_id = $1, name = $2, phone = $3, sort_order = $4 WHERE id = $5',
          [t.national_id, t.name, t.phone, order, match.id]);
      } else {
        await client.query(
          'INSERT INTO school_teachers (username, national_id, name, phone, sort_order) VALUES ($1, $2, $3, $4, $5)',
          [username, t.national_id, t.name, t.phone, order]);
      }
    }
    for (const t of existing) {
      if (!used.has(t.id)) await client.query('DELETE FROM school_teachers WHERE id = $1', [t.id]);
    }
    // ربط الخانات النصية القديمة بمعلم صار موجوداً بنفس الاسم
    const { byName } = await loadTeacherMaps(client, username);
    const loose = await client.query(
      "SELECT id, teacher_label FROM waiting_slots WHERE username = $1 AND teacher_id IS NULL", [username]);
    for (const s of loose.rows) {
      const hit = byName.get(normName(s.teacher_label));
      if (hit) await client.query('UPDATE waiting_slots SET teacher_id = $1, teacher_label = $2 WHERE id = $3', [hit.id, '', s.id]);
    }
  }

  const classes = parseClassesSheet(workbook);
  if (classes) {
    await client.query('DELETE FROM school_classes WHERE username = $1', [username]);
    let i = 0;
    for (const c of classes) {
      i += 1;
      await client.query('INSERT INTO school_classes (username, name, sort_order) VALUES ($1, $2, $3)', [username, c, i]);
    }
  }

  const info = parseInfoSheet(workbook);
  if (info) {
    const u = await client.query('SELECT school_name FROM users WHERE username = $1', [username]);
    const schoolName = info.school_name || (u.rows[0] && u.rows[0].school_name) || '';
    await client.query('DELETE FROM school_info WHERE username = $1', [username]);
    await client.query(
      'INSERT INTO school_info (username, education_dept, school_name, principal, deputy) VALUES ($1, $2, $3, $4, $5)',
      [username, info.education_dept, schoolName, info.principal, info.deputy]);
  }
}

async function insertAbsence(client, username, rec, byName) {
  const t = byName.get(normName(rec.teacher));
  const ins = await client.query(
    'INSERT INTO absence_records (username, abs_date, day_name, teacher_id, teacher_name) VALUES ($1, $2, $3, $4, $5) RETURNING id',
    [username, rec.date, rec.day, t ? t.id : null, t ? t.name : rec.teacher]);
  const id = ins.rows[0].id;
  for (const s of rec.slots) {
    const sub = s.sub_name ? byName.get(normName(s.sub_name)) : null;
    await client.query(
      'INSERT INTO absence_slots (absence_id, period, class_name, sub_id, sub_name) VALUES ($1, $2, $3, $4, $5)',
      [id, s.period, s.class_name, sub ? sub.id : null, sub ? sub.name : s.sub_name]);
  }
  return id;
}

async function setWaitingSlots(client, username, day, cells, byName) {
  await client.query('DELETE FROM waiting_slots WHERE username = $1 AND day_name = $2', [username, day]);
  for (const c of cells) {
    const t = byName.get(normName(c.value));
    await client.query(
      'INSERT INTO waiting_slots (username, day_name, row_no, period, teacher_id, teacher_label) VALUES ($1, $2, $3, $4, $5, $6)',
      [username, day, c.row_no, c.period, t ? t.id : null, t ? '' : c.value]);
  }
}

// استيراد كامل (أول مرة): الأساسيات + الانتظار + الغياب
async function importFullFromWorkbook(username, workbook) {
  await withTx(async (client) => {
    await client.query('DELETE FROM waiting_slots WHERE username = $1', [username]);
    await client.query(
      'DELETE FROM absence_slots WHERE absence_id IN (SELECT id FROM absence_records WHERE username = $1)', [username]);
    await client.query('DELETE FROM absence_records WHERE username = $1', [username]);

    await applyBaseFromWorkbook(client, username, workbook);
    const { byName } = await loadTeacherMaps(client, username);

    const cells = parseWaitingSheet(workbook);
    for (const day of CANONICAL_WAITING_DAYS) {
      await setWaitingSlots(client, username, day, cells.filter(c => c.day === day), byName);
    }
    for (const rec of parseReportSheet(workbook)) {
      await insertAbsence(client, username, rec, byName);
    }
    await client.query('UPDATE users SET waiting_migrated = TRUE WHERE username = $1', [username]);
  });
}

async function importBaseFromWorkbook(username, workbook) {
  await withTx(async (client) => { await applyBaseFromWorkbook(client, username, workbook); });
}

const _waitingReady = new Set();
const _waitingPending = new Map();

// نقل البيانات القديمة من الإكسل مرة واحدة لكل مدرسة (آمن لو وصل طلبان معاً)
function ensureWaitingReady(username) {
  if (_waitingReady.has(username)) return Promise.resolve();
  if (_waitingPending.has(username)) return _waitingPending.get(username);
  const p = (async () => {
    const r = await pool.query('SELECT waiting_migrated, excel_data FROM users WHERE username = $1', [username]);
    if (r.rows.length === 0) return;
    if (r.rows[0].waiting_migrated) { _waitingReady.add(username); return; }
    if (r.rows[0].excel_data) {
      const wb = xlsx.read(r.rows[0].excel_data, { type: 'buffer' });
      await importFullFromWorkbook(username, wb);
      console.log(`🔁 تم نقل بيانات الانتظار والغياب للمستخدم (${username}) من الإكسل إلى قاعدة البيانات.`);
    } else {
      await pool.query('UPDATE users SET waiting_migrated = TRUE WHERE username = $1', [username]);
    }
    _waitingReady.add(username);
  })().finally(() => _waitingPending.delete(username));
  _waitingPending.set(username, p);
  return p;
}

// بعد رفع ملف إكسل جديد: لو المدرسة لم تُنقل بعد نستورد كل شيء من الملف الجديد،
// وإلا نحدّث المعلمين والفصول والمعلومات فقط (الانتظار والغياب لا تُمسّ أبداً).
async function refreshFromUploadedExcel(username, buffer) {
  const wb = xlsx.read(buffer, { type: 'buffer' });
  const r = await pool.query('SELECT waiting_migrated FROM users WHERE username = $1', [username]);
  if (r.rows.length === 0) return;
  if (r.rows[0].waiting_migrated) {
    await importBaseFromWorkbook(username, wb);
  } else {
    await importFullFromWorkbook(username, wb);
  }
  _waitingReady.add(username);
}

async function prepareWaiting(req) {
  const username = requireSessionUser(req);
  await ensureWaitingReady(username);
  return username;
}

// عند تغيير اسم المستخدم من لوحة الأدمن تتبعه كل الجداول
async function renameSchoolUsername(oldName, newName) {
  if (!oldName || !newName || oldName === newName) return;
  for (const t of ['school_teachers', 'school_classes', 'school_info', 'waiting_slots', 'absence_records',
                   'monitoring_schedule', 'monitoring_attendance', 'accountability', 'wa_auth']) {
    await pool.query(`UPDATE ${t} SET username = $1 WHERE username = $2`, [newName, oldName]);
  }
  _waitingReady.delete(oldName);
  _waitingReady.add(newName);
}

async function deleteSchoolData(username) {
  await pool.query('DELETE FROM waiting_slots WHERE username = $1', [username]);
  await pool.query('DELETE FROM absence_slots WHERE absence_id IN (SELECT id FROM absence_records WHERE username = $1)', [username]);
  await pool.query('DELETE FROM absence_records WHERE username = $1', [username]);
  for (const t of ['school_teachers', 'school_classes', 'school_info']) {
    await pool.query(`DELETE FROM ${t} WHERE username = $1`, [username]);
  }
  _waitingReady.delete(username);
}

// ---------------------------- قراءة من القاعدة ---------------------------
async function dbTeachers(username) {
  const r = await pool.query(
    'SELECT id, national_id, name, phone FROM school_teachers WHERE username = $1 ORDER BY sort_order, id', [username]);
  return r.rows;
}

async function dbClasses(username) {
  const r = await pool.query('SELECT name FROM school_classes WHERE username = $1 ORDER BY sort_order, id', [username]);
  return r.rows.map(x => x.name);
}

async function dbSchoolInfo(username) {
  const r = await pool.query(
    'SELECT education_dept, school_name, principal, deputy FROM school_info WHERE username = $1', [username]);
  const i = r.rows[0] || {};
  return { B1: i.education_dept || '', B2: i.school_name || '', B3: i.principal || '', B4: i.deputy || '' };
}

async function dbWaitingMatrix(username) {
  const r = await pool.query(
    `SELECT w.day_name, w.row_no, w.period, w.teacher_label, t.name AS teacher_name
       FROM waiting_slots w LEFT JOIN school_teachers t ON t.id = w.teacher_id
      WHERE w.username = $1`, [username]);
  const schedule = {};
  CANONICAL_WAITING_DAYS.forEach(d => { schedule[d] = [0, 1, 2, 3].map(() => ['', '', '', '', '', '', '']); });
  r.rows.forEach(s => {
    const day = canonicalizeDayName(s.day_name);
    if (!schedule[day] || s.row_no < 0 || s.row_no > 3 || s.period < 1 || s.period > 7) return;
    schedule[day][s.row_no][s.period - 1] = s.teacher_name || s.teacher_label || '';
  });
  return schedule;
}

// سجلات الغياب بالشكل القديم تماماً (ترتيب المفاتيح مهم لدوال الـ PDF التي تعتمد Object.values)
async function dbAbsenceItems(username, filter = {}) {
  const recs = (await pool.query(
    `SELECT r.id, r.abs_date, r.day_name, r.teacher_name, t.name AS cur_name, t.phone AS cur_phone
       FROM absence_records r LEFT JOIN school_teachers t ON t.id = r.teacher_id
      WHERE r.username = $1 ORDER BY r.id`, [username])).rows;
  const slots = (await pool.query(
    `SELECT s.absence_id, s.period, s.class_name, s.sub_name, st.name AS cur_sub, st.phone AS sub_phone
       FROM absence_slots s
       JOIN absence_records r ON r.id = s.absence_id
       LEFT JOIN school_teachers st ON st.id = s.sub_id
      WHERE r.username = $1`, [username])).rows;
  const teachers = await dbTeachers(username);
  const phoneByName = new Map();
  teachers.forEach(t => { if (t.phone) phoneByName.set(normName(t.name), t.phone); });

  const slotMap = new Map();
  slots.forEach(s => {
    if (!slotMap.has(s.absence_id)) slotMap.set(s.absence_id, {});
    slotMap.get(s.absence_id)[s.period] = s;
  });

  const wantDate = filter.date ? normalizeDateLoose(filter.date) : '';
  const wantDay = filter.day ? normalizeArabicText(filter.day) : '';

  const items = [];
  recs.forEach(r => {
    if (wantDate && normalizeDateLoose(r.abs_date) !== wantDate) return;
    if (wantDay && normalizeArabicText(r.day_name) !== wantDay) return;
    const name = r.cur_name || r.teacher_name;
    const item = { 'التاريخ': r.abs_date, 'اليوم': r.day_name, 'المعلم الغائب': name };
    const subPhones = [];
    for (let p = 1; p <= 7; p++) {
      const s = (slotMap.get(r.id) || {})[p];
      const sub = s ? (s.cur_sub || s.sub_name || '') : '';
      item[`فصل ${p}`] = s ? (s.class_name || '') : '';
      item[`بديل ${p}`] = sub;
      subPhones.push(s ? (s.sub_phone || phoneByName.get(normName(sub)) || '') : '');
    }
    item._id = r.id;
    item.absentPhone = r.cur_phone || phoneByName.get(normName(name)) || '';
    item.subPhones = subPhones;
    items.push(item);
  });
  return items;
}

// =========================================================================
// 📊 مسارات المعلمين والفصول (من القاعدة)
// =========================================================================
app.get('/get-data', async (req, res) => {
  try {
    const username = await prepareWaiting(req);
    const rows = await dbTeachers(username);
    res.json(rows.map(t => ({ 'السجل المدني': t.national_id || '', Name_teacher: t.name, 'رقم الجوال': t.phone || '' })));
  } catch (e) { sendWaitingError(res, e); }
});

app.get('/get-classes', async (req, res) => {
  try {
    const username = await prepareWaiting(req);
    res.json((await dbClasses(username)).map(n => ({ Name_class: n })));
  } catch (e) { sendWaitingError(res, e); }
});

app.get('/api/teachers', async (req, res) => {
  try {
    const username = await prepareWaiting(req);
    res.json((await dbTeachers(username)).map(t => t.name));
  } catch (e) { sendWaitingError(res, e); }
});

// بيانات المعلمين الكاملة [السجل المدني، الاسم، الجوال] — تُستخدم في تقرير المساءلة والواتساب
app.get('/api/teachers-raw', async (req, res) => {
  try {
    const username = await prepareWaiting(req);
    res.json((await dbTeachers(username)).map(t => [t.national_id || '', t.name, t.phone || '']));
  } catch (e) { sendWaitingError(res, e); }
});

// =========================================================================
// ⏳ جدول الانتظار (من القاعدة)
// =========================================================================
app.get('/get-waiting-teachers', async (req, res) => {
  try {
    const username = await prepareWaiting(req);
    const day = canonicalizeDayName(req.query.day);
    const period = parseInt(req.query.period, 10);
    if (!(period >= 1 && period <= 7)) return res.json({ teachers: [] });
    const schedule = await dbWaitingMatrix(username);
    const dayRows = schedule[day];
    if (!dayRows) return res.json({ teachers: [] });
    res.json({ teachers: dayRows.map(r => r[period - 1]).filter(Boolean) });
  } catch (e) { sendWaitingError(res, e); }
});

app.post('/save-waiting-schedule-day', async (req, res) => {
  try {
    const { day, schedule } = req.body || {};
    if (!day || !Array.isArray(schedule)) {
      return res.status(400).json({ success: false, error: 'بيانات غير صالحة: يجب إرسال day و schedule' });
    }
    const username = await prepareWaiting(req);
    const canonicalDay = canonicalizeDayName(day);
    const { byName } = await loadTeacherMaps(pool, username);

    const cells = [];
    for (let r = 0; r < 4; r++) {
      const rowValues = Array.isArray(schedule[r]) ? schedule[r] : [];
      for (let c = 0; c < 7; c++) {
        const val = cellStr(rowValues[c]);
        if (val) cells.push({ row_no: r, period: c + 1, value: val });
      }
    }
    await withTx(async (client) => { await setWaitingSlots(client, username, canonicalDay, cells, byName); });
    res.json({ success: true });
  } catch (e) { sendWaitingError(res, e, 'خطأ أثناء الحفظ:'); }
});

app.get('/get-waiting-schedule-full', async (req, res) => {
  try {
    const username = await prepareWaiting(req);
    res.json({ success: true, schedule: await dbWaitingMatrix(username) });
  } catch (e) { sendWaitingError(res, e, 'خطأ أثناء قراءة الجدول:'); }
});

app.get('/api/get-rotation', async (req, res) => {
  try {
    const userExcel = await getUserExcelPath(req);
    const workbook = xlsx.readFile(userExcel);

    if (!workbook.SheetNames.includes(ROTATION_SHEET)) {
      return res.json([]);
    }

    const rows = xlsx.utils.sheet_to_json(workbook.Sheets[ROTATION_SHEET]);
    res.json(rows);
  } catch (e) {
    res.status(500).json({ error: "خطأ في قراءة جدول المناوبة: " + e.message });
  }
});

app.post('/api/save-rotation', async (req, res) => {
  const { weekTitle, weekData } = req.body;
  try {
    if (!weekTitle || !Array.isArray(weekData)) {
      return res.status(400).json({ success: false, error: "بيانات غير صالحة: يجب إرسال weekTitle و weekData" });
    }

    const username = req.session?.username || req.headers['x-username'] || req.body?.username;
    const userExcel = await getUserExcelPath(req);
    const workbook = xlsx.readFile(userExcel);

    const header = ["الأسبوع", "اليوم", "التاريخ", "المناوب"];
    let rows = [];
    if (workbook.SheetNames.includes(ROTATION_SHEET)) {
      rows = xlsx.utils.sheet_to_json(workbook.Sheets[ROTATION_SHEET]);
    }

    weekData.forEach(entry => {
      const existingIndex = rows.findIndex(r => r['الأسبوع'] === weekTitle && r['اليوم'] === entry.day);
      const newRow = {
        "الأسبوع": weekTitle,
        "اليوم": entry.day,
        "التاريخ": entry.date || "",
        "المناوب": entry.teacher || ""
      };
      if (existingIndex !== -1) {
        rows[existingIndex] = newRow;
      } else {
        rows.push(newRow);
      }
    });

    const newSheet = xlsx.utils.json_to_sheet(rows, { header });
    if (workbook.SheetNames.includes(ROTATION_SHEET)) {
      workbook.Sheets[ROTATION_SHEET] = newSheet;
    } else {
      xlsx.utils.book_append_sheet(workbook, newSheet, ROTATION_SHEET);
    }

    xlsx.writeFile(workbook, userExcel);
    await syncOrThrow(username, userExcel);

    res.json({ success: true, message: "تم حفظ بيانات الأسبوع بنجاح" });
  } catch (e) {
    res.status(500).json({ success: false, error: "خطأ أثناء الحفظ: " + e.message });
  }
});

// =========================================================================
// 📝 سجلات الغياب والبدلاء (من القاعدة)
// =========================================================================
app.post('/save-report', async (req, res) => {
  try {
    const b = req.body || {};
    const date = normalizeDateLoose(b.date);
    const absent = cellStr(b.absentTeacher);
    if (!date || !absent) {
      return res.status(400).json({ success: false, error: 'يرجى تحديد التاريخ واسم المعلم الغائب' });
    }
    const username = await prepareWaiting(req);
    const { byName } = await loadTeacherMaps(pool, username);
    const slots = [];
    for (let i = 1; i <= 7; i++) {
      const cls = cellStr(b[`class${i}_room`]);
      const sub = cellStr(b[`class${i}_teacher`]);
      if (cls || sub) slots.push({ period: i, class_name: cls, sub_name: sub });
    }
    const day = canonicalizeDayName(cellStr(b.day));
    const id = await withTx(client => insertAbsence(client, username, { date, day, teacher: absent, slots }, byName));
    res.json({ success: true, id });
  } catch (e) { sendWaitingError(res, e, 'خطأ أثناء الحفظ:'); }
});

// 🗑️ حذف سجل غياب: بالمعرّف إن وُجد، وإلا بالمطابقة (التاريخ أو اليوم + اسم المعلم)
app.post('/delete-absence-record', async (req, res) => {
  try {
    const username = await prepareWaiting(req);
    const teacherName = req.body.teacherName || req.body.absentTeacher || req.body.teacher;
    const recordId = parseInt(req.body.recordId ?? req.body.id, 10);

    let target = null;
    if (recordId > 0) {
      const r = await pool.query('SELECT id FROM absence_records WHERE id = $1 AND username = $2', [recordId, username]);
      target = r.rows[0] || null;
    }
    if (!target && teacherName) {
      const wantName = normName(teacherName);
      const wantDate = req.body.date ? normalizeDateLoose(req.body.date) : '';
      const wantDay = req.body.day ? normalizeArabicText(req.body.day) : '';
      const all = (await pool.query(
        `SELECT r.id, r.abs_date, r.day_name, r.teacher_name, t.name AS cur_name
           FROM absence_records r LEFT JOIN school_teachers t ON t.id = r.teacher_id
          WHERE r.username = $1 ORDER BY r.id`, [username])).rows;
      target = all.find(r => {
        if (normName(r.cur_name || r.teacher_name) !== wantName && normName(r.teacher_name) !== wantName) return false;
        const dateOk = wantDate && normalizeDateLoose(r.abs_date) === wantDate;
        const dayOk = wantDay && normalizeArabicText(r.day_name) === wantDay;
        // يقبل التاريخ أو اليوم (قد يصل اسم اليوم في حقل التاريخ في الصفحات القديمة)
        const dateAsDay = wantDate && normalizeArabicText(r.day_name) === normalizeArabicText(req.body.date);
        return (dateOk && (!wantDay || dayOk)) || dateAsDay || (!wantDate && dayOk);
      }) || null;
    }
    if (!target) {
      return res.status(404).json({ success: false, message: 'لم يتم العثور على السجل المطلوب حذفه.' });
    }
    await withTx(async (client) => {
      await client.query('DELETE FROM absence_slots WHERE absence_id = $1', [target.id]);
      await client.query('DELETE FROM absence_records WHERE id = $1 AND username = $2', [target.id, username]);
    });
    res.json({ success: true, message: 'تم حذف السجل بنجاح' });
  } catch (e) { sendWaitingError(res, e, 'حدث خطأ أثناء الحذف:'); }
});

// 📥 سجلات الغياب ليوم وتاريخ محددين (تستخدمها صفحة التقرير)
app.get('/api/absences', async (req, res) => {
  try {
    const username = await prepareWaiting(req);
    const records = await dbAbsenceItems(username, { date: req.query.date, day: req.query.day });
    res.json({ success: true, records });
  } catch (e) { sendWaitingError(res, e); }
});

// تبقى بنفس الشكل القديم: teachers = سجلات الغياب ، mainInfo = معلومات المدرسة
app.get('/get-report-data', async (req, res) => {
  try {
    const username = await prepareWaiting(req);
    const [teachers, mainInfo] = await Promise.all([dbAbsenceItems(username), dbSchoolInfo(username)]);
    res.json({ teachers, mainInfo });
  } catch (e) { sendWaitingError(res, e); }
});

// 📤 يبني ملف إكسل كاملاً من القاعدة (للتنزيل من لوحة الأدمن)
async function buildWorkbookFromDb(username) {
  await ensureWaitingReady(username);
  const u = (await pool.query(
    'SELECT excel_data, school_name, region, full_name FROM users WHERE username = $1', [username])).rows[0] || {};
  const wb = u.excel_data
    ? xlsx.read(u.excel_data, { type: 'buffer' })
    : createCleanSchoolWorkbook(u.school_name || '', u.region || '', u.full_name || '');

  const put = (name, aoa) => {
    const ws = xlsx.utils.aoa_to_sheet(aoa);
    const existing = wb.SheetNames.find(n => String(n).trim() === String(name).trim());
    if (existing) wb.Sheets[existing] = ws;
    else xlsx.utils.book_append_sheet(wb, ws, name);
  };

  const teachers = await dbTeachers(username);
  put(TEACHERS_SHEET, [['السجل المدني', 'Name_teacher', 'رقم الجوال'], ...teachers.map(t => [t.national_id || '', t.name, t.phone || ''])]);
  put(CLASSES_SHEET, [['Name_class'], ...(await dbClasses(username)).map(c => [c])]);

  const info = await dbSchoolInfo(username);
  put(MAIN_INFO_SHEET, [['إدارة التعليم', info.B1], ['اسم المدرسة', info.B2], ['مدير المدرسة', info.B3], ['وكيل المدرسة', info.B4]]);

  const schedule = await dbWaitingMatrix(username);
  const waitRows = [];
  CANONICAL_WAITING_DAYS.forEach(d => {
    waitRows.push([d]);
    schedule[d].forEach(r => waitRows.push(r));
    waitRows.push([]);
  });
  put(WAITING_SHEET, waitRows);

  const header = ['التاريخ', 'اليوم', 'المعلم الغائب', 'فصل 1', 'بديل 1', 'فصل 2', 'بديل 2', 'فصل 3', 'بديل 3', 'فصل 4', 'بديل 4', 'فصل 5', 'بديل 5', 'فصل 6', 'بديل 6', 'فصل 7', 'بديل 7'];
  const items = await dbAbsenceItems(username);
  put(REPORT_SHEET, [header, ...items.map(it => header.map(h => it[h] ?? ''))]);

  return wb;
}

// =========================================================================
// 📋 مسارات الإشراف الأسبوعي
// =========================================================================

app.get('/get-monitoring-teachers', async (req, res) => {
  try {
    const username = await prepareWaiting(req);
    res.json({ success: true, teachers: (await dbTeachers(username)).map(t => t.name) });
  } catch (e) {
    res.status(e.status || 500).json({ success: false, error: e.message });
  }
});

// -------------------------------------------------------------------------
// 🗄️ الإشراف الأسبوعي — يُحفظ في PostgreSQL (جدول monitoring_schedule)
// -------------------------------------------------------------------------
const MONITORING_DAY_KEYS = ['sun', 'mon', 'tue', 'wed', 'thu'];
const MONITORING_DAY_NAMES = { sun: 'الأحد', mon: 'الإثنين', tue: 'الثلاثاء', wed: 'الأربعاء', thu: 'الخميس' };

function requireSessionUser(req) {
  const username = req.session?.username;
  if (!username) {
    const err = new Error('انتهت الجلسة، يرجى تسجيل الدخول من جديد');
    err.status = 401;
    throw err;
  }
  return String(username).trim();
}

// يقبل قائمة أسماء أو نصاً مفصولاً بـ "،" ويعيد قائمة نظيفة بلا تكرار
function toTeacherList(val) {
  let list = [];
  if (Array.isArray(val)) list = val;
  else if (typeof val === 'string') list = val.split(/[،,]/);
  return [...new Set(list.map(s => String(s ?? '').trim()).filter(Boolean))];
}

// يحوّل الجسم المرسل { sun: { "المقصف": [...] } } إلى صفوف جاهزة للإدخال
function scheduleToRows(schedule) {
  const rows = [];
  if (!schedule || typeof schedule !== 'object') return rows;
  MONITORING_DAY_KEYS.forEach(day => {
    const dayObj = schedule[day];
    if (!dayObj || typeof dayObj !== 'object') return;
    Object.keys(dayObj).forEach(rawTask => {
      const task = String(rawTask).trim();
      if (!task) return;
      toTeacherList(dayObj[rawTask]).forEach(teacher => rows.push([day, task, teacher]));
    });
  });
  return rows;
}

// 🔁 نقل جدول الإشراف القديم من ورقة moni في الإكسل إلى القاعدة (مرة واحدة لكل مدرسة)
async function migrateMonitoringFromExcel(req, username) {
  const flag = await pool.query('SELECT monitoring_migrated FROM users WHERE username = $1', [username]);
  if (flag.rows.length === 0 || flag.rows[0].monitoring_migrated) return;

  let rows = [];
  try {
    const filePath = await getUserExcelPath(req);
    const workbook = xlsx.readFile(filePath);
    const worksheet = workbook.Sheets[MONITORING_SHEET];
    if (worksheet) {
      const schedule = {};
      xlsx.utils.sheet_to_json(worksheet).forEach(row => {
        const task = row['المهمة'] || row['task'];
        if (!task) return;
        MONITORING_DAY_KEYS.forEach(day => {
          const cell = row[MONITORING_DAY_NAMES[day]];
          if (cell) {
            schedule[day] = schedule[day] || {};
            schedule[day][task] = String(cell);
          }
        });
      });
      rows = scheduleToRows(schedule);
    }
  } catch (e) {
    console.error(`⚠️ تعذر قراءة ورقة الإشراف من الإكسل للمستخدم (${username}):`, e.message);
    return; // نعيد المحاولة في الطلب القادم
  }

  const client = await pool.connect();
  try {
    await client.query('BEGIN');
    // "حجز" عملية النقل حتى لا تتكرر إذا وصل طلبان في نفس اللحظة
    const claim = await client.query(
      'UPDATE users SET monitoring_migrated = TRUE WHERE username = $1 AND monitoring_migrated IS NOT TRUE RETURNING id',
      [username]
    );
    if (claim.rows.length === 0) { await client.query('ROLLBACK'); return; }

    for (const [day, task, teacher] of rows) {
      await client.query(
        `INSERT INTO monitoring_schedule (username, day_key, task, teacher_name)
         VALUES ($1, $2, $3, $4) ON CONFLICT DO NOTHING`,
        [username, day, task, teacher]
      );
    }
    await client.query('COMMIT');
    console.log(`📦 تم نقل ${rows.length} تكليف إشراف من الإكسل إلى القاعدة للمستخدم (${username})`);
  } catch (e) {
    await client.query('ROLLBACK').catch(() => {});
    throw e;
  } finally {
    client.release();
  }
}

// يعيد الجدول بالشكل { sun: { "المقصف": ["أحمد", ...] }, mon: {...}, ... }
async function loadMonitoringSchedule(req, username) {
  await migrateMonitoringFromExcel(req, username);
  const result = await pool.query(
    'SELECT day_key, task, teacher_name FROM monitoring_schedule WHERE username = $1 ORDER BY id',
    [username]
  );
  const schedule = { sun: {}, mon: {}, tue: {}, wed: {}, thu: {} };
  result.rows.forEach(r => {
    (schedule[r.day_key][r.task] = schedule[r.day_key][r.task] || []).push(r.teacher_name);
  });
  return schedule;
}

// يستبدل جدول المدرسة بالكامل داخل معاملة واحدة (إما يُحفظ كله أو لا شيء)
async function replaceMonitoringSchedule(username, rows) {
  const client = await pool.connect();
  try {
    await client.query('BEGIN');
    await client.query('DELETE FROM monitoring_schedule WHERE username = $1', [username]);
    for (const [day, task, teacher] of rows) {
      await client.query(
        `INSERT INTO monitoring_schedule (username, day_key, task, teacher_name)
         VALUES ($1, $2, $3, $4) ON CONFLICT DO NOTHING`,
        [username, day, task, teacher]
      );
    }
    // بعد أول حفظ لا نعود للإكسل أبداً
    await client.query('UPDATE users SET monitoring_migrated = TRUE WHERE username = $1', [username]);
    await client.query('COMMIT');
  } catch (e) {
    await client.query('ROLLBACK').catch(() => {});
    throw e;
  } finally {
    client.release();
  }
}

function sendMonitoringError(res, e, prefix) {
  console.error(prefix, e);
  res.status(e.status || 500).json({ success: false, message: prefix + ' ' + e.message, error: e.message });
}

// 📥 جلب جدول الإشراف
app.get('/api/monitoring', async (req, res) => {
  try {
    const username = requireSessionUser(req);
    const schedule = await loadMonitoringSchedule(req, username);
    res.json({ success: true, schedule });
  } catch (e) {
    sendMonitoringError(res, e, 'حدث خطأ أثناء جلب جدول الإشراف:');
  }
});

// 💾 حفظ جدول الإشراف بالكامل
app.put('/api/monitoring', async (req, res) => {
  try {
    const username = requireSessionUser(req);
    const rows = scheduleToRows(req.body?.schedule);
    await replaceMonitoringSchedule(username, rows);
    res.json({ success: true, message: 'تم حفظ جدول الإشراف بنجاح', count: rows.length });
  } catch (e) {
    sendMonitoringError(res, e, 'حدث خطأ أثناء حفظ جدول الإشراف:');
  }
});

// 🗑️ تفريغ جدول الإشراف
app.delete('/api/monitoring', async (req, res) => {
  try {
    const username = requireSessionUser(req);
    await replaceMonitoringSchedule(username, []);
    res.json({ success: true, message: 'تم تفريغ جدول الإشراف' });
  } catch (e) {
    sendMonitoringError(res, e, 'حدث خطأ أثناء تفريغ جدول الإشراف:');
  }
});

// ↩️ المسارات القديمة — باقية لأي صفحة أخرى تستخدمها، لكنها الآن تقرأ وتكتب من القاعدة
app.post('/save-monitoring', async (req, res) => {
  try {
    const username = requireSessionUser(req);
    await replaceMonitoringSchedule(username, scheduleToRows(req.body));
    res.json({ success: true });
  } catch (e) {
    res.status(e.status || 500).json({ success: false, error: 'خطأ أثناء الحفظ: ' + e.message });
  }
});

app.get('/get-monitoring-schedule', async (req, res) => {
  try {
    const username = requireSessionUser(req);
    const schedule = await loadMonitoringSchedule(req, username);
    // نفس الشكل القديم: الأسماء نص واحد مفصول بـ "، "
    MONITORING_DAY_KEYS.forEach(day => {
      Object.keys(schedule[day]).forEach(task => { schedule[day][task] = schedule[day][task].join('، '); });
    });
    res.json({ success: true, schedule });
  } catch (e) {
    res.json({ success: false, error: 'تعذر قراءة جدول الإشراف: ' + e.message });
  }
});

app.get('/reportmoni', (req, res) => {
  res.sendFile(path.join(__dirname, 'reportmoni.html'));
});

app.get('/get-reportmoni-data', async (req, res) => {
  try {
    const { day, task } = req.query;
    const username = requireSessionUser(req);

    await ensureWaitingReady(username);
    const mainInfo = await dbSchoolInfo(username);

    if (!MONITORING_DAY_KEYS.includes(day)) return res.json({ success: true, data: [], mainInfo });

    const schedule = await loadMonitoringSchedule(req, username);
    const dayTasks = schedule[day] || {};

    const filteredResults = Object.keys(dayTasks)
      .filter(t => !task || task === 'all' || normalizeArabicText(t) === normalizeArabicText(task))
      .filter(t => dayTasks[t].length > 0)
      .map(t => ({ task: t, teachers: dayTasks[t].join('، ') }));

    res.json({ success: true, data: filteredResults, mainInfo });
  } catch (error) {
    res.json({ success: false, error: 'حدث خطأ في قراءة جدول الإشراف: ' + error.message });
  }
});


// -------------------------------------------------------------------------
// ✅ أداء الإشراف الفعلي + 💬 تذكير المشرفين بالواتساب
// -------------------------------------------------------------------------
const UTC_DAY_TO_KEY = { 0: 'sun', 1: 'mon', 2: 'tue', 3: 'wed', 4: 'thu' };

// يتحقق من التاريخ (YYYY-MM-DD) ويعيد مفتاح اليوم، أو يرمي خطأ واضح
function parseMonitoringDate(dateStr) {
  const s = String(dateStr || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) {
    const err = new Error('التاريخ غير صحيح'); err.status = 400; throw err;
  }
  const d = new Date(s + 'T00:00:00Z');
  if (isNaN(d.getTime()) || d.toISOString().slice(0, 10) !== s) {
    const err = new Error('التاريخ غير صحيح'); err.status = 400; throw err;
  }
  const dayKey = UTC_DAY_TO_KEY[d.getUTCDay()];
  if (!dayKey) {
    const err = new Error('التاريخ المختار يوم إجازة (الجمعة أو السبت)'); err.status = 400; throw err;
  }
  return { date: s, dayKey };
}

function formatHijriDate(dateStr) {
  try {
    return new Intl.DateTimeFormat('ar-SA-u-ca-islamic-umalqura-nu-latn', { day: 'numeric', month: 'long', year: 'numeric' })
      .format(new Date(dateStr + 'T12:00:00Z'));
  } catch (e) {
    return dateStr;
  }
}

// 📞 أرقام جوالات المعلمين من ورقة teacher_name
// نبحث أولاً عن عمود عنوانه فيه "جوال/هاتف/phone"، وإلا نلتقط أي خلية شكلها رقم جوال سعودي
// (لا نعتمد على أي عمود رقمي عشوائي حتى لا نرسل لرقم الهوية بالخطأ)
const SAUDI_MOBILE_RE = /^(?:\+?966|00966|0)?5\d{8}$/;

async function getTeacherPhones(req) {
  const username = await prepareWaiting(req);
  const phones = new Map();
  (await dbTeachers(username)).forEach(t => {
    const phone = String(t.phone || '').trim();
    if (t.name && phone) phones.set(String(t.name).trim(), phone);
  });
  return phones;
}

// 📥 حالة الأداء لتاريخ معيّن
app.get('/api/monitoring/attendance', async (req, res) => {
  try {
    const username = requireSessionUser(req);
    const { date } = parseMonitoringDate(req.query.date);
    const result = await pool.query(
      `SELECT task, teacher_name, status FROM monitoring_attendance WHERE username = $1 AND att_date = $2`,
      [username, date]
    );
    res.json({ success: true, data: result.rows });
  } catch (e) {
    sendMonitoringError(res, e, 'حدث خطأ أثناء جلب حالة الإشراف:');
  }
});

// 🔄 تسجيل أدّى / لم يؤدِّ (status = null يلغي التسجيل)
app.put('/api/monitoring/attendance', async (req, res) => {
  try {
    const username = requireSessionUser(req);
    const { date, dayKey } = parseMonitoringDate(req.body?.date);
    const task = String(req.body?.task || '').trim();
    const teacher = String(req.body?.teacher_name || '').trim();
    const status = req.body?.status ?? null;

    if (!task || !teacher) return res.status(400).json({ success: false, message: 'بيانات المعلم أو المهمة ناقصة' });
    if (status !== null && status !== 'done' && status !== 'missed') {
      return res.status(400).json({ success: false, message: 'الحالة غير صحيحة' });
    }

    if (status === null) {
      await pool.query(
        'DELETE FROM monitoring_attendance WHERE username = $1 AND att_date = $2 AND task = $3 AND teacher_name = $4',
        [username, date, task, teacher]
      );
      return res.json({ success: true, message: 'تم إلغاء التسجيل' });
    }

    // لا نسجّل إلا لمعلم مكلّف فعلاً بهذه المهمة في هذا اليوم
    const assigned = await pool.query(
      'SELECT 1 FROM monitoring_schedule WHERE username = $1 AND day_key = $2 AND task = $3 AND teacher_name = $4',
      [username, dayKey, task, teacher]
    );
    if (assigned.rows.length === 0) {
      return res.status(400).json({ success: false, message: 'المعلم غير مكلّف بهذه المهمة في هذا اليوم حسب جدول الإشراف' });
    }

    const result = await pool.query(
      `INSERT INTO monitoring_attendance (username, att_date, day_key, task, teacher_name, status)
       VALUES ($1, $2, $3, $4, $5, $6)
       ON CONFLICT (username, att_date, task, teacher_name)
       DO UPDATE SET status = EXCLUDED.status, updated_at = CURRENT_TIMESTAMP
       RETURNING task, teacher_name, status`,
      [username, date, dayKey, task, teacher, status]
    );
    res.json({ success: true, message: 'تم التسجيل', data: result.rows[0] });
  } catch (e) {
    sendMonitoringError(res, e, 'حدث خطأ أثناء تسجيل حالة الإشراف:');
  }
});

// 📊 رصيد كل معلم (لصفحة الإحصائيات لاحقاً) — from / to اختياريان بصيغة YYYY-MM-DD
app.get('/api/monitoring/attendance/stats', async (req, res) => {
  try {
    const username = requireSessionUser(req);
    const from = req.query.from ? parseMonitoringDateLoose(req.query.from) : null;
    const to = req.query.to ? parseMonitoringDateLoose(req.query.to) : null;
    const result = await pool.query(
      `SELECT teacher_name,
              COUNT(*) FILTER (WHERE status = 'done')::int   AS done_days,
              COUNT(*) FILTER (WHERE status = 'missed')::int AS missed_days,
              MAX(att_date) AS last_date
       FROM monitoring_attendance
       WHERE username = $1
         AND ($2::date IS NULL OR att_date >= $2::date)
         AND ($3::date IS NULL OR att_date <= $3::date)
       GROUP BY teacher_name
       ORDER BY teacher_name`,
      [username, from, to]
    );
    res.json({ success: true, data: result.rows });
  } catch (e) {
    sendMonitoringError(res, e, 'حدث خطأ أثناء حساب أرصدة الإشراف:');
  }
});

function parseMonitoringDateLoose(s) {
  const v = String(s).trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(v) || isNaN(new Date(v + 'T00:00:00Z'))) {
    const err = new Error('صيغة التاريخ يجب أن تكون YYYY-MM-DD'); err.status = 400; throw err;
  }
  return v;
}

// -------------------------------------------------------------------------
// 📊 صفحة أداء المعلمين — تجمع مؤشرات كل معلم من مصادر متعددة
// (الإشراف الآن، والغياب والانتظار لاحقاً بعد تحويلهما لقاعدة البيانات)
// -------------------------------------------------------------------------
app.get('/performance', (req, res) => {
  res.sendFile(path.join(__dirname, 'performance.html'));
});

// GET /api/performance?from=YYYY-MM-DD&to=YYYY-MM-DD
app.get('/api/performance', async (req, res) => {
  try {
    const username = requireSessionUser(req);
    const from = req.query.from ? parseMonitoringDateLoose(req.query.from) : null;
    const to = req.query.to ? parseMonitoringDateLoose(req.query.to) : null;

    // كل معلمي المدرسة يظهرون حتى لو ليس لهم أي تسجيل بعد
    let names = [];
    try {
      await ensureWaitingReady(username);
      names = (await dbTeachers(username)).map(t => String(t.name).trim()).filter(Boolean);
    } catch (e) {
      console.warn('⚠️ تعذر قراءة قائمة المعلمين:', e.message);
    }

    const mon = await pool.query(
      `SELECT teacher_name,
              COUNT(*) FILTER (WHERE status = 'done')::int   AS done,
              COUNT(*) FILTER (WHERE status = 'missed')::int AS missed,
              to_char(MAX(att_date) FILTER (WHERE status = 'done'), 'YYYY-MM-DD') AS last_done
       FROM monitoring_attendance
       WHERE username = $1
         AND ($2::date IS NULL OR att_date >= $2::date)
         AND ($3::date IS NULL OR att_date <= $3::date)
       GROUP BY teacher_name`,
      [username, from, to]
    );

    const byName = new Map();
    const blank = (n) => ({
      teacher_name: n,
      monitoring: { done: 0, missed: 0, last_done: null },
      absence: null,  // 🔜 بعد تحويل الغياب لقاعدة البيانات
      waiting: null   // 🔜 بعد تحويل الانتظار لقاعدة البيانات
    });
    names.forEach(n => byName.set(n, blank(n)));
    mon.rows.forEach(r => {
      if (!byName.has(r.teacher_name)) byName.set(r.teacher_name, blank(r.teacher_name));
      byName.get(r.teacher_name).monitoring = {
        done: r.done,
        missed: r.missed,
        last_done: r.last_done || null
      };
    });

    res.json({ success: true, from, to, data: [...byName.values()] });
  } catch (e) {
    sendMonitoringError(res, e, 'حدث خطأ أثناء جلب أداء المعلمين:');
  }
});

// GET /api/performance/teacher?name=...&from=&to=  — تفاصيل معلم واحد (للنافذة المنبثقة)
app.get('/api/performance/teacher', async (req, res) => {
  try {
    const username = requireSessionUser(req);
    const name = String(req.query.name || '').trim();
    if (!name) return res.status(400).json({ success: false, message: 'اسم المعلم مطلوب' });
    const from = req.query.from ? parseMonitoringDateLoose(req.query.from) : null;
    const to = req.query.to ? parseMonitoringDateLoose(req.query.to) : null;

    const result = await pool.query(
      `SELECT to_char(att_date, 'YYYY-MM-DD') AS date, day_key, task, status
       FROM monitoring_attendance
       WHERE username = $1 AND teacher_name = $2
         AND ($3::date IS NULL OR att_date >= $3::date)
         AND ($4::date IS NULL OR att_date <= $4::date)
       ORDER BY att_date DESC, task`,
      [username, name, from, to]
    );
    res.json({ success: true, monitoring: result.rows });
  } catch (e) {
    sendMonitoringError(res, e, 'حدث خطأ أثناء جلب تفاصيل المعلم:');
  }
});

// 💬 إرسال تذكير واتساب لمشرفي اليوم
// body: { date, tasks?: [أسماء المهام المعروضة], only?: [أسماء معلمين] , template?: نص }
// المتغيرات في النص: {الاسم} {المهمة} {اليوم} {التاريخ}
const DEFAULT_MONITORING_TEMPLATE =
  'المكرم الأستاذ/ {الاسم} حفظه الله\n' +
  'السلام عليكم ورحمة الله وبركاته\n' +
  'نذكّركم بأن لديكم إشرافاً اليوم {اليوم} الموافق {التاريخ} على: {المهمة}\n' +
  'شاكرين لكم تعاونكم ،،،';

app.post('/api/monitoring/notify', async (req, res) => {
  try {
    const username = requireSessionUser(req);
    const { date, dayKey } = parseMonitoringDate(req.body?.date);
    const taskFilter = Array.isArray(req.body?.tasks) ? req.body.tasks.map(t => String(t).trim()) : null;
    const onlyFilter = Array.isArray(req.body?.only) ? req.body.only.map(t => String(t).trim()) : null;
    let template = String(req.body?.template || '').trim() || DEFAULT_MONITORING_TEMPLATE;
    if (template.length > 1500) template = template.slice(0, 1500);

    // المكلّفون من القاعدة (وليس من المتصفح) — ونجمع مهام المعلم الواحد في رسالة واحدة
    const schedule = await loadMonitoringSchedule(req, username);
    const byTeacher = new Map();
    Object.entries(schedule[dayKey] || {}).forEach(([task, teachers]) => {
      if (taskFilter && !taskFilter.includes(task)) return;
      teachers.forEach(t => {
        if (onlyFilter && !onlyFilter.includes(t)) return;
        if (!byTeacher.has(t)) byTeacher.set(t, []);
        byTeacher.get(t).push(task);
      });
    });

    if (byTeacher.size === 0) {
      return res.status(400).json({ success: false, message: 'لا يوجد مشرفون مكلّفون في هذا اليوم حسب الاختيار' });
    }

    const phones = await getTeacherPhones(req);
    const s = await waEnsure(username); // يرمي رسالة واضحة إذا لم يُربط واتساب المدرسة

    const dayName = MONITORING_DAY_NAMES[dayKey];
    const hijri = formatHijriDate(date);
    const results = [];

    for (const [teacher, tasks] of byTeacher) {
      const num = waNormalizePhone(phones.get(teacher));
      if (!num) { results.push({ teacher, sent: false, error: 'لا يوجد رقم جوال صحيح للمعلم في ملف المعلمين' }); continue; }
      try {
        const found = await s.sock.onWhatsApp(`${num}@s.whatsapp.net`);
        const target = found && found[0];
        if (!target || !target.exists) { results.push({ teacher, sent: false, error: 'الرقم غير مسجل في واتساب' }); continue; }
        const text = template
          .split('{الاسم}').join(teacher)
          .split('{المهمة}').join(tasks.join('، '))
          .split('{اليوم}').join(dayName)
          .split('{التاريخ}').join(hijri);
        await s.sock.sendMessage(target.jid, { text });
        s.lastUsed = Date.now();
        results.push({ teacher, sent: true });
      } catch (err) {
        results.push({ teacher, sent: false, error: err.message });
      }
      // مهلة بسيطة بين الرسائل لتقليل خطر حظر الرقم
      await new Promise(r => setTimeout(r, 1200 + Math.floor(Math.random() * 800)));
    }

    const sentCount = results.filter(r => r.sent).length;
    res.json({ success: true, message: `تم إرسال ${sentCount} من ${results.length} رسالة`, results });
  } catch (e) {
    sendMonitoringError(res, e, 'تعذر إرسال الرسائل:');
  }
});

// =========================================================================
// 📦 3. إعداد رفع الملفات وإنشاء الحسابات
// =========================================================================
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    const uploadDir = path.join(__dirname, 'uploads');
    if (!fs.existsSync(uploadDir)) fs.mkdirSync(uploadDir, { recursive: true });
    cb(null, uploadDir);
  },
  filename: (req, file, cb) => {
    const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1E9);
    const ext = path.extname(file.originalname);
    cb(null, `school_${req.body.username || req.session?.username || 'update'}_${uniqueSuffix}${ext}`);
  }
});

const upload = multer({ 
  storage,
  fileFilter: (req, file, cb) => {
    if (file.mimetype.includes('excel') || file.mimetype.includes('spreadsheetml') || file.originalname.endsWith('.xlsx') || file.originalname.endsWith('.xls')) {
      cb(null, true);
    } else {
      cb(new Error('يرجى رفع ملف إكسل بصيغة .xlsx أو .xls فقط!'));
    }
  }
});

app.get('/register', (req, res) => {
  res.sendFile(path.join(__dirname, 'register.html'));
});

app.post('/register', upload.any(), async (req, res) => {
  const cleanFiles = () => {
    if (req.files && req.files.length > 0) {
      req.files.forEach(f => fs.existsSync(f.path) && fs.unlinkSync(f.path));
    } else if (req.file && fs.existsSync(req.file.path)) {
      fs.unlinkSync(req.file.path);
    }
  };

  try {
    const { fullName, region, schoolName, username, password } = req.body;
    const cleanUsername = String(username).trim();

    if (!fullName || !region || !schoolName || !cleanUsername || !password) {
      cleanFiles();
      return res.status(400).json({ success: false, message: 'يرجى تعبئة جميع الحقول المطلوبة!' });
    }

    const checkResult = await pool.query('SELECT username FROM users WHERE username = $1', [cleanUsername]);
    if (checkResult.rows.length > 0) {
      cleanFiles();
      return res.status(400).json({ success: false, message: 'اسم المستخدم مسجل بالفعل، اختر اسماً آخر!' });
    }

    const uploadsDir = path.join(__dirname, 'uploads');
    if (!fs.existsSync(uploadsDir)) fs.mkdirSync(uploadsDir, { recursive: true });

    const userExcelFileName = `data_${cleanUsername}.xlsx`;
    const userExcelPath = path.join(uploadsDir, userExcelFileName);
    let fileBuffer;

    const uploadedFile = (req.files && req.files.length > 0) ? req.files[0] : (req.file || null);

    if (uploadedFile) {
      fileBuffer = fs.readFileSync(uploadedFile.path);
      fs.writeFileSync(userExcelPath, fileBuffer);
      if (fs.existsSync(uploadedFile.path)) fs.unlinkSync(uploadedFile.path);
    } else {
      const cleanWb = createCleanSchoolWorkbook(schoolName, region, fullName);
      xlsx.writeFile(cleanWb, userExcelPath);
      fileBuffer = fs.readFileSync(userExcelPath);
    }

    const hashedPassword = bcrypt.hashSync(password, 10);
    
    await pool.query(
      `INSERT INTO users (full_name, region, school_name, username, password, school_excel_file, excel_data) 
       VALUES ($1, $2, $3, $4, $5, $6, $7)`,
      [fullName, region, schoolName, cleanUsername, hashedPassword, userExcelFileName, fileBuffer]
    );

    return res.json({
      success: true,
      message: 'تم إنشاء الحساب بنجاح وتأمين حفظ البيانات بداخل قاعدة البيانات! يتم توجيهك لصفحة الدخول...'
    });

  } catch (error) {
    cleanFiles();
    console.error("خطأ التسجيل:", error);
    res.status(500).json({ success: false, message: 'حدث خطأ في السيرفر أثناء معالجة الطلب: ' + error.message });
  }
});

app.post('/update-school-excel', upload.any(), async (req, res) => {
  const cleanFiles = () => {
    if (req.files && req.files.length > 0) {
      req.files.forEach(f => fs.existsSync(f.path) && fs.unlinkSync(f.path));
    } else if (req.file && fs.existsSync(req.file.path)) {
      fs.unlinkSync(req.file.path);
    }
  };

  try {
    const uploadedFile = (req.files && req.files.length > 0) ? req.files[0] : (req.file || null);

    if (!uploadedFile) {
      return res.status(400).json({ success: false, message: "يرجى اختيار ملف الأكسل أولاً!" });
    }

    const rawUsername = req.session?.username || req.headers['x-username'] || req.body?.username || req.query?.username;
    const uploadedPath = uploadedFile.path;

    if (rawUsername) {
      const username = String(rawUsername).trim();
      const targetFileName = uploadedFile.filename;
      const fileBuffer = fs.readFileSync(uploadedPath);

      const uploadsDir = path.join(__dirname, 'uploads');
      if (!fs.existsSync(uploadsDir)) fs.mkdirSync(uploadsDir, { recursive: true });
      const userFilePath = path.join(uploadsDir, `data_${username}.xlsx`);
      fs.writeFileSync(userFilePath, fileBuffer);

      await pool.query(
        'UPDATE users SET school_excel_file = $1, excel_data = $2 WHERE username = $3', 
        [targetFileName, fileBuffer, username]
      );

      if (fs.existsSync(uploadedPath)) fs.unlinkSync(uploadedPath);

      // 🔄 تحديث المعلمين والفصول ومعلومات المدرسة في القاعدة من الملف الجديد
      // (جدول الانتظار وسجلات الغياب الموجودة في القاعدة لا تُمسّ)
      await refreshFromUploadedExcel(username, fileBuffer);
    } else {
      if (fs.existsSync(uploadedPath)) fs.unlinkSync(uploadedPath);
      return res.status(401).json({ success: false, message: "يرجى تسجيل الدخول أولاً لتحديث بيانات مدرستك." });
    }

    return res.json({ success: true, message: "تم تحديث واستبدال بيانات المدرسة بنجاح" });
  } catch (error) {
    cleanFiles();
    console.error("خطأ أثناء تحديث ملف الإكسل:", error);
    return res.status(500).json({ success: false, message: "فشل في معالجة وتحديث الملف: " + error.message });
  }
});

// =========================================================================
// ⚖️ 4. مسارات نظام المساءلات (Accountability API Routes)
// =========================================================================

// 💾 1. إضافة مساءلة جديدة
app.post('/api/accountability', async (req, res) => {
  try {
    const { teacher_name, violation_type, hijri_date, delay_minutes, details, status } = req.body;
    const username = req.session?.username || req.body?.username || null;

    if (!teacher_name || !violation_type || !hijri_date) {
      return res.status(400).json({ success: false, message: 'يرجى تعبئة الحقول الأساسية: اسم المعلم، نوع المخالفة، والتاريخ' });
    }

    const queryText = `
      INSERT INTO accountability (username, teacher_name, violation_type, hijri_date, delay_minutes, details, status)
      VALUES ($1, $2, $3, $4, $5, $6, $7)
      RETURNING *
    `;
    const values = [
      username,
      teacher_name,
      violation_type,
      hijri_date,
      parseInt(delay_minutes) || 0,
      details || '',
      status || 'بانتظار الرد'
    ];

    const result = await pool.query(queryText, values);
    return res.json({
      success: true,
      message: 'تم حفظ المساءلة بنجاح',
      data: result.rows[0]
    });
  } catch (error) {
    console.error('خطأ في حفظ المساءلة:', error);
    return res.status(500).json({ success: false, message: 'حدث خطأ أثناء حفظ المساءلة: ' + error.message });
  }
});

// 📥 2. جلب المساءلات الخاصة بالمدرسة
app.get('/api/accountability', async (req, res) => {
  try {
    const username = req.session?.username || req.query?.username || null;
    let queryText = 'SELECT * FROM accountability ORDER BY id DESC';
    let values = [];

    if (username) {
      queryText = 'SELECT * FROM accountability WHERE username = $1 OR username IS NULL ORDER BY id DESC';
      values = [username];
    }

    const result = await pool.query(queryText, values);
    return res.json({ success: true, data: result.rows });
  } catch (error) {
    console.error('خطأ في جلب المساءلات:', error);
    return res.status(500).json({ success: false, message: 'حدث خطأ أثناء جلب المساءلات: ' + error.message });
  }
});

// 🔄 3. تحديث حالة أو تفاصيل مساءلة
app.put('/api/accountability/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const { status, details } = req.body;

    const queryText = `
      UPDATE accountability 
      SET status = COALESCE($1, status),
          details = COALESCE($2, details)
      WHERE id = $3
      RETURNING *
    `;
    const result = await pool.query(queryText, [status, details, id]);

    if (result.rows.length === 0) {
      return res.status(404).json({ success: false, message: 'لم يتم العثور على المساءلة المطلوب تحديثها' });
    }

    return res.json({ success: true, message: 'تم تحديث المساءلة بنجاح', data: result.rows[0] });
  } catch (error) {
    console.error('خطأ في تحديث المساءلة:', error);
    return res.status(500).json({ success: false, message: 'حدث خطأ أثناء تحديث المساءلة: ' + error.message });
  }
});

// ❌ 4. حذف مساءلة
app.delete('/api/accountability/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const result = await pool.query('DELETE FROM accountability WHERE id = $1 RETURNING *', [id]);

    if (result.rows.length === 0) {
      return res.status(404).json({ success: false, message: 'لم يتم العثور على المساءلة' });
    }

    return res.json({ success: true, message: 'تم حذف المساءلة بنجاح' });
  } catch (error) {
    console.error('خطأ في حذف المساءلة:', error);
    return res.status(500).json({ success: false, message: 'حدث خطأ أثناء حذف المساءلة: ' + error.message });
  }
});

// =========================================================================
// 🛠️ 5. مسارات لوحة تحكم الأدمن (Admin Routes)
// =========================================================================

function requireAdmin(req, res, next) {
  if (req.session && req.session.isAdmin) {
    return next();
  }
  return res.status(401).json({ success: false, message: 'غير مصرح بالوصول! يرجى تسجيل الدخول كمسؤول بالنظام.' });
}

app.get('/admin', (req, res) => {
  res.sendFile(path.join(__dirname, 'admin.html'));
});

app.post('/api/admin/login', (req, res) => {
  const { username, password } = req.body;
  const ADMIN_USER = process.env.ADMIN_USER || 'admin';
  const ADMIN_PASS = process.env.ADMIN_PASS || 'admin123';

  if (username === ADMIN_USER && password === ADMIN_PASS) {
    req.session.isAdmin = true;
    return res.json({ success: true, message: 'تم تسجيل دخول الأدمن بنجاح' });
  } else {
    return res.status(401).json({ success: false, message: 'اسم المستخدم أو كلمة المرور غير صحيحة!' });
  }
});

app.get('/api/admin/check-auth', (req, res) => {
  if (req.session && req.session.isAdmin) {
    return res.json({ authenticated: true });
  }
  return res.json({ authenticated: false });
});

app.post('/api/admin/logout', (req, res) => {
  if (req.session) {
    req.session.isAdmin = false;
  }
  return res.json({ success: true, message: 'تم تسجيل خروج الأدمن بنجاح' });
});

app.get('/api/admin/schools', requireAdmin, async (req, res) => {
  try {
    const result = await pool.query('SELECT id, username, full_name, region, school_name, school_excel_file, (excel_data IS NOT NULL) AS has_excel FROM users ORDER BY id DESC');
    res.json({ success: true, schools: result.rows });
  } catch (e) {
    res.status(500).json({ success: false, error: e.message });
  }
});

// ✏️ مسار تعديل بيانات مدرسة (مع إمكانية رفع/استبدال ملف الإكسل الخاص بها)
app.put('/api/admin/schools/:id', requireAdmin, upload.single('excel_file'), async (req, res) => {
  try {
    const { id } = req.params;
    const { full_name, school_name, region, username, password } = req.body;
    const oldRow = await pool.query('SELECT username FROM users WHERE id = $1', [id]);
    const oldUsername = oldRow.rows[0] ? oldRow.rows[0].username : null;

    // 🔒 كلمة المرور: نشفّرها فقط لو الأدمن كتب واحدة جديدة فعلاً، وإلا نبقي القديمة كما هي
    let uploadedBuffer = null;
    let hashedPassword = null;
    if (password && password.trim() !== '') {
      hashedPassword = bcrypt.hashSync(password.trim(), 10);
    }

    if (req.file) {
      // 📁 تم رفع ملف إكسل جديد: نقرأه من القرص ونخزنه بقاعدة البيانات مباشرة
      const fileBuffer = fs.readFileSync(req.file.path);
      uploadedBuffer = fileBuffer;

      if (hashedPassword) {
        await pool.query(
          `UPDATE users
           SET full_name = $1, school_name = $2, region = $3, username = $4, password = $5,
               school_excel_file = $6, excel_data = $7
           WHERE id = $8`,
          [full_name, school_name, region, username, hashedPassword, req.file.originalname, fileBuffer, id]
        );
      } else {
        await pool.query(
          `UPDATE users
           SET full_name = $1, school_name = $2, region = $3, username = $4,
               school_excel_file = $5, excel_data = $6
           WHERE id = $7`,
          [full_name, school_name, region, username, req.file.originalname, fileBuffer, id]
        );
      }

      // تنظيف الملف المؤقت من القرص بعد تخزينه بقاعدة البيانات
      fs.unlink(req.file.path, () => {});
    } else {
      // ✏️ تعديل البيانات فقط بدون رفع ملف جديد
      if (hashedPassword) {
        await pool.query(
          `UPDATE users
           SET full_name = $1, school_name = $2, region = $3, username = $4, password = $5
           WHERE id = $6`,
          [full_name, school_name, region, username, hashedPassword, id]
        );
      } else {
        await pool.query(
          `UPDATE users
           SET full_name = $1, school_name = $2, region = $3, username = $4
           WHERE id = $5`,
          [full_name, school_name, region, username, id]
        );
      }
    }

    // اسم المستخدم الجديد تتبعه كل جداول المدرسة
    await renameSchoolUsername(oldUsername, username);
    // ملف إكسل جديد من الأدمن: يحدّث المعلمين والفصول والمعلومات في القاعدة
    if (uploadedBuffer) await refreshFromUploadedExcel(username, uploadedBuffer);

    res.json({ success: true, message: 'تم تحديث بيانات المدرسة بنجاح' });
  } catch (e) {
    console.error('خطأ أثناء تعديل بيانات المدرسة:', e);
    res.status(500).json({ success: false, error: e.message });
  }
});

// 🗑️ مسار حذف مدرسة بالكامل
app.delete('/api/admin/schools/:id', requireAdmin, async (req, res) => {
  try {
    const { id } = req.params;
    const del = await pool.query('SELECT username FROM users WHERE id = $1', [id]);
    await pool.query('DELETE FROM users WHERE id = $1', [id]);
    if (del.rows[0]) await deleteSchoolData(del.rows[0].username);
    res.json({ success: true, message: 'تم حذف المدرسة بنجاح' });
  } catch (e) {
    console.error('خطأ أثناء حذف المدرسة:', e);
    res.status(500).json({ success: false, error: e.message });
  }
});

// 📥 مسار تنزيل ملف الإكسل المباشر من PostgreSQL للأدمن
app.get('/api/admin/download-excel/:id', requireAdmin, async (req, res) => {
  try {
    const { id } = req.params;
    const result = await pool.query('SELECT username, school_name FROM users WHERE id = $1', [id]);

    if (result.rows.length === 0) {
      return res.status(404).json({ success: false, message: 'عذراً، المدرسة غير موجودة.' });
    }

    const school = result.rows[0];
    // الملف يُبنى الآن من قاعدة البيانات (المعلمون، الفصول، الانتظار، الغياب) وبقية الأوراق تبقى كما هي
    const wb = await buildWorkbookFromDb(school.username);
    const buffer = xlsx.write(wb, { type: 'buffer', bookType: 'xlsx' });
    const fileName = `school_${school.username || id}_data.xlsx`;

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${encodeURIComponent(fileName)}"`);
    res.setHeader('Cache-Control', 'no-store');

    return res.send(buffer);
  } catch (error) {
    console.error('خطأ في تنزيل ملف الإكسل للأدمن:', error);
    res.status(500).send('حدث خطأ أثناء جلب الملف من قاعدة البيانات: ' + error.message);
  }
});

// =========================================================================
// 💬 واتساب: رقم خاص لكل مدرسة (ربط دائم بالباركود) وإرسال الرسائل مباشرة
// المكتبة: @whiskeysockets/baileys (اتصال WebSocket خفيف بدون متصفح)
// تُحفظ جلسة كل مدرسة بقاعدة PostgreSQL فلا تضيع عند إعادة تشغيل السيرفر
// =========================================================================
const QRCode = require('qrcode');
const pino = require('pino');

let _waLib = null;
async function waLib() {
  if (_waLib) return _waLib;
  const mod = await import('@whiskeysockets/baileys');
  const pick = (k) => mod[k] ?? (mod.default ? mod.default[k] : undefined);
  _waLib = {
    makeWASocket: pick('makeWASocket') || (typeof mod.default === 'function' ? mod.default : (mod.default && mod.default.default)),
    DisconnectReason: pick('DisconnectReason'),
    initAuthCreds: pick('initAuthCreds'),
    BufferJSON: pick('BufferJSON'),
    proto: pick('proto'),
    fetchLatestBaileysVersion: pick('fetchLatestBaileysVersion'),
    makeCacheableSignalKeyStore: pick('makeCacheableSignalKeyStore')
  };
  return _waLib;
}

// 🔐 تخزين مفاتيح الجلسة في PostgreSQL بدل الملفات
async function waAuthState(username) {
  const { initAuthCreds, BufferJSON, proto } = await waLib();
  const read = async (k) => {
    const r = await pool.query('SELECT value FROM wa_auth WHERE username = $1 AND key = $2', [username, k]);
    return r.rows[0] ? JSON.parse(r.rows[0].value, BufferJSON.reviver) : null;
  };
  const write = (k, v) => pool.query(
    `INSERT INTO wa_auth (username, key, value) VALUES ($1, $2, $3)
     ON CONFLICT (username, key) DO UPDATE SET value = EXCLUDED.value`,
    [username, k, JSON.stringify(v, BufferJSON.replacer)]
  );
  const del = (k) => pool.query('DELETE FROM wa_auth WHERE username = $1 AND key = $2', [username, k]);

  const creds = (await read('creds')) || initAuthCreds();
  return {
    state: {
      creds,
      keys: {
        get: async (type, ids) => {
          const out = {};
          await Promise.all(ids.map(async (id) => {
            let v = await read(`${type}-${id}`);
            if (type === 'app-state-sync-key' && v) v = proto.Message.AppStateSyncKeyData.fromObject(v);
            out[id] = v;
          }));
          return out;
        },
        set: async (data) => {
          const jobs = [];
          for (const type in data) {
            for (const id in data[type]) {
              const v = data[type][id];
              jobs.push(v ? write(`${type}-${id}`, v) : del(`${type}-${id}`));
            }
          }
          await Promise.all(jobs);
        }
      }
    },
    saveCreds: () => write('creds', creds)
  };
}

const waWipe = (username) => pool.query('DELETE FROM wa_auth WHERE username = $1', [username]);

async function waIsLinked(username) {
  const r = await pool.query("SELECT value FROM wa_auth WHERE username = $1 AND key = 'creds'", [username]);
  if (!r.rows[0]) return false;
  try { return !!JSON.parse(r.rows[0].value).me; } catch (e) { return false; }
}

// حالة كل مدرسة بالذاكرة: disconnected | connecting | qr | connected
const waSessions = new Map();
function waGet(u) {
  if (!waSessions.has(u)) {
    waSessions.set(u, { state: 'disconnected', qr: null, phone: '', sock: null, starting: false, wasOpen: false, keepAlive: true, lastUsed: 0, error: null });
  }
  return waSessions.get(u);
}

async function waStart(username) {
  const s = waGet(username);
  if (s.sock || s.starting) return s;
  s.starting = true; s.state = 'connecting'; s.qr = null; s.keepAlive = true;
  try {
    const lib = await waLib();
    const { state, saveCreds } = await waAuthState(username);
    const logger = pino({ level: 'silent' });
    let version;
    try { version = (await lib.fetchLatestBaileysVersion()).version; } catch (e) { /* نستخدم الافتراضي */ }
    const keys = lib.makeCacheableSignalKeyStore ? lib.makeCacheableSignalKeyStore(state.keys, logger) : state.keys;

    const sock = lib.makeWASocket({
      version, logger,
      auth: { creds: state.creds, keys },
      printQRInTerminal: false,
      browser: ['Wakeel', 'Chrome', '1.0.0'],
      markOnlineOnConnect: false,
      syncFullHistory: false
    });
    s.sock = sock;
    sock.ev.on('creds.update', saveCreds);

    sock.ev.on('connection.update', async (u) => {
      if (s.sock !== sock) return;
      const { connection, lastDisconnect, qr } = u;
      if (qr) {
        try { s.qr = await QRCode.toDataURL(qr, { margin: 1, width: 280 }); s.state = 'qr'; } catch (e) {}
      }
      if (connection === 'open') {
        s.state = 'connected'; s.qr = null; s.wasOpen = true; s.error = null; s.lastUsed = Date.now();
        s.phone = ((sock.user && sock.user.id) || '').split(':')[0].split('@')[0];
      }
      if (connection === 'close') {
        const code = lastDisconnect && lastDisconnect.error && lastDisconnect.error.output && lastDisconnect.error.output.statusCode;
        s.sock = null; s.qr = null;
        if (code === lib.DisconnectReason.loggedOut) {
          await waWipe(username);              // فُصل الربط من الجوال
          s.state = 'disconnected'; s.phone = ''; s.wasOpen = false;
        } else if (code === lib.DisconnectReason.restartRequired || (s.wasOpen && s.keepAlive)) {
          s.state = 'connecting';              // يحدث مباشرة بعد مسح الباركود
          setTimeout(() => waStart(username).catch(() => {}), 1500);
        } else {
          s.state = 'disconnected';
        }
      }
    });
  } catch (e) {
    console.error('❌ فشل تشغيل واتساب:', e.message);
    s.sock = null; s.state = 'disconnected'; s.error = e.message;
  } finally {
    s.starting = false;
  }
  return s;
}

async function waEnsure(username) {
  if (!(await waIsLinked(username))) {
    throw new Error('لم يتم ربط رقم واتساب للمدرسة. اربط الرقم أولاً من صفحة إعداد الواتساب.');
  }
  const s = waGet(username);
  if (s.state !== 'connected') {
    await waStart(username);
    const t0 = Date.now();
    while (s.state !== 'connected' && Date.now() - t0 < 25000) {
      await new Promise((r) => setTimeout(r, 500));
    }
    if (s.state !== 'connected') throw new Error('تعذر الاتصال بواتساب حالياً، حاول بعد قليل.');
  }
  s.lastUsed = Date.now();
  return s;
}

function waNormalizePhone(raw) {
  let d = String(raw || '').replace(/\D/g, '');
  if (d.startsWith('00')) d = d.slice(2);
  if (d.startsWith('0') && d.length === 10) d = '966' + d.slice(1);   // 05xxxxxxxx
  else if (d.length === 9 && d.startsWith('5')) d = '966' + d;        // 5xxxxxxxx
  return d.length >= 10 && d.length <= 15 ? d : '';
}

const waAuth = (req, res, next) => {
  if (req.session && req.session.username) return next();
  return res.status(401).json({ success: false, error: 'يرجى تسجيل الدخول أولاً' });
};

// حالة الربط (تُستدعى من صفحة الإعداد)
app.get('/api/whatsapp/status', waAuth, async (req, res) => {
  try {
    const u = req.session.username;
    const s = waGet(u);
    const linked = await waIsLinked(u);
    if (linked && !s.sock && !s.starting && s.state !== 'connected') waStart(u).catch(() => {});
    res.json({ success: true, linked, state: s.state, qr: s.state === 'qr' ? s.qr : null, phone: s.phone || null, error: s.error || null });
  } catch (e) {
    res.status(500).json({ success: false, error: e.message });
  }
});

// بدء الربط (يولّد الباركود)
app.post('/api/whatsapp/connect', waAuth, async (req, res) => {
  try {
    const s = waGet(req.session.username);
    s.error = null;
    if (s.state !== 'connected') await waStart(req.session.username);
    res.json({ success: true });
  } catch (e) {
    res.status(500).json({ success: false, error: e.message });
  }
});

// فصل الرقم نهائياً
app.post('/api/whatsapp/logout', waAuth, async (req, res) => {
  const u = req.session.username;
  const s = waGet(u);
  try {
    s.keepAlive = false;
    const sock = s.sock;
    if (sock) { try { await sock.logout(); } catch (e) {} try { sock.end(undefined); } catch (e) {} }
    s.sock = null;
    await waWipe(u);
    s.state = 'disconnected'; s.phone = ''; s.qr = null; s.wasOpen = false;
    res.json({ success: true });
  } catch (e) {
    res.status(500).json({ success: false, error: e.message });
  }
});

// إرسال رسالة لمعلم عبر رقم المدرسة (يُستخدم من report.html)
app.post('/send-whatsapp', waAuth, async (req, res) => {
  try {
    const { phone, message } = req.body || {};
    const num = waNormalizePhone(phone);
    if (!num) return res.status(400).json({ success: false, error: 'رقم الجوال غير صحيح' });
    if (!message || !String(message).trim()) return res.status(400).json({ success: false, error: 'نص الرسالة فارغ' });

    const s = await waEnsure(req.session.username);
    const found = await s.sock.onWhatsApp(`${num}@s.whatsapp.net`);
    const target = found && found[0];
    if (!target || !target.exists) {
      return res.status(404).json({ success: false, error: 'هذا الرقم غير مسجل في واتساب' });
    }
    await s.sock.sendMessage(target.jid, { text: String(message) });
    s.lastUsed = Date.now();
    res.json({ success: true });
  } catch (e) {
    console.error('❌ خطأ إرسال واتساب:', e.message);
    res.status(500).json({ success: false, error: e.message });
  }
});

// توفير الذاكرة: نغلق الاتصال الخامل بعد 30 دقيقة (الربط يبقى محفوظاً ويعود تلقائياً عند الحاجة)
setInterval(() => {
  const now = Date.now();
  for (const [, s] of waSessions) {
    if (s.sock && s.state === 'connected' && now - s.lastUsed > 30 * 60 * 1000) {
      s.keepAlive = false;
      try { s.sock.end(undefined); } catch (e) {}
    }
  }
}, 5 * 60 * 1000).unref();

// =========================================================================
// 🚀 تشغيل السيرفر
// =========================================================================
app.listen(PORT, () => {
  console.log(`🚀 السيرفر يعمل بنجاح على المنفذ: ${PORT}`);
});