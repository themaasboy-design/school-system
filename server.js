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

app.get('/get-data', async (req, res) => {
  try {
    const userExcel = await getUserExcelPath(req);
    const workbook = xlsx.readFile(userExcel);
    const sheet = getSheet(workbook, TEACHERS_SHEET);
    res.json(xlsx.utils.sheet_to_json(sheet));
  } catch (e) {
    res.status(401).json({ error: e.message });
  }
});

app.get('/get-classes', async (req, res) => {
  try {
    const userExcel = await getUserExcelPath(req);
    const workbook = xlsx.readFile(userExcel);
    const sheet = getSheet(workbook, CLASSES_SHEET);
    res.json(xlsx.utils.sheet_to_json(sheet));
  } catch (e) {
    res.status(401).json({ error: e.message });
  }
});

app.get('/get-waiting-teachers', async (req, res) => {
  const { day, period } = req.query;
  try {
    const userExcel = await getUserExcelPath(req);
    const workbook = xlsx.readFile(userExcel);
    const sheet = getSheet(workbook, WAITING_SHEET);
    const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });

    let dayIndex = data.findIndex(row => row && normalizeArabicText(row[0]) === normalizeArabicText(day));
    if (dayIndex === -1) return res.json({ teachers: [] });

    const colIndex = parseInt(period) - 1;
    let resultTeachers = [];

    for (let i = 1; i <= 4; i++) {
      let row = data[dayIndex + i];
      if (row && row[colIndex] !== undefined) {
        resultTeachers.push(row[colIndex]);
      }
    }
    res.json({ teachers: resultTeachers });
  } catch (e) {
    res.status(500).json({ error: "خطأ في قراءة الملف: " + e.message });
  }
});

app.post('/save-waiting-schedule-day', async (req, res) => {
  const { day, schedule } = req.body;
  try {
    if (!day || !Array.isArray(schedule)) {
      return res.status(400).json({ success: false, error: "بيانات غير صالحة: يجب إرسال day و schedule" });
    }

    const username = req.session?.username || req.headers['x-username'] || req.body?.username;
    const userExcel = await getUserExcelPath(req);
    const workbook = xlsx.readFile(userExcel);

    const canonicalDay = canonicalizeDayName(day);

    let data = [];
    if (workbook.SheetNames.includes(WAITING_SHEET)) {
      data = xlsx.utils.sheet_to_json(workbook.Sheets[WAITING_SHEET], { header: 1 });
    }

    let dayIndex = data.findIndex(row => row && normalizeArabicText(row[0]) === normalizeArabicText(canonicalDay));

    if (dayIndex === -1) {
      dayIndex = data.length;
      data.push([canonicalDay]);
      for (let r = 0; r < 4; r++) data.push([]);
    }

    for (let r = 0; r < 4; r++) {
      const rowIndex = dayIndex + 1 + r;
      const rowValues = schedule[r] || [];
      if (!data[rowIndex]) data[rowIndex] = [];
      for (let c = 0; c < 7; c++) {
        data[rowIndex][c] = rowValues[c] || "";
      }
    }

    const newSheet = xlsx.utils.aoa_to_sheet(data);
    if (workbook.SheetNames.includes(WAITING_SHEET)) {
      workbook.Sheets[WAITING_SHEET] = newSheet;
    } else {
      xlsx.utils.book_append_sheet(workbook, newSheet, WAITING_SHEET);
    }

    xlsx.writeFile(workbook, userExcel);
    await syncOrThrow(username, userExcel);

    res.json({ success: true });
  } catch (e) {
    res.status(500).json({ success: false, error: "خطأ أثناء الحفظ: " + e.message });
  }
});

app.get('/get-waiting-schedule-full', async (req, res) => {
  try {
    const userExcel = await getUserExcelPath(req);
    const workbook = xlsx.readFile(userExcel);

    if (!workbook.SheetNames.includes(WAITING_SHEET)) {
      return res.json({ success: true, schedule: {} });
    }

    const data = xlsx.utils.sheet_to_json(workbook.Sheets[WAITING_SHEET], { header: 1 });
    let schedule = {};

    CANONICAL_WAITING_DAYS.forEach(day => {
      const dayIndex = data.findIndex(row => row && normalizeArabicText(row[0]) === normalizeArabicText(day));
      if (dayIndex === -1) {
        schedule[day] = [[], [], [], []];
        return;
      }
      let matrix = [];
      for (let r = 1; r <= 4; r++) {
        matrix.push(data[dayIndex + r] || []);
      }
      schedule[day] = matrix;
    });

    res.json({ success: true, schedule });
  } catch (e) {
    res.status(500).json({ success: false, error: "خطأ أثناء قراءة الجدول: " + e.message });
  }
});

app.get('/api/teachers', async (req, res) => {
  try {
    const userExcel = await getUserExcelPath(req);
    const workbook = xlsx.readFile(userExcel);
    const sheet = getSheet(workbook, TEACHERS_SHEET);
    const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });

    const teachersNames = data.slice(1)
      .map(row => row[1])
      .filter(cell => cell !== undefined && cell !== null && cell.toString().trim() !== "");

    res.json(teachersNames);
  } catch (e) {
    res.status(401).json({ error: e.message });
  }
});

// 📋 مسار يرجّع بيانات المعلمين كاملة (كل الأعمدة بترتيبها الخام) - يُستخدم بتقرير المساءلة
app.get('/api/teachers-raw', async (req, res) => {
  try {
    const userExcel = await getUserExcelPath(req);
    const workbook = xlsx.readFile(userExcel);
    const sheet = getSheet(workbook, TEACHERS_SHEET);
    const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });

    const rows = data.slice(1)
      .filter(row => row && row[1] !== undefined && row[1] !== null && row[1].toString().trim() !== "");

    res.json(rows);
  } catch (e) {
    res.status(401).json({ error: e.message });
  }
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

app.post('/save-report', async (req, res) => {
  const reportData = req.body;
  try {
    const username = req.session?.username || req.headers['x-username'] || req.body?.username;
    const userExcel = await getUserExcelPath(req);
    const workbook = xlsx.readFile(userExcel);
    const header = ["التاريخ", "اليوم", "المعلم الغائب", "فصل 1", "بديل 1", "فصل 2", "بديل 2", "فصل 3", "بديل 3", "فصل 4", "بديل 4", "فصل 5", "بديل 5", "فصل 6", "بديل 6", "فصل 7", "بديل 7"];

    let data = [];
    if (workbook.SheetNames.includes(REPORT_SHEET)) {
      data = xlsx.utils.sheet_to_json(workbook.Sheets[REPORT_SHEET], { header: 1 });
    } else {
      data.push(header);
    }

    let row = [reportData.date, reportData.day, reportData.absentTeacher];
    for (let i = 1; i <= 7; i++) {
      row.push(reportData[`class${i}_room`], reportData[`class${i}_teacher`]);
    }
    data.push(row);

    const newSheet = xlsx.utils.aoa_to_sheet(data);
    workbook.Sheets[REPORT_SHEET] = newSheet;
    if (!workbook.SheetNames.includes(REPORT_SHEET)) {
      xlsx.utils.book_append_sheet(workbook, newSheet, REPORT_SHEET);
    }

    xlsx.writeFile(workbook, userExcel);
    await syncOrThrow(username, userExcel);

    res.json({ success: true });
  } catch (e) {
    res.status(500).json({ error: "خطأ أثناء الحفظ: " + e.message });
  }
});

// 🗑️ مسار حذف سجل غياب (يدعم المطابقة عبر التاريخ أو اسم اليوم)
app.post('/delete-absence-record', async (req, res) => {
  const teacherName = req.body.teacherName || req.body.absentTeacher || req.body.teacher;
  const date = req.body.date || req.body.hijriDate || req.body.day;

  console.log(`🔍 طلب حذف وارد -> المعلم: "${teacherName}" | التاريخ/اليوم: "${date}"`);

  try {
    const username = req.session?.username || req.headers['x-username'] || req.body?.username;
    const userExcel = await getUserExcelPath(req);
    const workbook = xlsx.readFile(userExcel);

    if (!workbook.SheetNames.includes(REPORT_SHEET)) {
      return res.status(404).json({ success: false, message: 'ورقة التقرير غير موجودة في الملف' });
    }

    let data = xlsx.utils.sheet_to_json(workbook.Sheets[REPORT_SHEET], { header: 1 });
    if (data.length <= 1) {
      return res.json({ success: true, message: 'لا توجد سجلات للحذف' });
    }

    const header = data[0];
    const rows = data.slice(1);

    // 🧹 دالة تنظيف وتوحيد الأحرف والمسافات
    const cleanStr = (s) => {
      if (s === null || s === undefined) return '';
      return s.toString()
        .trim()
        .replace(/[٠-٩]/g, d => '٠١٢٣٤٥٦٧٨٩'.indexOf(d))
        .replace(/[أإآا]/g, 'ا')
        .replace(/ة/g, 'ه')
        .replace(/ى/g, 'ي')
        .replace(/[\/\-\s\u200f\u200e]/g, '');
    };

    const targetDateOrDay = cleanStr(date);
    const targetTeacher = cleanStr(teacherName);

    let matchFound = false;

    const filteredRows = rows.filter((row) => {
      const rowDate = cleanStr(row[0]);   // العمود الأول: التاريخ الرقمي
      const rowDay = cleanStr(row[1]);    // العمود الثاني: اسم اليوم (الأحد، الإثنين...)
      const rowTeacher = cleanStr(row[2]);// العمود الثالث: اسم المعلم

      // المطابقة المرنة: يفحص التاريخ أو اليوم + اسم المعلم
      const isDateOrDayMatch = (rowDate === targetDateOrDay || rowDay === targetDateOrDay);
      const isTeacherMatch = (rowTeacher === targetTeacher);

      if (isDateOrDayMatch && isTeacherMatch && !matchFound) {
        matchFound = true; // حذف السجل المطابق
        return false;
      }
      return true;
    });

    if (!matchFound) {
      console.warn(`❌ لم يتم العثور على سجل مطابق للمعلم "${teacherName}" والتاريخ/اليوم "${date}".`);
      return res.status(404).json({
        success: false,
        message: 'لم يتم العثور على السجل المطلوب حذفه.'
      });
    }

    const newData = [header, ...filteredRows];
    workbook.Sheets[REPORT_SHEET] = xlsx.utils.aoa_to_sheet(newData);
    xlsx.writeFile(workbook, userExcel);

    if (username) {
      await syncExcelToDb(username, userExcel);
    }

    console.log(`✅ تم حذف سجل المعلم (${teacherName}) بنجاح.`);
    res.json({ success: true, message: 'تم حذف السجل بنجاح' });

  } catch (e) {
    console.error("❌ خطأ أثناء حذف سجل الغياب:", e);
    res.status(500).json({ success: false, message: "حدث خطأ أثناء الحذف: " + e.message });
  }
});

app.get('/get-report-data', async (req, res) => {
  try {
    const userExcel = await getUserExcelPath(req);
    const workbook = xlsx.readFile(userExcel);

    let reportData = workbook.SheetNames.includes(REPORT_SHEET) 
      ? xlsx.utils.sheet_to_json(workbook.Sheets[REPORT_SHEET]) 
      : [];

    let mainInfo = { B1: "", B2: "", B3: "", B4: "" };
    if (workbook.SheetNames.includes(MAIN_INFO_SHEET)) {
      const infoData = xlsx.utils.sheet_to_json(workbook.Sheets[MAIN_INFO_SHEET], { header: 1 });
      mainInfo = {
        B1: infoData[0]?.[1] || "",
        B2: infoData[1]?.[1] || "",
        B3: infoData[2]?.[1] || "",
        B4: infoData[3]?.[1] || ""
      };
    }

    res.json({ teachers: reportData, mainInfo });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// =========================================================================
// 📋 مسارات الإشراف الأسبوعي
// =========================================================================

app.get('/get-monitoring-teachers', async (req, res) => {
  try {
    const userExcel = await getUserExcelPath(req);
    const workbook = xlsx.readFile(userExcel);
    const sheet = getSheet(workbook, TEACHERS_SHEET);
    const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });

    const teachersList = data.slice(1)
      .map(row => row[1])
      .filter(cell => cell !== undefined && cell !== null && cell.toString().trim() !== "");

    res.json({ success: true, teachers: teachersList });
  } catch (e) {
    res.status(500).json({ success: false, error: e.message });
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

    // معلومات المدرسة ما زالت في ورقة maininfo بالإكسل
    const filePath = await getUserExcelPath(req);
    const workbook = xlsx.readFile(filePath);
    const infoSheet = workbook.Sheets[MAIN_INFO_SHEET];
    let mainInfo = { B1: '', B2: '', B3: '', B4: '' };
    if (infoSheet) {
      mainInfo.B1 = infoSheet['B1'] ? infoSheet['B1'].v : '';
      mainInfo.B2 = infoSheet['B2'] ? infoSheet['B2'].v : '';
      mainInfo.B3 = infoSheet['B3'] ? infoSheet['B3'].v : '';
      mainInfo.B4 = infoSheet['B4'] ? infoSheet['B4'].v : '';
    }

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
  const workbook = xlsx.readFile(await getUserExcelPath(req));
  const rows = xlsx.utils.sheet_to_json(getSheet(workbook, TEACHERS_SHEET), { header: 1 });
  const header = (rows[0] || []).map(h => String(h ?? ''));
  const phoneCol = header.findIndex(h => /جوال|هاتف|موبايل|تواصل|phone|mobile/i.test(h));

  const phones = new Map();
  rows.slice(1).forEach(row => {
    if (!row || !row[1]) return;
    const name = String(row[1]).trim();
    let phone = '';
    if (phoneCol >= 0 && row[phoneCol] != null) phone = String(row[phoneCol]);
    if (!phone) {
      const cell = row.find((c, i) => i !== 1 && c != null && SAUDI_MOBILE_RE.test(String(c).replace(/[\s-]/g, '')));
      if (cell != null) phone = String(cell);
    }
    if (name && phone) phones.set(name, phone);
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

function readTeacherNamesFromExcel(workbook) {
  const rows = xlsx.utils.sheet_to_json(getSheet(workbook, TEACHERS_SHEET) || {}, { header: 1 });
  return rows.slice(1).map(r => (r && r[1] != null ? String(r[1]).trim() : '')).filter(Boolean);
}

// GET /api/performance?from=YYYY-MM-DD&to=YYYY-MM-DD
app.get('/api/performance', async (req, res) => {
  try {
    const username = requireSessionUser(req);
    const from = req.query.from ? parseMonitoringDateLoose(req.query.from) : null;
    const to = req.query.to ? parseMonitoringDateLoose(req.query.to) : null;

    // كل معلمي المدرسة يظهرون حتى لو ليس لهم أي تسجيل بعد
    let names = [];
    try {
      names = readTeacherNamesFromExcel(xlsx.readFile(await getUserExcelPath(req)));
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

    // 🔒 كلمة المرور: نشفّرها فقط لو الأدمن كتب واحدة جديدة فعلاً، وإلا نبقي القديمة كما هي
    let hashedPassword = null;
    if (password && password.trim() !== '') {
      hashedPassword = bcrypt.hashSync(password.trim(), 10);
    }

    if (req.file) {
      // 📁 تم رفع ملف إكسل جديد: نقرأه من القرص ونخزنه بقاعدة البيانات مباشرة
      const fileBuffer = fs.readFileSync(req.file.path);

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
    await pool.query('DELETE FROM users WHERE id = $1', [id]);
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
    const result = await pool.query(
      'SELECT username, school_name, excel_data FROM users WHERE id = $1',
      [id]
    );

    if (result.rows.length === 0 || !result.rows[0].excel_data) {
      return res.status(404).json({ success: false, message: 'عذراً، لا يوجد ملف إكسل مخزن لهذه المدرسة.' });
    }

    const school = result.rows[0];
    const fileName = `school_${school.username || id}_data.xlsx`;

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${encodeURIComponent(fileName)}"`);
    res.setHeader('Cache-Control', 'no-store');

    return res.send(school.excel_data);
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