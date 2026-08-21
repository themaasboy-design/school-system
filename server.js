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
    
    client.release();
    console.log("✅ تم الاتصال بقاعدة بيانات PostgreSQL وبناء الجدول وتأمين التخزين الدائم بنجاح!");
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
    }
  } catch (e) {
    console.error(`❌ فشل مزامنة ملف الإكسل لقاعدة البيانات: ${e.message}`);
  }
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
      // 🌟 تدمير الملف القديم إن وجد على القرص لمنع قراءة بيانات عالقة
      if (fs.existsSync(userFilePath)) {
        fs.unlinkSync(userFilePath);
      }
      // كتابة الملف الصحيح القادم من قاعدة البيانات الدائمة
      fs.writeFileSync(userFilePath, dbResult.rows[0].excel_data);
      console.log(`⚡ تم تحديث الملف للمستخدم (${cleanUsername}) من PostgreSQL`);
      return userFilePath;
    } else {
      // إذا لم يكن هناك ملف بالـ DB، احذف أي ملف قديم على القرص أولاً
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

// 🧹 دالة تطبيع النص العربي (توحيد الهمزات والمسافات) لمنع مشاكل عدم تطابق أسماء الأيام
function normalizeArabicText(str) {
  if (!str) return '';
  return str.toString().trim().replace(/[أإآا]/g, 'ا').replace(/ة/g, 'ه').replace(/\s+/g, '');
}

// 📅 قائمة الأيام الرسمية المعتمدة لتخزين جدول الانتظار (الشكل القياسي المستخدم داخل ملف الإكسل)
const CANONICAL_WAITING_DAYS = ['الأحد', 'الإثنين', 'الثلاثاء', 'الأربعاء', 'الخميس'];

// 🔁 دالة تحويل أي تهجئة ليوم (مثل "الاثنين" أو "الإثنين") إلى الاسم القياسي الموحّد
function canonicalizeDayName(day) {
  const normalizedInput = normalizeArabicText(day);
  const match = CANONICAL_WAITING_DAYS.find(d => normalizeArabicText(d) === normalizedInput);
  return match || day;
}

// =========================================================================
// 📊 مسارات جلب وحفظ البيانات الأساسية
// =========================================================================

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

// 💾 مسار حفظ جدول الانتظار ليوم محدد (يُستخدم في صفحة waiting_setting)
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

    // إذا اليوم غير موجود بالشيت، أنشئ له بلوك جديد بالاسم القياسي الموحّد (عنوان اليوم + 4 صفوف فارغة)
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
    if (username) await syncExcelToDb(username, userExcel);

    res.json({ success: true });
  } catch (e) {
    res.status(500).json({ success: false, error: "خطأ أثناء الحفظ: " + e.message });
  }
});

// 📥 مسار جلب جدول الانتظار الكامل لكل الأيام (يُستخدم في صفحة waiting_setting)
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

// 👩‍🏫 مسار جلب أسماء المعلمين (يُستخدم في صفحة rotation)
app.get('/api/teachers', async (req, res) => {
  try {
    const userExcel = await getUserExcelPath(req);
    const workbook = xlsx.readFile(userExcel);
    const sheet = getSheet(workbook, TEACHERS_SHEET);
    const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });

    // استهداف العمود الثاني (index 1) تحديداً وهو اسم المعلم، وتجاهل عمود السجل المدني (index 0)
    const teachersNames = data.slice(1)
      .map(row => row[1])
      .filter(cell => cell !== undefined && cell !== null && cell.toString().trim() !== "");

    res.json(teachersNames);
  } catch (e) {
    res.status(401).json({ error: e.message });
  }
});

// 📥 مسار جلب بيانات جدول المناوبة المحفوظة (rotation)
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

// 💾 مسار حفظ بيانات أسبوع من جدول المناوبة (rotation)
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
    if (username) await syncExcelToDb(username, userExcel);

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
    if (username) await syncExcelToDb(username, userExcel);

    res.json({ success: true });
  } catch (e) {
    res.status(500).json({ error: "خطأ أثناء الحفظ: " + e.message });
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

app.post('/save-monitoring', async (req, res) => {
  const weeklySchedule = req.body;
  try {
    const username = req.session?.username || req.headers['x-username'] || req.body?.username;
    const userExcel = await getUserExcelPath(req);
    const workbook = xlsx.readFile(userExcel);
    
    const headers = ["المهمة", "الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس"];
    const tasksList = ["المقصف الغربي", "المقصف الشرقي", "الساحة الغربية", "الساحة الشرقية", "الدور الثاني"];
    const dayKeys = ["sun", "mon", "tue", "wed", "thu"];

    let aoaData = [headers];

    tasksList.forEach(task => {
      let row = [task];
      dayKeys.forEach(day => {
        row.push(weeklySchedule[day]?.[task] || "");
      });
      aoaData.push(row);
    });

    const newSheet = xlsx.utils.aoa_to_sheet(aoaData);

    if (!workbook.SheetNames.includes(MONITORING_SHEET)) {
      xlsx.utils.book_append_sheet(workbook, newSheet, MONITORING_SHEET);
    } else {
      workbook.Sheets[MONITORING_SHEET] = newSheet;
    }

    xlsx.writeFile(workbook, userExcel);
    if (username) await syncExcelToDb(username, userExcel);

    res.json({ success: true });
  } catch (e) {
    res.status(500).json({ success: false, error: "خطأ أثناء الحفظ: " + e.message });
  }
});

app.get('/get-monitoring-schedule', async (req, res) => {
  try {
    const filePath = await getUserExcelPath(req);
    const workbook = xlsx.readFile(filePath);
    const worksheet = workbook.Sheets[MONITORING_SHEET];

    if (!worksheet) return res.json({ success: true, schedule: {} });

    const rawData = xlsx.utils.sheet_to_json(worksheet);
    let schedule = { sun: {}, mon: {}, tue: {}, wed: {}, thu: {} };
    const daysMapping = { 'الأحد': 'sun', 'الإثنين': 'mon', 'الثلاثاء': 'tue', 'الأربعاء': 'wed', 'الخميس': 'thu' };

    rawData.forEach(row => {
      const taskName = row['المهمة'] || row['task'];
      if (taskName) {
        Object.keys(daysMapping).forEach(arabicDay => {
          const englishKey = daysMapping[arabicDay];
          if (row[arabicDay]) {
            schedule[englishKey][taskName] = row[arabicDay];
          }
        });
      }
    });

    res.json({ success: true, schedule });
  } catch (error) {
    res.json({ success: false, error: "تعذر قراءة البيانات من ملف الإكسل: " + error.message });
  }
});

app.get('/reportmoni', (req, res) => {
  res.sendFile(path.join(__dirname, 'reportmoni.html'));
});

app.get('/get-reportmoni-data', async (req, res) => {
  try {
    const { day, task } = req.query;
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

    const worksheet = workbook.Sheets[MONITORING_SHEET];
    if (!worksheet) return res.json({ success: true, data: [], mainInfo });

    const rawData = xlsx.utils.sheet_to_json(worksheet);
    let filteredResults = [];

    function normalizeText(str) {
      if (!str) return '';
      return str.toString().trim().replace(/[أإآا]/g, 'ا').replace(/ة/g, 'ه').replace(/\s+/g, '');
    }

    const dayTargets = {
      'sun': 'الأحد', 'mon': 'الإثنين', 'tue': 'الثلاثاء', 'wed': 'الأربعاء', 'thu': 'الخميس'
    };
    const targetDayName = dayTargets[day];

    rawData.forEach(row => {
      const rowTask = row['المهمة'] || row['task'] || '';
      const teacherInDay = targetDayName ? row[targetDayName] : '';

      if (task !== 'all' && normalizeText(rowTask) !== normalizeText(task)) return;

      if (teacherInDay && teacherInDay.toString().trim() !== "") {
        filteredResults.push({
          task: rowTask,
          teachers: teacherInDay.toString().trim()
        });
      }
    });

    res.json({ success: true, data: filteredResults, mainInfo });
  } catch (error) {
    res.json({ success: false, error: "حدث خطأ في قراءة ملف الإكسل: " + error.message });
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

// 📌 مسار تسجيل مدرسة جديدة (يدعم استقبال الملف بأي اسم حقل وحفظه بالـ PostgreSQL)
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
      // 🛡️ إنشاء ملف إكسل مخصص ونظيف تماماً للمدرسة الجديدة بدون الاعتماد على أي ملفات مسبقة
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

// 📌 مسار تحديث واستبدال ملف الإكسل للمدرسة من صفحة الإعدادات
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
      // 🔒 لا يوجد اسم مستخدم = لا نعرف لأي مدرسة يتبع هذا الملف، نرفض الطلب بدل الكتابة على ملف مشترك
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
// 🛠️ 4. مسارات لوحة تحكم الأدمن (Admin Routes)
// =========================================================================

function requireAdmin(req, res, next) {
  if (req.session && req.session.isAdmin) {
    return next();
  }
  return res.status(401).json({ success: false, message: 'غير مصرح بالوصول! يرجى تسجيل الدخول كمسؤول للنظام.' });
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
    const result = await pool.query('SELECT id, full_name, school_name, region, username, password, school_excel_file AS excel_path, (excel_data IS NOT NULL) AS has_excel FROM users ORDER BY id DESC');
    res.json({ success: true, schools: result.rows });
  } catch (error) {
    res.status(500).json({ success: false, error: error.message });
  }
});

app.put('/api/admin/schools/:id', requireAdmin, upload.single('excel_file'), async (req, res) => {
  try {
    const { id } = req.params;
    const { full_name, school_name, region, username, password } = req.body;
    const cleanUsername = String(username).trim();

    let finalPassword = password;
    if (password && !password.startsWith('$2a$') && !password.startsWith('$2b$')) {
      finalPassword = bcrypt.hashSync(password, 10);
    }

    if (req.file) {
      const excelFileName = req.file.filename;
      const fileBuffer = fs.readFileSync(req.file.path);

      if (cleanUsername) {
        const uploadsDir = path.join(__dirname, 'uploads');
        if (!fs.existsSync(uploadsDir)) fs.mkdirSync(uploadsDir, { recursive: true });
        const userFilePath = path.join(uploadsDir, `data_${cleanUsername}.xlsx`);
        fs.writeFileSync(userFilePath, fileBuffer);
      }

      await pool.query(
        `UPDATE users 
         SET full_name = $1, school_name = $2, region = $3, username = $4, password = $5, school_excel_file = $6, excel_data = $7 
         WHERE id = $8`,
        [full_name, school_name, region, cleanUsername, finalPassword, excelFileName, fileBuffer, id]
      );

      if (fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
    } else {
      await pool.query(
        `UPDATE users 
         SET full_name = $1, school_name = $2, region = $3, username = $4, password = $5 
         WHERE id = $6`,
        [full_name, school_name, region, cleanUsername, finalPassword, id]
      );
    }

    res.json({ success: true, message: 'تم تحديث بيانات المدرسة بنجاح' });
  } catch (error) {
    if (req.file && fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
    res.status(500).json({ success: false, error: error.message });
  }
});

app.delete('/api/admin/schools/:id', requireAdmin, async (req, res) => {
  try {
    const { id } = req.params;
    await pool.query('DELETE FROM users WHERE id = $1', [id]);
    res.json({ success: true, message: 'تم حذف المدرسة بنجاح' });
  } catch (error) {
    res.status(500).json({ success: false, error: error.message });
  }
});

app.get('/check-db', async (req, res) => {
  try {
    const result = await pool.query('SELECT username, school_name, OCTET_LENGTH(excel_data) AS excel_size_bytes FROM users;');
    res.json({ 
      success: true, 
      schools: result.rows 
    });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
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

    return res.send(school.excel_data);
  } catch (error) {
    console.error('خطأ في تنزيل ملف الإكسل للأدمن:', error);
    res.status(500).send('حدث خطأ أثناء جلب الملف من قاعدة البيانات: ' + error.message);
  }
});

// مسار تنزيل ملف الإكسل waiting_data.xlsx من المجلد الرئيسي
app.get('/download-template', (req, res) => {
    const filePath = path.join(__dirname,'templates', 'waiting_data.xlsx');
    res.download(filePath, 'waiting_data.xlsx', (err) => {
        if (err) {
            console.error('خطأ في تحميل الملف:', err);
            res.status(404).send('تعذر العثور على ملف النموذج');
        }
    });
});

// =========================================================================
// 🚀 تشغيل السيرفر
// =========================================================================

app.listen(PORT, '0.0.0.0', () => {
  console.log(`Server is running on port ${PORT}`);
});