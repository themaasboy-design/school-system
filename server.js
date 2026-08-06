// =========================================================================
// 📦 1. استدعاء المكتبات وإعداد التطبيق
// =========================================================================
process.env.PUPPETEER_CACHE_DIR = require('path').join(__dirname, '.cache');

const express = require('express');
const path = require('path');
const fs = require('fs');
const xlsx = require('xlsx');
const bodyParser = require('body-parser');
const bcrypt = require('bcryptjs');
const multer = require('multer');
const session = require('express-session');
const { Pool } = require('pg');
const { Client, LocalAuth } = require('whatsapp-web.js');
const puppeteer = require('puppeteer');
const qrcode = require('qrcode');

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
// ⚙️ إعدادات أسماء الملف والورقات والدوال المساعدة
// =========================================================================
const EXCEL_FILE = path.join(__dirname, 'waiting_data.xlsx');

const TEACHERS_SHEET = 'teacher_name';
const CLASSES_SHEET = 'classes';
const WAITING_SHEET = ' Waiting_table';
const REPORT_SHEET = 'report';
const MAIN_INFO_SHEET = 'maininfo';
const MONITORING_SHEET = 'moni';

// 🛠️ دالة تنظيف وتوحيد النصوص العربية
function normalizeArabicText(str) {
  if (!str) return '';
  return str.toString().trim()
    .replace(/[أإآا]/g, 'ا')
    .replace(/ة/g, 'ه')
    .replace(/\s+/g, '');
}

// 📍 دالة تحديد صف البداية الثابت لكل يوم (تطابق صفوف الإكسل المطلوبة بالكامل)
function getDayStartRow(dayStr) {
  const norm = normalizeArabicText(dayStr);
  if (norm.includes('احد')) return 0;       // صف 1 بالإكسل (Index 0)
  if (norm.includes('ثنين')) return 6;      // صف 7 بالإكسل (Index 6)
  if (norm.includes('ثلاثاء')) return 12;   // صف 13 بالإكسل (Index 12)
  if (norm.includes('اربعاء')) return 18;   // صف 19 بالإكسل (Index 18)
  if (norm.includes('خميس')) return 24;     // صف 25 بالإكسل (Index 24)
  return -1;
}

// 🛠️ دالة تحديد وإعادة بناء ملف المدرسة
async function getUserExcelPath(req) {
  const rawUsername = req.session?.username || req.headers['x-username'] || req.query?.username || req.body?.username;
  
  if (!rawUsername) {
    throw new Error("UNAUTHORIZED: يرجى تسجيل الدخول أولاً للوصول لبيانات المدرسة.");
  }

  const username = String(rawUsername).trim();
  const uploadsDir = path.join(__dirname, 'uploads');
  if (!fs.existsSync(uploadsDir)) {
    fs.mkdirSync(uploadsDir, { recursive: true });
  }

  const userFilePath = path.join(uploadsDir, `data_${username}.xlsx`);

  try {
    const dbResult = await pool.query('SELECT excel_data FROM users WHERE username = $1', [username]);
    if (dbResult.rows.length > 0 && dbResult.rows[0].excel_data) {
      fs.writeFileSync(userFilePath, dbResult.rows[0].excel_data);
      console.log(`⚡ تم تحديث ومزامنة ملف الإكسل للمستخدم (${username}) مباشرة من PostgreSQL!`);
      return userFilePath;
    }
  } catch (err) {
    console.error('خطأ في استعادة الملف من DB:', err.message);
  }

  if (!fs.existsSync(userFilePath)) {
    const templatePath = path.join(__dirname, 'template.xlsx');
    if (fs.existsSync(templatePath)) {
      fs.copyFileSync(templatePath, userFilePath);
    } else if (fs.existsSync(EXCEL_FILE)) {
      fs.copyFileSync(EXCEL_FILE, userFilePath);
    } else {
      const newWb = xlsx.utils.book_new();
      xlsx.utils.book_append_sheet(newWb, xlsx.utils.aoa_to_sheet([["المعلم", "الاسم"]]), TEACHERS_SHEET);
      xlsx.utils.book_append_sheet(newWb, xlsx.utils.aoa_to_sheet([["الفصل"]]), CLASSES_SHEET);
      xlsx.utils.book_append_sheet(newWb, xlsx.utils.aoa_to_sheet([["اليوم"]]), WAITING_SHEET);
      xlsx.utils.book_append_sheet(newWb, xlsx.utils.aoa_to_sheet([]), REPORT_SHEET);
      xlsx.utils.book_append_sheet(newWb, xlsx.utils.aoa_to_sheet([]), MAIN_INFO_SHEET);
      xlsx.utils.book_append_sheet(newWb, xlsx.utils.aoa_to_sheet([]), MONITORING_SHEET);
      xlsx.writeFile(newWb, userFilePath);
    }

    await syncExcelToDb(username, userFilePath);
  }

  return userFilePath;
}

function getSheet(workbook, sheetIdentifier) {
  if (typeof sheetIdentifier === 'number') {
    return workbook.Sheets[workbook.SheetNames[sheetIdentifier]];
  }
  return workbook.Sheets[sheetIdentifier];
}

// =========================================================================
// 📊 مسارات جلب وحفظ البيانات الأساسية والداشبورد
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

// 📌 جلب معلمي الانتظار للداشبورد بناءً على الخريطة الثابتة للصفوف والأعمدة
app.get('/get-waiting-teachers', async (req, res) => {
  const { day, period } = req.query;
  try {
    const userExcel = await getUserExcelPath(req);
    const workbook = xlsx.readFile(userExcel);
    const sheet = getSheet(workbook, WAITING_SHEET);
    const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });

    const startRow = getDayStartRow(day);
    if (startRow === -1) return res.json({ teachers: [] });

    const colIndex = parseInt(period) - 1; // الحصة 1 -> Index 0 (العمود A)
    let resultTeachers = [];

    // قراءة الأسطر الأربعة المخصصة لكل يوم من موقعها المباشر
    for (let i = 0; i < 4; i++) {
      let row = data[startRow + 2 + i];
      if (row && row[colIndex] !== undefined && row[colIndex] !== null) {
        let teacherName = row[colIndex].toString().trim();
        if (teacherName !== "") {
          resultTeachers.push(teacherName);
        }
      }
    }
    res.json({ teachers: resultTeachers });
  } catch (e) {
    res.status(500).json({ error: "خطأ في قراءة الملف: " + e.message });
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

    const dayTargets = {
      'sun': 'الأحد', 'mon': 'الإثنين', 'tue': 'الثلاثاء', 'wed': 'الأربعاء', 'thu': 'الخميس'
    };
    const targetDayName = dayTargets[day];

    rawData.forEach(row => {
      const rowTask = row['المهمة'] || row['task'] || '';
      const teacherInDay = targetDayName ? row[targetDayName] : '';

      if (task !== 'all' && normalizeArabicText(rowTask) !== normalizeArabicText(task)) return;

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
      const templatePath = path.join(__dirname, 'template.xlsx');
      if (fs.existsSync(templatePath)) {
        fs.copyFileSync(templatePath, userExcelPath);
      } else if (fs.existsSync(EXCEL_FILE)) {
        fs.copyFileSync(EXCEL_FILE, userExcelPath);
      } else {
        const newWb = xlsx.utils.book_new();
        xlsx.utils.book_append_sheet(newWb, xlsx.utils.aoa_to_sheet([["المعلم", "الاسم"]]), TEACHERS_SHEET);
        xlsx.utils.book_append_sheet(newWb, xlsx.utils.aoa_to_sheet([["الفصل"]]), CLASSES_SHEET);
        xlsx.utils.book_append_sheet(newWb, xlsx.utils.aoa_to_sheet([["اليوم"]]), WAITING_SHEET);
        xlsx.utils.book_append_sheet(newWb, xlsx.utils.aoa_to_sheet([]), REPORT_SHEET);
        xlsx.utils.book_append_sheet(newWb, xlsx.utils.aoa_to_sheet([[schoolName], [region]]), MAIN_INFO_SHEET);
        xlsx.utils.book_append_sheet(newWb, xlsx.utils.aoa_to_sheet([]), MONITORING_SHEET);
        xlsx.writeFile(newWb, userExcelPath);
      }
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
      fs.copyFileSync(uploadedPath, EXCEL_FILE);
      if (fs.existsSync(uploadedPath)) fs.unlinkSync(uploadedPath);
    }

    return res.json({ success: true, message: "تم تحديث واستبدال بيانات المدرسة بنجاح" });
  } catch (error) {
    cleanFiles();
    console.error("خطأ أثناء تحديث ملف الإكسل:", error);
    return res.status(500).json({ success: false, message: "فشل في معالجة وتحديث الملف: " + error.message });
  }
});

// =========================================================================
// 🗓️ مسارات إعداد جدول الانتظار (waiting_setting) - مواقع صفوف وأعمدة ثابتة
// =========================================================================

// 1️⃣ جلب جدول الانتظار الكامل للقراءة والصفحة
app.get('/get-waiting-schedule-full', async (req, res) => {
  try {
    const userExcel = await getUserExcelPath(req);
    const workbook = xlsx.readFile(userExcel);
    
    if (!workbook.SheetNames.includes(WAITING_SHEET)) {
      return res.json({ success: true, schedule: {} });
    }

    const sheet = getSheet(workbook, WAITING_SHEET);
    const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });

    const daysList = ['الأحد', 'الإثنين', 'الثلاثاء', 'الأربعاء', 'الخميس'];
    let fullSchedule = {};

    daysList.forEach(day => {
      fullSchedule[day] = [
        ["", "", "", "", "", "", ""],
        ["", "", "", "", "", "", ""],
        ["", "", "", "", "", "", ""],
        ["", "", "", "", "", "", ""]
      ];

      const startRow = getDayStartRow(day);

      if (startRow !== -1) {
        for (let r = 0; r < 4; r++) {
          let row = data[startRow + 2 + r];
          if (row) {
            for (let c = 0; c < 7; c++) {
              let val = row[c] !== undefined ? row[c] : "";
              fullSchedule[day][r][c] = val.toString().trim();
            }
          }
        }
      }
    });

    res.json({ success: true, schedule: fullSchedule });
  } catch (e) {
    res.status(500).json({ success: false, error: e.message });
  }
});

// 2️⃣ حفظ جدول اليوم المحدد مباشرة في الصفوف الثابتة المخصصة له
app.post('/save-waiting-schedule-day', async (req, res) => {
  try {
    const { day, schedule } = req.body;
    const username = req.session?.username || req.headers['x-username'] || req.body?.username;
    const userExcel = await getUserExcelPath(req);
    const workbook = xlsx.readFile(userExcel);

    const daysList = ['الأحد', 'الإثنين', 'الثلاثاء', 'الأربعاء', 'الخميس'];
    let sheetData = [];

    if (workbook.SheetNames.includes(WAITING_SHEET)) {
      sheetData = xlsx.utils.sheet_to_json(workbook.Sheets[WAITING_SHEET], { header: 1 });
    }

    // تجهيز وهيكلة الصفوف المحددة لجميع الأيام (من الصف 1 إلى الصف 30)
    const periodHeaders = ["الحصة الأولى", "الحصة الثانية", "الحصة الثالثة", "الحصة الرابعة", "الحصة الخامسة", "الحصة السادسة", "الحصة السابعة"];
    
    daysList.forEach(d => {
      const sRow = getDayStartRow(d);
      if (sRow !== -1) {
        if (!sheetData[sRow]) sheetData[sRow] = [`جدول الانتظار ليوم ${d}`];
        if (!sheetData[sRow + 1]) sheetData[sRow + 1] = periodHeaders;
        for (let i = 0; i < 4; i++) {
          if (!sheetData[sRow + 2 + i]) sheetData[sRow + 2 + i] = ["", "", "", "", "", "", ""];
        }
      }
    });

    // إدراج وحفظ بيانات اليوم في مكانه الثابت والدقيق
    const startRow = getDayStartRow(day);

    if (startRow !== -1) {
      for (let r = 0; r < 4; r++) {
        let rowIndex = startRow + 2 + r;
        if (!sheetData[rowIndex]) {
          sheetData[rowIndex] = [];
        }

        for (let c = 0; c < 7; c++) {
          let cellVal = (schedule && schedule[r] && schedule[r][c]) ? schedule[r][c] : "";
          sheetData[rowIndex][c] = cellVal; // الحفظ يبدأ من العمود A (Index 0)
        }
      }
    }

    const newSheet = xlsx.utils.aoa_to_sheet(sheetData);
    workbook.Sheets[WAITING_SHEET] = newSheet;
    if (!workbook.SheetNames.includes(WAITING_SHEET)) {
      xlsx.utils.book_append_sheet(workbook, newSheet, WAITING_SHEET);
    }

    xlsx.writeFile(workbook, userExcel);
    if (username) await syncExcelToDb(username, userExcel);

    res.json({ success: true, message: "تم الحفظ بنجاح" });
  } catch (e) {
    res.status(500).json({ success: false, error: e.message });
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

// 📥 تنزيل ملف الإكسل مباشرة للمدرسة من قاعدة البيانات
app.get('/api/admin/download-excel/:id', requireAdmin, async (req, res) => {
  try {
    const { id } = req.params;
    const result = await pool.query('SELECT school_name, username, school_excel_file, excel_data FROM users WHERE id = $1', [id]);

    if (result.rows.length === 0) {
      return res.status(404).json({ success: false, message: 'المدرسة غير موجودة' });
    }

    const school = result.rows[0];
    if (!school.excel_data) {
      return res.status(404).json({ success: false, message: 'لا يوجد ملف إكسل مخزن لهذه المدرسة' });
    }

    const fileName = school.school_excel_file || `data_${school.username}.xlsx`;

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${encodeURIComponent(fileName)}"`);
    return res.send(school.excel_data);
  } catch (err) {
    console.error('خطأ في تنزيل ملف الإكسل:', err);
    res.status(500).json({ success: false, error: err.message });
  }
});

// =========================================================================
// 📱 5. واتساب - كل مدرسة برقمها الخاص
// =========================================================================

// تخزين جلسات الواتساب لكل مدرسة { username: { client, isReady, qrData } }
const schoolWaClients = {};

function getPuppeteerConfig() {
  return {
    headless: true,
    executablePath: puppeteer.executablePath(),
    args: [
      '--no-sandbox',
      '--disable-setuid-sandbox',
      '--disable-dev-shm-usage',
      '--disable-accelerated-2d-canvas',
      '--no-first-run',
      '--no-zygote',
      '--disable-gpu'
    ]
  };
}

// 🔗 تهيئة عميل واتساب لمدرسة معينة
function initSchoolWhatsApp(username) {
  if (schoolWaClients[username]?.client) {
    return; // الجلسة موجودة مسبقاً
  }

  const client = new Client({
    authStrategy: new LocalAuth({
      clientId: `school_${username}`,
      dataPath: './whatsapp-sessions'
    }),
    puppeteer: getPuppeteerConfig()
  });

  schoolWaClients[username] = {
    client,
    isReady: false,
    qrData: null
  };

  client.on('qr', async (qr) => {
    console.log(`📱 QR جاهز للمدرسة: ${username}`);
    try {
      schoolWaClients[username].qrData = await qrcode.toDataURL(qr);
      schoolWaClients[username].isReady = false;
    } catch (e) {
      console.error('خطأ في توليد QR:', e.message);
    }
  });

  client.on('ready', () => {
    console.log(`✅ واتساب المدرسة "${username}" متصل!`);
    schoolWaClients[username].isReady = true;
    schoolWaClients[username].qrData = null;
  });

  client.on('auth_failure', () => {
    console.error(`❌ فشل مصادقة واتساب المدرسة: ${username}`);
    schoolWaClients[username].isReady = false;
  });

  client.on('disconnected', () => {
    console.log(`⚠️ انقطع اتصال واتساب المدرسة: ${username}`);
    schoolWaClients[username].isReady = false;
    schoolWaClients[username].client = null;
    // إعادة تهيئة تلقائية بعد 10 ثوانٍ
    setTimeout(() => initSchoolWhatsApp(username), 10000);
  });

  client.initialize();
}

// 🔌 بدء تهيئة واتساب المدرسة (يُستدعى من زر الربط في الواجهة)
app.post('/api/whatsapp/connect', (req, res) => {
  const username = req.session?.username;
  if (!username) return res.status(401).json({ success: false, message: 'يرجى تسجيل الدخول أولاً' });

  if (!schoolWaClients[username]?.client) {
    initSchoolWhatsApp(username);
  }
  res.json({ success: true, message: 'جاري تهيئة الواتساب...' });
});

// 🖼️ جلب صورة QR بصيغة JSON (للواجهة)
app.get('/whatsapp/qr-image', (req, res) => {
  const username = req.session?.username;
  if (!username) return res.status(401).json({ error: 'غير مصرح' });

  const state = schoolWaClients[username];
  if (state?.isReady) return res.json({ connected: true, qr: null });
  if (state?.qrData)  return res.json({ connected: false, qr: state.qrData });
  res.json({ connected: false, qr: null });
});

// 📡 عرض QR أو حالة الاتصال
app.get('/whatsapp/qr', async (req, res) => {
  const username = req.session?.username;
  if (!username) {
    return res.status(401).send('<h3 style="color:red;text-align:center;font-family:sans-serif;margin-top:50px">يرجى تسجيل الدخول أولاً</h3>');
  }

  // تهيئة العميل إذا لم يكن موجوداً
  if (!schoolWaClients[username]?.client) {
    initSchoolWhatsApp(username);
  }

  const state = schoolWaClients[username];

  if (state?.isReady) {
    return res.send(`
      <div style="text-align:center;padding:40px;font-family:sans-serif;direction:rtl">
        <h2 style="color:green">✅ الواتساب متصل ومستعد للإرسال!</h2>
        <p>رقم الواتساب المرتبط بمدرستك جاهز لإرسال الرسائل للمعلمين.</p>
      </div>
    `);
  }

  if (state?.qrData) {
    return res.send(`
      <div style="text-align:center;padding:30px;font-family:sans-serif;direction:rtl">
        <h2>📱 امسح الرمز بتطبيق واتساب</h2>
        <p>افتح واتساب ← النقاط الثلاث ← الأجهزة المرتبطة ← ربط جهاز</p>
        <img src="${state.qrData}" style="width:260px;border:1px solid #ccc;padding:10px;border-radius:10px" />
        <p style="color:orange">⏳ سيتم تحديث الصفحة تلقائياً...</p>
        <script>setTimeout(() => location.reload(), 5000);</script>
      </div>
    `);
  }

  return res.send(`
    <div style="text-align:center;padding:40px;font-family:sans-serif;direction:rtl">
      <h2 style="color:orange">⏳ جاري تحضير الرمز...</h2>
      <p>انتظر 10 ثوانٍ ثم حدّث الصفحة</p>
      <script>setTimeout(() => location.reload(), 8000);</script>
    </div>
  `);
});

// 📊 جلب حالة اتصال واتساب المدرسة (للواجهة)
app.get('/api/whatsapp/status', (req, res) => {
  const username = req.session?.username;
  if (!username) return res.status(401).json({ error: 'غير مصرح' });

  const state = schoolWaClients[username];
  res.json({
    connected: state?.isReady || false,
    hasQr: !!state?.qrData,
    initialized: !!state?.client
  });
});

// 🚀 إرسال رسالة لرقم محدد
const handleSendWhatsApp = async (req, res) => {
  const username = req.session?.username;
  if (!username) return res.status(401).json({ success: false, message: 'يرجى تسجيل الدخول' });

  const state = schoolWaClients[username];
  if (!state?.isReady) {
    return res.status(400).json({
      success: false,
      message: 'واتساب مدرستك غير متصل! اذهب لصفحة /whatsapp/qr لربطه أولاً.'
    });
  }

  const { phone, message } = req.body;
  if (!phone || !message) {
    return res.status(400).json({ success: false, message: 'يرجى إدخال رقم الهاتف والرسالة' });
  }

  try {
    let cleanPhone = phone.toString().replace(/[^0-9]/g, '');
    if (cleanPhone.startsWith('05')) {
      cleanPhone = '966' + cleanPhone.substring(1);
    } else if (cleanPhone.startsWith('5') && cleanPhone.length === 9) {
      cleanPhone = '966' + cleanPhone;
    }

    await state.client.sendMessage(`${cleanPhone}@c.us`, message);
    res.json({ success: true, message: 'تم إرسال الرسالة بنجاح!' });
  } catch (error) {
    console.error('❌ خطأ في الإرسال:', error);
    res.status(500).json({ success: false, error: 'فشل الإرسال: ' + error.message });
  }
};

app.post('/send-whatsapp', handleSendWhatsApp);
app.post('/api/send-whatsapp', handleSendWhatsApp);

// 📢 إرسال رسالة لجميع معلمي المدرسة (من ملف الإكسل)
// ⚠️ تأكد من رقم عمود الهاتف في ملف الإكسل الخاص بك (0 = العمود الأول، 1 = الثاني، 2 = الثالث...)
const PHONE_COLUMN_INDEX = 2; // عدّل هذا الرقم حسب موقع عمود الهاتف في ملفك

app.post('/api/whatsapp/send-all', async (req, res) => {
  const username = req.session?.username;
  if (!username) return res.status(401).json({ success: false, message: 'غير مصرح' });

  const state = schoolWaClients[username];
  if (!state?.isReady) {
    return res.status(400).json({ success: false, message: 'واتساب غير متصل! اذهب لصفحة /whatsapp/qr أولاً.' });
  }

  const { message } = req.body;
  if (!message) return res.status(400).json({ success: false, message: 'يرجى إدخال نص الرسالة' });

  try {
    const userExcel = await getUserExcelPath(req);
    const workbook = xlsx.readFile(userExcel);
    const sheet = getSheet(workbook, TEACHERS_SHEET);
    const teachers = xlsx.utils.sheet_to_json(sheet, { header: 1 }).slice(1); // تخطي رأس الجدول

    const results = [];
    for (const row of teachers) {
      const phone = row[PHONE_COLUMN_INDEX];
      const name  = row[1] || 'غير معروف';

      if (!phone) continue;

      try {
        let cleanPhone = phone.toString().replace(/[^0-9]/g, '');
        if (cleanPhone.startsWith('05')) {
          cleanPhone = '966' + cleanPhone.substring(1);
        } else if (cleanPhone.startsWith('5') && cleanPhone.length === 9) {
          cleanPhone = '966' + cleanPhone;
        }

        await state.client.sendMessage(`${cleanPhone}@c.us`, message);
        results.push({ name, phone, status: '✅ تم الإرسال' });

        // انتظار ثانية ونصف بين كل رسالة لتجنب الحظر
        await new Promise(r => setTimeout(r, 1500));
      } catch (e) {
        results.push({ name, phone, status: '❌ فشل: ' + e.message });
      }
    }

    res.json({ success: true, results });
  } catch (e) {
    res.status(500).json({ success: false, error: e.message });
  }
});

// =========================================================================
// 🚀 6. تشغيل السيرفر
// =========================================================================
app.listen(PORT, () => {
  console.log(`🚀 السيرفر يعمل بنجاح على المنفذ ${PORT}`);
});
