// =========================================================================
// 📦 1. استدعاء المكتبات وإعداد التطبيق
// =========================================================================
const express = require('express');
const path = require('path');
const fs = require('fs');
const xlsx = require('xlsx');
const bodyParser = require('body-parser');
const initSqlJs = require('sql.js');
const bcrypt = require('bcryptjs');
const multer = require('multer');
const session = require('express-session');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.static(__dirname));
app.use(bodyParser.json());
app.use(express.urlencoded({ extended: true }));

// 🎯 إعداد الجلسات (Session) لحفظ هوية المدرسة المسجلة
app.use(session({
  secret: 'school_management_secret_key_2026',
  resave: false,
  saveUninitialized: false,
  cookie: { 
    maxAge: 24 * 60 * 60 * 1000, // صلاحية الجلسة 24 ساعة
    httpOnly: true 
  }
}));

// منع التخزين المؤقت (Cache) للبيانات
app.use((req, res, next) => {
  res.set('Cache-Control', 'no-store, no-cache, must-revalidate, private');
  next();
});

// 🏠 [تمت الإضافة]: توجيه الزائر تلقائياً من الرابط الرئيسي إلى صفحة تسجيل الدخول
app.get('/', (req, res) => {
  res.redirect('/login.html');
});

// =========================================================================
// 🗄️ 2. التعامل مع قاعدة البيانات SQLite (sql.js)
// =========================================================================
async function getDb() {
  const SQL = await initSqlJs();
  const dbPath = path.join(__dirname, 'app_data.db');
  let db;

  if (fs.existsSync(dbPath)) {
    const filebuffer = fs.readFileSync(dbPath);
    db = new SQL.Database(filebuffer);
  } else {
    db = new SQL.Database();
    db.run(`
      CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE,
        password TEXT,
        full_name TEXT,
        region TEXT,
        school_name TEXT,
        school_excel_file TEXT
      )
    `);
    const data = db.export();
    fs.writeFileSync(dbPath, Buffer.from(data));
  }

  // التأكد من وجود كافة الأعمدة المطلوبة
  const columns = ['full_name', 'region', 'school_name', 'school_excel_file'];
  let tableUpdated = false;
  columns.forEach(col => {
    try {
      db.run(`ALTER TABLE users ADD COLUMN ${col} TEXT`);
      tableUpdated = true;
    } catch (e) {
      // العمود موجود مسبقاً
    }
  });

  if (tableUpdated) {
    const data = db.export();
    fs.writeFileSync(dbPath, Buffer.from(data));
  }

  return db;
}

// 🔐 مسار تسجيل الدخول
app.post('/login', async (req, res) => {
  const { username, password } = req.body;

  try {
    const db = await getDb();
    const stmt = db.prepare('SELECT * FROM users WHERE username = ?');
    stmt.bind([username]);

    if (stmt.step()) {
      const user = stmt.getAsObject();
      stmt.free();

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
    } else {
      stmt.free();
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
// ⚙️ إعدادات أسماء الملف والورقات
// =========================================================================
const EXCEL_FILE = path.join(__dirname, 'waiting_data.xlsx');

const TEACHERS_SHEET = 'teacher_name';
const CLASSES_SHEET = 'classes';
const WAITING_SHEET = ' Waiting_table';
const REPORT_SHEET = 'report';
const MAIN_INFO_SHEET = 'maininfo';
const MONITORING_SHEET = 'moni';

// 🛠️ دالة تحديد وإنشاء ملف المدرسة
async function getUserExcelPath(req) {
  const username = req.session?.username || req.headers['x-username'] || req.query?.username || req.body?.username;
  
  if (!username) {
    throw new Error("UNAUTHORIZED: يرجى تسجيل الدخول أولاً للوصول لبيانات المدرسة.");
  }

  const uploadsDir = path.join(__dirname, 'uploads');
  if (!fs.existsSync(uploadsDir)) {
    fs.mkdirSync(uploadsDir, { recursive: true });
  }

  let schoolFileName = null;
  try {
    const db = await getDb();
    const stmt = db.prepare('SELECT school_excel_file FROM users WHERE username = ?');
    stmt.bind([username]);
    if (stmt.step()) {
      const user = stmt.getAsObject();
      schoolFileName = user.school_excel_file;
    }
    stmt.free();
  } catch (e) {
    console.error('خطأ في تحديد ملف المستخدم من DB:', e.message);
  }

  if (!schoolFileName || schoolFileName.trim() === '') {
    schoolFileName = `data_${username}.xlsx`;
  }

  const userFilePath = path.join(uploadsDir, schoolFileName);

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

    let dayIndex = data.findIndex(row => row && row[0] === day);
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

app.post('/save-report', async (req, res) => {
  const reportData = req.body;
  try {
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
// 📦 رفع الملفات وإنشاء الحسابات
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

app.post('/register', upload.none(), async (req, res) => {
  try {
    const { fullName, region, schoolName, username, password } = req.body;

    if (!fullName || !region || !schoolName || !username || !password) {
      return res.status(400).json({ success: false, message: 'يرجى تعبئة جميع الحقول المطلوبة!' });
    }

    const db = await getDb();
    const checkStmt = db.prepare('SELECT username FROM users WHERE username = ?');
    checkStmt.bind([username]);
    if (checkStmt.step()) {
      checkStmt.free();
      return res.status(400).json({ success: false, message: 'اسم المستخدم مسجل بالفعل، اختر اسماً آخر!' });
    }
    checkStmt.free();

    const uploadsDir = path.join(__dirname, 'uploads');
    if (!fs.existsSync(uploadsDir)) fs.mkdirSync(uploadsDir, { recursive: true });

    const userExcelFileName = `data_${username}.xlsx`;
    const userExcelPath = path.join(uploadsDir, userExcelFileName);
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

    const hashedPassword = bcrypt.hashSync(password, 10);
    const insertStmt = db.prepare(`
      INSERT INTO users (full_name, region, school_name, username, password, school_excel_file) 
      VALUES (?, ?, ?, ?, ?, ?)
    `);
    
    insertStmt.run([fullName, region, schoolName, username, hashedPassword, userExcelFileName]);
    insertStmt.free();

    const data = db.export();
    fs.writeFileSync(path.join(__dirname, 'app_data.db'), Buffer.from(data));

    return res.json({
      success: true,
      message: 'تم إنشاء الحساب بنجاح وتجهيز ملف إكسل خاص بمدرستك! يتم توجيهك لصفحة الدخول...'
    });

  } catch (error) {
    res.status(500).json({ success: false, message: 'حدث خطأ في السيرفر أثناء معالجة الطلب: ' + error.message });
  }
});

app.post('/update-school-excel', upload.single('excelFile'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ success: false, message: "يرجى اختيار ملف الأكسل أولاً!" });
    }

    const username = req.body.username || req.session?.username;
    const uploadedPath = req.file.path;

    if (username) {
      const targetFileName = req.file.filename;

      const db = await getDb();
      const updateStmt = db.prepare('UPDATE users SET school_excel_file = ? WHERE username = ?');
      updateStmt.run([targetFileName, username]);
      updateStmt.free();

      const data = db.export();
      fs.writeFileSync(path.join(__dirname, 'app_data.db'), Buffer.from(data));
    } else {
      fs.copyFileSync(uploadedPath, EXCEL_FILE);
      if (fs.existsSync(uploadedPath)) fs.unlinkSync(uploadedPath);
    }

    return res.json({ success: true, message: "تم تحديث واستبدال بيانات المدرسة بنجاح" });
  } catch (error) {
    if (req.file && fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
    return res.status(500).json({ success: false, message: "فشل في معالجة وتحديث الملف: " + error.message });
  }
});

// =========================================================================
// 🛠️ 3. مسارات لوحة تحكم الأدمن المكتملة الشاملة (Admin Routes)
// =========================================================================

// 1. جلب قائمة كل المدارس والبيانات الكاملة (بما فيها كلمَة المرور و المَلف)
app.get('/api/admin/schools', async (req, res) => {
  try {
    const db = await getDb();
    const stmt = db.prepare('SELECT id, full_name, school_name, region, username, password, school_excel_file AS excel_path FROM users');
    let schools = [];
    while (stmt.step()) {
      schools.push(stmt.getAsObject());
    }
    stmt.free();
    res.json({ success: true, schools });
  } catch (error) {
    res.status(500).json({ success: false, error: error.message });
  }
});

// 2. تعديل كافة بيانات المدرسة + إمكانية رفع/استبدال ملف الإكسل مباشرة
app.put('/api/admin/schools/:id', upload.single('excel_file'), async (req, res) => {
  try {
    const { id } = req.params;
    const { full_name, school_name, region, username, password } = req.body;
    const db = await getDb();

    // تشفير كلمة المرور إذا تم تغييرها وتوفيرها بصيغة جديدة غير مشفرة
    let finalPassword = password;
    if (password && !password.startsWith('$2a$') && !password.startsWith('$2b$')) {
      finalPassword = bcrypt.hashSync(password, 10);
    }

    let excelFileName = req.file ? req.file.filename : null;

    if (excelFileName) {
      // تعديل مع رفع/استبدال ملف إكسل جديد
      const updateStmt = db.prepare(`
        UPDATE users 
        SET full_name = ?, school_name = ?, region = ?, username = ?, password = ?, school_excel_file = ? 
        WHERE id = ?
      `);
      updateStmt.run([full_name, school_name, region, username, finalPassword, excelFileName, id]);
      updateStmt.free();
    } else {
      // تعديل البيانات بدون استبدال الملف الحالي
      const updateStmt = db.prepare(`
        UPDATE users 
        SET full_name = ?, school_name = ?, region = ?, username = ?, password = ? 
        WHERE id = ?
      `);
      updateStmt.run([full_name, school_name, region, username, finalPassword, id]);
      updateStmt.free();
    }

    // حفظ التغييرات فوراً في قاعدة البيانات SQLite
    const data = db.export();
    fs.writeFileSync(path.join(__dirname, 'app_data.db'), Buffer.from(data));

    res.json({ success: true, message: 'تم تحديث بيانات المدرسة وملف الإكسل بنجاح' });
  } catch (error) {
    res.status(500).json({ success: false, error: error.message });
  }
});

// 3. حذف مدرسة نهائياً من قاعدة البيانات
app.delete('/api/admin/schools/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const db = await getDb();
    const deleteStmt = db.prepare('DELETE FROM users WHERE id = ?');
    deleteStmt.run([id]);
    deleteStmt.free();

    // حفظ التغييرات فوراً في قاعدة البيانات SQLite
    const data = db.export();
    fs.writeFileSync(path.join(__dirname, 'app_data.db'), Buffer.from(data));

    res.json({ success: true, message: 'تم حذف المدرسة بنجاح' });
  } catch (error) {
    res.status(500).json({ success: false, error: error.message });
  }
});

// =========================================================================
// 🚀 تشغيل السيرفر
// =========================================================================

app.listen(PORT, '0.0.0.0', () => {
    console.log(`Server is running on port ${PORT}`);
});