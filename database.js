const initSqlJs = require('sql.js');
const fs = require('fs');
const bcrypt = require('bcryptjs');

async function initDatabase() {
  // تحميل محرك SQLite
  const SQL = await initSqlJs();
  
  let db;
  const dbPath = 'app_data.db';

  // إذا كان الملف موجوداً مسبقاً نقرؤه، وإلا ننشئ قاعدة جديدة
  if (fs.existsSync(dbPath)) {
    const filebuffer = fs.readFileSync(dbPath);
    db = new SQL.Database(filebuffer);
    console.log('تم تحميل قاعدة البيانات الموجودة بنجاح.');
  } else {
    db = new SQL.Database();
    console.log('تم إنشاء قاعدة بيانات جديدة.');
  }

  // 1. إنشاء جدول المستخدمين والمدارس بالهيكلية الجديدة (جعل ملف الإكسل اختياري)
  db.run(`
    CREATE TABLE IF NOT EXISTS users (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      full_name TEXT,
      region TEXT,
      school_name TEXT,
      username TEXT UNIQUE NOT NULL,
      password TEXT NOT NULL,
      school_excel_file TEXT
    )
  `);

  // 🔄 في حال كان ملف قاعدة البيانات موجوداً مسبقاً بالهيكل القديم، نضيف الأعمدة الجديدة تلقائياً
  const columnsToAdd = [
    { name: 'full_name', type: 'TEXT' },
    { name: 'region', type: 'TEXT' },
    { name: 'school_name', type: 'TEXT' }
  ];

  columnsToAdd.forEach(col => {
    try {
      db.run(`ALTER TABLE users ADD COLUMN ${col.name} ${col.type}`);
    } catch (e) {
      // العمود موجود مسبقاً (يتم التغاضي عن الخطأ)
    }
  });

  // 2. إضافة مستخدم تجريبي للطلب الأول
  const rawPassword = '123456';
  const hashedPassword = bcrypt.hashSync(rawPassword, 10);

  try {
    db.run(
      `INSERT OR IGNORE INTO users (full_name, region, school_name, username, password, school_excel_file) VALUES (?, ?, ?, ?, ?, ?)`,
      ['مدرسة الأمل', 'الرياض', 'مدرسة الأمل الثانوية', 'school_al_amal', hashedPassword, 'al_amal_data.xlsx']
    );
    console.log('تم إضافة/التحقق من المدرسة التجريبية بنجاح!');
  } catch (err) {
    console.error('خطأ أثناء إضافة المستخدم:', err.message);
  }

  // حفظ التغييرات إلى الملف على القرص
  const data = db.export();
  const buffer = Buffer.from(data);
  fs.writeFileSync(dbPath, buffer);
  console.log('تم حفظ قاعدة البيانات في الملف app_data.db بنجاح.');
}

initDatabase();