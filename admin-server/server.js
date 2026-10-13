'use strict';

const express = require('express');
const session = require('express-session');
const helmet = require('helmet');
const multer = require('multer');
const XLSX = require('xlsx');
const Database = require('better-sqlite3');
const crypto = require('node:crypto');
const path = require('node:path');

const app = express();
const PORT = Number(process.env.PORT || 3000);
const ADMIN_PASSWORD = process.env.ADMIN_PASSWORD;
const SESSION_SECRET = process.env.SESSION_SECRET;
const DB_PATH = process.env.DB_PATH || path.join(__dirname, 'data', 'admin.sqlite');

if (!ADMIN_PASSWORD || ADMIN_PASSWORD.length < 16 || !SESSION_SECRET || SESSION_SECRET.length < 32) {
  console.error('Set ADMIN_PASSWORD (16+ chars) and SESSION_SECRET (32+ chars) in server environment.');
  process.exit(1);
}

const fs = require('node:fs');
fs.mkdirSync(path.dirname(DB_PATH), { recursive: true });
const db = new Database(DB_PATH);
db.pragma('journal_mode = WAL');
db.exec(`
  CREATE TABLE IF NOT EXISTS price_state (
    code TEXT PRIMARY KEY,
    regular_price INTEGER NOT NULL CHECK (regular_price >= 0),
    sale_price INTEGER CHECK (sale_price IS NULL OR sale_price >= 0),
    updated_at TEXT NOT NULL
  );
  CREATE TABLE IF NOT EXISTS change_log (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    batch_id TEXT NOT NULL,
    change_type TEXT NOT NULL,
    code TEXT NOT NULL,
    old_value TEXT,
    new_value TEXT,
    changed_at TEXT NOT NULL
  );
`);

app.disable('x-powered-by');
app.use(helmet({ contentSecurityPolicy: false }));
app.use(express.urlencoded({ extended: false, limit: '20kb' }));
app.use(express.json({ limit: '100kb' }));
app.use(session({
  name: 'ty_admin',
  secret: SESSION_SECRET,
  resave: false,
  saveUninitialized: false,
  cookie: { httpOnly: true, sameSite: 'strict', secure: process.env.NODE_ENV === 'production', maxAge: 30 * 60 * 1000 }
}));

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 5 * 1024 * 1024, files: 1 }
});
const pending = new Map();

function requireAdmin(req, res, next) {
  if (req.session && req.session.isAdmin) return next();
  return res.status(401).json({ error: 'ابتدا وارد پنل مدیریت شوید.' });
}
function cleanCode(value) {
  return String(value ?? '').trim().replace(/[۰-۹]/g, d => String('۰۱۲۳۴۵۶۷۸۹'.indexOf(d))).replace(/[٠-٩]/g, d => String('٠١٢٣٤٥٦٧٨٩'.indexOf(d)));
}
function findHeader(headers, accepted) {
  return headers.find(h => accepted.includes(String(h).trim().replace(/\u200c/g, ' ')));
}
function parsePriceFile(buffer) {
  const workbook = XLSX.read(buffer, { type: 'buffer', cellDates: false });
  const first = workbook.Sheets[workbook.SheetNames[0]];
  if (!first) throw new Error('شیت اکسل پیدا نشد.');
  const rows = XLSX.utils.sheet_to_json(first, { defval: '', raw: false });
  if (!rows.length) throw new Error('فایل اکسل خالی است.');
  const headers = Object.keys(rows[0]);
  const codeHeader = findHeader(headers, ['کد کالا', 'کد', 'کد محصول']);
  const priceHeader = findHeader(headers, ['قیمت به ریال', 'قیمت', 'قیمت (ریال)']);
  if (!codeHeader || !priceHeader) throw new Error('ستون‌های «کد کالا» و «قیمت به ریال» لازم هستند.');
  const seen = new Set();
  const valid = [];
  const errors = [];
  rows.forEach((row, index) => {
    const code = cleanCode(row[codeHeader]);
    const rawPrice = String(row[priceHeader] ?? '').trim().replace(/[٬,\s]/g, '').replace(/[۰-۹]/g, d => String('۰۱۲۳۴۵۶۷۸۹'.indexOf(d))).replace(/[٠-٩]/g, d => String('٠١٢٣٤٥٦٧٨٩'.indexOf(d)));
    const price = Number(rawPrice);
    const line = index + 2;
    if (!code) { errors.push({ line, reason: 'کد کالا خالی است' }); return; }
    if (!/^\d+$/.test(rawPrice) || !Number.isSafeInteger(price) || price < 0) { errors.push({ line, code, reason: 'قیمت باید عدد صحیح و غیرمنفی به ریال باشد' }); return; }
    if (seen.has(code)) { errors.push({ line, code, reason: 'کد کالا در فایل تکراری است؛ ردیف تکراری خودکار اعمال نمی‌شود' }); return; }
    seen.add(code);
    valid.push({ code, price });
  });
  return { valid, errors, totalRows: rows.length };
}

app.get('/health', (_req, res) => res.json({ status: 'ok', mode: 'staging-only', catalogPublishing: false }));
app.get('/', (_req, res) => res.type('html').send(`<!doctype html><html lang="fa" dir="rtl"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>پنل آزمایشی تکنو یدک</title><body style="font-family:Tahoma,Arial;max-width:760px;margin:40px auto;padding:20px;line-height:2"><h1>پنل آزمایشی مدیریت کاتالوگ</h1><p>این سرویس از کاتالوگ عمومی جداست و در حال حاضر قیمت‌ها را به سایت منتشر نمی‌کند.</p><form method="post" action="/login"><label>رمز مدیر <input name="password" type="password" required autocomplete="current-password"></label> <button>ورود</button></form><p>پس از ورود، از مسیر <code>/admin</code> فایل اکسل را بررسی کنید.</p></body></html>`));
app.post('/login', (req, res) => {
  const supplied = String(req.body.password || '');
  const a = Buffer.from(supplied);
  const b = Buffer.from(ADMIN_PASSWORD);
  const ok = a.length === b.length && crypto.timingSafeEqual(a, b);
  if (!ok) return res.status(401).send('رمز صحیح نیست.');
  req.session.regenerate(err => {
    if (err) return res.status(500).send('خطای ایجاد نشست.');
    req.session.isAdmin = true;
    res.redirect('/admin');
  });
});
app.post('/logout', requireAdmin, (req, res) => req.session.destroy(() => res.redirect('/')));

app.get('/admin', requireAdmin, (_req, res) => res.type('html').send(`<!doctype html><html lang="fa" dir="rtl"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>مدیریت آزمایشی</title><body style="font-family:Tahoma,Arial;max-width:900px;margin:30px auto;padding:20px;line-height:2"><h1>مدیریت آزمایشی قیمت‌ها</h1><p><strong>هشدار:</strong> این نسخه فقط پیش‌نمایش و ثبت در پایگاه دادهٔ آزمایشی را فراهم می‌کند؛ به سایت عمومی متصل نیست.</p><h2>بررسی اکسل قیمت عادی</h2><form method="post" action="/api/preview/regular" enctype="multipart/form-data"><input type="file" name="xlsx" accept=".xlsx,.xls,.csv" required><button>نمایش پیش‌نمایش</button></form><h2>بررسی اکسل قیمت تخفیفی</h2><form method="post" action="/api/preview/sale" enctype="multipart/form-data"><input type="file" name="xlsx" accept=".xlsx,.xls,.csv" required><button>نمایش پیش‌نمایش</button></form><p>فایل باید ستون‌های «کد کالا» و «قیمت به ریال» داشته باشد. این نسخه تا زمان پیاده‌سازی صفحهٔ تأیید امن، هیچ تغییر قیمت را اعمال نمی‌کند.</p><form method="post" action="/logout"><button>خروج</button></form></body></html>`));

app.post('/api/preview/:type', requireAdmin, upload.single('xlsx'), (req, res) => {
  if (!['regular', 'sale'].includes(req.params.type)) return res.status(400).json({ error: 'نوع فایل نامعتبر است.' });
  if (!req.file) return res.status(400).json({ error: 'فایل دریافت نشد.' });
  try {
    const parsed = parsePriceFile(req.file.buffer);
    const previewId = crypto.randomUUID();
    pending.set(previewId, { type: req.params.type, parsed, createdAt: Date.now() });
    for (const [key, item] of pending) if (Date.now() - item.createdAt > 10 * 60 * 1000) pending.delete(key);
    return res.json({
      previewId,
      type: req.params.type,
      totalRows: parsed.totalRows,
      validRows: parsed.valid.length,
      errorCount: parsed.errors.length,
      errors: parsed.errors.slice(0, 100),
      sample: parsed.valid.slice(0, 20),
      note: 'فقط پیش‌نمایش است؛ در این نسخه هیچ قیمتی ذخیره یا در سایت منتشر نمی‌شود.'
    });
  } catch (error) {
    return res.status(400).json({ error: error.message || 'خواندن فایل ناموفق بود.' });
  }
});

app.listen(PORT, () => console.log(`Staging admin server listening on ${PORT}; catalog publishing is disabled.`));
