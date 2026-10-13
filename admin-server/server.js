'use strict';
const express = require('express');
const session = require('express-session');
const helmet = require('helmet');
const multer = require('multer');
const XLSX = require('xlsx');
const Database = require('better-sqlite3');
const crypto = require('node:crypto');
const path = require('node:path');
const fs = require('node:fs');

const app = express();
const PORT = Number(process.env.PORT || 3000);
const ADMIN_PASSWORD = process.env.ADMIN_PASSWORD;
const SESSION_SECRET = process.env.SESSION_SECRET;
const DB_PATH = process.env.DB_PATH || path.join(__dirname, 'data', 'admin.sqlite');
if (!ADMIN_PASSWORD || ADMIN_PASSWORD.length < 16 || !SESSION_SECRET || SESSION_SECRET.length < 32) {
  console.error('Set ADMIN_PASSWORD (16+ chars) and SESSION_SECRET (32+ chars) in server environment.');
  process.exit(1);
}
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
  name: 'ty_admin', secret: SESSION_SECRET, resave: false, saveUninitialized: false,
  cookie: { httpOnly: true, sameSite: 'strict', secure: process.env.NODE_ENV === 'production', maxAge: 30 * 60 * 1000 }
}));
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 5 * 1024 * 1024, files: 1 } });
const pending = new Map();
function requireAdmin(req, res, next) {
  if (req.session && req.session.isAdmin) return next();
  return res.status(401).json({ error: 'ابتدا وارد پنل مدیریت شوید.' });
}
function normalizeDigits(v) {
  return String(v ?? '').replace(/[۰-۹]/g, d => String('۰۱۲۳۴۵۶۷۸۹'.indexOf(d))).replace(/[٠-٩]/g, d => String('٠١٢٣٤٥٦٧٨٩'.indexOf(d)));
}
function normalizeHeader(v) {
  return String(v ?? '').replace(/\u200c/g, ' ').replace(/[يى]/g, 'ی').replace(/ك/g, 'ک').replace(/[()]/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase();
}
function cleanCode(value) { return normalizeDigits(value).replace(/\s+/g, '').trim(); }
function findHeader(headers, accepted) {
  const choices = accepted.map(normalizeHeader);
  return headers.find(h => choices.includes(normalizeHeader(h)));
}
function parsePriceFile(buffer) {
  const workbook = XLSX.read(buffer, { type: 'buffer', cellDates: false });
  const first = workbook.Sheets[workbook.SheetNames[0]];
  if (!first) throw new Error('شیت اکسل پیدا نشد.');
  const rows = XLSX.utils.sheet_to_json(first, { defval: '', raw: false });
  if (!rows.length) throw new Error('فایل اکسل خالی است یا ردیف عنوان ندارد.');
  const headers = Object.keys(rows[0]);
  const codeHeader = findHeader(headers, ['کد کالا', 'کد', 'کد محصول', 'کدکالا']);
  const priceHeader = findHeader(headers, ['قیمت به ریال', 'قیمت', 'قیمت (ریال)', 'قیمت ریال']);
  if (!codeHeader || !priceHeader) throw new Error('ستون‌های «کد کالا» و «قیمت به ریال» لازم هستند. عنوان ستون‌ها را بررسی کنید.');
  const seen = new Set(), valid = [], errors = [];
  rows.forEach((row, index) => {
    const code = cleanCode(row[codeHeader]);
    const rawPrice = normalizeDigits(row[priceHeader]).trim().replace(/[٬,،\s]/g, '');
    const price = Number(rawPrice);
    const line = index + 2;
    if (!code) { errors.push({ line, code: '', price: String(row[priceHeader] ?? ''), reason: 'کد کالا خالی است' }); return; }
    if (!/^\d+$/.test(rawPrice) || !Number.isSafeInteger(price) || price < 0) {
      errors.push({ line, code, price: String(row[priceHeader] ?? ''), reason: 'قیمت باید عدد صحیح و غیرمنفی به ریال باشد' }); return;
    }
    if (seen.has(code)) { errors.push({ line, code, price, reason: 'کد کالا در فایل تکراری است؛ این ردیف خودکار اعمال نمی‌شود' }); return; }
    seen.add(code);
    valid.push({ line, code, price });
  });
  return { valid, errors, totalRows: rows.length, headers: { code: codeHeader, price: priceHeader } };
}
const loginPage = `<!doctype html><html lang="fa" dir="rtl"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>پنل آزمایشی تکنو یدک</title><style>body{font-family:Tahoma,Arial;max-width:760px;margin:30px auto;padding:18px;line-height:2;background:#f2f6fa;color:#17344b}main{background:white;border-radius:16px;padding:24px;box-shadow:0 4px 20px #1232}input,button{font:inherit;padding:10px;border-radius:8px}button{background:#153a56;color:white;border:0;cursor:pointer}</style><main><h1>پنل آزمایشی مدیریت کاتالوگ</h1><p>این پنل جدا از سایت عمومی است. در این مرحله فقط فایل اکسل را بررسی می‌کند و هیچ قیمتی منتشر نمی‌کند.</p><form method="post" action="/login"><label>رمز مدیر <input name="password" type="password" required autocomplete="current-password"></label> <button>ورود</button></form></main></html>`;
const adminPage = `<!doctype html><html lang="fa" dir="rtl"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>مدیریت آزمایشی قیمت‌ها</title>
<style>
*{box-sizing:border-box}body{font-family:Tahoma,Arial;max-width:1100px;margin:20px auto;padding:12px;line-height:1.9;background:#f2f6fa;color:#17344b}main{background:#fff;border-radius:16px;padding:20px;box-shadow:0 4px 20px #1232}h1{font-size:24px}.warn{background:#fff5dc;padding:12px;border-radius:10px}.panels{display:grid;grid-template-columns:repeat(auto-fit,minmax(260px,1fr));gap:14px}.panel{border:1px solid #dce6ee;border-radius:12px;padding:14px}input,button{font:inherit;padding:10px;border-radius:8px;max-width:100%}button{background:#153a56;color:white;border:0;cursor:pointer}button:disabled{opacity:.5}#result{margin-top:20px;overflow-wrap:anywhere}.stats{display:flex;gap:10px;flex-wrap:wrap}.stat{background:#edf4f8;border-radius:9px;padding:10px}.tablewrap{overflow-x:auto}table{border-collapse:collapse;width:100%;font-size:14px;background:white}th,td{padding:8px;border-bottom:1px solid #e2e8ed;text-align:right}th{background:#edf4f8}.bad{color:#a22}.good{color:#176b43}.muted{color:#637788}form.logout{margin-top:18px}@media(max-width:500px){main{padding:12px}h1{font-size:20px}table{font-size:12px}}
</style><main><h1>مدیریت آزمایشی قیمت‌ها</h1><p class="warn"><b>مرحله آزمایشی:</b> این صفحه فقط پیش‌نمایش می‌سازد. هیچ قیمت در پایگاه داده یا سایت عمومی ثبت/منتشر نمی‌شود.</p>
<div class="panels"><section class="panel"><h2>فایل قیمت عادی</h2><p>ستون‌های لازم: «کد کالا» و «قیمت به ریال»</p><form class="previewForm" data-type="regular"><input type="file" name="xlsx" accept=".xlsx,.xls,.csv" required><p><button>بررسی و پیش‌نمایش قیمت عادی</button></p></form></section>
<section class="panel"><h2>فایل قیمت تخفیفی</h2><p>فایل تخفیف جدا بررسی می‌شود و قیمت اصلی را جایگزین نمی‌کند.</p><form class="previewForm" data-type="sale"><input type="file" name="xlsx" accept=".xlsx,.xls,.csv" required><p><button>بررسی و پیش‌نمایش تخفیف</button></p></form></section></div>
<div id="result" aria-live="polite"></div><form class="logout" method="post" action="/logout"><button>خروج از پنل</button></form>
<script>
const esc=v=>String(v??'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');
document.querySelectorAll('.previewForm').forEach(form=>form.addEventListener('submit',async e=>{
 e.preventDefault();const button=form.querySelector('button');const result=document.getElementById('result');button.disabled=true;button.textContent='در حال بررسی...';result.innerHTML='<p>فایل در حال بررسی است...</p>';
 try{const data=new FormData(form);const res=await fetch('/api/preview/'+form.dataset.type,{method:'POST',body:data});const out=await res.json();if(!res.ok)throw new Error(out.error||'بررسی فایل ناموفق بود');
 let html='<h2>نتیجه بررسی</h2><div class="stats"><div class="stat">کل ردیف‌ها: <b>'+out.totalRows+'</b></div><div class="stat good">ردیف معتبر: <b>'+out.validRows+'</b></div><div class="stat bad">خطاها: <b>'+out.errorCount+'</b></div></div><p class="muted">ستون کد شناسایی‌شده: '+esc(out.headers.code)+' | ستون قیمت: '+esc(out.headers.price)+'</p><p class="warn">'+esc(out.note)+'</p>';
 if(out.sample.length){html+='<h3>نمونه ردیف‌های معتبر (حداکثر ۲۰ ردیف)</h3><div class="tablewrap"><table><thead><tr><th>ردیف اکسل</th><th>کد کالا</th><th>قیمت (ریال)</th></tr></thead><tbody>'+out.sample.map(x=>'<tr><td>'+x.line+'</td><td>'+esc(x.code)+'</td><td>'+Number(x.price).toLocaleString('fa-IR')+'</td></tr>').join('')+'</tbody></table></div>'}
 if(out.errors.length){html+='<h3 class="bad">خطاها (حداکثر ۱۰۰ مورد نمایش داده می‌شود)</h3><div class="tablewrap"><table><thead><tr><th>ردیف اکسل</th><th>کد کالا</th><th>قیمت ورودی</th><th>علت</th></tr></thead><tbody>'+out.errors.map(x=>'<tr><td>'+x.line+'</td><td>'+esc(x.code)+'</td><td>'+esc(x.price)+'</td><td class="bad">'+esc(x.reason)+'</td></tr>').join('')+'</tbody></table></div>'}
 result.innerHTML=html;
 }catch(err){result.innerHTML='<p class="bad">'+esc(err.message)+'</p>'}
 finally{button.disabled=false;button.textContent=form.dataset.type==='regular'?'بررسی و پیش‌نمایش قیمت عادی':'بررسی و پیش‌نمایش تخفیف'}
}));
</script></main></html>`;
app.get('/health', (_req, res) => res.json({ status: 'ok', mode: 'staging-only', catalogPublishing: false }));
app.get('/', (_req, res) => res.type('html').send(loginPage));
app.post('/login', (req, res) => {
  const supplied = String(req.body.password || '');
  const a = Buffer.from(supplied), b = Buffer.from(ADMIN_PASSWORD);
  const ok = a.length === b.length && crypto.timingSafeEqual(a, b);
  if (!ok) return res.status(401).send('رمز صحیح نیست.');
  req.session.regenerate(err => {
    if (err) return res.status(500).send('خطای ایجاد نشست.');
    req.session.isAdmin = true;
    res.redirect('/admin');
  });
});
app.post('/logout', requireAdmin, (req, res) => req.session.destroy(() => res.redirect('/')));
app.get('/admin', requireAdmin, (_req, res) => res.type('html').send(adminPage));
app.post('/api/preview/:type', requireAdmin, upload.single('xlsx'), (req, res) => {
  if (!['regular', 'sale'].includes(req.params.type)) return res.status(400).json({ error: 'نوع فایل نامعتبر است.' });
  if (!req.file) return res.status(400).json({ error: 'فایل دریافت نشد.' });
  try {
    const parsed = parsePriceFile(req.file.buffer);
    const previewId = crypto.randomUUID();
    pending.set(previewId, { type: req.params.type, parsed, createdAt: Date.now() });
    for (const [key, item] of pending) if (Date.now() - item.createdAt > 10 * 60 * 1000) pending.delete(key);
    return res.json({
      previewId, type: req.params.type, totalRows: parsed.totalRows, validRows: parsed.valid.length,
      errorCount: parsed.errors.length, errors: parsed.errors.slice(0, 100), sample: parsed.valid.slice(0, 20),
      headers: parsed.headers,
      note: 'فقط پیش‌نمایش است؛ در این نسخه هیچ قیمتی ذخیره یا در سایت منتشر نمی‌شود.'
    });
  } catch (error) {
    return res.status(400).json({ error: error.message || 'خواندن فایل ناموفق بود.' });
  }
});
app.listen(PORT, () => console.log(`Staging admin server listening on ${PORT}; catalog publishing is disabled.`));
