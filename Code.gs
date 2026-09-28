/** DIYA BRIDAL STUDIO & ACADEMY CRM - Backend (Google Apps Script)
 *  Deploy: Deploy > New deployment > Web app > Execute as: Me > Who has access: Anyone. Copy the /exec URL into index.html (CONFIG.API_URL).
 *  First run: select setup() and click Run once. Default login: admin / admin123 (change it!) */

const HEADERS = {
  Users: ['id','username','passHash','role','name','created'],
  Customers: ['id','name','mobile','whatsapp','email','address','notes','created','updated'],
  Services: ['id','name','price','image','description','active','created','updated'],
  Appointments: ['id','customerId','customerName','serviceId','serviceName','date','time','staff','status','notes','created','updated'],
  Invoices: ['id','invoiceNo','customerId','customerName','items','total','method','status','pdfUrl','staff','date','time','created'],
  Payments: ['id','invoiceNo','amount','method','status','date','time','staff','created'],
  Students: ['id','name','photo','mobile','whatsapp','email','address','course','batch','joined','fee','paid','balance','status','created','updated'],
  Courses: ['id','name','duration','fee','description','modules','status','created','updated'],
  Course_Modules: ['id','courseId','module'],
  Batches: ['id','course','start','end','trainer','status','created'],
  Attendance: ['id','batch','date','studentId','status','created'],
  Student_Fees: ['id','studentId','amount','date','receiptUrl','staff','created'],
  Certificates: ['id','certNo','studentId','studentName','course','completion','pdfUrl','created'],
  Settings: ['key','value']
};
const P = PropertiesService.getScriptProperties();

/** Static site (GitHub Pages) calls this Web App with POST text/plain JSON: {fn, args}. */
const PUBLIC_FNS = {login: login, api: api, verifyCertificate: verifyCertificate, saveInvoicePdf: saveInvoicePdf};
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function doGet() { return json_({ok: true, app: 'Diya CRM API'}); }
function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    const f = PUBLIC_FNS[req.fn];               // whitelist: nothing else can be called
    if (!f) throw new Error('Unknown function');
    return json_({ok: true, data: f.apply(null, req.args || [])});
  } catch (err) { return json_({ok: false, error: String(err.message || err)}); }
}

function setup() {
  let ss = P.getProperty('SSID') ? SpreadsheetApp.openById(P.getProperty('SSID')) : null;
  if (!ss) { ss = SpreadsheetApp.create('DIYA CRM Database'); P.setProperty('SSID', ss.getId()); }
  Object.keys(HEADERS).forEach(n => {
    let sh = ss.getSheetByName(n) || ss.insertSheet(n);
    if (sh.getLastRow() === 0) sh.appendRow(HEADERS[n]);
    sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).setNumberFormat('@'); // keep dates/IDs as plain text
  });
  const d = ss.getSheetByName('Sheet1'); if (d) ss.deleteSheet(d);
  let root = DriveApp.createFolder('DIYA CRM');
  if (!P.getProperty('ROOT')) {
    P.setProperty('ROOT', root.getId());
    ['Invoices','Receipts','Certificates','Student Photos','Customer Photos','QR Code'].forEach(f => root.createFolder(f));
  }
  if (rows_('Users').length === 0) append_('Users', {id: id_('U'), username: 'admin', passHash: hash_('admin123'), role: 'admin', name: 'Owner'});
  return ss.getUrl();
}

/* ---------- helpers ---------- */
const ss_ = () => SpreadsheetApp.openById(P.getProperty('SSID'));
const hash_ = s => Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s));
const id_ = p => p + Utilities.getUuid().slice(0, 8).toUpperCase();
const now_ = () => Utilities.formatDate(new Date(), 'Asia/Kolkata', "yyyy-MM-dd'T'HH:mm:ss");

function rows_(name) {
  const v = ss_().getSheetByName(name).getDataRange().getValues();
  const h = v.shift();
  return v.map(r => Object.fromEntries(h.map((k, i) => [k, r[i] instanceof Date ? Utilities.formatDate(r[i], 'Asia/Kolkata', "yyyy-MM-dd'T'HH:mm:ss") : r[i]])));
}
function append_(name, rec) {
  const h = HEADERS[name];
  if (h.includes('created') && !rec.created) rec.created = now_();
  ss_().getSheetByName(name).appendRow(h.map(k => rec[k] === undefined ? '' : rec[k]));
  return rec;
}
function update_(name, rec) {
  const sh = ss_().getSheetByName(name), h = HEADERS[name], data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) if (data[i][0] === rec.id) {
    if (h.includes('updated')) rec.updated = now_();
    h.forEach((k, c) => { if (rec[k] !== undefined && k !== 'id' && k !== 'created') sh.getRange(i + 1, c + 1).setValue(rec[k]); });
    return rec;
  }
  throw new Error('Record not found');
}
function folder_(name) {
  const root = DriveApp.getFolderById(P.getProperty('ROOT'));
  const it = root.getFoldersByName(name);
  return it.hasNext() ? it.next() : root.createFolder(name);
}
function clean_(o) { Object.keys(o).forEach(k => { if (typeof o[k] === 'string') o[k] = o[k].replace(/[<>]/g, '').trim(); }); return o; }

/* ---------- auth ---------- */
function login(username, password) {
  const u = rows_('Users').find(x => x.username === String(username).trim() && x.passHash === hash_(password));
  if (!u) throw new Error('Wrong username or password');
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('T' + token, JSON.stringify({name: u.name, role: u.role, username: u.username}), 21600);
  return {token, name: u.name, role: u.role};
}
function user_(token) {
  const s = CacheService.getScriptCache().get('T' + token);
  if (!s) throw new Error('Session expired. Please log in again.');
  return JSON.parse(s);
}

/* ---------- single API entry ---------- */
const ADMIN_ONLY = ['remove', 'saveSettings', 'saveQR'];
const TABLES = ['Customers','Services','Appointments','Students','Courses','Batches','Attendance','Student_Fees'];

function api(token, action, p) {
  const u = user_(token);
  if (ADMIN_ONLY.includes(action) && u.role !== 'admin') throw new Error('Admin only');
  p = p || {};
  const t = p.table;
  if (t && !TABLES.includes(t) && !['Invoices','Payments','Certificates'].includes(t)) throw new Error('Invalid table');
  switch (action) {
    case 'list': return rows_(t);
    case 'save': {
      const rec = clean_(p.rec);
      if (t === 'Services' && rec.id && u.role !== 'admin') throw new Error('Only admin can change services');
      if (t === 'Customers' && !/^\d{10}$/.test(String(rec.mobile))) throw new Error('Enter a 10 digit mobile number');
      if (t === 'Customers' && !rec.id && rows_(t).some(c => String(c.mobile) === String(rec.mobile))) throw new Error('This mobile number already exists');
      if (!rec.name && !['Appointments','Batches'].includes(t)) throw new Error('Name is required');
      if (t === 'Students') {
        if (!rec.id) { rec.paid = 0; rec.balance = Number(rec.fee || 0); }
        else { const ex = rows_('Students').find(s => s.id === rec.id); if (ex) rec.balance = Number(rec.fee || 0) - Number(ex.paid || 0); }
      }
      if (rec.id) return update_(t, rec);
      rec.id = id_(t[0]); return append_(t, rec);
    }
    case 'remove': {
      const sh = ss_().getSheetByName(t), d = sh.getDataRange().getValues();
      for (let i = 1; i < d.length; i++) if (d[i][0] === p.id) { sh.deleteRow(i + 1); return true; }
      throw new Error('Record not found');
    }
    case 'settings': { const o = {}; rows_('Settings').forEach(r => o[r.key] = r.value); return o; }
    case 'saveSettings': {
      const sh = ss_().getSheetByName('Settings'), d = sh.getDataRange().getValues();
      Object.keys(p.data).forEach(k => {
        const i = d.findIndex(r => r[0] === k);
        if (i > 0) sh.getRange(i + 1, 2).setValue(p.data[k]); else sh.appendRow([k, p.data[k]]);
      });
      return true;
    }
    case 'saveQR': {
      const blob = Utilities.newBlob(Utilities.base64Decode(p.base64), p.mime, 'owner-upi-qr');
      const f = folder_('QR Code').createFile(blob);
      f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      const url = 'https://drive.google.com/thumbnail?id=' + f.getId() + '&sz=w600';
      api(token, 'saveSettings', {data: {upiQr: url}}); return url;
    }
    case 'invoice': return invoice_(u, p);
    case 'feeAdd': return feeAdd_(u, p);
    case 'attSave': return attSave_(p);
    case 'certGen': return certGen_(u, p);
    case 'report': return report_();
    case 'dashboard': return dashboard_();
    default: throw new Error('Unknown action');
  }
}

/* ---------- billing ---------- */
function invoice_(u, p) {
  if (!p.items || !p.items.length) throw new Error('Select at least one service');
  if (!['Cash','UPI','Card'].includes(p.method)) throw new Error('Select a payment method');
  const services = rows_('Services');
  // Prices always come from the sheet, never from the browser.
  const items = p.items.map(id => {
    const s = services.find(x => x.id === id && String(x.active) !== 'false');
    if (!s) throw new Error('Service not available');
    return {name: s.name, price: Number(s.price)};
  });
  const total = items.reduce((a, b) => a + b.price, 0);
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const year = new Date().getFullYear();
    const n = rows_('Invoices').length + 1;
    const invoiceNo = 'INV/' + year + '/' + ('000' + n).slice(-4);
    const date = now_().slice(0, 10), time = now_().slice(11);
    const c = rows_('Customers').find(x => x.id === p.customerId);
    if (!c) throw new Error('Select a customer');
    const inv = append_('Invoices', {id: id_('I'), invoiceNo, customerId: c.id, customerName: c.name, items: JSON.stringify(items),
      total, method: p.method, status: 'PAID', staff: u.name, date, time});
    append_('Payments', {id: id_('P'), invoiceNo, amount: total, method: p.method, status: 'PAID', date, time, staff: u.name});
    return {invoiceNo, total, items, date, time, customer: c, method: p.method, status: 'PAID'};
  } finally { lock.releaseLock(); }
}

/** Saves an invoice PDF to Drive. html = invoice markup built by the browser. */
function saveInvoicePdf(token, invoiceNo, html) {
  user_(token);
  const blob = Utilities.newBlob(html, 'text/html', invoiceNo.replace(/\//g, '-') + '.html').getAs('application/pdf');
  const f = folder_('Invoices').createFile(blob).setName(invoiceNo.replace(/\//g, '-') + '.pdf');
  const sh = ss_().getSheetByName('Invoices'), d = sh.getDataRange().getValues(), col = HEADERS.Invoices.indexOf('pdfUrl') + 1;
  for (let i = 1; i < d.length; i++) if (d[i][1] === invoiceNo) sh.getRange(i + 1, col).setValue(f.getUrl());
  return f.getUrl();
}

/* ---------- certificates: unique numbers DBSA/YYYY/NNN ---------- */
function nextCertNo_() {
  const y = new Date().getFullYear(), pre = 'DBSA/' + y + '/';
  const max = rows_('Certificates').filter(x => String(x.certNo).startsWith(pre)).reduce((m, x) => Math.max(m, Number(String(x.certNo).split('/')[2]) || 0), 0);
  return pre + ('00' + (max + 1)).slice(-3);
}
function verifyCertificate(certNo) {   // public: used by verification page
  const c = rows_('Certificates').find(x => String(x.certNo).toUpperCase() === String(certNo).trim().toUpperCase());
  return c ? {status: 'VALID', studentName: c.studentName, course: c.course, certNo: c.certNo, completion: c.completion} : {status: 'NOT FOUND'};
}

/* ---------- dashboard ---------- */
const d10_ = v => String(v || '').slice(0, 10);
const settings_ = () => { const o = {}; rows_('Settings').forEach(r => o[r.key] = r.value); return o; };
function dashboard_() {
  const today = now_().slice(0, 10), inv = rows_('Invoices'), ap = rows_('Appointments'), st = rows_('Students');
  return {
    customers: rows_('Customers').length, students: st.length,
    todayAppointments: ap.filter(a => d10_(a.date) === today),
    todaySales: inv.filter(i => d10_(i.date) === today).reduce((a, b) => a + Number(b.total), 0),
    pending: st.reduce((a, s) => a + Number(s.balance || 0), 0),
    recentBills: inv.slice(-5).reverse(), recentStudents: st.slice(-5).reverse()
  };
}

/* ---------- student fees + receipt ---------- */
function pdfFile_(folder, name, html) {
  return folder_(folder).createFile(Utilities.newBlob(html, 'text/html', 'x.html').getAs('application/pdf')).setName(name);
}
function feeAdd_(u, p) {
  const amt = Number(p.amount);
  if (!(amt > 0)) throw new Error('Enter a valid amount');
  const st = rows_('Students').find(s => s.id === p.studentId);
  if (!st) throw new Error('Student not found');
  if (amt > Number(st.balance)) throw new Error('Amount is more than the balance');
  const paid = Number(st.paid || 0) + amt, bal = Number(st.fee) - paid, rid = id_('F'), date = now_().slice(0, 10);
  const html = '<div style="font-family:serif;padding:30px"><h2 style="color:#6b1d3a">DIYA BRIDAL STUDIO & ACADEMY</h2><h3>Fee Receipt</h3><p>Receipt: ' + rid +
    '<br>Date: ' + date + '<br>Student: ' + st.name + '<br>Course: ' + st.course + '</p><p>Amount received: Rs. ' + amt + '<br>Total paid: Rs. ' + paid +
    '<br>Balance: Rs. ' + bal + '</p><p>Received by: ' + u.name + '</p></div>';
  const f = pdfFile_('Receipts', 'Receipt-' + rid + '.pdf', html);
  append_('Student_Fees', {id: rid, studentId: st.id, amount: amt, date, receiptUrl: f.getUrl(), staff: u.name});
  update_('Students', {id: st.id, paid, balance: bal});
  return {receiptUrl: f.getUrl(), paid, balance: bal};
}

/* ---------- attendance (one row per student per batch per day) ---------- */
function attSave_(p) {
  if (!p.batch || !p.date) throw new Error('Select batch and date');
  const sh = ss_().getSheetByName('Attendance'), d = sh.getDataRange().getValues();
  for (let i = d.length - 1; i > 0; i--) if (d[i][1] === p.batch && d10_(d[i][2]) === p.date) sh.deleteRow(i + 1);
  (p.marks || []).forEach(m => {
    if (['Present', 'Absent', 'Leave'].includes(m.status)) append_('Attendance', {id: id_('A'), batch: p.batch, date: p.date, studentId: m.studentId, status: m.status});
  });
  return true;
}

/* ---------- certificate: unique number + PDF in Drive ---------- */
function certGen_(u, p) {
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const st = rows_('Students').find(s => s.id === p.studentId);
    if (!st) throw new Error('Student not found');
    if (st.status !== 'Completed') throw new Error('Set the student status to Completed first');
    if (rows_('Certificates').some(x => x.studentId === st.id)) throw new Error('Certificate already issued for this student');
    const certNo = nextCertNo_(), completion = now_().slice(0, 10), set = settings_();
    const html = '<div style="border:8px double #c9a45c;padding:50px;text-align:center;font-family:serif;color:#3b2a22">' +
      '<h1 style="color:#6b1d3a;letter-spacing:2px">DIYA BRIDAL STUDIO & ACADEMY</h1><h2>CERTIFICATE OF COMPLETION</h2>' +
      '<p>This is to certify that</p><h1>' + st.name + '</h1><p>has successfully completed</p><h2>' + st.course + '</h2>' +
      '<p>Certificate Number: <b>' + certNo + '</b><br>Completion Date: ' + completion + '</p><br><br>' +
      '<p>_______________________<br>Authorized Signature<br>' + (set.signature || '') + '</p></div>';
    const f = pdfFile_('Certificates', certNo.replace(/\//g, '-') + '.pdf', html);
    f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    append_('Certificates', {id: id_('C'), certNo, studentId: st.id, studentName: st.name, course: st.course, completion, pdfUrl: f.getUrl()});
    return {certNo, pdfUrl: f.getUrl()};
  } finally { lock.releaseLock(); }
}

/* ---------- reports ---------- */
function report_() {
  const inv = rows_('Invoices'), st = rows_('Students'), today = now_().slice(0, 10);
  const byDay = {}, byMonth = {}, bySvc = {}, byCourse = {};
  inv.forEach(i => {
    const d = d10_(i.date), t = Number(i.total);
    byDay[d] = (byDay[d] || 0) + t; byMonth[d.slice(0, 7)] = (byMonth[d.slice(0, 7)] || 0) + t;
    try { JSON.parse(i.items).forEach(x => { bySvc[x.name] = (bySvc[x.name] || 0) + Number(x.price); }); } catch (e) {}
  });
  st.forEach(s => { const k = s.course || '-'; byCourse[k] = (byCourse[k] || 0) + 1; });
  const last = o => Object.keys(o).sort().slice(-7).map(k => [k, o[k]]);
  return {
    today: byDay[today] || 0, month: byMonth[today.slice(0, 7)] || 0,
    totalPayments: rows_('Payments').reduce((a, x) => a + Number(x.amount), 0),
    daily: last(byDay), monthly: last(byMonth), services: Object.entries(bySvc).sort((a, b) => b[1] - a[1]),
    students: {total: st.length, active: st.filter(s => s.status === 'Active').length, completed: st.filter(s => s.status === 'Completed').length},
    courses: Object.entries(byCourse), pendingFees: st.reduce((a, s) => a + Number(s.balance || 0), 0), certs: rows_('Certificates').length
  };
}
