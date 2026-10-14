const express = require('express'), Database = require('better-sqlite3'), bcrypt = require('bcryptjs'),
  jwt = require('jsonwebtoken'), path = require('path');
const crypto = require('crypto'), nodemailer = require('nodemailer'), ExcelJS = require('exceljs'), PDF = require('pdfkit');
const SECRET = process.env.JWT_SECRET;
if (!SECRET || SECRET.length < 16) { console.error('Définissez JWT_SECRET (16 caractères minimum).'); process.exit(1); }
const db = new Database(process.env.DB_PATH || 'akwe.db');
db.pragma('journal_mode = WAL');
db.exec(`
CREATE TABLE IF NOT EXISTS users(id INTEGER PRIMARY KEY, name TEXT, email TEXT UNIQUE, hash TEXT);
CREATE TABLE IF NOT EXISTS sales(uid INT, id TEXT, a TEXT, q REAL, p REAL, c REAL, m TEXT, d TEXT, PRIMARY KEY(uid,id));
CREATE TABLE IF NOT EXISTS moves(uid INT, id TEXT, t TEXT, k TEXT, a REAL, n TEXT, d TEXT, PRIMARY KEY(uid,id));
CREATE TABLE IF NOT EXISTS momo(uid INT, id TEXT, o TEXT, a REAL, r TEXT, d TEXT, s TEXT, PRIMARY KEY(uid,id));`);

db.exec('CREATE TABLE IF NOT EXISTS resets(uid INT, h TEXT PRIMARY KEY, exp INT)');
for (const c of ['consent TEXT', 'tv INTEGER DEFAULT 0']) { try { db.exec('ALTER TABLE users ADD COLUMN ' + c); } catch {} }
const sha = t => crypto.createHash('sha256').update(t).digest('hex');
const today = () => new Date().toISOString().slice(0, 10);
/* limitation des tentatives (en mémoire) */
const hits = new Map();
const tooMany = (k, max, win) => { const n = Date.now(), h = (hits.get(k) || []).filter(t => n - t < win); hits.set(k, h); return h.length >= max; };
const hit = k => hits.set(k, [...(hits.get(k) || []), Date.now()]);
setInterval(() => { const n = Date.now(); for (const [k, h] of hits) if (!h.some(t => n - t < 36e5)) hits.delete(k); }, 6e5).unref();
/* e-mail : SMTP_URL (ex. smtps://user:pass@smtp.exemple.com) ; sinon le lien est affiché dans la console */
const tr = process.env.SMTP_URL ? nodemailer.createTransport(process.env.SMTP_URL) : null;
const mail = async (to, link) => {
  if (!tr) return console.log('[RESET] ' + to + ' -> ' + link);
  await tr.sendMail({ from: process.env.MAIL_FROM || 'Akwè Caisse <no-reply@localhost>', to, subject: 'Réinitialisation de votre mot de passe',
    text: 'Pour choisir un nouveau mot de passe, ouvrez ce lien (valable 1 heure) :\n' + link + '\n\nSi vous n\'êtes pas à l\'origine de cette demande, ignorez ce message.' });
};
const num = v => (Number.isFinite(+v) ? +v : 0), str = (v, n = 200) => String(v ?? '').slice(0, n);
const dt = v => (/^\d{4}-\d{2}-\d{2}$/.test(v) ? v : today());
const sign = (id, tv = 0) => jwt.sign({ uid: id, tv }, SECRET, { expiresIn: '30d' });
const guard = (q, s, n) => {
  try {
    const p = jwt.verify((q.headers.authorization || '').slice(7), SECRET), u = db.prepare('SELECT tv FROM users WHERE id=?').get(p.uid);
    if (!u || (u.tv || 0) !== p.tv) throw 0;
    q.uid = p.uid; n();
  }
  catch { s.status(401).json({ error: 'Non autorisé' }); }
};

/* ---- Import des SMS Mobile Money (MTN MoMo / Moov Money) ---- */
const AM = /(\d[\d .,]*)\s*(?:FCFA|F\s?CFA|XOF)/i;
function parseSms(text) {
  const parts = String(text).split(/\n{2,}/).flatMap(c => {
    const l = c.split('\n').map(s => s.trim()).filter(Boolean);
    return l.length > 1 && l.every(s => AM.test(s)) ? l : [l.join(' ')];
  });
  return parts.map(s => {
    if (!/re[cç]u|received|d[ée]p[ôo]t|deposit|cr[ée]dit/i.test(s)) return null; // paiements entrants seulement
    const a = s.match(AM); if (!a) return null;
    const amount = parseInt(a[1].trim().replace(/\s/g, '').replace(/[.,]\d{1,2}$/, '').replace(/[.,]/g, ''), 10);
    if (!(amount > 0)) return null;
    const ref = ((s.match(/(?:transaction\s*)?(?:id|ref|r[ée]f[ée]rence)(?:\s+de\s+transaction)?\s*[:.\-]?\s*([A-Z0-9.\-]{6,})/i) || [])[1] || '').replace(/[.\-]+$/, '');
    const from = ((s.match(/\bde\s+([A-Za-zÀ-ÿ' ]{2,40}?)\s*(?:\(|\d|\.|,|le\s)/) || [])[1] || '').trim();
    const dm = s.match(/(\d{2})[\/-](\d{2})[\/-](\d{4})/);
    return {
      o: /moov/i.test(s) ? 'Moov Money' : 'MTN MoMo', a: amount,
      r: ref ? (from ? from + ' · ' + ref : ref) : from, d: dm ? `${dm[3]}-${dm[2]}-${dm[1]}` : today(),
    };
  }).filter(Boolean);
}

/* ---- Données par période, rapports ---- */
const per = v => (['jour', 'semaine', 'mois'].includes(v) ? v : 'jour');
function data(uid, p) {
  const days = p === 'jour' ? 1 : p === 'semaine' ? 7 : 30;
  const from = new Date(Date.now() - (days - 1) * 864e5).toISOString().slice(0, 10);
  const g = t => db.prepare(`SELECT * FROM ${t} WHERE uid=? AND d>=? ORDER BY d, rowid`).all(uid, from);
  return { from, sl: g('sales'), mv: g('moves'), mo: g('momo') };
}
function report(uid, p) {
  const { from, sl, mv } = data(uid, p);
  const v = sl.reduce((t, x) => t + x.q * x.p, 0), cg = sl.reduce((t, x) => t + x.q * x.c, 0);
  const ei = mv.filter(m => m.t === 'in').reduce((t, m) => t + m.a, 0);
  const out = mv.filter(m => m.t === 'out'), so = out.reduce((t, m) => t + m.a, 0);
  const ch = out.filter(m => m.k !== 'Achat marchandises').reduce((t, m) => t + m.a, 0);
  return { periode: p, depuis: from, ventes: v, autres_entrees: ei, cout_produits: cg, charges: ch, sorties: so, benefice: v - cg - ch + ei };
}

const app = express();
app.set('trust proxy', 1);
app.use((q, s, n) => { s.set({ 'X-Content-Type-Options': 'nosniff', 'Referrer-Policy': 'same-origin', 'X-Frame-Options': 'DENY' }); n(); });
app.use(express.json({ limit: '2mb' }));
app.use(express.static(path.join(__dirname, 'public')));

app.post('/api/register', (q, s) => {
  const ip = 'r:' + q.ip;
  if (tooMany(ip, 10, 36e5)) return s.status(429).json({ error: 'Trop de créations de compte. Réessayez plus tard.' });
  hit(ip);
  const { name, email, password, consent } = q.body || {};
  if (consent !== true) return s.status(400).json({ error: 'Vous devez accepter les conditions et la politique de confidentialité.' });
  if (!/^\S+@\S+\.\S+$/.test(email || '')) return s.status(400).json({ error: 'E-mail invalide.' });
  if (String(password || '').length < 8) return s.status(400).json({ error: 'Mot de passe : 8 caractères minimum.' });
  if (db.prepare('SELECT 1 FROM users WHERE email=?').get(email.toLowerCase())) return s.status(409).json({ error: 'Ce compte existe déjà.' });
  const r = db.prepare('INSERT INTO users(name,email,hash,consent) VALUES(?,?,?,?)').run(str(name, 80), email.toLowerCase(), bcrypt.hashSync(password, 10), new Date().toISOString());
  s.json({ token: sign(r.lastInsertRowid, 0) });
});
app.post('/api/login', (q, s) => {
  const { email, password } = q.body || {}, e = String(email || '').toLowerCase(), k = 'l:' + q.ip + ':' + e;
  if (tooMany(k, 5, 9e5)) return s.status(429).json({ error: 'Trop de tentatives. Réessayez dans 15 minutes.' });
  const u = db.prepare('SELECT * FROM users WHERE email=?').get(e);
  if (!u || !bcrypt.compareSync(String(password || ''), u.hash)) { hit(k); return s.status(401).json({ error: 'E-mail ou mot de passe incorrect.' }); }
  hits.delete(k);
  s.json({ token: sign(u.id, u.tv || 0) });
});
app.post('/api/forgot', async (q, s) => {
  const k = 'f:' + q.ip;
  if (tooMany(k, 5, 36e5)) return s.status(429).json({ error: 'Trop de demandes. Réessayez plus tard.' });
  hit(k);
  const email = String((q.body || {}).email || '').toLowerCase(), u = db.prepare('SELECT id FROM users WHERE email=?').get(email);
  if (u) {
    const t = crypto.randomBytes(32).toString('hex');
    db.prepare('INSERT INTO resets VALUES(?,?,?)').run(u.id, sha(t), Date.now() + 36e5);
    try { await mail(email, (process.env.BASE_URL || `${q.protocol}://${q.get('host')}`) + '/#reset=' + t); } catch (e) { console.error('mail:', e.message); }
  }
  s.json({ ok: true });
});
app.post('/api/reset', (q, s) => {
  const { token, password } = q.body || {};
  if (String(password || '').length < 8) return s.status(400).json({ error: 'Mot de passe : 8 caractères minimum.' });
  const r = db.prepare('SELECT uid FROM resets WHERE h=? AND exp>?').get(sha(String(token || '')), Date.now());
  if (!r) return s.status(400).json({ error: 'Lien invalide ou expiré.' });
  db.prepare('UPDATE users SET hash=?, tv=COALESCE(tv,0)+1 WHERE id=?').run(bcrypt.hashSync(password, 10), r.uid);
  db.prepare('DELETE FROM resets WHERE uid=?').run(r.uid);
  s.json({ ok: true });
});

app.get('/api/data', guard, (q, s) => {
  const g = (t, c) => db.prepare(`SELECT ${c} FROM ${t} WHERE uid=? ORDER BY rowid DESC`).all(q.uid);
  s.json({ sales: g('sales', 'id,a,q,p,c,m,d'), moves: g('moves', 'id,t,k,a,n,d'), momo: g('momo', 'id,o,a,r,d,s') });
});

const iS = db.prepare('INSERT OR REPLACE INTO sales VALUES(?,?,?,?,?,?,?,?)'),
  iM = db.prepare('INSERT OR REPLACE INTO moves VALUES(?,?,?,?,?,?,?)'),
  iO = db.prepare('INSERT OR REPLACE INTO momo VALUES(?,?,?,?,?,?,?)');
const sync = db.transaction((u, b) => {
  for (const t of ['sales', 'moves', 'momo']) db.prepare(`DELETE FROM ${t} WHERE uid=?`).run(u);
  const rv = a => (Array.isArray(a) ? a.slice(0, 50000).reverse() : []);
  rv(b.sales).forEach(x => iS.run(u, str(x.id, 20), str(x.a), num(x.q), num(x.p), num(x.c), str(x.m, 20), dt(x.d)));
  rv(b.moves).forEach(x => iM.run(u, str(x.id, 20), x.t === 'in' ? 'in' : 'out', str(x.k, 60), num(x.a), str(x.n), dt(x.d)));
  rv(b.momo).forEach(x => iO.run(u, str(x.id, 20), str(x.o, 20), num(x.a), str(x.r), dt(x.d), x.s ? str(x.s, 20) : null));
});
app.post('/api/sync', guard, (q, s) => { sync(q.uid, q.body || {}); s.json({ ok: true }); });

app.post('/api/sms/parse', guard, (q, s) => s.json({ items: parseSms((q.body || {}).text || '') }));
app.get('/api/report', guard, (q, s) => s.json(report(q.uid, per(q.query.per))));

app.get('/api/export.xlsx', guard, async (q, s) => {
  const p = per(q.query.per), d = data(q.uid, p), r = report(q.uid, p), wb = new ExcelJS.Workbook();
  const sheet = (n, cols, rows) => {
    const w = wb.addWorksheet(n);
    w.columns = cols.map(([h, k, wd, f]) => ({ header: h, key: k, width: wd, style: f ? { numFmt: '#,##0' } : {} }));
    w.addRows(rows); w.getRow(1).font = { bold: true };
  };
  sheet('Résumé', [['Indicateur', 'k', 34], ['Montant (F)', 'v', 18, 1]], [
    { k: 'Depuis le', v: r.depuis }, { k: 'Ventes', v: r.ventes }, { k: 'Autres entrées', v: r.autres_entrees },
    { k: 'Coût des produits vendus', v: r.cout_produits }, { k: 'Charges', v: r.charges }, { k: 'Sorties de caisse', v: r.sorties }, { k: 'Bénéfice', v: r.benefice }]);
  sheet('Ventes', [['Date', 'd', 12], ['Article', 'a', 30], ['Quantité', 'q', 10], ['Prix unitaire', 'p', 14, 1], ['Coût unitaire', 'c', 14, 1], ['Total', 't', 14, 1], ['Paiement', 'm', 14]], d.sl.map(x => ({ ...x, t: x.q * x.p })));
  sheet('Caisse', [['Date', 'd', 12], ['Type', 't', 10], ['Catégorie', 'k', 24], ['Montant', 'a', 14, 1], ['Note', 'n', 30]], d.mv.map(x => ({ ...x, t: x.t === 'in' ? 'Entrée' : 'Sortie' })));
  sheet('Mobile Money', [['Date', 'd', 12], ['Opérateur', 'o', 14], ['Montant', 'a', 14, 1], ['Référence', 'r', 34], ['Rapproché', 's', 12]], d.mo.map(x => ({ ...x, s: x.s ? 'Oui' : 'Non' })));
  s.set({ 'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'Content-Disposition': `attachment; filename="akwe-${p}.xlsx"` });
  await wb.xlsx.write(s); s.end();
});
app.get('/api/export.pdf', guard, (q, s) => {
  const p = per(q.query.per), d = data(q.uid, p), r = report(q.uid, p), u = db.prepare('SELECT name FROM users WHERE id=?').get(q.uid);
  const F = n => Math.round(n).toLocaleString('fr-FR').replace(/[\u202f\u00a0]/g, ' ') + ' F', D = x => x.split('-').reverse().join('/');
  const doc = new PDF({ margin: 40, size: 'A4' });
  s.set({ 'Content-Type': 'application/pdf', 'Content-Disposition': `attachment; filename="akwe-${p}.pdf"` });
  doc.pipe(s);
  doc.fontSize(20).text('Rapport Akwè Caisse');
  doc.fontSize(10).fillColor('#555').text(`${u.name || ''} - du ${D(d.from)} au ${D(today())}`).moveDown();
  doc.fillColor('#000').fontSize(12);
  [['Ventes', r.ventes], ['Autres entrées', r.autres_entrees], ['Coût des produits vendus', r.cout_produits], ['Charges', r.charges], ['Sorties de caisse', r.sorties]].forEach(([k, v]) => doc.text(`${k} : ${F(v)}`));
  doc.font('Helvetica-Bold').text(`Bénéfice : ${F(r.benefice)}`).font('Helvetica').moveDown();
  const sec = (t, rows) => { doc.font('Helvetica-Bold').fontSize(13).text(t).font('Helvetica').fontSize(10); rows.length ? rows.slice(0, 400).forEach(x => doc.text(x)) : doc.text('Aucune donnée.'); doc.moveDown(); };
  sec('Ventes', d.sl.map(x => `${D(x.d)}  ${x.a} x${x.q}  ${F(x.q * x.p)}  (${x.m})`));
  sec('Caisse', d.mv.map(x => `${D(x.d)}  ${x.t === 'in' ? '+' : '-'}${F(x.a)}  ${x.k}${x.n ? ' - ' + x.n : ''}`));
  sec('Mobile Money', d.mo.map(x => `${D(x.d)}  ${F(x.a)}  ${x.o}  ${x.r || ''}  ${x.s ? '[rapproché]' : '[non rapproché]'}`));
  doc.end();
});
app.delete('/api/account', guard, (q, s) => {
  const u = db.prepare('SELECT hash FROM users WHERE id=?').get(q.uid);
  if (!bcrypt.compareSync(String((q.body || {}).password || ''), u.hash)) return s.status(403).json({ error: 'Mot de passe incorrect.' });
  db.transaction(() => { for (const t of ['sales', 'moves', 'momo', 'resets']) db.prepare(`DELETE FROM ${t} WHERE uid=?`).run(q.uid); db.prepare('DELETE FROM users WHERE id=?').run(q.uid); })();
  s.json({ ok: true });
});

app.listen(process.env.PORT || 3000, () => console.log('Akwè Caisse prêt sur le port ' + (process.env.PORT || 3000)));
