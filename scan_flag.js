'use strict';
// Scan DMS list attachments and flag (star) INBOX emails whose lists contain
// addresses in Sergiev Posad, Khotkovo, Krasnozavodsk, Peresvet,
// or Sergievo-Posadsky district.
//
// Passwords for encrypted archives are taken from:
//   1) another email with the same sender + subject (e.g. Sberbank),
//   2) another email from the same sender,
//   3) a password newsletter indexed across all mail folders (e.g. Zetta).
//
// Env:
//   IMAP_USER, IMAP_PASS        credentials (required)
//   IMAP_HOST (default imap.rambler.ru), IMAP_PORT (993)
//   IMAP_LIMIT                  max newest INBOX messages to inspect (0 = all)
//   DRY_RUN=1                   do not set flags, do not update state
//   RESCAN=1                    ignore scan state
const { ImapFlow } = require('imapflow');
const { simpleParser } = require('mailparser');
const XLSX = require('xlsx');
const fs = require('fs');
const os = require('os');
const path = require('path');
const { execFileSync } = require('child_process');

const SEVENZIP = 'C:/Program Files/7-Zip/7z.exe';
const STATE_FILE = path.join(__dirname, 'scan_state.json');
const REPORT_FILE = path.join(__dirname, 'last_scan.json');
const CONFIG_FILE = path.join(__dirname, 'config.json');
const LOG_FILE = path.join(__dirname, 'scan.log');
const LOCK_FILE = path.join(__dirname, 'scan.lock');

let CONFIG = {};
try {
  CONFIG = JSON.parse(fs.readFileSync(CONFIG_FILE, 'utf8'));
} catch {
  /* config is optional when IMAP_USER / IMAP_PASS are provided */
}
const DRY_RUN = process.env.DRY_RUN === '1';
const RESCAN = process.env.RESCAN === '1';
// --recheck-no-address: re-evaluate only the mails previously left as 'no_address'
// (they were stored before the policy-number rule existed).
const RECHECK = process.argv.includes('--recheck-no-address');
const LIMIT = process.env.IMAP_LIMIT ? Number(process.env.IMAP_LIMIT) : 0;
const MAX_RETRIES = 5;

function logLine(text) {
  try {
    fs.appendFileSync(LOG_FILE, `[${new Date().toISOString()}] ${text}\n`);
  } catch {
    /* ignore logging errors */
  }
}

function lockOwnerAlive() {
  try {
    const pid = Number(String(fs.readFileSync(LOCK_FILE, 'utf8')).trim());
    if (!pid) return true; // unknown owner: fall back to the age check
    try {
      process.kill(pid, 0);
      return true;
    } catch (e) {
      return e.code === 'EPERM'; // alive, just owned by another user
    }
  } catch {
    return true;
  }
}

function acquireLock() {
  try {
    const fd = fs.openSync(LOCK_FILE, 'wx');
    fs.writeSync(fd, String(process.pid));
    fs.closeSync(fd);
    return true;
  } catch (e) {
    if (e.code !== 'EEXIST') throw e;
    try {
      const age = Date.now() - fs.statSync(LOCK_FILE).mtimeMs;
      // A crashed run leaves the lock behind; do not make the next 20 minutes wait.
      if (!lockOwnerAlive()) {
        logLine('removing lock of a dead run');
        fs.unlinkSync(LOCK_FILE);
        return acquireLock();
      }
      if (age > 20 * 60 * 1000) {
        fs.unlinkSync(LOCK_FILE); // hung run
        return acquireLock();
      }
    } catch {
      /* ignore */
    }
    return false;
  }
}

function releaseLock() {
  try {
    fs.unlinkSync(LOCK_FILE);
  } catch {
    /* ignore */
  }
}

const TARGETS = [
  // город и район: «г. Сергиев Посад», «в Сергиевом Посаде», «Сергиево-Посадский р-н»
  {
    name: 'Сергиев Посад / Сергиево-Посадский',
    // город в любой падежной форме, либо районный adjective + «район/округ»;
    // «Сергиево-Посадская улица» в другом городе сюда не попадает
    re: /сергиев[а-я]*[\s-]*посад(?!ск)|сергиев[а-я]*[\s-]*посадск[а-я]*\s+(?:район|р-н|округ|городск|муницип)/,
  },
  // Хотьково; исключаем прилагательные «Хотьковская/Хотьковский»
  { name: 'Хотьково', re: /хотьков(?!ск)/ },
  // Краснозаводск; исключаем «Краснозаводская улица» и подобные
  { name: 'Краснозаводск', re: /краснозаводск(?!ая|ий|ое|ой|ую|им|их)/ },
  // Пересвет; исключаем фамилию «Пересветов»
  { name: 'Пересвет', re: /пересвет(?!ов)/ },
];

// Wide rule: any settlement of the Sergievo-Posad district counts as ours, but only
// when the same address cell also carries a 1413xx index or names the district —
// village names repeat in other districts, this guard keeps false hits out.
const SETTLEMENTS_FILE = path.join(__dirname, 'sergiev_posad_settlements.json');
let DISTRICT_SETTLEMENTS = [];
try {
  DISTRICT_SETTLEMENTS = JSON.parse(fs.readFileSync(SETTLEMENTS_FILE, 'utf8')).settlements.map((s) => s.name);
} catch {
  /* list is optional */
}
const escapeRe = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
const DISTRICT_SETTLEMENT_RE = DISTRICT_SETTLEMENTS.length
  ? new RegExp(
    '(?:^|[^а-яa-zё-])('
      + DISTRICT_SETTLEMENTS.map((n) => escapeRe(String(n).toLowerCase().replace(/ё/g, 'е'))).join('|')
      + ')(?![а-яa-zё-])',
    'i'
  )
  : null;
const DISTRICT_INDEX_RE = /(?:^|[^0-9])1413\d\d(?:[^0-9]|$)/;
const DISTRICT_NAME_RE = /сергиев[а-я]*[\s-]*посадск|сергиевопосадск/;

// A bare settlement name is ambiguous: «д Гаврилково» exists both in the
// Sergievo-Posad district and in Tver oblast, «Богородское» is also a Moscow
// city district. So a settlement hit only counts when the same cell does not
// name another region, another city district or Moscow itself.
const OTHER_REGION_RE = /(?:^|[\s,])([а-яё-]+)\s+(?:обл\.?|область|край)(?![а-я])/i;
const REPUBLIC_RE = /(?:^|[\s,])республика\s+([а-яё-]+)/i;
const MOSCOW_CITY_RE = /(?:^|[\s,])(?:город\s+)?москва(?:\s+г\.?)?(?:[\s,]|$)/i;
const OKRUG_RE = /(?:г\.\s?о\.|городской\s+округ|муниципальный\s+округ)\s+([а-яё][а-яё-]*)/gi;
const CITY_RE = /(?:^|[\s,])г\.?\s+([а-яё][а-яё-]*)/gi;
const OUR_CITY_RE = /^(сергиев|посад|хотьков|краснозаводск|пересвет)/;
const DISTRICT_SETTLEMENT_SET = new Set(
  DISTRICT_SETTLEMENTS.map((n) => String(n).toLowerCase().replace(/ё/g, 'е'))
);
function geographyElsewhere(s) {
  const region = s.match(OTHER_REGION_RE);
  if (region && !/^московск/.test(normText(region[1]))) return true;
  if (REPUBLIC_RE.test(s)) return true;
  if (MOSCOW_CITY_RE.test(s)) return true;
  const okrugs = [...s.matchAll(OKRUG_RE)]
    .map((m) => normText(m[1]))
    .filter((n) => !/^(муниципальн|городск)/.test(n));
  if (okrugs.length && !okrugs.some((n) => /сергиев/.test(n))) return true;
  const cities = [...s.matchAll(CITY_RE)].map((m) => normText(m[1]));
  if (cities.some((c) => !OUR_CITY_RE.test(c) && !DISTRICT_SETTLEMENT_SET.has(c))) return true;
  return false;
}

const SPREADSHEET_RE = /\.(xlsx|xlsm|xls)$/i;
const ARCHIVE_RE = /\.(zip|rar|7z)$/i;
// Any settlement marker, regardless of the column: tells whether the list carries
// city data at all (a list without a single city is treated as junk).
const SETTLEMENT_ANY_RE = /(^|[\s,;])(г\.|город|пгт|пос\.|посёлок|поселок|с\.|село|рп|г\.о\.|городской округ|москва|россия)/i;
const ADDR_HEADER_RE = /адрес|город|насел|прописк|место жит|address/i;
const PW_SENDER_RE = /parol|password/i;
const PW_SUBJECT_RE = /парол|password/i;
// Guarantee letters are out of scope: «СРОЧНО: ...» emails and ГП*.zip files.
// Data-change letters (new dates/requisites) are not registers either: they are
// parsed and applied by lists/parse_changes.js, so the scanner must not star them.
const SKIP_SUBJECT_RE = /^\s*срочно\b|изменени[ея]\s+(сроков|условий|реквизитов|данных)/i;
const SKIP_ATTACHMENT_RE = /^ГП\d/i;
// Emails with full addresses but no local patients are moved here.
const DEST_FOLDER = 'дмс нет пациентов';
// Only DMS list emails are processed; acts, invoices, price lists are ignored.
const LIST_SUBJECT_RE = /спис|прикрепл|открепл|застрахован|дмс/i;
const LIST_ATTACHMENT_RE = /спис|прикрепл|открепл|застрахован|snyat|прикр|откр/i;

// Sender-based routing (matched against the display name, case-insensitive).
const LAB_ROUTES = [
  { target: 'invitro', label: 'ИНВИТРО', re: /^независимая\s+лаборатория\s+инвитро$/i },
  { target: 'trash', label: 'Корзина (INVITRO LAB)', re: /^invitro\s+lab$/i },
  { target: 'trash', label: 'Корзина (Лаборатория СИТИЛАБ)', re: /^лаборатория\s+ситилаб$/i },
];

// Insurer -> Archimed organization id (DefOrgId from the LoadRisk profile XML).
const ORGIDS = {
  'Ингосстрах': 9, 'Согаз': 5, 'Альфа': 7, 'ВСК': 6, 'Югория': 24,
  'Ренессанс': 15, 'Зетта': 10, 'Росгосстрах': 8, 'Сбербанк страхование': 37,
  'Абсолют': 13, 'Ресо-Гарантия': 14, 'Энергогарант': 16,
  'Регион-Медсервис': 18, 'Согласие': 36,
};
const INSURER_RE = [
  [/ингосстрах|игс_/i, 'Ингосстрах'],
  [/согаз|v8_/i, 'Согаз'],
  [/альфа|8050322/i, 'Альфа'],
  [/югория|ugsk/i, 'Югория'],
  [/ренессанс|renins|П650/i, 'Ренессанс'],
  [/зетта|zetta/i, 'Зетта'],
  [/росгосстрах|rgs/i, 'Росгосстрах'],
  [/сбербанк|sberins/i, 'Сбербанк страхование'],
  [/вск/i, 'ВСК'],
  [/абсолют|absolut/i, 'Абсолют'],
  [/ресо|reso/i, 'Ресо-Гарантия'],
  [/энергогарант/i, 'Энергогарант'],
  [/регион[\s-]*медсервис/i, 'Регион-Медсервис'],
  [/согласие|soglasie/i, 'Согласие'],
];

const OP_OTKR_RE = /откреп|snyat|снят|[_\s]откр/i;
const OP_PRIKR_RE = /прикреп|прикр/i;
// 1 = прикрепление, 2 = открепление, null = неизвестно
function opTypeOf(parsed) {
  const text = [parsed.subject || '', ...(parsed.attachments || []).map((a) => a.filename || '')].join(' ');
  if (OP_OTKR_RE.test(text)) return 2;
  if (OP_PRIKR_RE.test(text)) return 1;
  return null;
}

function detectInsurerOf(parsed) {
  const text = [
    parsed.subject || '',
    ...(parsed.attachments || []).map((a) => a.filename || ''),
    parsed.from?.text || '',
  ].join(' ');
  for (const [re, name] of INSURER_RE) if (re.test(text)) return name;
  return null;
}

// Ask the database whether any of these policies is already known for the insurer.
function policyCheck(insurer, policies) {
  const orgid = ORGIDS[insurer] || 0;
  if (!orgid) return { orgid: 0, total: 0, found: [] };
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'dms-pol-'));
  const file = path.join(dir, 'check.json');
  fs.writeFileSync(file, JSON.stringify({ insurer, orgid, policies }), 'utf8');
  try {
    const out = execFileSync('powershell', ['-NoProfile', '-ExecutionPolicy', 'Bypass',
      '-File', 'C:/Users/Администратор/.claude/tools/db_policy_check.ps1', '-Json', file],
    { encoding: 'utf8' });
    return JSON.parse(out);
  } catch (e) {
    // database unreachable: treat as 'cannot judge' so the letter is never moved by mistake
    return { orgid, total: 0, found: [], error: String(e.message || e).slice(0, 200) };
  } finally {
    try { fs.rmSync(dir, { recursive: true, force: true }); } catch { /* ignore */ }
  }
}

const normText = (v) => String(v).toLowerCase().replace(/ё/g, 'е');
const normSubject = (s) => normText(s).replace(/^(re|fw|fwd)\s*:\s*/g, '').replace(/\s+/g, ' ').trim();
const senderOf = (parsed) => (parsed.from?.value?.[0]?.address || '').toLowerCase();
const domainOf = (parsed) => {
  const a = senderOf(parsed);
  return a.includes('@') ? a.split('@')[1] : '';
};

function htmlToText(html) {
  return String(html)
    .replace(/<(style|script)[\s\S]*?<\/\1>/gi, ' ')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/g, ' ')
    .replace(/&quot;/g, '"')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/\s+/g, ' ');
}

function extractPassword(text) {
  const src = String(text || '');
  let m = src.match(/(?:пароль|password)[^:：\n]{0,200}[:：]\s*(\S+)/i);
  if (m) return m[1].replace(/[.,;:!?)\]]+$/, '');
  const idx = src.toLowerCase().search(/парол|password/);
  if (idx >= 0) {
    for (const tok of src.slice(idx, idx + 400).split(/\s+/)) {
      const t = tok.replace(/^[.,;:!?(\[]+|[.,;:!?)\]]+$/g, '');
      if (/^[\x21-\x7E]{6,24}$/.test(t) && /[A-Za-z]/.test(t) && /[0-9]/.test(t)) return t;
    }
  }
  return null;
}

function routeFor(parsed) {
  const from = parsed.from?.value?.[0];
  const name = normText(from?.name || '').replace(/\s+/g, ' ').trim();
  for (const r of LAB_ROUTES) if (r.re.test(name)) return r;
  return null;
}

function matchLocalities(value) {
  const s = normText(value);
  const hits = [];
  for (const t of TARGETS) if (t.re.test(s)) hits.push(t.name);
  // «1413xx» — indexes of the Sergievo-Posad district only. Ингосстрах often writes
  // «Московская Обл, О. Сергиево-Посадский, Ул Воробьевская» without naming a city.
  if (DISTRICT_INDEX_RE.test(s)) hits.push('Индекс 1413xx');
  // wide rule: any settlement of the district counts, but only when the cell
  // does not place it in another region / city district (homonymous names).
  if (DISTRICT_SETTLEMENT_RE) {
    const m = s.match(DISTRICT_SETTLEMENT_RE);
    const explicit = hits.length > 0 || DISTRICT_NAME_RE.test(s);
    if (m && (explicit || !geographyElsewhere(s))) hits.push('Нас. пункт: ' + m[1]);
  }
  return hits;
}

// True when the text names a settlement or a city district, e.g.
// «г. Пушкино», «пгт Скоропусковский», «с. Муханово», «д Акулово»,
// «Москва», «г.о. Пушкинский». A house number («д. 25») does not count.
function hasSettlement(value) {
  const s = normText(value);
  if (!s.trim()) return false;
  if (/(?:^|[,\s])москва(?:[,\s]|$)/.test(s)) return true;
  // no settlement marker, just «, БРЯНСК, КРАХМАЛЕВА, д.33 кв.220» (Renaissance):
  // city, street and house are all required, so placeholder «., ., ., д..» is not an address
  if (/(?:^|[,\s])[а-яa-z][а-яa-z-]{2,}\s*,\s*[^\s,][^,]{2,}\s*,\s*(?:д\.|дом)(?![а-яa-z])/.test(s)) return true;
  // postfix form used by some insurers: «Пушкино г», «Сергиев Посад г.», «Хотьково пгт»
  if (/(?:^|[,\s])[а-яa-z][а-яa-z-]*\s+(?:г|гор|пгт|рп)(?:\.|,|\s|$)/.test(s)) return true;
  if (/(?:^|[,\s])(?:городской\s+округ|г\.\s?о\.)/.test(s)) return true;
  if (
    /(?:^|[,\s])(?:город|гор|пгт|рп|р\.\s?п\.?|пос|поселок|сел|село|дер|деревня|станица|хутор|аул|мкр|микрорайон)(?![а-яa-z])/.test(
      s
    )
  ) {
    return true;
  }
  // single-letter markers: «г. X», «г X», «с. X», «д X», «п. X», «х. X»
  if (/(?:^|[,\s])[гсдхп]\s*\.\s*[а-яa-z]/.test(s)) return true;
  if (/(?:^|[,\s])[гсдхп]\s+[а-яa-z]/.test(s)) return true;
  return false;
}

// Returns { hits, fullAddresses, noAddressSheets }.
// A cell counts as a full address when its column is «Город/Населённый пункт»
// or the value itself names a settlement/district.
// A real table header, not a letterhead line like «... Адрес: Россия, ...».
function isAddrHeader(text) {
  const t = String(text ?? '').replace(/\s+/g, ' ').trim();
  if (!t || t.length > 40) return false;
  if (t.endsWith(':')) return false;
  return ADDR_HEADER_RE.test(t);
}

// Policy numbers of a sheet: a header cell with «полис» (but not «серия полиса»)
// and the values below it. Used when a detachment list has no addresses at all.
function collectPolicies(rows) {
  const out = new Set();
  for (let r = 0; r < Math.min(rows.length, 80); r++) {
    const row = rows[r] || [];
    for (let c = 0; c < row.length; c++) {
      const header = String(row[c] || '');
      if (!/полис/i.test(header) || /серия/i.test(header)) continue;
      for (let i = r + 1; i < rows.length; i++) {
        const v = String((rows[i] || [])[c] ?? '').trim();
        if (!v || /полис/i.test(v)) continue;
        if (v.length < 3 || v.length > 40) continue;
        out.add(v);
      }
    }
  }
  return out;
}

function scanWorkbook(buffer) {
  const wb = XLSX.read(buffer, { type: 'buffer' });
  const hits = [];
  const noAddressSheets = [];
  const policies = new Set();
  let fullAddresses = 0;
  let cityCells = 0;

  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    if (!ws) continue;
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '' });
    const hasData = rows.some((row) => (row || []).some((cell) => String(cell ?? '').trim() !== ''));
    if (!hasData) continue;

    let headerRow = -1;
    let addrCols = [];
    for (let r = 0; r < Math.min(rows.length, 80); r++) {
      const cols = [];
      (rows[r] || []).forEach((cell, c) => {
        const header = String(cell || '');
        if (isAddrHeader(header)) cols.push({ index: c, isLocality: /город|насел/i.test(header) });
      });
      if (cols.length) {
        headerRow = r;
        addrCols = cols;
        break;
      }
    }

    for (const p of collectPolicies(rows)) policies.add(p);
    for (const row of rows) {
      for (const cell of row || []) {
        const v = String(cell ?? "");
        if (v.length > 2 && SETTLEMENT_ANY_RE.test(v) && !/адрес|город|насел|место жит/i.test(v)) cityCells++;
      }
    }

    if (headerRow < 0) {
      // No address column: count target strings anywhere so misses stay visible.
      let auditCount = 0;
      const auditSamples = [];
      for (const row of rows) {
        for (const cell of row || []) {
          if (matchLocalities(cell).length) {
            auditCount++;
            if (auditSamples.length < 2) auditSamples.push(String(cell).trim().slice(0, 160));
          }
        }
      }
      noAddressSheets.push({ sheet: sheetName, auditCount, auditSamples });
      continue;
    }

    const checkCell = (value, r, col) => {
      if (value === null || value === undefined || String(value).trim() === '') return;
      const text = String(value);
      if (col.isLocality || hasSettlement(text)) fullAddresses++;
      const localities = matchLocalities(text);
      if (localities.length) {
        hits.push({ sheet: sheetName, row: r + 1, col: col.index + 1, value: text.trim(), localities });
      }
    };

    for (let r = headerRow + 1; r < rows.length; r++) {
      for (const col of addrCols) checkCell((rows[r] || [])[col.index], r, col);
    }
  }

  return { hits, fullAddresses, noAddressSheets, cityCells, policies: [...policies] };
}

function listFiles(dir) {
  return fs
    .readdirSync(dir, { recursive: true })
    .map((f) => path.join(dir, String(f)))
    .filter((f) => {
      try {
        return fs.statSync(f).isFile();
      } catch {
        return false;
      }
    });
}

function extractArchive(archivePath, password) {
  const dest = fs.mkdtempSync(path.join(os.tmpdir(), 'dms-'));
  const args = ['x', '-y', '-sccUTF-8', `-o${dest}`, `-p${password || ''}`, archivePath];
  try {
    execFileSync(SEVENZIP, args, { encoding: 'utf8', stdio: 'pipe' });
    return { dest, ok: true };
  } catch (e) {
    return { dest, ok: false, error: `${e.stdout || ''}\n${e.stderr || ''}`.trim() || e.message };
  }
}

// Folders that must never be opened: ХЕЛИКС is a laboratory, it only sends
// newsletters and has no DMS registers, so the scan leaves it completely alone.
const SKIP_FOLDERS = new Set(['хеликс']);

// Password newsletters from all folders except INBOX.
async function buildPasswordIndex(client) {
  const boxes = await client.list();
  const candidates = [];
  for (const box of boxes) {
    if (box.path === 'INBOX') continue;
    if (SKIP_FOLDERS.has(String(box.path).toLowerCase())) continue;
    let lock = null;
    try {
      lock = await client.getMailboxLock(box.path);
      const total = client.mailbox.exists;
      if (!total) continue;
      const start = Math.max(1, total - 100 + 1);
      const uids = [];
      for await (const msg of client.fetch(`${start}:*`, { envelope: true })) {
        const subject = msg.envelope?.subject || '';
        const from = (msg.envelope?.from || []).map((a) => a.address || '').join(',').toLowerCase();
        if (PW_SENDER_RE.test(from) || PW_SUBJECT_RE.test(subject)) uids.push(msg.uid);
      }
      for (const uid of uids) {
        const msg = await client.fetchOne(String(uid), { source: true }, { uid: true });
        if (!msg || !msg.source) continue;
        const p = await simpleParser(msg.source);
        const password = extractPassword(`${p.text || ''}\n${p.html ? htmlToText(p.html) : ''}`);
        if (password) {
          candidates.push({
            sender: senderOf(p),
            domain: domainOf(p),
            subject: normSubject(p.subject),
            date: p.date ? p.date.getTime() : 0,
            password,
            source: `${box.path}: ${p.subject || ''}`,
          });
        }
      }
    } catch {
      /* skip unreadable folder */
    } finally {
      if (lock) lock.release();
    }
  }
  return candidates;
}

// Ordered password candidates: the letter's own body first, then its own subject's
// newest letters, then the sender, then the indexed newsletters. Insurance companies
// rotate the password inside the same subject («Список ЗЛ. Вега-М»), so a single
// guess is not enough — the caller tries them in order until 7-Zip accepts one.
function findPasswords(mail, allMails, index) {
  const wantSender = senderOf(mail);
  const wantSubject = normSubject(mail.subject);
  const readPass = (m) => extractPassword(`${m.text || ''}\n${m.html ? htmlToText(m.html) : ''}`);
  const out = [];
  const seen = new Set();
  const add = (password, from) => {
    if (!password || seen.has(password)) return;
    seen.add(password);
    out.push({ password, from });
  };
  const newestFirst = (list) =>
    [...list].sort((a, b) => (b.date ? b.date.getTime() : 0) - (a.date ? a.date.getTime() : 0));

  // 0) the letter itself
  add(readPass(mail), `${mail.subject} (это письмо)`);

  // 1) same sender + same subject (Sberbank pattern)
  for (const m of newestFirst(allMails)) {
    if (senderOf(m) !== wantSender || normSubject(m.subject) !== wantSubject) continue;
    add(readPass(m), `${m.subject} (то же письмо, тема)`);
  }
  // 2) same sender, any subject, in INBOX
  for (const m of newestFirst(allMails)) {
    if (senderOf(m) !== wantSender) continue;
    add(readPass(m), m.subject);
  }
  // 3) indexed password newsletters: same sender, else same domain, newest before the email
  const target = mail.date ? mail.date.getTime() : Date.now();
  const bySender = index.filter((c) => c.sender === wantSender);
  const byDomain = index.filter((c) => domainOf(mail) && c.domain === domainOf(mail));
  for (const pool of [bySender, byDomain]) {
    if (!pool.length) continue;
    const before = pool.filter((c) => c.date <= target + 86400000).sort((a, b) => b.date - a.date);
    if (before.length) {
      for (const c of before) add(c.password, c.source);
      continue;
    }
    const after = pool.sort((a, b) => a.date - b.date);
    if (after.length && after[0].date - target < 40 * 86400000) add(after[0].password, after[0].source);
  }
  return out;
}

function noAddressWarning(label, miss) {
  const extra =
    miss.auditCount > 0
      ? ` (ВНИМАНИЕ: целевые строки на листе есть — ${miss.auditCount}: ${miss.auditSamples.join(' || ')})`
      : '';
  return `${label}: лист «${miss.sheet}» — нет колонки адреса, пропущен${extra}`;
}

function scanMailAttachments(parsed, allMails, pwIndex, tempDirs) {
  const results = [];
  const errors = [];
  const warnings = [];
  const archives = [];
  const found = { passwords: [], fullAddresses: 0, cityCells: 0, policies: new Set() };

  for (const att of parsed.attachments || []) {
    const name = att.filename || '';
    if (SKIP_ATTACHMENT_RE.test(name)) continue; // guarantee letter
    if (SPREADSHEET_RE.test(name)) {
      try {
        const wb = scanWorkbook(att.content);
        found.fullAddresses += wb.fullAddresses;
        found.cityCells += wb.cityCells;
        for (const p of wb.policies) found.policies.add(p);
        for (const hit of wb.hits) results.push({ file: name, ...hit });
        for (const miss of wb.noAddressSheets) warnings.push(noAddressWarning(name, miss));
      } catch (e) {
        errors.push(`${name}: ${e.message}`);
      }
    } else if (ARCHIVE_RE.test(name)) {
      archives.push(att);
    }
  }

  for (const att of archives) {
    const name = att.filename || '';
    const work = fs.mkdtempSync(path.join(os.tmpdir(), 'dms-att-'));
    tempDirs.push(work);
    const archivePath = path.join(work, name.replace(/[\\/]/g, '_'));
    fs.writeFileSync(archivePath, att.content);
    const candidates = findPasswords(parsed, allMails, pwIndex);
    const attempts = candidates.length ? candidates : [{ password: '', from: null }];
    let ex = null;
    let pw = null;
    for (const cand of attempts) {
      ex = extractArchive(archivePath, cand.password);
      tempDirs.push(ex.dest);
      if (ex.ok) {
        pw = cand;
        break;
      }
    }
    if (!ex.ok) {
      const reason =
        (ex.error || '')
          .split(/\r?\n/)
          .map((l) => l.trim())
          .filter((l) => /ERROR|Cannot|Wrong|Data Error|Encrypted/i.test(l))
          .slice(0, 1)
          .join('') || 'ошибка распаковки';
      const tried = attempts.filter((c) => c.password).map((c) => `«${c.password}»`).join(', ');
      errors.push(`${name}: пароль ${tried || 'не найден'} — ${reason}`);
      continue;
    }
    if (pw && pw.password) found.passwords.push({ file: name, password: pw.password, source: pw.from });
    for (const file of listFiles(ex.dest)) {
      if (!SPREADSHEET_RE.test(file)) continue;
      try {
        const wb = scanWorkbook(fs.readFileSync(file));
        found.fullAddresses += wb.fullAddresses;
        found.cityCells += wb.cityCells;
        for (const p of wb.policies) found.policies.add(p);
        for (const hit of wb.hits) results.push({ file: `${name}/${path.basename(file)}`, ...hit });
        for (const miss of wb.noAddressSheets) {
          warnings.push(noAddressWarning(`${name}/${path.basename(file)}`, miss));
        }
      } catch (e) {
        errors.push(`${name}/${path.basename(file)}: ${e.message}`);
      }
    }
  }

  return { results, errors, warnings, passwords: found.passwords, fullAddresses: found.fullAddresses, policies: [...found.policies], cityCells: found.cityCells };
}

async function ensureFolder(client, name) {
  const wanted = name.toLowerCase();
  const boxes = await client.list();
  const hit = boxes.find(
    (b) => (b.path || '').toLowerCase() === wanted || (b.name || '').toLowerCase() === wanted
  );
  if (hit) return hit.path;
  await client.mailboxCreate(name);
  logLine(`created folder: ${name}`);
  return name;
}

function loadState() {
  try {
    const s = JSON.parse(fs.readFileSync(STATE_FILE, 'utf8'));
    if (!s.scanned) s.scanned = {};
    if (!s.retry) s.retry = {};
    return s;
  } catch {
    return { uidValidity: null, scanned: {}, retry: {} };
  }
}

async function run() {
  const client = new ImapFlow({
    host: process.env.IMAP_HOST || 'imap.rambler.ru',
    port: Number(process.env.IMAP_PORT || 993),
    secure: true,
    auth: {
      user: process.env.IMAP_USER || CONFIG.user,
      pass: process.env.IMAP_PASS || CONFIG.password,
    },
    logger: false,
  });
  // Without a listener a dropped TLS connection (write ECONNABORTED) is an
  // unhandled 'error' event that kills node before the lock is released.
  client.on('error', (e) => {
    logLine(`IMAP error: ${e.code || ''} ${e.message}`.trim());
  });

  const state = loadState();
  const tempDirs = [];
  const report = {
    dryRun: DRY_RUN,
    checked: 0,
    flagged: [],
    moved: [],
    routed: [],
    noFull: 0,
    noFullList: [],
    policyFlagged: [],
    trashed: [],
    policyMissed: [],
    skipped: 0,
    passwordsUsed: [],
    errors: [],
    warnings: [],
    scannedAt: new Date().toISOString(),
  };

  await client.connect();
  // Index password newsletters first (opens other folders, releases them).
  const pwIndex = await buildPasswordIndex(client);

  let destPath = null;
  let trashPath = 'Trash';
  try {
    const boxes = await client.list();
    const tbox = boxes.find((b) => (b.specialUse || '') === '\Trash');
    if (tbox) trashPath = tbox.path;
  } catch { /* keep default */ }
  const routePaths = { invitro: null, trash: null };
  if (!DRY_RUN) {
    try {
      destPath = await ensureFolder(client, DEST_FOLDER);
    } catch (e) {
      logLine(`warn: cannot create folder ${DEST_FOLDER}: ${e.message}`);
    }
    try {
      routePaths.invitro = await ensureFolder(client, 'ИНВИТРО');
    } catch (e) {
      logLine(`warn: folder ИНВИТРО: ${e.message}`);
    }
    try {
      const boxes = await client.list();
      const trash =
        boxes.find((b) => b.specialUse === '\\Trash') ||
        boxes.find((b) => (b.path || '').toLowerCase() === 'trash');
      routePaths.trash = trash ? trash.path : null;
    } catch (e) {
      logLine(`warn: trash folder: ${e.message}`);
    }
  }

  const lock = await client.getMailboxLock('INBOX');
  try {
    const uidValidity = String(client.mailbox.uidValidity);
    if (state.uidValidity !== uidValidity) state.scanned = {};
    state.uidValidity = uidValidity;

    const total = client.mailbox.exists;
    const start = LIMIT > 0 ? Math.max(1, total - LIMIT + 1) : 1;
    const mails = [];
    for await (const msg of client.fetch(`${start}:*`, { source: true, envelope: true })) {
      if (!msg || !msg.source) {
        logLine(`warn: no message source for uid ${msg && msg.uid}`);
        continue;
      }
      const parsed = await simpleParser(msg.source);
      mails.push({ uid: msg.uid, seq: msg.seq, parsed });
    }

    const allParsed = mails.map((m) => m.parsed);
    for (const mail of mails) {
      const parsed = mail.parsed;

      // Sender-based routing (Invitro / Citilab) takes precedence over list logic.
      const route = routeFor(parsed);
      if (route) {
        const info = {
          uid: mail.uid,
          date: parsed.date ? parsed.date.toISOString() : null,
          from: parsed.from?.text || '',
          subject: parsed.subject || '',
          destination: route.label,
        };
        if (DRY_RUN) {
          report.routed.push(info);
          continue;
        }
        const targetPath = routePaths[route.target];
        if (!targetPath) {
          report.errors.push({ subject: parsed.subject, error: `папка «${route.label}» недоступна` });
          continue;
        }
        try {
          await client.messageMove(String(mail.uid), targetPath, { uid: true });
          report.routed.push(info);
          state.scanned[mail.uid] = { status: 'routed', destination: route.label, at: report.scannedAt };
          delete state.retry[mail.uid];
        } catch (e) {
          report.errors.push({
            subject: parsed.subject,
            error: `не перемещено в «${route.label}»: ${e.message}`,
          });
        }
        continue;
      }

      const prevState = state.scanned[mail.uid];
      const rechecking = RECHECK && prevState && prevState.status === 'no_address';
      if (RECHECK && !rechecking) { report.skipped++; continue; }
      if (!RESCAN && prevState && !rechecking) {
        report.skipped++;
        continue;
      }
      if (SKIP_SUBJECT_RE.test(parsed.subject || '')) {
        report.skipped++;
        continue; // guarantee letter, not a list
      }
      const hasRelevantAttachment = (parsed.attachments || []).some(
        (a) => SPREADSHEET_RE.test(a.filename || '') || ARCHIVE_RE.test(a.filename || '')
      );
      if (!hasRelevantAttachment) continue;
      const looksLikeList =
        LIST_SUBJECT_RE.test(parsed.subject || '') ||
        (parsed.attachments || []).some((a) => LIST_ATTACHMENT_RE.test(a.filename || ''));
      if (!looksLikeList) continue;

      report.checked++;
      const { results, errors, warnings, passwords, fullAddresses, policies, cityCells } = scanMailAttachments(
        parsed,
        allParsed,
        pwIndex,
        tempDirs
      );
      for (const e of errors) report.errors.push({ subject: parsed.subject, error: e });
      for (const w of warnings) report.warnings.push({ subject: parsed.subject, warning: w });
      for (const p of passwords) report.passwordsUsed.push({ subject: parsed.subject, ...p });

      const localities = [...new Set(results.flatMap((r) => r.localities))];

      // 1) Local patients found: star the email and keep it in INBOX.
      if (localities.length) {
        const info = {
          uid: mail.uid,
          date: parsed.date ? parsed.date.toISOString() : null,
          from: parsed.from?.text || '',
          subject: parsed.subject || '',
          localities,
          matches: results.slice(0, 5),
          matchCount: results.length,
          flagged: false,
        };
        if (!DRY_RUN) {
          await client.messageFlagsAdd(mail.uid, ['\\Flagged'], { uid: true });
          info.flagged = true;
        }
        report.flagged.push(info);
        if (!DRY_RUN) {
          state.scanned[mail.uid] = { status: 'flagged', flagged: true, at: report.scannedAt };
          delete state.retry[mail.uid];
        }
        continue;
      }

      // 2) Could not read every attachment: leave the email alone and retry later.
      if (errors.length) {
        report.skipped++;
        if (!DRY_RUN) {
          if ((state.retry[mail.uid] || 0) + 1 < MAX_RETRIES) {
            state.retry[mail.uid] = (state.retry[mail.uid] || 0) + 1;
          } else {
            state.scanned[mail.uid] = { status: 'error', at: report.scannedAt };
            delete state.retry[mail.uid];
          }
        }
        continue;
      }

      // 3) No complete address anywhere (typical for detachment lists): decide by
      //    policy numbers. Ours -> star it; foreign -> move out; no policies -> leave.
      if (fullAddresses === 0) {
        // Прикрепление without a single city: the addresses are junk, the list is useless.
        if (opTypeOf(parsed) === 1 && cityCells === 0) {
          const info = { uid: mail.uid, subject: parsed.subject || '' };
          if (!DRY_RUN) {
            try {
              await client.messageMove(String(mail.uid), trashPath, { uid: true });
              state.scanned[mail.uid] = { status: 'trashed_no_city', at: report.scannedAt };
              delete state.retry[mail.uid];
            } catch (e) {
              report.errors.push({ subject: parsed.subject, error: `не перемещено в корзину: ${e.message}` });
              continue;
            }
          }
          report.trashed.push(info);
          continue;
        }
        const insurer = detectInsurerOf(parsed);
        const check = insurer && policies.length
          ? policyCheck(insurer, policies)
          : { orgid: 0, total: 0, found: [] };
        if (check.found.length) {
          if (!DRY_RUN) {
            await client.messageFlagsAdd(mail.uid, ['\\Flagged'], { uid: true });
            state.scanned[mail.uid] = { status: 'flagged_by_policy', at: report.scannedAt };
            delete state.retry[mail.uid];
          }
          report.policyFlagged.push({
            uid: mail.uid,
            subject: parsed.subject || '',
            insurer,
            policies: policies.length,
            found: check.found.length,
          });
          continue;
        }
        if (check.total === 0) {
          report.skipped++;
          report.noFull++;
          report.noFullList.push({
            subject: parsed.subject || '(без темы)',
            from: parsed.from?.text || '',
          });
          if (!DRY_RUN) state.scanned[mail.uid] = { status: 'no_address', at: report.scannedAt };
          continue;
        }
        // policies found but none belong to us -> fall through to the move below
        report.policyMissed.push({
          uid: mail.uid,
          subject: parsed.subject || '',
          insurer,
          policies: policies.length,
        });
      }

      // 4) Full addresses present, no local patients: move out of INBOX.
      const moveInfo = {
        uid: mail.uid,
        date: parsed.date ? parsed.date.toISOString() : null,
        from: parsed.from?.text || '',
        subject: parsed.subject || '',
        fullAddresses,
        moved: false,
      };
      if (DRY_RUN) {
        report.moved.push(moveInfo);
        continue;
      }
      if (!destPath) {
        report.errors.push({
          subject: parsed.subject,
          error: `папка «${DEST_FOLDER}» недоступна — письмо оставлено в INBOX`,
        });
        continue;
      }
      try {
        await client.messageMove(String(mail.uid), destPath, { uid: true });
        moveInfo.moved = true;
        report.moved.push(moveInfo);
        state.scanned[mail.uid] = { status: 'moved', at: report.scannedAt };
        delete state.retry[mail.uid];
      } catch (e) {
        report.errors.push({ subject: parsed.subject, error: `не перемещено: ${e.message}` });
      }
    }
  } finally {
    lock.release();
    try { await client.logout(); } catch (e) { logLine('logout: ' + e.message); }
  }

  for (const dir of tempDirs) {
    try {
      fs.rmSync(dir, { recursive: true, force: true });
    } catch {
      /* ignore */
    }
  }

  const toJson = (obj) => JSON.stringify(obj, (k, v) => (typeof v === 'bigint' ? v.toString() : v), 2);
  if (!DRY_RUN) fs.writeFileSync(STATE_FILE, toJson(state));
  fs.writeFileSync(REPORT_FILE, toJson(report));

  console.log(
    `Режим: ${DRY_RUN ? 'ПРОВЕРКА (без флагов и перемещений)' : 'БОЕВОЙ (звёздочки и перемещение)'}`
  );
  console.log(`Проверено писем с файлами списков: ${report.checked}`);
  console.log(`Найдено писем с целевыми адресами: ${report.flagged.length}`);
  for (const f of report.flagged) {
    console.log(`\n★ ${f.date?.slice(0, 16).replace('T', ' ')} | ${f.from}`);
    console.log(`  ${f.subject}`);
    console.log(`  Адреса: ${f.localities.join(', ')} | совпадений: ${f.matchCount}`);
    for (const m of f.matches) console.log(`   - [${m.file}] ${m.value}`);
  }
  if (report.moved.length) {
    console.log(
      `\n${DRY_RUN ? 'Будет перемещено' : 'Перемещено'} в «${DEST_FOLDER}»: ${report.moved.length}`
    );
    for (const m of report.moved) {
      console.log(`  → ${m.date?.slice(0, 16).replace('T', ' ')} | ${m.from}`);
      console.log(`     ${m.subject} | полных адресов: ${m.fullAddresses}`);
    }
  }
  if (report.routed.length) {
    console.log(`\n${DRY_RUN ? 'Будет перемещено' : 'Перемещено'} по отправителю: ${report.routed.length}`);
    for (const r of report.routed) {
      console.log(`  → ${r.destination} | ${r.date?.slice(0, 16).replace('T', ' ')} | ${r.from}`);
      console.log(`     ${r.subject}`);
    }
  }
  if (report.trashed && report.trashed.length) {
    console.log(`\n${DRY_RUN ? 'Будет удалено' : 'Удалено'} в корзину (прикрепление без города): ${report.trashed.length}`);
    for (const t of report.trashed) console.log(`  → ${t.subject}`);
  }
  if (report.policyFlagged.length) {
    console.log(`\nПомечено по номеру полиса (без адресов): ${report.policyFlagged.length}`);
    for (const p of report.policyFlagged) {
      console.log(`  ★ ${p.insurer}: полисов ${p.policies}, наших ${p.found} | ${p.subject}`);
    }
  }
  if (report.policyMissed.length) {
    console.log(`\nПеремещено по номерам полисов (наших нет): ${report.policyMissed.length}`);
    for (const p of report.policyMissed) {
      console.log(`  → ${p.insurer}: полисов ${p.policies} | ${p.subject}`);
    }
  }
  console.log(`\nБез полного адреса (оставлены в INBOX): ${report.noFull}`);
  for (const n of report.noFullList.slice(0, 20)) console.log(`  · ${n.subject}`);
  if (report.passwordsUsed.length) {
    console.log(`\nИспользованные пароли (${report.passwordsUsed.length}):`);
    for (const p of report.passwordsUsed) console.log(`  - ${p.file}: «${p.password}» ← ${p.source}`);
  }
  if (report.errors.length) {
    console.log(`\nОшибки (${report.errors.length}):`);
    for (const e of report.errors) console.log(`  - ${e.subject}: ${e.error}`);
  }
  if (report.warnings.length) {
    console.log(`\nПредупреждения (${report.warnings.length}):`);
    for (const w of report.warnings) console.log(`  - ${w.subject}: ${w.warning}`);
  }
  console.log(`\nОтчёт: ${REPORT_FILE}`);

  logLine(
    `${DRY_RUN ? 'DRY ' : ''}checked=${report.checked} flagged=${report.flagged.length} ` +
      `moved=${report.moved.length} routed=${report.routed.length} noFull=${report.noFull} ` +
      `errors=${report.errors.length}` +
      (report.flagged.length ? ` | flagged: ${report.flagged.map((f) => f.subject).join(' / ')}` : '') +
      (report.moved.length ? ` | moved: ${report.moved.map((f) => f.subject).join(' / ')}` : '') +
      (report.routed.length
        ? ` | routed: ${report.routed.map((r) => `${r.destination}: ${r.subject}`).join(' / ')}`
        : '')
  );
}

async function main() {
  if (!acquireLock()) {
    console.log('Предыдущий запуск ещё выполняется — пропуск.');
    logLine('skip: previous run still in progress');
    return;
  }
  try {
    await run();
  } finally {
    releaseLock();
  }
}

// quick offline check of the address rules: SCANFLAG_SELFTEST=1 node scan_flag.js
if (process.env.SCANFLAG_SELFTEST === '1') {
  const samples = [
    '141300, Московская обл, г. Сергиев Посад, ул. Ленина, д. 5',
    '141370, Московская обл, г. Хотьково, ул. Мира, 3',
    '141322, Московская обл, Сергиево-Посадский р-н, с. Абрамцево, ул. Лесная, 1',
    '141355, д. Махра, ул. Центральная, 12',
    '141355, Сергиево-Посадский г.о., д. Махра, ул. Центральная, 12',
    '141200, г. Пушкино, ул. Ленина, 5',
    'Москва, ул. Тверская, 1',
    'г. Пушкино, д. Махра, ул. Ленина, 5',
    'д. Ивашково, ул. Полевая, 4',
    '141355, Ивашково, ул. Полевая, 4',
    // одноимённые населённые пункты в других регионах/округах — не наши
    'ТВЕРСКАЯ ОБЛ, д Гаврилково, Эжк эдем, д. 6, кв. 164',
    'РОССИЯ, 107258, Город Москва, вн.тер.г. муниципальный округ Богородское, ул Глебовская, д. 10А',
    'г. Москва, ш Дмитровское, д. 94, корп. 1, кв. 136',
    'Московская область, г.о. Одинцовский, п Назарьево, д. 23, кв. 20',
    // тот же Назарьево, но в нашем округе — наш
    'Московская область, г.о. Сергиево-Посадский, п Назарьево, д. 23, кв. 20',
    // без региона и индекса название принимается как раньше
    'д. Гаврилково, д. 6, кв. 1',
  ];
  for (const s of samples) console.log(JSON.stringify(matchLocalities(s)).padEnd(45), s);
  process.exit(0);
}

if (process.env.SCANFLAG_EXPORT === '1') {
  // re-usable from scripts (e.g. a dry check of an archive folder)
  module.exports = { scanWorkbook, scanMailAttachments, matchLocalities, opTypeOf, detectInsurerOf };
} else {
  main().catch((e) => {
  logLine(`ERROR: ${e.message}`);
  console.error('ERROR:', e.message);
  if (process.env.DEBUG) console.error(e.stack);
  process.exit(1);
  });
}
