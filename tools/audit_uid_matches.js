'use strict';
// Read-only: show every row of one email's list attachments that the scanner
// considers local (same rules as scan_flag.js), with the patient name column.
// Env: UID (required), PASSWORD (optional), FOLDER (default INBOX)
const { ImapFlow } = require('imapflow');
const { simpleParser } = require('mailparser');
const XLSX = require('xlsx');
const fs = require('fs');
const os = require('os');
const path = require('path');
const { execFileSync } = require('child_process');

const CONFIG = JSON.parse(fs.readFileSync(path.join(__dirname, '..', 'config.json'), 'utf8'));
const UID = process.env.UID;
const FOLDER = process.env.FOLDER || 'INBOX';
const PASSWORD = process.env.PASSWORD || '';
const SEVENZIP = 'C:/Program Files/7-Zip/7z.exe';

const SETTLEMENTS_FILE = path.join(__dirname, '..', 'sergiev_posad_settlements.json');
let DISTRICT_SETTLEMENTS = [];
try {
  DISTRICT_SETTLEMENTS = JSON.parse(fs.readFileSync(SETTLEMENTS_FILE, 'utf8')).settlements.map((s) => s.name);
} catch { /* optional */ }
const escapeRe = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
const DISTRICT_SETTLEMENT_RE = DISTRICT_SETTLEMENTS.length
  ? new RegExp('(?:^|[^а-яa-zё-])(' + DISTRICT_SETTLEMENTS.map((n) => escapeRe(String(n).toLowerCase().replace(/ё/g, 'е'))).join('|') + ')(?![а-яa-zё-])', 'i')
  : null;
const DISTRICT_INDEX_RE = /(?:^|[^0-9])1413\d\d(?:[^0-9]|$)/;
const DISTRICT_NAME_RE = /сергиев[а-я]*[\s-]*посадск|сергиевопосадск/;
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
const TARGETS = [
  { name: 'Сергиев Посад / Сергиево-Посадский', re: /сергиев[а-я]*[\s-]*посад(?!ск)|сергиев[а-я]*[\s-]*посадск[а-я]*\s+(?:район|р-н|округ|городск|муницип)/ },
  { name: 'Хотьково', re: /хотьков(?!ск)/ },
  { name: 'Краснозаводск', re: /краснозаводск(?!ая|ий|ое|ой|ую|им|их)/ },
  { name: 'Пересвет', re: /пересвет(?!ов)/ },
];
const normText = (v) => String(v).toLowerCase().replace(/ё/g, 'е');
function matchLocalities(value) {
  const s = normText(value);
  const hits = [];
  for (const t of TARGETS) if (t.re.test(s)) hits.push(t.name);
  if (DISTRICT_INDEX_RE.test(s)) hits.push('Индекс 1413xx');
  if (DISTRICT_SETTLEMENT_RE) {
    const m = s.match(DISTRICT_SETTLEMENT_RE);
    const explicit = hits.length > 0 || DISTRICT_NAME_RE.test(s);
    if (m && (explicit || !geographyElsewhere(s))) hits.push('Нас. пункт: ' + m[1]);
  }
  return hits;
}
const ADDR_HEADER_RE = /адрес|город|насел|прописк|место жит|address/i;
function isAddrHeader(text) {
  const t = String(text ?? '').replace(/\s+/g, ' ').trim();
  if (!t || t.length > 40) return false;
  if (t.endsWith(':')) return false;
  return ADDR_HEADER_RE.test(t);
}
const NAME_HEADER_RE = /фио|фамил|застрахован|пациент|сотрудник/i;

function extractFiles(work, file) {
  if (!/\.(zip|rar|7z)$/i.test(file)) return [file];
  const dest = path.join(work, 'unpacked');
  fs.mkdirSync(dest, { recursive: true });
  try {
    execFileSync(SEVENZIP, ['x', '-y', '-sccUTF-8', `-o${dest}`, `-p${PASSWORD}`, file], { encoding: 'utf8', stdio: 'pipe' });
    return fs.readdirSync(dest, { recursive: true }).map((f) => path.join(dest, String(f)));
  } catch (e) {
    console.log('ОШИБКА распаковки:', `${e.stdout || ''}${e.stderr || ''}`.slice(0, 300));
    return [];
  }
}

(async () => {
  const client = new ImapFlow({
    host: 'imap.rambler.ru', port: 993, secure: true,
    auth: { user: CONFIG.user, pass: CONFIG.password }, logger: false,
  });
  await client.connect();
  const lock = await client.getMailboxLock(FOLDER);
  try {
    const msg = await client.fetchOne(String(UID), { source: true }, { uid: true });
    if (!msg || !msg.source) throw new Error('письмо не найдено');
    const p = await simpleParser(msg.source);
    console.log('Тема:', p.subject);
    console.log('От:', p.from?.text);
    console.log('Дата:', p.date?.toISOString());
    const work = fs.mkdtempSync(path.join(os.tmpdir(), 'audituid-'));
    for (const att of p.attachments || []) {
      if (!/\.(xlsx|xlsm|xls|zip|rar|7z)$/i.test(att.filename || '')) continue;
      const file = path.join(work, (att.filename || 'file').replace(/[\\/]/g, '_'));
      fs.writeFileSync(file, att.content);
      for (const f of extractFiles(work, file)) {
        if (!/\.(xlsx|xlsm|xls)$/i.test(f)) continue;
        const wb = XLSX.read(fs.readFileSync(f), { type: 'buffer' });
        for (const sheetName of wb.SheetNames) {
          const ws = wb.Sheets[sheetName];
          if (!ws) continue;
          const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '' });
          let headerRow = -1;
          let addrCols = [];
          let nameCol = -1;
          for (let r = 0; r < Math.min(rows.length, 80); r++) {
            const cols = [];
            (rows[r] || []).forEach((cell, c) => { if (isAddrHeader(String(cell || ''))) cols.push({ index: c }); });
            if (cols.length) {
              headerRow = r; addrCols = cols;
              (rows[r] || []).forEach((cell, c) => { if (nameCol < 0 && NAME_HEADER_RE.test(String(cell || ''))) nameCol = c; });
              break;
            }
          }
          console.log(`\n=== ${path.basename(f)} | лист «${sheetName}» | строк ${rows.length} | строка заголовка ${headerRow + 1} | колонки адреса: ${addrCols.map((c) => c.index + 1).join(',')} | имя: ${nameCol >= 0 ? nameCol + 1 : '?'}`);
          if (headerRow < 0) { console.log('  нет колонки адреса'); continue; }
          console.log('  заголовки:', (rows[headerRow] || []).map((v, i) => `${i + 1}=${String(v).slice(0, 25)}`).join(' | '));
          const suspect = /серг|1413|хотьк|красноз|пересв|посад/i;
          let addrCells = 0;
          for (let r = headerRow + 1; r < rows.length; r++) {
            const row = rows[r] || [];
            for (const col of addrCols) {
              const text = String(row[col.index] ?? '');
              if (!text.trim()) continue;
              addrCells++;
              const loc = matchLocalities(text);
              if (loc.length) {
                const fio = nameCol >= 0 ? `${row[nameCol] ?? ''} ${row[nameCol + 1] ?? ''} ${row[nameCol + 2] ?? ''}`.trim() : '';
                console.log(`  СТРОКА ${r + 1}: ${fio || '(без ФИО)'} | ${loc.join(', ')} | ${text.trim()}`);
              }
            }
          }
          console.log(`  всего непустых адресных ячеек: ${addrCells}`);
          // Independent sweep: any cell mentioning our towns/index, even outside address columns.
          for (let r = 0; r < rows.length; r++) {
            for (const cell of rows[r] || []) {
              const v = String(cell ?? '');
              if (v.trim() && suspect.test(v)) {
                console.log(`  [сырой поиск] СТРОКА ${r + 1}, колонка ${(rows[r] || []).indexOf(cell) + 1}: ${(rows[r] || []).map((x) => String(x)).join(' | ').slice(0, 300)}`);
              }
            }
          }
        }
      }
    }
  } finally {
    lock.release();
    await client.logout();
  }
})().catch((e) => { console.error(e.message); process.exit(1); });
