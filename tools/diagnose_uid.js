'use strict';
// Read-only diagnostic: dump address-column detection for given INBOX UIDs.
const { ImapFlow } = require('imapflow');
const { simpleParser } = require('mailparser');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const CONFIG = JSON.parse(fs.readFileSync(path.join(__dirname, '..', 'config.json'), 'utf8'));
const SPREADSHEET_RE = /\.(xlsx|xlsm|xls)$/i;
const ARCHIVE_RE = /\.(zip|rar|7z)$/i;
const ADDR_HEADER_RE = /адрес|город|насел|прописк|место жит|address/i;

function isAddrHeader(text) {
  const t = String(text ?? '').replace(/\s+/g, ' ').trim();
  if (!t || t.length > 40) return false;
  if (t.endsWith(':')) return false;
  return ADDR_HEADER_RE.test(t);
}

const uids = process.argv.slice(2);
(async () => {
  const client = new ImapFlow({
    host: 'imap.rambler.ru', port: 993, secure: true,
    auth: { user: CONFIG.user, pass: CONFIG.password }, logger: false,
  });
  await client.connect();
  const lock = await client.getMailboxLock('INBOX');
  try {
    for (const uid of uids) {
      const msg = await client.fetchOne(uid, { source: true }, { uid: true });
      const p = await simpleParser(msg.source);
      console.log(`\n===== UID ${uid} | ${p.subject} | ${p.date?.toISOString()}`);
      for (const a of p.attachments || []) {
        if (!SPREADSHEET_RE.test(a.filename || '') && !ARCHIVE_RE.test(a.filename || '')) continue;
        console.log(`--- file: ${a.filename} (${a.size} bytes)`);
        if (ARCHIVE_RE.test(a.filename || '')) { console.log('  archive, skipped by diagnostic'); continue; }
        const wb = XLSX.read(a.content, { type: 'buffer' });
        for (const sn of wb.SheetNames) {
          const ws = wb.Sheets[sn];
          const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '' });
          console.log(`  [sheet] ${sn} rows=${rows.length}`);
          let headerRow = -1, addrCols = [];
          for (let r = 0; r < Math.min(rows.length, 80); r++) {
            const cols = [];
            (rows[r] || []).forEach((cell, c) => {
              if (isAddrHeader(String(cell || ''))) cols.push({ c, h: String(cell).trim() });
            });
            if (cols.length) { headerRow = r; addrCols = cols; break; }
          }
          console.log(`  headerRow=${headerRow} addrCols=${JSON.stringify(addrCols)}`);
          const from = headerRow >= 0 ? Math.max(0, headerRow - 2) : 0;
          for (let r = from; r < rows.length; r++) {
            console.log(`   r${r + 1}: ${JSON.stringify((rows[r] || []).slice(0, 20))}`);
          }
          if (headerRow >= 0) {
            const samples = [];
            for (let r = headerRow + 1; r < rows.length && samples.length < 8; r++) {
              samples.push(addrCols.map((col) => String((rows[r] || [])[col.c] ?? '').trim()));
            }
            console.log('  address-column samples:');
            for (const s of samples) console.log(`    ${JSON.stringify(s)}`);
          }
        }
      }
    }
  } finally { lock.release(); await client.logout(); }
})().catch((e) => { console.error(e); process.exit(1); });
