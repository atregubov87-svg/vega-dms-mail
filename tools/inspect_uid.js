'use strict';
// Inspect one email's attachments: extract archives and show sheet layouts.
const { ImapFlow } = require('imapflow');
const { simpleParser } = require('mailparser');
const XLSX = require('xlsx');
const fs = require('fs');
const os = require('os');
const path = require('path');
const { execFileSync } = require('child_process');

const UID = process.env.INSPECT_UID;
const PASSWORD = process.env.PASSWORD || '';
const SEVENZIP = 'C:/Program Files/7-Zip/7z.exe';

(async () => {
  const client = new ImapFlow({
    host: 'imap.rambler.ru',
    port: 993,
    secure: true,
    auth: { user: process.env.IMAP_USER, pass: process.env.IMAP_PASS },
    logger: false,
  });
  await client.connect();
  const lock = await client.getMailboxLock('INBOX');
  try {
    const msg = await client.fetchOne(String(UID), { source: true }, { uid: true });
    if (!msg || !msg.source) throw new Error('no source');
    const p = await simpleParser(msg.source);
    console.log('Тема:', p.subject);
    console.log('От:', p.from?.text);
    for (const att of p.attachments || []) {
      console.log('\n=== Вложение:', att.filename, att.size, 'байт');
      const work = fs.mkdtempSync(path.join(os.tmpdir(), 'inspect-'));
      const file = path.join(work, (att.filename || 'file').replace(/[\\/]/g, '_'));
      fs.writeFileSync(file, att.content);
      let files = [file];
      if (/\.(zip|rar|7z)$/i.test(att.filename || '')) {
        const dest = path.join(work, 'unpacked');
        fs.mkdirSync(dest, { recursive: true });
        try {
          execFileSync(SEVENZIP, ['x', '-y', '-sccUTF-8', `-o${dest}`, `-p${PASSWORD}`, file], {
            encoding: 'utf8',
            stdio: 'pipe',
          });
          files = fs.readdirSync(dest, { recursive: true }).map((f) => path.join(dest, String(f)));
          console.log('Распаковано:', files.map((f) => path.relative(dest, f)).join(', '));
        } catch (e) {
          console.log('ОШИБКА распаковки:', `${e.stdout || ''}${e.stderr || ''}`.slice(0, 400));
        }
      }
      let shown = 0;
      for (const f of files) {
        if (!/\.(xlsx|xlsm|xls)$/i.test(f)) continue;
        if (process.env.ONLY_FIRST === '1' && shown > 0) continue;
        shown++;
        try {
          const wb = XLSX.read(fs.readFileSync(f), { type: 'buffer' });
          for (const name of wb.SheetNames) {
            const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: false, defval: '' });
            console.log(`-- лист «${name}», строк: ${rows.length}`);
            for (let r = 0; r < Math.min(10, rows.length); r++) {
              const vals = (rows[r] || []).map((v) => String(v ?? '').replace(/\s+/g, ' ').trim()).filter(Boolean);
              if (vals.length) console.log(`   ${r + 1}: ${vals.slice(0, 14).join(' | ').slice(0, 350)}`);
            }
          }
        } catch (e) {
          console.log('  ошибка чтения', path.basename(f), e.message);
        }
      }
    }
  } finally {
    lock.release();
    await client.logout();
  }
})();
