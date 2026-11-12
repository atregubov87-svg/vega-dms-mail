'use strict';
// Read-only: print sheet layout (first rows) of spreadsheet attachments
// in emails whose subject matches IMAP_FILTER.
const { ImapFlow } = require('imapflow');
const { simpleParser } = require('mailparser');
const XLSX = require('xlsx');

(async () => {
  const client = new ImapFlow({
    host: process.env.IMAP_HOST || 'imap.rambler.ru',
    port: Number(process.env.IMAP_PORT || 993),
    secure: true,
    auth: { user: process.env.IMAP_USER, pass: process.env.IMAP_PASS },
    logger: false,
  });
  await client.connect();
  const lock = await client.getMailboxLock('INBOX');
  try {
    const filter = (process.env.IMAP_FILTER || '').toLowerCase();
    const rowsLimit = Number(process.env.ROWS || 15);
    const uids = [];
    for await (const msg of client.fetch('1:*', { envelope: true })) {
      if ((msg.envelope?.subject || '').toLowerCase().includes(filter)) uids.push(msg.uid);
    }
    for (const uid of uids.slice(-3)) {
      const msg = await client.fetchOne(String(uid), { source: true }, { uid: true });
      const p = await simpleParser(msg.source);
      console.log('='.repeat(70));
      console.log('UID:', uid, '|', p.subject);
      for (const att of p.attachments || []) {
        if (!/\.(xlsx|xlsm|xls)$/i.test(att.filename || '')) continue;
        console.log(`\n### ${att.filename} (${att.size} байт)`);
        let wb;
        try {
          wb = XLSX.read(att.content, { type: 'buffer' });
        } catch (e) {
          console.log('  ошибка чтения:', e.message);
          continue;
        }
        for (const name of wb.SheetNames) {
          const ws = wb.Sheets[name];
          const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '' });
          console.log(`  -- лист «${name}»: ${ws['!ref'] || 'пусто'}, строк: ${rows.length}`);
          for (let r = 0; r < Math.min(rowsLimit, rows.length); r++) {
            const vals = (rows[r] || []).map((v) => String(v ?? '').replace(/\s+/g, ' ').trim()).filter((v) => v !== '');
            if (vals.length) console.log(`   ${r + 1}: ${vals.slice(0, 12).join(' | ').slice(0, 300)}`);
          }
        }
      }
      console.log();
    }
  } finally {
    lock.release();
    await client.logout();
  }
})().catch((e) => {
  console.error('ERROR:', e.message);
  process.exit(1);
});
