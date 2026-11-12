'use strict';
// Read-only: print text and attachment names of emails whose subject matches IMAP_FILTER.
const { ImapFlow } = require('imapflow');
const { simpleParser } = require('mailparser');

(async () => {
  const client = new ImapFlow({
    host: process.env.IMAP_HOST || 'imap.rambler.ru',
    port: Number(process.env.IMAP_PORT || 993),
    secure: true,
    auth: { user: process.env.IMAP_USER, pass: process.env.IMAP_PASS },
    logger: false,
  });
  await client.connect();
  const lock = await client.getMailboxLock(process.env.IMAP_BOX || 'INBOX');
  try {
    const filter = (process.env.IMAP_FILTER || '').toLowerCase();
    const uids = [];
    for await (const msg of client.fetch('1:*', { envelope: true })) {
      const subject = (msg.envelope?.subject || '').toLowerCase();
      const from = (msg.envelope?.from || []).map((a) => a.address).join(',').toLowerCase();
      if (!filter || subject.includes(filter) || from.includes(filter)) {
        uids.push({ uid: msg.uid, subject: msg.envelope?.subject || '' });
      }
    }
    for (const u of uids) {
      const msg = await client.fetchOne(String(u.uid), { source: true }, { uid: true });
      if (!msg || !msg.source) continue;
      const p = await simpleParser(msg.source);
      console.log('='.repeat(70));
      console.log('UID:', u.uid);
      console.log('Тема:', p.subject);
      console.log('От:', p.from?.text);
      console.log('Дата:', p.date?.toISOString());
      console.log('Вложения:');
      for (const a of p.attachments || []) console.log(`  - ${a.filename} (${a.size} байт)`);
      console.log('--- текст ---');
      const text = (p.text || '').trim();
      console.log((text || '(пусто)').slice(0, 1500));
      if (!text && p.html) {
        const plain = String(p.html).replace(/<[^>]+>/g, ' ').replace(/&nbsp;/g, ' ').replace(/\s+/g, ' ').trim();
        console.log('--- html (текст) ---');
        console.log(plain.slice(0, 1500));
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
