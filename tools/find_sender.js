'use strict';
// Read-only: list INBOX messages whose From/Subject contains the given text.
const { ImapFlow } = require('imapflow');
const fs = require('fs');
const path = require('path');

const CONFIG = JSON.parse(fs.readFileSync(path.join(__dirname, '..', 'config.json'), 'utf8'));
const needle = (process.argv[2] || 'spiski').toLowerCase();
const matchAll = needle === '-';
const box = process.argv[3] || 'INBOX';

(async () => {
  const client = new ImapFlow({
    host: process.env.IMAP_HOST || 'imap.rambler.ru',
    port: 993,
    secure: true,
    auth: { user: CONFIG.user, pass: CONFIG.password },
    logger: false,
  });
  await client.connect();
  const lock = await client.getMailboxLock(box);
  try {
    const found = [];
    for await (const msg of client.fetch('1:*', { envelope: true, flags: true, internalDate: true, bodyStructure: true })) {
      const env = msg.envelope || {};
      const from = (env.from || []).map((a) => `${a.name || ''} <${a.address || ''}>`).join(', ');
      const subject = env.subject || '';
      if (!matchAll && !`${from} ${subject}`.toLowerCase().includes(needle)) continue;
      const atts = [];
      const walk = (n) => {
        if (!n) return;
        if (n.disposition === 'attachment' || (n.childNodes || []).length === 0) {
          if (n.parameters && n.parameters.name) atts.push(n.parameters.name);
        }
        for (const c of n.childNodes || []) walk(c);
      };
      walk(msg.bodyStructure);
      found.push({
        uid: msg.uid,
        date: msg.internalDate ? new Date(msg.internalDate).toISOString() : null,
        from,
        subject,
        flags: [...(msg.flags || [])],
        attachments: atts,
      });
    }
    console.log(JSON.stringify(found, null, 2));
    console.error(`total matches in ${box}: ${found.length}`);
  } finally {
    lock.release();
    await client.logout();
  }
})().catch((e) => {
  console.error(e);
  process.exit(1);
});
