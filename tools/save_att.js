'use strict';
// Save attachments of one email to its own fresh directory under work/raw/<uid>/.
// Usage: node save_att.js <folder> <uid>
const { ImapFlow } = require('imapflow');
const { simpleParser } = require('mailparser');
const fs = require('fs');
const path = require('path');

const CONFIG = JSON.parse(fs.readFileSync(path.join(__dirname, '..', 'config.json'), 'utf8'));
const OUT_ROOT = 'C:/Users/Администратор/vegasrv/work/raw';

(async () => {
  const box = process.argv[2] || 'INBOX';
  const uid = process.argv[3];
  if (!uid) throw new Error('uid required');
  const client = new ImapFlow({
    host: 'imap.rambler.ru', port: 993, secure: true,
    auth: { user: CONFIG.user, pass: CONFIG.password }, logger: false,
  });
  await client.connect();
  const lock = await client.getMailboxLock(box);
  try {
    const msg = await client.fetchOne(uid, { source: true }, { uid: true });
    const p = await simpleParser(msg.source);
    const dest = path.join(OUT_ROOT, String(uid));
    if (fs.existsSync(dest)) throw new Error(`directory already exists: ${dest}`);
    fs.mkdirSync(dest, { recursive: true });
    console.log(`SUBJECT: ${p.subject}`);
    for (const a of p.attachments || []) {
      const name = String(a.filename || 'attachment').replace(/[\\/:*?"<>|]/g, '_');
      fs.writeFileSync(path.join(dest, name), a.content);
      console.log(`SAVED: ${path.join(dest, name)} (${a.size})`);
    }
  } finally { lock.release(); await client.logout(); }
})().catch((e) => { console.error(e); process.exit(1); });
