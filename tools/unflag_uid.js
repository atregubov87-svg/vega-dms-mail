'use strict';
// Remove the IMAP \Flagged star from one email.
// Env: MSG_UID (required), FOLDER (default INBOX)
const { ImapFlow } = require('imapflow');
const fs = require('fs');
const path = require('path');

const CONFIG = JSON.parse(fs.readFileSync(path.join(__dirname, '..', 'config.json'), 'utf8'));
const UID = process.env.MSG_UID;
const FOLDER = process.env.FOLDER || 'INBOX';

(async () => {
  if (!UID) throw new Error('MSG_UID is required');
  const client = new ImapFlow({
    host: 'imap.rambler.ru', port: 993, secure: true,
    auth: { user: CONFIG.user, pass: CONFIG.password }, logger: false,
  });
  await client.connect();
  const lock = await client.getMailboxLock(FOLDER);
  try {
    const before = await client.fetchOne(String(UID), { flags: true, envelope: true }, { uid: true });
    if (!before) throw new Error('письмо не найдено');
    await client.messageFlagsRemove(String(UID), ['\\Flagged'], { uid: true });
    const after = await client.fetchOne(String(UID), { flags: true }, { uid: true });
    console.log('Тема:', before.envelope?.subject);
    console.log('Флаги до :', [...before.flags].join(' '));
    console.log('Флаги после:', [...after.flags].join(' '));
  } finally {
    lock.release();
    await client.logout();
  }
})().catch((e) => { console.error(e.message); process.exit(1); });
