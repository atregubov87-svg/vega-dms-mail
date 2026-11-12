'use strict';
// Read-only: list newest emails in IMAP_BOX.
const { ImapFlow } = require('imapflow');

(async () => {
  const client = new ImapFlow({
    host: process.env.IMAP_HOST || 'imap.rambler.ru',
    port: Number(process.env.IMAP_PORT || 993),
    secure: true,
    auth: { user: process.env.IMAP_USER, pass: process.env.IMAP_PASS },
    logger: false,
  });
  await client.connect();
  const box = process.env.IMAP_BOX || 'INBOX';
  const lock = await client.getMailboxLock(box);
  try {
    const total = client.mailbox.exists;
    const n = Number(process.env.IMAP_LIMIT || 20);
    const start = Math.max(1, total - n + 1);
    console.log(`Ящик: ${box}, писем: ${total}`);
    const out = [];
    for await (const msg of client.fetch(`${start}:*`, { envelope: true })) {
      out.push({
        uid: msg.uid,
        date: msg.envelope?.date ? new Date(msg.envelope.date).toISOString() : null,
        from: (msg.envelope?.from || []).map((a) => a.address).join(','),
        subject: msg.envelope?.subject || '',
      });
    }
    for (const m of out.reverse()) console.log(`${m.date?.slice(0, 16).replace('T', ' ')} | ${m.from} | ${m.subject}`);
  } finally {
    lock.release();
    await client.logout();
  }
})().catch((e) => {
  console.error('ERROR:', e.message);
  process.exit(1);
});
