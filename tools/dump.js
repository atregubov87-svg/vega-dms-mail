'use strict';
// Read-only: dump text bodies + attachment names for the newest INBOX messages.
// Env: IMAP_USER, IMAP_PASS, IMAP_LIMIT (default 5), IMAP_FILTER (subject substring)
const { ImapFlow } = require('imapflow');
const { simpleParser } = require('mailparser');

async function main() {
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
    const total = client.mailbox.exists;
    const n = Number(process.env.IMAP_LIMIT || 5);
    const start = Math.max(1, total - n + 1);
    const filter = (process.env.IMAP_FILTER || '').toLowerCase();
    const out = [];
    for await (const msg of client.fetch(`${start}:*`, { source: true, envelope: true })) {
      const parsed = await simpleParser(msg.source);
      const subject = parsed.subject || msg.envelope?.subject || '';
      if (filter && !subject.toLowerCase().includes(filter)) continue;
      out.push({
        subject,
        from: parsed.from?.text || '',
        date: parsed.date ? parsed.date.toISOString() : null,
        text: (parsed.text || '').trim().slice(0, 4000),
        attachments: (parsed.attachments || []).map((a) => ({ filename: a.filename, size: a.size })),
      });
    }
    out.reverse();
    console.log(JSON.stringify(out, null, 2));
  } finally {
    lock.release();
  }
  await client.logout();
}

main().catch((e) => {
  console.error('IMAP ERROR:', e.message);
  process.exit(1);
});
