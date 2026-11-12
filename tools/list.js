'use strict';
// Read-only IMAP peek: list mailboxes and recent INBOX messages.
// Credentials come from IMAP_USER / IMAP_PASS environment variables.
const { ImapFlow } = require('imapflow');

async function main() {
  const client = new ImapFlow({
    host: process.env.IMAP_HOST || 'imap.rambler.ru',
    port: Number(process.env.IMAP_PORT || 993),
    secure: true,
    auth: { user: process.env.IMAP_USER, pass: process.env.IMAP_PASS },
    logger: false,
  });

  await client.connect();
  const boxes = (await client.list()).map((b) => ({
    path: b.path,
    name: b.name,
    special: b.specialUse || null,
  }));

  const limit = Number(process.env.IMAP_LIMIT || 10);
  const lock = await client.getMailboxLock('INBOX');
  let messages = [];
  try {
    const total = client.mailbox.exists;
    if (total > 0) {
      const start = Math.max(1, total - limit + 1);
      for await (const msg of client.fetch(`${start}:*`, {
        envelope: true,
        flags: true,
        internalDate: true,
        bodyStructure: false,
      })) {
        const env = msg.envelope || {};
        messages.push({
          seq: msg.seq,
          date: msg.internalDate ? new Date(msg.internalDate).toISOString() : null,
          from: (env.from || []).map((a) => a.address || a.name).join(', '),
          subject: env.subject || '',
          seen: msg.flags ? msg.flags.has('\\Seen') : null,
        });
      }
    }
    messages = messages.reverse(); // newest first
    console.log(JSON.stringify({ mailboxes: boxes, inboxTotal: total, messages }, null, 2));
  } finally {
    lock.release();
  }
  await client.logout();
}

main().catch((e) => {
  console.error('IMAP ERROR:', e.message);
  process.exit(1);
});
