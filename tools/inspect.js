'use strict';
// Read-only: show attachment metadata for the newest INBOX messages.
const { ImapFlow } = require('imapflow');

function collect(part, out) {
  if (!part) return;
  if (part.disposition === 'attachment' || part.dispositionParameters?.filename) {
    const name = part.dispositionParameters?.filename || part.parameters?.name || '';
    out.push({
      filename: name,
      type: `${part.type || ''}/${part.subtype || ''}`,
      size: part.size || 0,
      encoding: part.encoding || '',
    });
  }
  for (const child of part.childNodes || []) collect(child, out);
}

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
    const rows = [];
    for await (const msg of client.fetch(`${start}:*`, { envelope: true, bodyStructure: true, internalDate: true })) {
      const atts = [];
      collect(msg.bodyStructure, atts);
      rows.push({
        subject: msg.envelope?.subject || '',
        from: (msg.envelope?.from || []).map((a) => a.address).join(', '),
        date: msg.internalDate ? new Date(msg.internalDate).toISOString() : null,
        attachments: atts,
      });
    }
    rows.reverse();
    console.log(JSON.stringify(rows, null, 2));
  } finally {
    lock.release();
  }
  await client.logout();
}

main().catch((e) => {
  console.error('IMAP ERROR:', e.message);
  process.exit(1);
});
