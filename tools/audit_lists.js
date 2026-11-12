'use strict';
// Read-only audit: which INBOX emails with spreadsheet/archive attachments
// are recognised as DMS lists and which are skipped (and why).
const { ImapFlow } = require('imapflow');
const { simpleParser } = require('mailparser');
const fs = require('fs');
const path = require('path');

const CONFIG = JSON.parse(fs.readFileSync(path.join(__dirname, '..', 'config.json'), 'utf8'));
const SPREADSHEET_RE = /\.(xlsx|xlsm|xls)$/i;
const ARCHIVE_RE = /\.(zip|rar|7z)$/i;
const LIST_SUBJECT_RE = /спис|прикрепл|открепл|застрахован|дмс/i;
const LIST_ATTACHMENT_RE = /спис|прикрепл|открепл|застрахован|snyat|прикр|откр/i;
const SKIP_SUBJECT_RE = /^\s*срочно\b/i;
const SKIP_ATTACHMENT_RE = /^ГП\d/i;

(async () => {
  const client = new ImapFlow({
    host: 'imap.rambler.ru',
    port: 993,
    secure: true,
    auth: { user: CONFIG.user, pass: CONFIG.password },
    logger: false,
  });
  await client.connect();
  const lock = await client.getMailboxLock('INBOX');
  const rows = [];
  let total = 0;
  try {
    total = client.mailbox.exists;
    for await (const msg of client.fetch('1:*', { source: true, envelope: true })) {
      if (!msg || !msg.source) {
        rows.push({ uid: msg && msg.uid, subject: '(нет source)', attachments: [], isList: false, reason: 'no source' });
        continue;
      }
      const p = await simpleParser(msg.source);
      const atts = (p.attachments || []).filter(
        (a) => SPREADSHEET_RE.test(a.filename || '') || ARCHIVE_RE.test(a.filename || '')
      );
      if (!atts.length) continue;
      const isList =
        LIST_SUBJECT_RE.test(p.subject || '') || atts.some((a) => LIST_ATTACHMENT_RE.test(a.filename || ''));
      let reason = '';
      if (!isList) reason = 'тема и имена файлов не похожи на список';
      else if (SKIP_SUBJECT_RE.test(p.subject || '')) reason = 'гарантийное письмо (СРОЧНО)';
      else if (atts.every((a) => SKIP_ATTACHMENT_RE.test(a.filename || ''))) reason = 'только ГП*.zip';
      rows.push({
        uid: msg.uid,
        date: p.date ? p.date.toISOString() : '',
        from: (p.from?.value || []).map((a) => a.address).join(','),
        subject: p.subject || '(без темы)',
        attachments: atts.map((a) => a.filename || ''),
        isList,
        reason,
      });
    }
  } finally {
    lock.release();
    await client.logout();
  }

  const lists = rows.filter((r) => r.isList);
  const other = rows.filter((r) => !r.isList);
  console.log(`Всего писем в INBOX: ${total}`);
  console.log(`Писем с файлами xls/xlsx/zip: ${rows.length}`);
  console.log(`Распознано как списки: ${lists.length}`);
  console.log(`Не распознано: ${other.length}`);
  console.log('\n=== НЕ распознаны как списки ===');
  for (const r of other) {
    console.log(`${r.date.slice(0, 16).replace('T', ' ')} | ${r.from} | ${r.subject}`);
    console.log(`    файлы: ${r.attachments.join(', ')}`);
  }
  console.log('\n=== Распознаны как списки ===');
  for (const r of lists) {
    console.log(`${r.date.slice(0, 16).replace('T', ' ')} | ${r.from} | ${r.subject}${r.reason ? ` [${r.reason}]` : ''}`);
  }
})().catch((e) => {
  console.error('ERROR:', e.message);
  process.exit(1);
});
