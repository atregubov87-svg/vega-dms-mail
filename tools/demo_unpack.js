'use strict';
// Demo: download the newest ZIP attachment, find its password in an email
// (same sender + same subject, body contains "Пароль"), and unpack with 7-Zip.
const { ImapFlow } = require('imapflow');
const { simpleParser } = require('mailparser');
const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');

const SEVENZIP = 'C:/Program Files/7-Zip/7z.exe';
const OUT = path.join(__dirname, 'samples');

const normSubject = (s) =>
  (s || '').toLowerCase().replace(/^(re|fw|fwd)\s*:\s*/g, '').replace(/\s+/g, ' ').trim();
const addr = (p) => (p?.value?.[0]?.address || '').toLowerCase();

async function fetchRecent(limit) {
  const client = new ImapFlow({
    host: process.env.IMAP_HOST || 'imap.rambler.ru',
    port: Number(process.env.IMAP_PORT || 993),
    secure: true,
    auth: { user: process.env.IMAP_USER, pass: process.env.IMAP_PASS },
    logger: false,
  });
  await client.connect();
  const lock = await client.getMailboxLock('INBOX');
  const list = [];
  try {
    const total = client.mailbox.exists;
    const start = Math.max(1, total - limit + 1);
    for await (const msg of client.fetch(`${start}:*`, { source: true })) {
      list.push(await simpleParser(msg.source));
    }
  } finally {
    lock.release();
  }
  await client.logout();
  return list; // oldest -> newest
}

function findPassword(mails, target) {
  const wantAddr = addr(target.from);
  const wantSubject = normSubject(target.subject);
  for (const m of mails) {
    if (addr(m.from) !== wantAddr) continue;
    if (normSubject(m.subject) !== wantSubject) continue;
    const body = `${m.text || ''}\n${m.html || ''}`;
    const mt = body.match(/парол[ья][^:\n]{0,40}[:：]\s*([^\s<]+)/i) || body.match(/password\s*[:：]\s*(\S+)/i);
    if (mt) return mt[1].trim();
  }
  return null;
}

(async () => {
  const mails = await fetchRecent(20);
  let target = null;
  for (let i = mails.length - 1; i >= 0 && !target; i--) {
    const att = (mails[i].attachments || []).find((a) => /\.(zip|rar|7z)$/i.test(a.filename || ''));
    if (att) target = { mail: mails[i], att };
  }
  if (!target) throw new Error('No archive attachment found in recent mail');

  const password = findPassword(mails, target.mail);
  console.log('Письмо:', target.mail.subject);
  console.log('От:', addr(target.mail.from));
  console.log('Архив:', target.att.filename);
  console.log('Пароль найден:', password ? 'да' : 'нет', password ? `(${password})` : '');

  const dir = path.join(OUT, 'sberins');
  fs.mkdirSync(dir, { recursive: true });
  const zipPath = path.join(dir, target.att.filename);
  fs.writeFileSync(zipPath, target.att.content);
  console.log('Сохранено:', zipPath, target.att.size, 'байт');

  const dest = path.join(dir, 'unpacked');
  fs.mkdirSync(dest, { recursive: true });
  const args = ['x', '-y', `-o${dest}`];
  if (password) args.push(`-p${password}`);
  args.push(zipPath);
  try {
    const out = execFileSync(SEVENZIP, args, { encoding: 'utf8' });
    console.log('7-Zip: OK');
    console.log(out.split(/\r?\n/).filter((l) => /Extracting|Everything is Ok/.test(l)).join('\n'));
  } catch (e) {
    console.error('7-Zip ошибка:', (e.stdout || '') + (e.stderr || '') || e.message);
    process.exitCode = 2;
  }
  console.log('Файлы в папке:');
  for (const f of fs.readdirSync(dest, { recursive: true })) console.log(' -', f);
})().catch((e) => {
  console.error('ERROR:', e.message);
  process.exit(1);
});
