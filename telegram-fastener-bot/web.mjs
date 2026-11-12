// Web order form: colleagues pick sizes in a browser instead of typing a Telegram message. The form is turned
// into the bot's text order format and goes through the same parseOrder / generate / sendMail path, so every
// business rule and template stays in one place. Runs as its own process beside bot.mjs.
import fs from 'node:fs/promises';
import http from 'node:http';
import os from 'node:os';
import path from 'node:path';
import crypto from 'node:crypto';
import {execFile, spawn} from 'node:child_process';
import {promisify} from 'node:util';
import {fileURLToPath} from 'node:url';
import {parseOrder} from './parser.mjs';
import {describeItem, generate} from './excel.mjs';
import {mailBody, mailSubject, sendMail} from './mail.mjs';
import {orderText} from './items.mjs';

const root = path.dirname(fileURLToPath(import.meta.url));
const readJson = async name => {
  try { return JSON.parse((await fs.readFile(path.join(root, name), 'utf8')).replace(/^﻿/, '')); }
  catch (e) { if (e.code !== 'ENOENT') throw e; }
};

// web.json: {"port": 8080, "code": ""}. The user chose on 2026-11-12 to share the link without an access code;
// writing a code here brings the access-code field back for everyone.
const cfg = {port: 8080, code: '', ...await readJson('web.json')};

const one = s => String(s ?? '').replace(/\s+/g, ' ').trim();

// What the Excel file will list, for the confirmation screen.
const review = order => ({
  project: order.project, material: order.material, requester: order.requester,
  sections: order.sections.map(sec => ({title: sec.title, items: sec.items.map(item => ({...describeItem(order.material, item), qty: item.qty}))})),
});

// One order at a time: Excel generation and Outlook are not meant to run in parallel.
let chain = Promise.resolve();
const queued = fn => { const p = chain.then(fn); chain = p.catch(() => {}); return p; };
// A confirm that is sent twice (double tap, retry after a timeout) must not create two orders.
const submitted = new Map();
let recent = [];

async function placeOrder(form) {
  const source = orderText(form), order = parseOrder(source);
  const id = `web_${Date.now()}_${crypto.randomBytes(8).toString('hex')}`;
  const dir = path.join(root, 'outputs', id), now = new Date();
  await fs.mkdir(dir, {recursive: true});
  await fs.writeFile(path.join(dir, 'request.json'), JSON.stringify({via: 'web', receivedAt: now.toISOString(), order, source}, null, 2));
  let file;
  try { file = await generate(order, id); }
  catch (e) { console.error(`Excel generation failed for ${id}: ${String(e.message).split(/\r?\n/)[0]}`); throw Object.assign(Error('Excel generation failed. Your order has been saved on the host computer; please contact the administrator.'), {status: 500}); }
  // The weekly summary lists orders that have a delivery record; a web order is delivered as a download.
  await fs.writeFile(path.join(dir, 'delivery.json'), JSON.stringify({status: 'sent', via: 'web'}));
  console.log(`Web order generated ${id}`);
  const result = {id, filename: path.basename(file), download: `/file/${id}`, subject: mailSubject(order), email: 'off'};
  const to = (await readJson('email.json'))?.to;
  if (to) {
    try {
      const sent = await sendMail({to, subject: result.subject, body: mailBody(order, order.requester, file, now, 'web order form'), attachment: file}, dir);
      await fs.writeFile(path.join(dir, 'email.json'), JSON.stringify({...sent, to, subject: result.subject, at: new Date().toISOString(), via: 'web'}));
      result.email = sent.status === 'sent' ? 'sent' : 'queued';
      console.log(`Web order emailed ${id} (${sent.status})`);
    } catch (e) {
      result.email = 'failed';
      console.error(`Email failed for ${id}: ${e.message.split(/\r?\n/)[0]} ${new Date().toISOString()}`);
    }
  }
  return result;
}

// Ten wrong access codes from one address lock it out for ten minutes.
const failures = new Map();
// Behind a tunnel on this PC every request arrives from loopback; the tunnel names the real caller in a header.
function caller(req) {
  const ip = req.socket.remoteAddress || '';
  if (!/^(::1|127\.|::ffff:127\.)/.test(ip)) return ip;
  return one(req.headers['cf-connecting-ip'] || String(req.headers['x-forwarded-for'] || '').split(',')[0]).slice(0, 64) || ip;
}
function allowed(req, code) {
  if (!cfg.code) return;
  if (failures.size > 5000) failures.clear();
  const ip = caller(req), f = failures.get(ip);
  if (f && f.count >= 10 && Date.now() - f.at < 600e3) throw Object.assign(Error('Too many wrong access codes. Try again in 10 minutes.'), {status: 429});
  const a = Buffer.from(String(code ?? '')), b = Buffer.from(String(cfg.code));
  if (a.length === b.length && crypto.timingSafeEqual(a, b)) { failures.delete(ip); return; }
  failures.set(ip, {count: (f && Date.now() - f.at < 600e3 ? f.count : 0) + 1, at: Date.now()});
  throw Object.assign(Error('Wrong access code. Ask the administrator for the code.'), {status: 401});
}

async function body(req) {
  const chunks = [];
  let size = 0;
  for await (const chunk of req) {
    if ((size += chunk.length) > 300e3) throw Object.assign(Error('Order too large.'), {status: 413});
    chunks.push(chunk);
  }
  try { return JSON.parse(Buffer.concat(chunks).toString('utf8')); }
  catch { throw Object.assign(Error('Bad request.'), {status: 400}); }
}

const server = http.createServer(async (req, res) => {
  res.setHeader('X-Robots-Tag', 'noindex, nofollow');
  res.setHeader('X-Content-Type-Options', 'nosniff');
  res.setHeader('X-Frame-Options', 'DENY');
  const json = (status, data) => { res.writeHead(status, {'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store'}); res.end(JSON.stringify(data)); };
  try {
    const url = new URL(req.url, 'http://x');
    if (req.method === 'GET' && url.pathname === '/') {
      res.writeHead(200, {'Content-Type': 'text/html; charset=utf-8', 'Cache-Control': 'no-store'});
      const page = await fs.readFile(path.join(root, 'web.html'), 'utf8');
      // The address bar shows the changing tunnel address, so the page names the link worth saving.
      const link = /^https:\/\/[\w./-]+$/.test(cfg.link || '') ? `<p class="note">Save or share this permanent link, not the address shown above: <a href="${cfg.link}">${cfg.link}</a></p>` : '';
      res.end((cfg.code ? page : page.replace('<body>', '<body class="nocode">')).replace('<!--link-->', link));
      return;
    }
    const fileId = req.method === 'GET' && url.pathname.match(/^\/file\/(web_\d+_[0-9a-f]{16})$/);
    if (fileId) {
      const dir = path.join(root, 'outputs', fileId[1]);
      const name = (await fs.readdir(dir).catch(() => [])).find(f => f.endsWith('.xlsx'));
      if (!name) { json(404, {error: 'File not found.'}); return; }
      res.writeHead(200, {'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'Content-Disposition': `attachment; filename="${name}"`});
      res.end(await fs.readFile(path.join(dir, name)));
      return;
    }
    if (req.method === 'POST' && (url.pathname === '/api/preview' || url.pathname === '/api/order')) {
      const data = await body(req);
      allowed(req, data.code);
      if (url.pathname === '/api/preview') { json(200, review(parseOrder(orderText(data.order)))); return; }
      const nonce = one(data.nonce).slice(0, 64);
      if (!nonce) throw Error('Bad request.');
      if (!submitted.has(nonce)) {
        // Anyone with the link can order, and every order emails the administrator: cap what a stranger could send.
        const hour = Date.now() - 3600e3;
        recent = recent.filter(t => t > hour);
        if (recent.length >= 30) throw Object.assign(Error('Too many orders in the last hour. Please try again later or tell the administrator.'), {status: 429});
        recent.push(Date.now());
        // Check the order before queueing so a mistake is answered at once and may be corrected and resent.
        parseOrder(orderText(data.order));
        const job = queued(() => placeOrder(data.order));
        submitted.set(nonce, job);
        job.catch(() => submitted.delete(nonce));
      }
      json(200, await submitted.get(nonce));
      return;
    }
    json(404, {error: 'Not found.'});
  } catch (e) {
    if (!e.status) e.status = 400; // parser and form messages are written for the person ordering
    if (!res.headersSent) json(e.status, {error: e.message}); else res.end();
  }
});
// Outside access: when cloudflared.exe sits beside this file, a free Cloudflare quick tunnel publishes the form at a
// random https://….trycloudflare.com address. The address changes whenever the tunnel restarts, so it is kept in
// web-url.json, where the bot's /web command reads it. The file is removed while the tunnel is down.
const urlFile = path.join(root, 'web-url.json');
// Permanent link (user decision 2026-11-12): web.json "pages" names a local clone of the repository's gh-pages
// branch, whose index.html (web-link.html here) forwards visitors to the address in its url.json. Each new tunnel
// address is committed and pushed there with the Git credentials saved on this PC, so the GitHub Pages link in
// web.json "link" keeps working after a restart (GitHub takes about a minute to publish the change).
let current = '';
async function publish(url) {
  if (!cfg.pages) return;
  const git = (...args) => promisify(execFile)(cfg.git || 'git', args, {cwd: cfg.pages, windowsHide: true, timeout: 90000, env: {...process.env, GIT_TERMINAL_PROMPT: '0', GCM_INTERACTIVE: 'never'}});
  for (let attempt = 1; attempt <= 6 && current === url; attempt++) {
    try {
      await fs.writeFile(path.join(cfg.pages, 'url.json'), JSON.stringify({url, at: new Date().toISOString()}));
      await git('add', 'url.json');
      await git('commit', '-q', '-m', 'Update order form address');
      await git('push', '-q', 'origin', 'HEAD:gh-pages');
      console.log(`Permanent link updated. ${new Date().toISOString()}`);
      return;
    } catch (e) {
      console.error(`Permanent link update failed (attempt ${attempt}): ${String(e.stderr || e.message).replace(/\s+/g, ' ').slice(0, 300)} ${new Date().toISOString()}`);
      await new Promise(resolve => setTimeout(resolve, 30000));
    }
  }
}
async function tunnel(delay = 5000) {
  const exe = path.join(root, 'cloudflared.exe');
  await fs.rm(urlFile, {force: true});
  if (!await fs.access(exe).then(() => true, () => false)) return;
  const started = Date.now();
  const child = spawn(exe, ['tunnel', '--no-autoupdate', '--url', `http://localhost:${cfg.port}`], {windowsHide: true, stdio: ['ignore', 'ignore', 'pipe']});
  let seen = false, text = '';
  child.stderr.on('data', chunk => {
    if (seen) return;
    text = (text + chunk).slice(-4000);
    const url = text.match(/https:\/\/[a-z0-9-]+\.trycloudflare\.com/);
    if (!url) return;
    seen = true;
    fs.writeFile(urlFile, JSON.stringify({url: url[0], at: new Date().toISOString()})).catch(() => {});
    console.log(`Outside address: ${url[0]}  ${new Date().toISOString()}`);
    current = url[0];
    publish(url[0]);
  });
  const again = () => {
    fs.rm(urlFile, {force: true}).catch(() => {});
    // A tunnel that ran for a while restarts quickly; one that keeps failing backs off to 5 minutes.
    const wait = Date.now() - started > 60000 ? 5000 : Math.min(delay * 2, 300000);
    console.error(`Tunnel stopped; restarting in ${wait / 1000}s. ${new Date().toISOString()}`);
    setTimeout(() => tunnel(wait), wait);
  };
  child.once('error', again);
  child.once('exit', () => child.listenerCount('error') && (child.removeAllListeners('error'), again()));
  process.once('exit', () => child.kill());
}

server.on('error', e => { console.error(`Web order form could not start on port ${cfg.port}: ${e.code}`); process.exit(1); });
server.listen(cfg.port, '0.0.0.0', () => {
  const ips = Object.values(os.networkInterfaces()).flat().filter(x => x.family === 'IPv4' && !x.internal).map(x => x.address);
  console.log(`Web order form online: ${['localhost', ...ips].map(ip => `http://${ip}:${cfg.port}`).join('  ')}  ${new Date().toISOString()}`);
  tunnel();
});
