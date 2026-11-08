import { readFileSync, writeFileSync, mkdirSync, rmSync } from 'node:fs';
import { createServer, type Server } from 'node:http';
import { createHash, randomBytes } from 'node:crypto';
import { spawn } from 'node:child_process';
import { homedir } from 'node:os';
import { join } from 'node:path';

const E = process.env;
export const CID = E.OUTLOOK_MCP_CLIENT_ID ?? '';
const SECRET = E.OUTLOOK_MCP_CLIENT_SECRET ?? '';
const A = `${E.OUTLOOK_MCP_AUTH ?? 'https://login.microsoftonline.com'}/${E.OUTLOOK_MCP_TENANT ?? 'common'}/oauth2/v2.0`;
export const SC = 'offline_access User.Read Calendars.ReadWrite MailboxSettings.Read';
export const DIR = E.OUTLOOK_MCP_HOME ?? join(homedir(), '.outlook-calendar-lite-mcp');
const F = join(DIR, `${E.OUTLOOK_MCP_PROFILE ?? 'default'}.json`);

export type T = { at: string; exp: number; rt: string };
type DC = { device_code: string; user_code: string; verification_uri: string; interval: number; expires_in: number; message: string };

export class AuthNeeded extends Error {}
export const need = () => { if (!CID) throw new Error('OUTLOOK_MCP_CLIENT_ID ausente (client id do app Azure)'); };
const sleep = (ms: number) => new Promise(r => setTimeout(r, ms));
export const b64u = (b: Buffer) => b.toString('base64url');
export const rnd = (n = 32) => b64u(randomBytes(n));
export const s256 = (v: string) => b64u(createHash('sha256').update(v).digest());

// ---- Microsoft (usado pelos dois modos)
export async function msPost(p: string, b: Record<string, string>): Promise<any> {
  if (SECRET && b.grant_type !== 'urn:ietf:params:oauth:grant-type:device_code') b.client_secret = SECRET;
  const r = await fetch(A + p, { method: 'POST', headers: { 'content-type': 'application/x-www-form-urlencoded' }, body: new URLSearchParams(b) });
  return r.json();
}
export const msAuthUrl = (redirect: string, challenge: string, state: string) =>
  `${A}/authorize?${new URLSearchParams({ client_id: CID, response_type: 'code', redirect_uri: redirect, scope: SC, code_challenge: challenge, code_challenge_method: 'S256', state, prompt: 'select_account' })}`;
export const msCode = (code: string, verifier: string, redirect: string) =>
  msPost('/token', { grant_type: 'authorization_code', client_id: CID, code, code_verifier: verifier, redirect_uri: redirect, scope: SC });
export const msRefresh = (rt: string) => msPost('/token', { grant_type: 'refresh_token', client_id: CID, refresh_token: rt, scope: SC });
export const toT = (r: any, old?: T | null): T => ({ at: r.access_token, exp: Date.now() + (r.expires_in ?? 3600) * 1e3, rt: r.refresh_token ?? old?.rt ?? '' });
const fail = (r: any) => new Error(`login: ${r?.error ?? 'erro'} ${r?.error_description ?? ''}`.trim());

export const page = (t: string) => `<!doctype html><meta charset="utf-8"><title>outlook-calendar-lite-mcp</title><body style="font:16px system-ui;max-width:420px;margin:15vh auto;text-align:center"><h2>${t}</h2><p>Pode fechar esta aba.</p></body>`;

// ---- modo local (stdio): token em arquivo
let t: T | null = (() => { try { return JSON.parse(readFileSync(F, 'utf8')) as T; } catch { return null; } })();
let pend: { msg: string; p: Promise<void> } | null = null;

function save(r: any) {
  t = toT(r, t);
  mkdirSync(DIR, { recursive: true, mode: 0o700 });
  writeFileSync(F, JSON.stringify(t), { mode: 0o600 });
}

function open(url: string) {
  if (E.OUTLOOK_MCP_OPEN === 'test') { fetch(url).catch(() => {}); return true; } // testes: segue os redirects sozinho
  const [c, a] = process.platform === 'darwin' ? ['open', [url]] : process.platform === 'win32' ? ['cmd', ['/c', 'start', '""', url.replace(/&/g, '^&')]] : ['xdg-open', [url]];
  try { spawn(c, a as string[], { detached: true, stdio: 'ignore' }).on('error', () => {}).unref(); return true; } catch { return false; }
}

// auth code + PKCE com redirect em localhost: abre a escolha de conta no navegador
async function browser(): Promise<{ url: string; p: Promise<void> }> {
  const v = rnd(), st = rnd(16);
  let done!: (e?: Error) => void, redir = '';
  const p = new Promise<void>((ok, ko) => { done = e => (e ? ko(e) : ok()); });
  const srvs: Server[] = [];
  const close = () => srvs.forEach(s => s.close());
  const h = async (q: any, w: any) => {
    const u = new URL(q.url ?? '/', 'http://localhost');
    if (u.pathname !== '/cb') { w.writeHead(404).end(); return; }
    const end = (m: string, e?: Error) => { w.writeHead(200, { 'content-type': 'text/html; charset=utf-8' }).end(page(m)); clearTimeout(to); close(); done(e); };
    if (u.searchParams.get('state') !== st) return end('Falha no login', new Error('login: state inválido'));
    if (!u.searchParams.get('code')) return end('Login cancelado', fail({ error: u.searchParams.get('error'), error_description: u.searchParams.get('error_description') }));
    const r = await msCode(u.searchParams.get('code')!, v, redir).catch(e => ({ error: String(e) }));
    if (!r.access_token) return end('Falha no login', fail(r));
    save(r);
    end('Conta conectada ao outlook-calendar-lite-mcp');
  };
  const s4 = createServer(h); srvs.push(s4);
  await new Promise<void>(r => s4.listen(0, '127.0.0.1', () => r()));
  const port = (s4.address() as any).port;
  const s6 = createServer(h); srvs.push(s6); // localhost pode resolver para ::1
  s6.on('error', () => {}).listen(port, '::1');
  redir = `http://localhost:${port}/cb`;
  const to = setTimeout(() => { close(); done(new Error('login: tempo esgotado')); }, 10 * 6e4);
  return { url: msAuthUrl(redir, s256(v), st), p };
}

async function deviceStart(): Promise<DC> {
  const d = await msPost('/devicecode', { client_id: CID, scope: SC });
  if (!d.device_code) throw fail(d);
  return d;
}
async function devicePoll(d: DC) {
  let iv = (d.interval || 5) * 1e3;
  const end = Date.now() + d.expires_in * 1e3;
  while (Date.now() < end) {
    await sleep(iv);
    const r = await msPost('/token', { grant_type: 'urn:ietf:params:oauth:grant-type:device_code', client_id: CID, device_code: d.device_code });
    if (r.access_token) return save(r);
    if (r.error === 'authorization_pending') continue;
    if (r.error === 'slow_down') { iv += 5e3; continue; }
    throw fail(r);
  }
  throw new Error('login: código expirou');
}

// inicia o login se ainda não há conta; não bloqueia. devolve a mensagem para o usuário
export async function startLogin(): Promise<string> {
  need();
  if (pend) return pend.msg;
  let msg: string, p: Promise<void>;
  if (E.OUTLOOK_MCP_LOGIN !== 'device') {
    const b = await browser();
    const opened = open(b.url);
    msg = `${opened ? 'abri o navegador para conectar a conta Microsoft. se não abriu, acesse' : 'para conectar a conta Microsoft, acesse'}: ${b.url}\ndepois de escolher a conta, repita a chamada.`;
    p = b.p;
  } else {
    const d = await deviceStart();
    msg = `login necessário: abra ${d.verification_uri} e digite o código ${d.user_code}. depois repita a chamada.`;
    p = devicePoll(d);
  }
  pend = { msg, p: p.finally(() => { pend = null; }) };
  pend.p.catch(e => process.stderr.write(`${e.message}\n`));
  return msg;
}

export const hasLocal = () => !!t?.rt;

export async function tok(force = false): Promise<string> {
  need();
  if (t && !force && t.exp > Date.now() + 6e4) return t.at;
  if (t?.rt) {
    const r = await msRefresh(t.rt);
    if (r.access_token) { save(r); return t!.at; }
  }
  throw new AuthNeeded(await startLogin());
}

export async function login() {
  need();
  const m = await startLogin();
  process.stderr.write(m.split('\n')[0] + '\n');
  await pend?.p;
  process.stderr.write(t ? 'ok\n' : 'falhou\n');
}

export function logout() { rmSync(F, { force: true }); t = null; }
export async function relogin() { logout(); return startLogin(); }
