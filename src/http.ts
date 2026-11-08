// modo remoto (claude.ai): MCP Streamable HTTP + servidor OAuth que repassa o login para a Microsoft
import { createServer, type IncomingMessage, type ServerResponse } from 'node:http';
import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';
import { join } from 'node:path';
import { handler, type Tool, type Info } from './mcp.ts';
import { newIds, type Ctx } from './ctx.ts';
import { DIR, need, rnd, s256, msAuthUrl, msCode, msRefresh, toT, page, AuthNeeded, type T } from './auth.ts';

const E = process.env;
const PORT = +(E.PORT ?? 8787);
const BASE = (E.OUTLOOK_MCP_URL ?? `http://localhost:${PORT}`).replace(/\/$/, '');
const CB = BASE + '/ms/cb';
const ALLOW = (E.OUTLOOK_MCP_ALLOW ?? '').toLowerCase().split(',').map(s => s.trim()).filter(Boolean);
const AT_TTL = 3600, RT_TTL = 90 * 864e2;
const FS = join(DIR, 'http-store.json');

type Store = {
  clients: Record<string, { redirect_uris: string[]; name?: string }>;
  users: Record<string, { t: T; mail: string }>;
  at: Record<string, { uid: string; exp: number }>;
  rt: Record<string, { uid: string; cid: string; exp: number }>;
};
const S: Store = (() => { try { return JSON.parse(readFileSync(FS, 'utf8')); } catch { return { clients: {}, users: {}, at: {}, rt: {} }; } })();
function persist() {
  const n = Date.now();
  for (const k in S.at) if (S.at[k].exp < n) delete S.at[k];
  for (const k in S.rt) if (S.rt[k].exp < n) delete S.rt[k];
  mkdirSync(DIR, { recursive: true, mode: 0o700 });
  writeFileSync(FS, JSON.stringify(S), { mode: 0o600 });
}

const pendA = new Map<string, { cid: string; redir: string; ch: string; state: string; v: string; exp: number }>();
const codes = new Map<string, { cid: string; redir: string; ch: string; uid: string; exp: number }>();
const ctxs = new Map<string, Ctx>();

function ctxOf(uid: string): Ctx {
  let c = ctxs.get(uid);
  if (c) return c;
  c = {
    ids: newIds(), c: {}, caps: {},
    async tok(force = false) {
      const u = S.users[uid];
      if (!u) throw new AuthNeeded('conta desconectada; reconecte o conector');
      if (!force && u.t.exp > Date.now() + 6e4) return u.t.at;
      const r = await msRefresh(u.t.rt);
      if (!r.access_token) throw new AuthNeeded('sessão Microsoft expirou; desconecte e reconecte o conector');
      u.t = toT(r, u.t); persist();
      return u.t.at;
    },
  };
  ctxs.set(uid, c);
  return c;
}

const okRedir = (u: string) => { try { const x = new URL(u); return x.protocol === 'https:' || ['localhost', '127.0.0.1', '[::1]'].includes(x.hostname); } catch { return false; } };

// cliente registrado (DCR) ou client id metadata document (client_id = URL https)
async function client(cid: string): Promise<{ redirect_uris: string[] } | null> {
  if (S.clients[cid]) return S.clients[cid];
  if (!/^https:\/\//.test(cid)) return null;
  try {
    const r = await fetch(cid, { signal: AbortSignal.timeout(5000) });
    const d = await r.json() as any;
    return d?.client_id === cid && Array.isArray(d.redirect_uris) ? { redirect_uris: d.redirect_uris } : null;
  } catch { return null; }
}

const read = (r: IncomingMessage) => new Promise<string>((ok, ko) => {
  let s = '';
  r.on('data', c => { s += c; if (s.length > 1e6) { ko(new Error('corpo grande demais')); r.destroy(); } });
  r.on('end', () => ok(s)); r.on('error', ko);
});
const json = (w: ServerResponse, s: number, b: unknown, h: Record<string, string> = {}) => { w.writeHead(s, { 'content-type': 'application/json', 'cache-control': 'no-store', ...h }); w.end(JSON.stringify(b)); };
const html = (w: ServerResponse, s: number, t: string) => { w.writeHead(s, { 'content-type': 'text/html; charset=utf-8' }); w.end(page(t)); };
const go = (w: ServerResponse, u: string) => { w.writeHead(302, { location: u }); w.end(); };
const withQ = (u: string, q: Record<string, string>) => { const x = new URL(u); for (const [k, v] of Object.entries(q)) x.searchParams.set(k, v); return x.toString(); };
const form = (raw: string, ct = '') => (ct.includes('json') ? JSON.parse(raw || '{}') : Object.fromEntries(new URLSearchParams(raw)));

const PRM = { resource: BASE + '/mcp', authorization_servers: [BASE], bearer_methods_supported: ['header'], resource_name: 'outlook-calendar-lite-mcp' };
const ASM = {
  issuer: BASE, authorization_endpoint: BASE + '/authorize', token_endpoint: BASE + '/token', registration_endpoint: BASE + '/register',
  response_types_supported: ['code'], grant_types_supported: ['authorization_code', 'refresh_token'], code_challenge_methods_supported: ['S256'],
  token_endpoint_auth_methods_supported: ['none'], client_id_metadata_document_supported: true,
};

export function serveHttp(tools: Tool[], info: Info) {
  need();
  const h = handler(tools, info);
  const srv = createServer(async (r, w) => {
    const u = new URL(r.url ?? '/', BASE), p = u.pathname, q = u.searchParams;
    try {
      if (r.method === 'OPTIONS') { w.writeHead(204, { 'access-control-allow-origin': '*', 'access-control-allow-headers': '*', 'access-control-allow-methods': 'GET,POST,DELETE' }); return w.end(); }
      if (p.startsWith('/.well-known/oauth-protected-resource')) return json(w, 200, PRM);
      if (p === '/.well-known/oauth-authorization-server' || p === '/.well-known/openid-configuration') return json(w, 200, ASM);

      if (p === '/register' && r.method === 'POST') {
        const b = form(await read(r), 'json');
        const ru: string[] = Array.isArray(b.redirect_uris) ? b.redirect_uris : [];
        if (!ru.length || !ru.every(okRedir)) return json(w, 400, { error: 'invalid_redirect_uri' });
        const cid = rnd(16);
        S.clients[cid] = { redirect_uris: ru, name: String(b.client_name ?? '') };
        persist();
        return json(w, 201, { client_id: cid, client_id_issued_at: Math.floor(Date.now() / 1e3), redirect_uris: ru, client_name: b.client_name, token_endpoint_auth_method: 'none', grant_types: ['authorization_code', 'refresh_token'], response_types: ['code'] });
      }

      if (p === '/authorize' && r.method === 'GET') {
        const cid = q.get('client_id') ?? '', redir = q.get('redirect_uri') ?? '', cl = await client(cid);
        if (!cl) return html(w, 400, 'Cliente desconhecido');
        if (!cl.redirect_uris.includes(redir)) return html(w, 400, 'redirect_uri não registrado');
        const state = q.get('state') ?? '', ch = q.get('code_challenge') ?? '';
        if (q.get('response_type') !== 'code' || q.get('code_challenge_method') !== 'S256' || !ch) return go(w, withQ(redir, { error: 'invalid_request', error_description: 'requer code + PKCE S256', ...(state ? { state } : {}) }));
        const pid = rnd(16), v = rnd();
        pendA.set(pid, { cid, redir, ch, state, v, exp: Date.now() + 6e5 });
        return go(w, msAuthUrl(CB, s256(v), pid));
      }

      if (p === '/ms/cb' && r.method === 'GET') {
        const pa = pendA.get(q.get('state') ?? '');
        if (!pa || pa.exp < Date.now()) return html(w, 400, 'Login expirado, tente conectar de novo');
        pendA.delete(q.get('state')!);
        const back = (e: string, d: string) => go(w, withQ(pa.redir, { error: e, error_description: d, ...(pa.state ? { state: pa.state } : {}) }));
        if (!q.get('code')) return back('access_denied', q.get('error_description') ?? 'login cancelado');
        const t = await msCode(q.get('code')!, pa.v, CB);
        if (!t.access_token) return back('server_error', `microsoft: ${t.error ?? ''}`);
        const me = await (await fetch((E.OUTLOOK_MCP_GRAPH ?? 'https://graph.microsoft.com/v1.0') + '/me?$select=mail,userPrincipalName', { headers: { authorization: `Bearer ${t.access_token}` } })).json() as any;
        const mail = String(me?.mail || me?.userPrincipalName || '').toLowerCase();
        if (!mail) return back('server_error', 'não foi possível ler a conta');
        if (ALLOW.length && !ALLOW.some(a => mail === a || mail.endsWith('@' + a))) return back('access_denied', 'conta não autorizada neste servidor');
        S.users[mail] = { t: toT(t, S.users[mail]?.t), mail };
        ctxs.delete(mail);
        persist();
        const code = rnd();
        codes.set(code, { cid: pa.cid, redir: pa.redir, ch: pa.ch, uid: mail, exp: Date.now() + 6e4 });
        return go(w, withQ(pa.redir, { code, ...(pa.state ? { state: pa.state } : {}), iss: BASE }));
      }

      if (p === '/token' && r.method === 'POST') {
        const b = form(await read(r), r.headers['content-type'] ?? '');
        const issue = (uid: string, cid: string) => {
          const at = rnd(), rt = rnd(), n = Date.now();
          S.at[at] = { uid, exp: n + AT_TTL * 1e3 };
          S.rt[rt] = { uid, cid, exp: n + RT_TTL * 1e3 };
          persist();
          return json(w, 200, { access_token: at, token_type: 'Bearer', expires_in: AT_TTL, refresh_token: rt });
        };
        if (b.grant_type === 'authorization_code') {
          const c = codes.get(b.code ?? '');
          codes.delete(b.code ?? '');
          if (!c || c.exp < Date.now() || c.cid !== b.client_id || c.redir !== b.redirect_uri || s256(String(b.code_verifier ?? '')) !== c.ch) return json(w, 400, { error: 'invalid_grant' });
          return issue(c.uid, c.cid);
        }
        if (b.grant_type === 'refresh_token') {
          const x = S.rt[b.refresh_token ?? ''];
          if (!x || x.exp < Date.now() || (b.client_id && b.client_id !== x.cid) || !S.users[x.uid]) return json(w, 400, { error: 'invalid_grant' });
          delete S.rt[b.refresh_token];
          return issue(x.uid, x.cid);
        }
        return json(w, 400, { error: 'unsupported_grant_type' });
      }

      if (p === '/mcp') {
        const tk = /^Bearer (.+)$/i.exec(r.headers.authorization ?? '')?.[1] ?? '';
        const a = S.at[tk];
        if (!a || a.exp < Date.now() || !S.users[a.uid]) return json(w, 401, { error: 'invalid_token' }, { 'www-authenticate': `Bearer resource_metadata="${BASE}/.well-known/oauth-protected-resource", error="invalid_token"` });
        if (r.method === 'GET') { w.writeHead(405, { allow: 'POST, DELETE' }); return w.end(); }
        if (r.method === 'DELETE') { w.writeHead(200); return w.end(); }
        if (r.method !== 'POST') { w.writeHead(405); return w.end(); }
        let m: any;
        try { m = JSON.parse(await read(r)); } catch { return json(w, 400, { jsonrpc: '2.0', id: null, error: { code: -32700, message: 'parse error' } }); }
        const c = ctxOf(a.uid), arr = Array.isArray(m);
        const out = (await Promise.all((arr ? m : [m]).map((x: any) => h(x, c)))).filter(Boolean);
        if (!out.length) { w.writeHead(202); return w.end(); }
        const init = (arr ? m : [m]).some((x: any) => x?.method === 'initialize');
        return json(w, 200, arr ? out : out[0], init ? { 'mcp-session-id': rnd(16) } : {});
      }

      if (p === '/' && r.method === 'GET') { w.writeHead(200, { 'content-type': 'text/plain' }); return w.end(`outlook-calendar-lite-mcp: conecte em ${BASE}/mcp`); }
      w.writeHead(404); w.end();
    } catch (e: any) {
      process.stderr.write(`http ${p}: ${e?.message ?? e}\n`);
      if (!w.headersSent) json(w, 500, { error: 'server_error' });
    }
  });
  srv.listen(PORT, () => process.stderr.write(`outlook-calendar-lite-mcp http em :${PORT} — URL do conector: ${BASE}/mcp\n`));
  return srv;
}
