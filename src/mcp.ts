import { createInterface } from 'node:readline';
import { createHash, randomBytes } from 'node:crypto';
import { als, type Ctx } from './ctx.ts';

export type S = { type?: string; description?: string; enum?: unknown[]; items?: S; properties?: Record<string, S>; required?: string[]; default?: unknown; minimum?: number; maximum?: number; oneOf?: S[] };
export type Tool = {
  name: string;
  description: string;
  inputSchema: S;
  annotations?: Record<string, unknown>;
  cf?: (a: any) => Promise<string | null>; // presente = pede confirmação; devolve o preview (null = nada a confirmar)
  run: (a: any) => Promise<string>;
};
export type Res = { uri: string; name: string; description?: string; mimeType: string; read: () => string };
export type Info = { name: string; version: string; instructions?: string; resources?: Res[] };
type Msg = { jsonrpc: '2.0'; id?: string | number | null; method?: string; params?: any; result?: any; error?: any };

const VERS = ['2025-11-25', '2025-06-18', '2025-03-26', '2024-11-05'];

// validação mínima do schema (required, tipo, enum)
function chk(s: S, v: unknown, p = 'args'): string | null {
  if (s.oneOf) return s.oneOf.some(x => !chk(x, v, p)) ? null : `${p}: formato inválido`;
  const ty = Array.isArray(v) ? 'array' : v === null ? 'null' : typeof v;
  if (!s.type) return null;
  if (s.type === 'integer' ? !Number.isInteger(v) : s.type !== ty) return `${p}: esperado ${s.type}`;
  if (s.enum && !s.enum.includes(v)) return `${p}: use ${s.enum.join('|')}`;
  if (s.type === 'array' && s.items) for (const [i, x] of (v as unknown[]).entries()) { const e = chk(s.items, x, `${p}[${i}]`); if (e) return e; }
  if (s.type === 'object') {
    const o = v as Record<string, unknown>;
    for (const r of s.required ?? []) if (o[r] === undefined || o[r] === '') return `${p}.${r}: obrigatório`;
    for (const [k, x] of Object.entries(o)) { const ps = s.properties?.[k]; if (ps && x !== undefined) { const e = chk(ps, x, `${p}.${k}`); if (e) return e; } }
  }
  return null;
}

const SEC = randomBytes(16);
const ctok = (n: string, a: any) => { const { cf, ...r } = a; return createHash('sha256').update(SEC).update(n + JSON.stringify(r, Object.keys(r).sort())).digest('hex').slice(0, 6); };

async function call(t: Tool, a: any, c: Ctx): Promise<string> {
  if (!t.cf) return t.run(a);
  const pv = await t.cf(a);
  if (pv === null) return t.run(a); // ex.: evento só na agenda do usuário, ninguém é avisado
  if (c.caps?.elicitation && c.ask) {
    const r = await c.ask('elicitation/create', {
      message: pv,
      requestedSchema: { type: 'object', properties: { confirmar: { type: 'boolean', title: 'Confirmar?' } }, required: ['confirmar'] },
    });
    return r?.action === 'accept' && r?.content?.confirmar === true ? t.run(a) : 'cancelado pelo usuário';
  }
  const k = ctok(t.name, a);
  if (a.cf === k) return t.run(a);
  return `${pv}\n\nCONFIRMAÇÃO NECESSÁRIA: mostre isto ao usuário; se ele aprovar, repita a chamada idêntica com cf="${k}".`;
}

// processa uma mensagem JSON-RPC no contexto do usuário; null = sem resposta (notificação)
export function handler(tools: Tool[], info: Info, hooks: { onInit?: (c: Ctx) => void } = {}) {
  const T = new Map(tools.map(t => [t.name, t]));
  const cfp: S = { type: 'string', description: 'código de confirmação devolvido pela 1ª chamada' };
  const list = tools.map(t => ({
    name: t.name,
    description: t.description,
    inputSchema: t.cf ? { ...t.inputSchema, properties: { ...t.inputSchema.properties, cf: cfp } } : t.inputSchema,
    ...(t.annotations ? { annotations: t.annotations } : {}),
  }));
  const H: Record<string, (p: any, c: Ctx) => Promise<unknown> | unknown> = {
    initialize: (p, c) => {
      c.caps = p?.capabilities ?? {};
      hooks.onInit?.(c);
      return { protocolVersion: VERS.includes(p?.protocolVersion) ? p.protocolVersion : VERS[1], capabilities: { tools: {}, ...(info.resources?.length ? { resources: {} } : {}) }, serverInfo: { name: info.name, version: info.version }, ...(info.instructions ? { instructions: info.instructions } : {}) };
    },
    ping: () => ({}),
    'tools/list': () => ({ tools: list }),
    'resources/list': () => ({ resources: (info.resources ?? []).map(({ read, ...r }) => r) }),
    'resources/read': p => {
      const r = info.resources?.find(x => x.uri === p?.uri);
      if (!r) throw Object.assign(new Error(`recurso desconhecido: ${p?.uri}`), { code: -32002 });
      return { contents: [{ uri: r.uri, mimeType: r.mimeType, text: r.read() }] };
    },
    'prompts/list': () => ({ prompts: [] }),
    'tools/call': async (p, c) => {
      const t = T.get(p?.name);
      if (!t) throw Object.assign(new Error(`tool desconhecida: ${p?.name}`), { code: -32602 });
      const a = p?.arguments ?? {};
      const e = chk(t.inputSchema, a);
      if (e) return { content: [{ type: 'text', text: `erro: ${e}` }], isError: true };
      try { return { content: [{ type: 'text', text: await call(t, a, c) }] }; }
      catch (x: any) { return { content: [{ type: 'text', text: `erro: ${x?.message ?? x}` }], isError: true }; }
    },
  };
  return async (m: Msg, c: Ctx): Promise<Msg | null> => {
    if (!m || m.jsonrpc !== '2.0' || typeof m.method !== 'string') return m?.id !== undefined ? { jsonrpc: '2.0', id: m.id ?? null, error: { code: -32600, message: 'invalid request' } } : null;
    if (m.id === undefined) return null;
    const h = H[m.method];
    if (!h) return { jsonrpc: '2.0', id: m.id, error: { code: -32601, message: `método não suportado: ${m.method}` } };
    try { return { jsonrpc: '2.0', id: m.id, result: await als.run(c, () => h(m.params, c)) }; }
    catch (x: any) { return { jsonrpc: '2.0', id: m.id, error: { code: x?.code ?? -32603, message: String(x?.message ?? x) } }; }
  };
}

// transporte stdio (Claude Desktop / Code)
export function stdio(tools: Tool[], info: Info, c: Ctx, hooks: { onInit?: (c: Ctx) => void } = {}) {
  const out = (o: unknown) => process.stdout.write(JSON.stringify(o) + '\n');
  const pend = new Map<string, (r: Msg) => void>();
  let sid = 0;
  c.ask = (method, params) => new Promise((ok, ko) => {
    const id = 's' + ++sid;
    const t = setTimeout(() => { pend.delete(id); ko(new Error('sem resposta do cliente')); }, 6e5);
    pend.set(id, r => { clearTimeout(t); r.error ? ko(new Error(r.error.message)) : ok(r.result); });
    out({ jsonrpc: '2.0', id, method, params });
  });
  const h = handler(tools, info, hooks);
  const rl = createInterface({ input: process.stdin });
  rl.on('line', async l => {
    if (!l.trim()) return;
    let m: Msg;
    try { m = JSON.parse(l); } catch { return out({ jsonrpc: '2.0', id: null, error: { code: -32700, message: 'parse error' } }); }
    if (!m.method && m.id !== undefined) { const f = pend.get(String(m.id)); if (f) { pend.delete(String(m.id)); f(m); } return; }
    const r = await h(m, c);
    if (r) out(r);
  });
  rl.on('close', () => process.exit(0));
}
