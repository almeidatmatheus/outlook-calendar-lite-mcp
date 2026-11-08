import { cx } from './ctx.ts';
const tok = (f: boolean) => cx().tok(f);

const B = process.env.OUTLOOK_MCP_GRAPH ?? 'https://graph.microsoft.com/v1.0';
export const TXT = { Prefer: 'outlook.body-content-type="text"' };
// id imutável: não muda quando o evento muda de calendário, então pode ser salvo fora do MCP
const IMM = 'IdType="ImmutableId"';
const pref = (h: Record<string, string>) => ({ ...h, Prefer: [IMM, h.Prefer].filter(Boolean).join(', ') });
const sleep = (ms: number) => new Promise(r => setTimeout(r, ms));

async function req(m: string, p: string, b?: unknown, h: Record<string, string> = {}): Promise<Response> {
  let f = false;
  for (let i = 0; ; i++) {
    const r = await fetch(p.startsWith('http') ? p : B + p, {
      method: m,
      headers: { authorization: `Bearer ${await tok(f)}`, ...(b === undefined ? {} : { 'content-type': 'application/json' }), ...pref(h) },
      body: b === undefined ? undefined : JSON.stringify(b),
    });
    if (r.status === 401 && !f) { f = true; continue; }
    if ([429, 503, 504].includes(r.status) && i < 3) { await sleep((+(r.headers.get('retry-after') ?? 0) || 2 ** i) * 1e3); continue; }
    if (!r.ok) {
      const e = await r.json().catch(() => ({})) as any;
      throw new Error(`graph ${r.status} ${e?.error?.code ?? ''}: ${e?.error?.message ?? r.statusText}`);
    }
    return r;
  }
}

export async function g(m: string, p: string, b?: unknown, h?: Record<string, string>): Promise<any> {
  const r = await req(m, p, b, h);
  return r.status === 204 || r.headers.get('content-length') === '0' ? null : r.json().catch(() => null);
}

export const graw = async (p: string) => Buffer.from(await (await req('GET', p)).arrayBuffer());

export type BR = { method: string; url: string; body?: unknown };
export async function batch(rs: BR[]): Promise<{ status: number; body: any }[]> {
  const o: { status: number; body: any }[] = [];
  for (let i = 0; i < rs.length; i += 20) {
    const ch = rs.slice(i, i + 20);
    const r = await g('POST', '/$batch', {
      requests: ch.map((x, j) => ({ id: String(j), method: x.method, url: x.url, headers: { Prefer: IMM, ...(x.body === undefined ? {} : { 'content-type': 'application/json' }) }, ...(x.body === undefined ? {} : { body: x.body }) })),
    });
    const by = new Map<string, any>((r?.responses ?? []).map((x: any) => [x.id, x]));
    ch.forEach((_, j) => { const x = by.get(String(j)); o.push({ status: x?.status ?? 0, body: x?.body }); });
  }
  return o;
}
