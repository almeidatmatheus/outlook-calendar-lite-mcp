import type { S } from './mcp.ts';
import { g } from './g.ts';

import { cx } from './ctx.ts';
import { cell, tsv, nm } from './fmt.ts';

// helpers compartilhados pelas ferramentas de calendário
export const norm = (s: string) => s.normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().trim().replace(/\s+/g, '_');

// ---- conta, fuso e horário de trabalho
// tz = nome IANA (contas de data e exibição); tzG = fuso como a caixa informa (é o que o Graph aceita com certeza)
type MS = { tz: string; tzG: string; dias: string[]; ini: string; fim: string };
const valido = (z?: string) => { if (!z) return false; try { new Intl.DateTimeFormat('en-US', { timeZone: z }); return true; } catch { return false; } };
// fuso do Windows ("E. South America Standard Time") -> IANA ("America/Sao_Paulo"): as duas listas do Graph têm o mesmo rótulo
async function iana(win: string) {
  const [w, i] = await Promise.all([g('GET', '/me/outlook/supportedTimeZones'), g('GET', "/me/outlook/supportedTimeZones(TimeZoneStandard=microsoft.graph.timeZoneStandard'Iana')")]);
  const rot = (w?.value ?? []).find((x: any) => x.alias === win)?.displayName;
  return (i?.value ?? []).map((x: any) => x.alias).find((a: string, k: number) => i.value[k].displayName === rot && valido(a));
}
export async function ms(): Promise<MS> {
  const C = cx().c;
  if (!C.ms) {
    const r = await g('GET', '/me/mailboxSettings').catch(() => null);
    const w = r?.workingHours ?? {}, raw = String(r?.timeZone ?? '');
    const env = process.env.OUTLOOK_MCP_TZ, local = Intl.DateTimeFormat().resolvedOptions().timeZone;
    let tz = valido(env) ? env! : valido(raw) ? raw : '';
    if (!tz && raw) tz = (await iana(raw).catch(() => '')) ?? '';
    tz ||= valido(local) ? local : 'UTC';
    C.ms = {
      tz, tzG: raw || tz,
      dias: w.daysOfWeek?.length ? w.daysOfWeek : ['monday', 'tuesday', 'wednesday', 'thursday', 'friday'],
      ini: String(w.startTime ?? '09:00').slice(0, 5), fim: String(w.endTime ?? '18:00').slice(0, 5),
    };
  }
  return C.ms;
}
export async function me(): Promise<{ mail: string; nome: string }> {
  const C = cx().c;
  if (!C.me) { const r = await g('GET', '/me?$select=mail,userPrincipalName,displayName'); C.me = { mail: String(r?.mail || r?.userPrincipalName || '').toLowerCase(), nome: r?.displayName ?? '' }; }
  return C.me;
}
// cabeçalho que faz o Graph devolver e entender horários no fuso do usuário
export const PH = async (texto = false) => ({ Prefer: `outlook.timezone="${(await ms()).tzG}"${texto ? ', outlook.body-content-type="text"' : ''}` });

// ---- horas: trabalha com "hora de parede" local (AAAA-MM-DDTHH:mm) e converte para UTC só nas buscas
function off(tz: string, t: number) {
  const p = Object.fromEntries(new Intl.DateTimeFormat('en-US', { timeZone: tz, hourCycle: 'h23', year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit' }).formatToParts(new Date(t)).map(x => [x.type, x.value]));
  return (Date.UTC(+p.year, +p.month - 1, +p.day, +p.hour, +p.minute, +p.second) - Math.floor(t / 1e3) * 1e3) / 6e4;
}
export const utc = (local: string, tz: string) => { const w = Date.parse(local + 'Z'); let t = w - off(tz, w) * 6e4; t = w - off(tz, t) * 6e4; return new Date(t).toISOString().replace(/\.\d{3}Z$/, 'Z'); };
export const agora = (tz: string) => { const t = Date.now(); return new Date(t + off(tz, t) * 6e4).toISOString().slice(0, 16); };
export const dias = (d: string, n: number) => new Date(Date.parse(d.slice(0, 10) + 'T00:00Z') + n * 864e5).toISOString().slice(0, 10);
export const dow = (d: string) => new Date(d.slice(0, 10) + 'T00:00Z').getUTCDay();
const DS = ['dom', 'seg', 'ter', 'qua', 'qui', 'sex', 'sab'];
export const DOW_EN = ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'];
export const dia = (d: string) => `${DS[dow(d)]} ${d.slice(5, 10)}`;
export const min = (a: string, b: string) => Math.round((Date.parse(b.slice(0, 16) + 'Z') - Date.parse(a.slice(0, 16) + 'Z')) / 6e4);
export const soma = (l: string, m: number) => new Date(Date.parse(l.slice(0, 16) + 'Z') + m * 6e4).toISOString().slice(0, 16);

// "2026-11-10 14:00" | "2026-11-10T14:00" | "2026-11-10" -> forma canônica
export function quando(v: string, campo = 'inicio'): { l: string; diaTodo: boolean } {
  const s = String(v ?? '').trim().replace(' ', 'T');
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return { l: s + 'T00:00', diaTodo: true };
  if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}(:\d{2})?$/.test(s) && !isNaN(Date.parse(s.slice(0, 16) + 'Z'))) return { l: s.slice(0, 16), diaTodo: false };
  throw new Error(`${campo} inválido: ${v} (use AAAA-MM-DDTHH:mm ou AAAA-MM-DD)`);
}

export const PER = ['hoje', 'amanha', 'semana', 'proxima_semana', '7d', '30d'];
// janela [início, fim) em hora local
export async function janela(a: { periodo?: string; de?: string; ate?: string }, padrao = 'semana'): Promise<[string, string, string]> {
  const { tz } = await ms(), hoje = agora(tz).slice(0, 10);
  if (a.de || a.ate) {
    const de = a.de ? quando(a.de, 'de').l.slice(0, 10) : hoje, ate = a.ate ? quando(a.ate, 'ate').l.slice(0, 10) : dias(de, 7);
    if (ate < de) throw new Error('ate antes de de');
    return [de + 'T00:00', dias(ate, 1) + 'T00:00', tz];
  }
  const p = a.periodo ?? padrao, w = (dow(hoje) + 6) % 7; // 0 = segunda
  const J: Record<string, [string, string]> = {
    hoje: [hoje, dias(hoje, 1)], amanha: [dias(hoje, 1), dias(hoje, 2)],
    semana: [hoje, dias(hoje, 7 - w)], proxima_semana: [dias(hoje, 7 - w), dias(hoje, 14 - w)],
    '7d': [hoje, dias(hoje, 7)], '30d': [hoje, dias(hoje, 30)],
  };
  const [x, y] = J[p] ?? J[padrao];
  return [x + 'T00:00', y + 'T00:00', tz];
}

// ---- calendários
export async function cal(nome?: string): Promise<string> {
  if (!nome) return '/me/calendar';
  const r = await g('GET', '/me/calendars?$select=id,name&$top=100');
  const c = (r?.value ?? []).find((x: any) => norm(x.name) === norm(nome));
  if (!c) throw new Error(`calendário desconhecido: ${nome} (use cx para ver os calendários)`);
  return `/me/calendars/${c.id}`;
}

// ---- eventos
export const SEL = 'id,subject,start,end,isAllDay,location,organizer,responseStatus,isOnlineMeeting,type,showAs,isCancelled,seriesMasterId,isOrganizer';
const RESP: Record<string, string> = { organizer: 'org', accepted: 'ok', tentativelyAccepted: 'tlv', declined: 'nao', notResponded: '?', none: '-' };
const DISP: Record<string, string> = { free: 'livre', tentative: 'prov', busy: 'ocup', oof: 'fora', workingElsewhere: 'remoto', unknown: '-' };
export const resp = (e: any) => (e.isOrganizer ? 'org' : RESP[e.responseStatus?.response] ?? '-');
export const disp = (e: any) => DISP[e.showAs] ?? '-';
export const loc = (e: any) => e.location?.displayName ?? '';

// "ter 11-10 14:00–15:00 (60 min)" ou "ter 11-10 dia todo"
export function hora(e: any) {
  const a = String(e.start?.dateTime ?? ''), b = String(e.end?.dateTime ?? '');
  if (e.isAllDay) { const n = min(a, b) / 1440; return `${dia(a)} dia todo${n > 1 ? ` (${n} dias, até ${dia(dias(b, -1))})` : ''}`; }
  return `${dia(a)} ${a.slice(11, 16)}–${a.slice(0, 10) === b.slice(0, 10) ? '' : dia(b) + ' '}${b.slice(11, 16)} (${min(a, b)} min)`;
}

export const H = ['id', 'dia', 'ini', 'fim', 'titulo', 'local', 'org', 'resp', 'on', 'rec', 'disp'];
export function table(es: any[]) {
  const M: Map<string, any> = (cx().c.ev ??= new Map());
  return tsv(H, es.map(e => {
    M.set(e.id, e);
    const a = String(e.start?.dateTime ?? ''), b = String(e.end?.dateTime ?? '');
    return [e.id, dia(a), e.isAllDay ? 'dia todo' : a.slice(11, 16), e.isAllDay ? (min(a, b) > 1440 ? dia(dias(b, -1)) : '-') : b.slice(11, 16),
      cell(`${e.isCancelled ? '[cancelado] ' : ''}${e.subject ?? ''}`, 60), cell(loc(e), 30), e.isOrganizer ? 'eu' : cell(nm(e.organizer), 20), resp(e), e.isOnlineMeeting ? 1 : 0, e.type === 'occurrence' || e.type === 'exception' ? 1 : 0, disp(e)];
  }));
}

// calendarView paginado (Graph limita por página)
export async function view(base: string, de: string, ate: string, tz: string, max: number, sel = SEL) {
  let u: string | undefined = `${base}/calendarView?startDateTime=${utc(de, tz)}&endDateTime=${utc(ate, tz)}&$select=${sel}&$orderby=${encodeURIComponent('start/dateTime')}&$top=${Math.min(max, 100)}`;
  const o: any[] = [];
  let mais = false;
  while (u && o.length < max) {
    const r = await g('GET', u, undefined, await PH());
    o.push(...(r?.value ?? []));
    u = r?.['@odata.nextLink'];
  }
  if (u || o.length > max) mais = true;
  return { es: o.slice(0, max), mais };
}

export async function ev(id: string, extra = '') {
  return g('GET', `/me/events/${id}?$select=${SEL},attendees${extra}`, undefined, await PH(!!extra));
}

export const rcp = (xs: unknown) => {
  const v = (Array.isArray(xs) ? xs : xs ? String(xs).split(/\s*[,;]\s*/) : []).map(String).filter(Boolean);
  for (const a of v) if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(a)) throw new Error(`email inválido: ${a}`);
  return v;
};
export const att = (xs: string[], type: 'required' | 'optional' = 'required') => xs.map(address => ({ emailAddress: { address }, type }));

// recorrência simples a partir do dia de início
export const REC = ['diaria', 'dias_uteis', 'semanal', 'mensal', 'anual'];
export function rec(tipo: string, ini: string, ate?: string, vezes?: number) {
  const d = ini.slice(0, 10), w = DOW_EN[dow(d)], dd = +d.slice(8, 10), mm = +d.slice(5, 7);
  const pattern = ({
    diaria: { type: 'daily', interval: 1 },
    dias_uteis: { type: 'weekly', interval: 1, daysOfWeek: DOW_EN.slice(1, 6), firstDayOfWeek: 'monday' },
    semanal: { type: 'weekly', interval: 1, daysOfWeek: [w], firstDayOfWeek: 'monday' },
    mensal: { type: 'absoluteMonthly', interval: 1, dayOfMonth: dd },
    anual: { type: 'absoluteYearly', interval: 1, dayOfMonth: dd, month: mm },
  } as Record<string, unknown>)[tipo];
  const range = ate ? { type: 'endDate', startDate: d, endDate: quando(ate, 'rec_ate').l.slice(0, 10) } : vezes ? { type: 'numbered', startDate: d, numberOfOccurrences: vezes } : { type: 'noEnd', startDate: d };
  return { pattern, range };
}
export function recTxt(r: any) {
  if (!r?.pattern) return '';
  const p = r.pattern, n = r.range;
  const base = ({ daily: 'diária', weekly: `semanal (${(p.daysOfWeek ?? []).map((x: string) => DS[DOW_EN.indexOf(x)]).join(',')})`, absoluteMonthly: `mensal (dia ${p.dayOfMonth})`, relativeMonthly: 'mensal', absoluteYearly: 'anual', relativeYearly: 'anual' } as Record<string, string>)[p.type] ?? p.type;
  return `${base}${p.interval > 1 ? ` a cada ${p.interval}` : ''}${n?.type === 'endDate' ? ` até ${n.endDate}` : n?.type === 'numbered' ? `, ${n.numberOfOccurrences} vezes` : ''}`;
}

// ---- schemas
export const str = (description: string, x: Partial<S> = {}): S => ({ type: 'string', description, ...x });
export const ALVO = str('id do evento (coluna id de ls/find)');
export const obj = (properties: Record<string, S>, required: string[] = []): S => ({ type: 'object', properties, required });
export const RO = { readOnlyHint: true, openWorldHint: false };
export const LISTA_EMAILS: S = { type: 'array', items: { type: 'string' }, description: 'emails' };
