// login Microsoft + Graph de calendário simulados, para testes sem conta real
import { createServer, type IncomingMessage, type ServerResponse } from 'node:http';

type M = Record<string, any>;
export const calls: { m: string; p: string; b?: any; prefer?: string }[] = [];
export const subs: M[] = []; // assinaturas de notificação recebidas
export const ctl = { patch404: false }; // força o PATCH a falhar como assinatura vencida
const TZ = 'E. South America Standard Time', OFF = -3; // a caixa informa o fuso no formato do Windows // São Paulo não tem horário de verão
let n = 0, tokTries = 0, throttle = 1;
const nid = () => 'AAMkAGI2EV' + 'x'.repeat(120) + ++n;
const loc = (t = Date.now()) => new Date(t + OFF * 36e5).toISOString().slice(0, 16);
const HOJE = loc().slice(0, 10);
export const dia = (k: number) => new Date(Date.parse(HOJE + 'T00:00Z') + k * 864e5).toISOString().slice(0, 10);
const dt = (d: string, h: string) => ({ dateTime: `${d}T${h}:00.0000000`, timeZone: TZ });
const pe = (name: string, address: string) => ({ emailAddress: { name, address } });
const EU = pe('Eu', 'eu@minha.com');
const at = (p: any, response = 'none', type = 'required') => ({ ...p, type, status: { response } });

export const evs: M[] = [
  { id: 'M1', subject: 'Daily do time', type: 'seriesMaster', start: dt(dia(-7), '10:00'), end: dt(dia(-7), '10:15'), recurrence: { pattern: { type: 'weekly', interval: 1, daysOfWeek: ['monday', 'wednesday'] }, range: { type: 'noEnd', startDate: dia(-7) } }, isOrganizer: true, organizer: EU, attendees: [at(pe('Ana Lima', 'ana@acme.com'), 'accepted')], hidden: true },
  { id: nid(), subject: 'Daily do time', type: 'occurrence', seriesMasterId: 'M1', start: dt(HOJE, '10:00'), end: dt(HOJE, '10:15'), isOrganizer: true, organizer: EU, responseStatus: { response: 'organizer' }, isOnlineMeeting: true, onlineMeeting: { joinUrl: 'https://teams.microsoft.com/l/meetup-join/abc' }, showAs: 'busy', attendees: [at(pe('Ana Lima', 'ana@acme.com'), 'accepted')], body: { content: 'Pauta: bloqueios do dia' } },
  { id: nid(), subject: 'Reunião com cliente ACME', type: 'singleInstance', start: dt(HOJE, '14:00'), end: dt(HOJE, '15:00'), location: { displayName: 'Sala 3' }, isOrganizer: false, organizer: pe('Ana Lima', 'ana@acme.com'), responseStatus: { response: 'notResponded' }, showAs: 'tentative', attendees: [at(EU), at(pe('Bob Silva', 'bob@acme.com'), 'accepted', 'optional')], body: { content: 'Apresentar a proposta.\n\nEsta mensagem é confidencial e destinada apenas ao destinatário autorizado.' } },
  { id: nid(), subject: 'Feriado', type: 'singleInstance', isAllDay: true, start: dt(dia(1), '00:00'), end: dt(dia(2), '00:00'), isOrganizer: true, organizer: EU, responseStatus: { response: 'organizer' }, showAs: 'free', attendees: [] },
  { id: nid(), subject: 'Dentista', type: 'singleInstance', start: dt(dia(1), '09:00'), end: dt(dia(1), '09:30'), isOrganizer: true, organizer: EU, responseStatus: { response: 'organizer' }, showAs: 'busy', attendees: [] },
];
const find = (id: string) => evs.find(e => e.id === id);
const pub = (e: M) => { const { hidden, cal, ...r } = e; return r; };
const L = (s: string) => s.slice(0, 16);
// UTC da busca -> hora local do mock
const toLoc = (u: string) => loc(Date.parse(u));

function graph(m: string, url: URL, b: any): [number, any] {
  const p = url.pathname.replace(/^\/v1\.0/, ''), q = url.searchParams;
  let x: RegExpMatchArray | null;
  if (m === 'GET' && p === '/me') return [200, { displayName: 'Eu', mail: 'eu@minha.com' }];
  if (m === 'GET' && p === '/me/mailboxSettings') return [200, { timeZone: TZ, workingHours: { daysOfWeek: ['monday', 'tuesday', 'wednesday', 'thursday', 'friday'], startTime: '09:00:00.0000000', endTime: '18:00:00.0000000', timeZone: { name: TZ } } }];
  if (m === 'GET' && p === "/me/outlook/supportedTimeZones(TimeZoneStandard=microsoft.graph.timeZoneStandard'Iana')") return [200, { value: [{ alias: 'America/New_York', displayName: '(UTC-05:00) Eastern Time (US & Canada)' }, { alias: 'America/Sao_Paulo', displayName: '(UTC-03:00) Brasilia' }] }];
  if (m === 'GET' && p === '/me/outlook/supportedTimeZones') return [200, { value: [{ alias: 'Eastern Standard Time', displayName: '(UTC-05:00) Eastern Time (US & Canada)' }, { alias: 'E. South America Standard Time', displayName: '(UTC-03:00) Brasilia' }] }];
  if (m === 'GET' && p === '/me/calendar') return [200, { id: 'cal1', name: 'Calendário' }];
  if (m === 'GET' && (x = p.match(/^\/me\/calendars\/([^/]+)$/))) return [200, { id: x[1], name: x[1] === 'cal2' ? 'Pessoal' : 'Calendário' }];
  if (m === 'GET' && p === '/subscriptions') return [200, { value: subs.map(y => ({ ...y, clientState: null })) }];
  if (m === 'POST' && p === '/subscriptions') {
    if (String(b.notificationUrl).includes('expira-ruim')) return [400, { error: { code: 'InvalidRequest', message: 'Subscription expiration can only be 10070 minutes in the future. Expiration date time validation failed.' } }];
    if (String(b.notificationUrl).includes('fora-do-ar')) return [400, { error: { code: 'ValidationError', message: 'Subscription validation request failed. Notification endpoint must respond with 200 OK to validation request.' } }];
    const y = { ...b, id: 'sub-' + ++n, expirationDateTime: b.expirationDateTime.replace('Z', '.0000000Z') };
    subs.push(y);
    return [201, { ...y, clientState: null }];
  }
  if ((x = p.match(/^\/subscriptions\/([^/]+)$/))) {
    const y = subs.find(z => z.id === x![1]);
    if (!y || (m === 'PATCH' && ctl.patch404)) return [404, { error: { code: 'ResourceNotFound', message: 'subscription not found' } }];
    if (m === 'PATCH') { y.expirationDateTime = b.expirationDateTime; return [200, { ...y, clientState: null }]; }
    if (m === 'DELETE') { subs.splice(subs.indexOf(y), 1); return [204, null]; }
  }
  if (m === 'GET' && p === '/me/calendars') return [200, { value: [{ id: 'cal1', name: 'Calendário', canEdit: true, isDefaultCalendar: true, owner: { address: 'eu@minha.com' } }, { id: 'cal2', name: 'Pessoal', canEdit: true, isDefaultCalendar: false, owner: { address: 'eu@minha.com' } }] }];
  if (m === 'GET' && /^\/me(\/calendar|\/calendars\/[^/]+)?\/calendarView$/.test(p)) {
    if (throttle-- > 0) return [429, { error: { code: 'TooManyRequests', message: 'slow' } }];
    const a = toLoc(q.get('startDateTime')!), z = toLoc(q.get('endDateTime')!), cal = p.match(/calendars\/([^/]+)/)?.[1] ?? 'cal1';
    const v = evs.filter(e => !e.hidden && (e.cal ?? 'cal1') === cal && L(e.start.dateTime) < z && L(e.end.dateTime) > a).sort((x, y) => x.start.dateTime.localeCompare(y.start.dateTime));
    const top = +(q.get('$top') ?? 100), skip = +(q.get('$skip') ?? 0);
    const next = skip + top < v.length ? (() => { const u = new URL(url); u.searchParams.set('$skip', String(skip + top)); return u.toString(); })() : undefined;
    return [200, { value: v.slice(skip, skip + top).map(pub), ...(next ? { '@odata.nextLink': next } : {}) }];
  }
  if (m === 'POST' && p === '/me/calendar/getSchedule') {
    if (b.schedules.includes('falha@x.com')) return [400, { error: { code: 'ErrorInvalidRequest', message: 'not supported for consumer accounts' } }];
    const a = L(b.startTime.dateTime), z = L(b.endTime.dateTime);
    return [200, { value: b.schedules.map((s: string) => {
      if (s === 'x@externo.com') return { scheduleId: s, error: { message: 'no access' } };
      const items = s === 'eu@minha.com'
        ? evs.filter(e => !e.hidden && L(e.start.dateTime) < z && L(e.end.dateTime) > a).map(e => ({ status: e.showAs ?? 'busy', start: e.start, end: e.end }))
        : [{ status: 'busy', start: dt(dia(1), '13:00'), end: dt(dia(1), '14:00') }];
      return { scheduleId: s, scheduleItems: items };
    }) }];
  }
  if (m === 'POST' && (x = p.match(/^\/me(?:\/calendar|\/calendars\/([^/]+))\/events$/))) {
    const e = { ...b, id: nid(), cal: x[1] ?? 'cal1', type: b.recurrence ? 'seriesMaster' : 'singleInstance', isOrganizer: true, organizer: EU, responseStatus: { response: 'organizer' }, showAs: b.showAs ?? 'busy', attendees: (b.attendees ?? []).map((y: any) => ({ ...y, status: { response: 'none' } })), ...(b.isOnlineMeeting ? { onlineMeeting: { joinUrl: 'https://teams.microsoft.com/l/meetup-join/novo' } } : {}) };
    evs.push(e);
    return [201, pub(e)];
  }
  if ((x = p.match(/^\/me\/events\/([^/]+)\/(cancel|accept|decline|tentativelyAccept)$/)) && m === 'POST') {
    const e = find(x[1]);
    if (!e) return [404, { error: { code: 'ErrorItemNotFound', message: 'not found' } }];
    if (x[2] === 'cancel') { if (!e.isOrganizer) return [400, { error: { code: 'ErrorAccessDenied', message: 'only organizer' } }]; e.isCancelled = true; }
    else e.responseStatus = { response: { accept: 'accepted', decline: 'declined', tentativelyAccept: 'tentativelyAccepted' }[x[2]] };
    return [202, null];
  }
  if ((x = p.match(/^\/me\/events\/([^/]+)$/))) {
    const e = find(x[1]);
    if (!e) return [404, { error: { code: 'ErrorItemNotFound', message: 'not found' } }];
    if (m === 'GET') return [200, pub(e)];
    if (m === 'PATCH') { Object.assign(e, b); return [200, pub(e)]; }
    if (m === 'DELETE') { evs.splice(evs.indexOf(e), 1); return [204, null]; }
  }
  return [404, { error: { code: 'NoRoute', message: `${m} ${p}` } }];
}

const read = (r: IncomingMessage) => new Promise<string>(ok => { let s = ''; r.on('data', c => (s += c)); r.on('end', () => ok(s)); });
const send = (w: ServerResponse, s: number, b: any) => {
  if (b === null) { w.writeHead(s, { 'content-length': '0' }); return w.end(); }
  w.writeHead(s, { 'content-type': 'application/json', ...(s === 429 ? { 'retry-after': '0' } : {}) }); w.end(JSON.stringify(b));
};

export function start(): Promise<number> {
  const srv = createServer(async (r, w) => {
    const url = new URL(r.url!, 'http://h'), raw = await read(r);
    if (url.pathname.endsWith('/authorize')) { // usuário "escolhe a conta" e volta com código
      calls.push({ m: 'GET', p: url.pathname + url.search });
      const back = new URL(url.searchParams.get('redirect_uri')!);
      back.searchParams.set('code', 'MSCODE'); back.searchParams.set('state', url.searchParams.get('state')!);
      w.writeHead(302, { location: back.toString() }); return w.end();
    }
    if (url.pathname.endsWith('/devicecode')) return send(w, 200, { device_code: 'dc', user_code: 'ABC123', verification_uri: 'https://microsoft.com/devicelogin', interval: 0.02, expires_in: 60, message: 'abra e digite ABC123' });
    if (url.pathname.endsWith('/token')) {
      const f = new URLSearchParams(raw);
      if (f.get('grant_type')!.includes('device_code') && tokTries++ < 1) return send(w, 400, { error: 'authorization_pending' });
      if (f.get('grant_type') === 'authorization_code' && (f.get('code') !== 'MSCODE' || !f.get('code_verifier') || !f.get('redirect_uri'))) return send(w, 400, { error: 'invalid_grant' });
      return send(w, 200, { access_token: 'AT', refresh_token: 'RT', expires_in: 3600 });
    }
    if (r.headers.authorization !== 'Bearer AT') return send(w, 401, { error: { code: 'InvalidAuth', message: 'no' } });
    const b = raw ? JSON.parse(raw) : undefined;
    calls.push({ m: r.method!, p: url.pathname + url.search, b, prefer: String(r.headers.prefer ?? '') });
    const [s, body] = graph(r.method!, url, b);
    send(w, s, body);
  });
  return new Promise(ok => srv.listen(0, '127.0.0.1', () => ok((srv.address() as any).port)));
}
