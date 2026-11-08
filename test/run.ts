// sobe login + Graph de calendário simulados, roda o servidor MCP real via stdio e http e exercita todas as tools
import { spawn } from 'node:child_process';
import { mkdtempSync, readFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { createInterface } from 'node:readline';
import { start, calls, evs, dia } from './mock.ts';

const port = await start();
const home = mkdtempSync(join(tmpdir(), 'oc-'));
const ENV = { OUTLOOK_MCP_CLIENT_ID: 'cid', OUTLOOK_MCP_AUTH: `http://127.0.0.1:${port}`, OUTLOOK_MCP_GRAPH: `http://127.0.0.1:${port}/v1.0`, OUTLOOK_MCP_OPEN: 'test', OUTLOOK_MCP_TZ: '' };
let fails = 0;
const ok = (c: boolean, msg: string, extra = '') => { console.log(`${c ? 'ok  ' : 'FAIL'} ${msg}`); if (!c) { fails++; if (extra) console.log(extra); } };
const sleep = (ms: number) => new Promise(s => setTimeout(s, ms));
const lastCall = (f: (c: (typeof calls)[0]) => boolean) => [...calls].reverse().find(f);

function client(elicit?: (p: any) => any, env: Record<string, string> = {}) {
  const p = spawn(process.execPath, ['src/main.ts'], { env: { ...process.env, ...ENV, OUTLOOK_MCP_HOME: home, ...env }, stdio: ['pipe', 'pipe', 'inherit'] });
  const W = new Map<number, (r: any) => void>(); let id = 0;
  createInterface({ input: p.stdout! }).on('line', l => {
    const m = JSON.parse(l);
    if (m.method === 'elicitation/create') return p.stdin!.write(JSON.stringify({ jsonrpc: '2.0', id: m.id, result: elicit!(m.params) }) + '\n');
    W.get(m.id)?.(m);
  });
  const rq = (method: string, params: any = {}) => new Promise<any>(r => { const i = ++id; W.set(i, r); p.stdin!.write(JSON.stringify({ jsonrpc: '2.0', id: i, method, params }) + '\n'); });
  const call = async (name: string, args: any = {}) => { const r = await rq('tools/call', { name, arguments: args }); return { t: r.result?.content?.[0]?.text ?? JSON.stringify(r.error), e: !!r.result?.isError }; };
  return { p, rq, call };
}
const cfOf = (t: string) => t.match(/cf="(\w+)"/)?.[1];
const idOf = (t: string, titulo: string) => t.split('\n').find(l => l.includes(titulo))?.split('|')[0];

const A = client();
const init = await A.rq('initialize', { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 't', version: '0' } });
ok(init.result?.serverInfo?.name === 'outlook-calendar-lite-mcp' && !!init.result?.instructions, 'initialize');
const tl = (await A.rq('tools/list')).result.tools;
ok(tl.map((t: any) => t.name).join(',') === 'ls,rd,find,free,cr,ed,rm,rsp,cx', `tools/list: ${tl.map((t: any) => t.name).join(',')}`);
ok(['cr', 'ed', 'rm', 'rsp'].every(n => !!tl.find((t: any) => t.name === n).inputSchema.properties.cf), 'cr/ed/rm/rsp aceitam cf');

await sleep(400);
const az = calls.find(c => c.p.includes('/authorize'))?.p ?? '';
ok(az.includes('Calendars.ReadWrite') && az.includes('prompt=select_account') && az.includes('code_challenge_method=S256'), 'initialize abre login Microsoft com permissão de calendário', az);
ok(JSON.parse(readFileSync(join(home, 'default.json'), 'utf8')).rt === 'RT', 'conta conectada: token salvo');
let r: { t: string; e: boolean };

// ---- leitura
r = await A.call('ls', { periodo: 'hoje' });
const cv = lastCall(c => c.p.includes('calendarView'));
ok(!r.e && r.t.startsWith('#agenda') && r.t.includes('America/Sao_Paulo') && r.t.includes('agora:') && r.t.includes('Daily do time') && r.t.includes('Reunião com cliente ACME'), 'ls hoje: fuso do Windows convertido para America/Sao_Paulo; cabeçalho com agora + retry 429', r.t);
ok(!!cv?.p.includes(`startDateTime=${dia(0)}T03:00:00Z`), 'ls: meia-noite de São Paulo vira 03:00 UTC na busca', cv?.p);
ok(!!cv?.prefer?.includes('outlook.timezone="E. South America Standard Time"'), 'ls: pede horários no fuso da caixa, no formato que o Graph aceita (Prefer)', cv?.prefer);
const linhas = r.t.split('\n');
const daily = linhas.find(l => l.includes('Daily'))!.split('|'), acme = linhas.find(l => l.includes('ACME'))!.split('|');
ok(daily[2] === '10:00' && daily[3] === '10:15' && daily[6] === 'eu' && daily[7] === 'org' && daily[8] === '1' && daily[9] === '1', 'ls: daily recorrente, online, organizada por mim', daily.join('|'));
ok(acme[5] === 'Sala 3' && acme[6] === 'Ana Lima' && acme[7] === '?' && acme[10] === 'prov', 'ls: convite da Ana sem resposta', acme.join('|'));
ok(daily[0].startsWith('AAMk') && acme[0].startsWith('AAMk'), 'ls: coluna id traz o id real do Graph');
ok(calls.filter(c => c.p.startsWith('/v1.0/me')).every(c => c.prefer?.includes('IdType="ImmutableId"')), 'toda chamada pede id imutável ao Graph');
console.log(r.t);
const ID: Record<string, string> = { daily: daily[0], acme: acme[0] };

r = await A.call('ls', { periodo: 'amanha' });
ok(r.t.includes('|dia todo|') && r.t.includes('Feriado') && r.t.includes('Dentista'), 'ls amanha: dia inteiro', r.t);
ID.dent = idOf(r.t, 'Dentista')!; ID.fer = idOf(r.t, 'Feriado')!;

r = await A.call('rd', { alvo: ID.acme });
ok(!r.e && r.t.includes('local:Sala 3') && r.t.includes('Bob Silva <bob@acme.com> ok (opcional)') && r.t.includes('Apresentar a proposta') && !r.t.includes('confidencial') && r.t.includes('quando:') && r.t.includes('14:00–15:00 (60 min)'), 'rd: participantes, resposta de cada um, descrição limpa', r.t);
r = await A.call('rd', { alvo: ID.daily });
ok(r.t.includes('rec:semanal (seg,qua)') && r.t.includes('online:https://teams.microsoft.com/') && r.t.includes('org:eu'), 'rd: recorrência da série e link do Teams', r.t);
console.log(r.t);

r = await A.call('find', { termo: 'cliente' });
ok(!r.e && r.t.split('\n').length === 2 && r.t.includes('ACME'), 'find por título', r.t);
r = await A.call('find', { termo: 'bob' });
ok(r.t.includes('ACME'), 'find por participante', r.t);
r = await A.call('find', { termo: 'nada disso' });
ok(r.t === '0 resultados', 'find sem resultado');

r = await A.call('free', { periodo: 'amanha', duracao: 60, participantes: ['ana@acme.com'], todos_dias: true });
ok(!r.e && r.t.includes('|09:30|13:00|210') && r.t.includes('|14:00|18:00|240') && !r.t.includes('aviso'), 'free: cruza minha agenda e a da Ana no horário de trabalho (feriado marcado como livre não bloqueia)', r.t);
console.log(r.t);
r = await A.call('free', { periodo: 'amanha', participantes: ['x@externo.com'], todos_dias: true });
ok(r.t.includes('sem acesso à agenda de x@externo.com'), 'free: avisa quando não vê a agenda de alguém', r.t);
r = await A.call('free', { periodo: 'amanha', participantes: ['falha@x.com'], todos_dias: true, das: '08:00', ate_as: '10:00' });
ok(r.t.includes('conta corporativa') && r.t.includes('|08:00|09:00|60') && r.t.includes('|09:30|10:00|30'), 'free: conta pessoal usa só a sua agenda; das/ate_as', r.t);

// ---- criar
r = await A.call('cr', { titulo: 'Foco', inicio: `${dia(1)}T09:15`, duracao: 60 });
const cr1 = lastCall(c => c.m === 'POST' && c.p.endsWith('/events'));
ok(/^criado:AAMk\S+\|/.test(r.t) && r.t.includes('09:15–10:15') && r.t.includes('conflito:') && r.t.includes('Dentista') && !r.t.includes('CONFIRMAÇÃO'), 'cr sem participantes: sem confirmação, avisa conflito com o dentista', r.t);
ok(!!cr1 && cr1.b.start.dateTime === `${dia(1)}T09:15:00` && cr1.b.start.timeZone === 'E. South America Standard Time' && cr1.b.end.dateTime === `${dia(1)}T10:15:00` && cr1.b.attendees.length === 0, 'cr: horário no fuso do usuário', JSON.stringify(cr1?.b));
const nPost = calls.filter(c => c.m === 'POST' && c.p.endsWith('/events')).length;
r = await A.call('cr', { titulo: 'Kickoff', inicio: `${dia(2)} 15:00`, fim: `${dia(2)} 16:30`, participantes: ['ana@acme.com'], opcionais: ['bob@acme.com'], online: true, local: 'Sala 1' });
const cf1 = cfOf(r.t);
ok(!!cf1 && r.t.includes('CRIAR "Kickoff"') && r.t.includes('convite para: ana@acme.com, bob@acme.com (opcional)') && calls.filter(c => c.m === 'POST' && c.p.endsWith('/events')).length === nPost, 'cr com participantes: preview, nada criado', r.t);
r = await A.call('cr', { titulo: 'Kickoff', inicio: `${dia(2)} 15:00`, fim: `${dia(2)} 17:00`, participantes: ['ana@acme.com'], opcionais: ['bob@acme.com'], online: true, local: 'Sala 1', cf: cf1 });
ok(r.t.includes('CONFIRMAÇÃO'), 'cr: cf não vale se args mudam');
r = await A.call('cr', { titulo: 'Kickoff', inicio: `${dia(2)} 15:00`, fim: `${dia(2)} 16:30`, participantes: ['ana@acme.com'], opcionais: ['bob@acme.com'], online: true, local: 'Sala 1', cf: cf1 });
const cr2 = lastCall(c => c.m === 'POST' && c.p.endsWith('/events'));
ok(r.t.includes('criado:') && r.t.includes('online:https://teams') && cr2?.b.attendees[1].type === 'optional' && cr2.b.isOnlineMeeting === true && cr2.b.location.displayName === 'Sala 1', 'cr confirmado: convite com Teams', r.t);
r = await A.call('cr', { titulo: 'Viagem', inicio: dia(3), fim: dia(5) });
const cr3 = lastCall(c => c.m === 'POST' && c.p.endsWith('/events'));
ok(cr3?.b.isAllDay === true && cr3.b.end.dateTime === `${dia(6)}T00:00:00` && r.t.includes('dia todo (3 dias'), 'cr dia inteiro de vários dias', r.t);
r = await A.call('cr', { titulo: 'Academia', inicio: `${dia(1)}T07:00`, duracao: 60, rec: 'semanal', rec_vezes: 4, calendario: 'pessoal' });
const cr4 = lastCall(c => c.m === 'POST' && c.p.endsWith('/events'));
ok(!!cr4 && cr4.p.includes('/calendars/cal2/events') && cr4.b.recurrence.pattern.type === 'weekly' && cr4.b.recurrence.range.numberOfOccurrences === 4, 'cr recorrente em outro calendário', JSON.stringify(cr4?.b.recurrence));
r = await A.call('cr', { titulo: 'x', inicio: 'amanhã às 3' });
ok(r.e && r.t.includes('AAAA-MM-DDTHH:mm'), 'cr: data inválida explica o formato', r.t);
r = await A.call('cr', { titulo: 'x', inicio: `${dia(1)}T10:00`, fim: `${dia(1)}T09:00` });
ok(r.e && r.t.includes('fim precisa ser depois'), 'cr: fim antes do início', r.t);
r = await A.call('cr', { titulo: 'x', inicio: `${dia(1)}T10:00`, participantes: ['nao-email'] });
ok(r.e && r.t.includes('email inválido'), 'cr: valida email', r.t);

// ---- alterar
r = await A.call('ed', { alvo: ID.dent, inicio: `${dia(1)}T11:00` });
const ed1 = lastCall(c => c.m === 'PATCH');
ok(r.t.startsWith('alterado') && !!ed1 && ed1.b.start.dateTime === `${dia(1)}T11:00:00` && ed1.b.end.dateTime === `${dia(1)}T11:30:00`, 'ed sem participantes: sem confirmação, mantém a duração', r.t);
r = await A.call('ed', { alvo: ID.daily, titulo: 'Daily (novo horário)', inicio: `${dia(0)}T10:30` });
const cf2 = cfOf(r.t);
ok(!!cf2 && r.t.includes('ALTERAR "Daily do time"') && r.t.includes('novo horário:') && r.t.includes('1 participante(s) serão avisados'), 'ed com participantes: preview', r.t);
r = await A.call('ed', { alvo: ID.daily, titulo: 'Daily (novo horário)', inicio: `${dia(0)}T10:30`, cf: cf2 });
ok(r.t.startsWith('alterado') && lastCall(c => c.m === 'PATCH')?.b.subject === 'Daily (novo horário)', 'ed confirmado', r.t);
r = await A.call('ed', { alvo: ID.daily, serie: true, add_participantes: ['carla@acme.com'] });
r = await A.call('ed', { alvo: ID.daily, serie: true, add_participantes: ['carla@acme.com'], cf: cfOf(r.t) });
const ed3 = lastCall(c => c.m === 'PATCH');
ok(!!ed3 && ed3.p.endsWith('/me/events/M1') && ed3.b.attendees.length === 2, 'ed série inteira: convida mais alguém na série', JSON.stringify(ed3));
r = await A.call('ed', { alvo: ID.acme, titulo: 'x' });
ok(r.e && r.t.includes('use rsp'), 'ed: convite de outra pessoa não é alterável', r.t);
r = await A.call('ed', { alvo: ID.dent });
ok(r.e && r.t.includes('nada para alterar'), 'ed sem mudanças', r.t);

// ---- responder
r = await A.call('rsp', { alvo: ID.acme, resposta: 'aceitar', comentario: 'Combinado!' });
const cf3 = cfOf(r.t);
ok(!!cf3 && r.t.includes('ACEITAR "Reunião com cliente ACME"') && r.t.includes('Ana Lima <ana@acme.com>'), 'rsp: preview da resposta ao organizador', r.t);
r = await A.call('rsp', { alvo: ID.acme, resposta: 'aceitar', comentario: 'Combinado!', cf: cf3 });
const rs1 = lastCall(c => c.p.endsWith('/accept'));
ok(r.t === 'respondido' && rs1?.b.sendResponse === true && rs1.b.comment === 'Combinado!', 'rsp aceitar', r.t);
r = await A.call('rsp', { alvo: ID.acme, resposta: 'talvez', avisar: false });
ok(r.t === 'respondido' && lastCall(c => c.p.endsWith('/tentativelyAccept'))?.b.sendResponse === false, 'rsp sem avisar: sem confirmação', r.t);
r = await A.call('rsp', { alvo: ID.dent, resposta: 'aceitar' });
ok(r.e && r.t.includes('você organiza'), 'rsp: evento próprio', r.t);

// ---- cancelar / excluir
r = await A.call('rm', { alvo: ID.daily, mensagem: 'Hoje não teremos daily.' });
const cf4 = cfOf(r.t);
ok(r.t.includes('CANCELAR') && r.t.includes('recebem o cancelamento') && !calls.some(c => c.p.endsWith('/cancel')), 'rm de reunião minha com participantes: preview de cancelamento', r.t);
r = await A.call('rm', { alvo: ID.daily, mensagem: 'Hoje não teremos daily.', cf: cf4 });
ok(r.t === 'cancelado' && lastCall(c => c.p.endsWith('/cancel'))?.b.Comment === 'Hoje não teremos daily.', 'rm confirmado: cancelamento com mensagem', r.t);
r = await A.call('rm', { alvo: ID.acme });
ok(r.t.includes('EXCLUIR da sua agenda') && r.t.includes('não é avisado') && r.t.includes('use rsp'), 'rm de convite de terceiro: explica que o organizador não é avisado', r.t);
r = await A.call('rm', { alvo: ID.acme, cf: cfOf(r.t) });
ok(r.t === 'excluido' && calls.some(c => c.m === 'DELETE'), 'rm confirmado: excluído', r.t);

// ---- conta e erros
r = await A.call('cx');
ok(!r.e && r.t.includes('conta:eu@minha.com') && r.t.includes('fuso:America/Sao_Paulo') && r.t.includes('trabalho:mon,tue,wed,thu,fri 09:00-18:00') && r.t.includes('Pessoal'), 'cx: conta, fuso, horário de trabalho, calendários', r.t);
r = await A.call('ls', { periodo: 'xyz' });
ok(r.e && r.t.includes('hoje|amanha'), 'enum inválido rejeitado', r.t);
r = await A.call('rd', {});
ok(r.e && r.t.includes('alvo: obrigatório'), 'required', r.t);
r = await A.call('rd', { alvo: 'naoexiste' });
ok(r.e && r.t.includes('404'), 'id inexistente: erro do Graph', r.t);
r = await A.call('ls', { calendario: 'inexistente' });
ok(r.e && r.t.includes('calendário desconhecido'), 'calendário desconhecido', r.t);
const nAuth = calls.filter(c => c.p.includes('/authorize')).length;
r = await A.call('cx', { acao: 'trocar' });
ok(r.t.includes('abri o navegador'), 'cx trocar: abre escolha de conta de novo', r.t);
await sleep(400);
ok(calls.filter(c => c.p.includes('/authorize')).length === nAuth + 1, 'cx trocar: novo login feito');
r = await A.call('rd', { alvo: ID.dent });
ok(!r.e && r.t.includes(`id:${ID.dent}`), 'o id continua valendo depois de trocar de sessão', r.t);
A.p.kill();

// login por código de dispositivo (servidor sem navegador)
const C = client(undefined, { OUTLOOK_MCP_LOGIN: 'device', OUTLOOK_MCP_HOME: mkdtempSync(join(tmpdir(), 'oc-')) });
await C.rq('initialize', { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 't', version: '0' } });
r = await C.call('ls');
ok(r.e && r.t.includes('ABC123'), 'device code: devolve link + código', r.t);
await sleep(300);
r = await C.call('ls', { periodo: 'hoje' });
ok(!r.e && r.t.startsWith('#agenda'), 'device code: funciona após login', r.t);
C.p.kill();

// cliente com elicitation: confirmação vai direto ao usuário
let asked = 0;
const B = client(() => (++asked === 1 ? { action: 'accept', content: { confirmar: true } } : { action: 'decline' }));
await B.rq('initialize', { protocolVersion: '2025-06-18', capabilities: { elicitation: {} }, clientInfo: { name: 't', version: '0' } });
await sleep(400);
r = await B.call('rd', { alvo: ID.dent });
ok(!r.e && r.t.includes('Dentista'), 'id salvo funciona em outro processo, sem listar antes', r.t);
const n0 = evs.length;
r = await B.call('cr', { titulo: 'Com Ana', inicio: `${dia(4)}T10:00`, participantes: ['ana@acme.com'] });
ok(r.t.includes('criado:') && asked === 1 && evs.length === n0 + 1, 'elicitation aceita -> cria e convida', r.t);
r = await B.call('cr', { titulo: 'Com Bob', inicio: `${dia(4)}T11:00`, participantes: ['bob@acme.com'] });
ok(r.t.includes('cancelado') && evs.length === n0 + 1, 'elicitation recusada -> nada criado', r.t);
r = await B.call('cr', { titulo: 'Só meu', inicio: `${dia(4)}T12:00` });
ok(r.t.includes('criado:') && asked === 2, 'sem participantes: nem pergunta', r.t);
B.p.kill();

// ---- modo remoto: simula o claude.ai conectando o conector
const HP = 18789, BASE = `http://127.0.0.1:${HP}`;
const hs = spawn(process.execPath, ['src/main.ts', 'http'], {
  env: { ...process.env, ...ENV, OUTLOOK_MCP_HOME: mkdtempSync(join(tmpdir(), 'oc-')), PORT: String(HP), OUTLOOK_MCP_URL: BASE, OUTLOOK_MCP_ALLOW: 'minha.com' },
  stdio: ['ignore', 'ignore', 'pipe'],
});
await new Promise<void>(ok => hs.stderr!.on('data', d => String(d).includes('http em') && ok()));
const F = (p: string, o: RequestInit = {}) => fetch(p.startsWith('http') ? p : BASE + p, { redirect: 'manual', ...o });
let x = await F('/mcp', { method: 'POST', body: '{}' });
ok(x.status === 401 && (x.headers.get('www-authenticate') ?? '').includes(`resource_metadata="${BASE}/.well-known/oauth-protected-resource"`), 'http: 401 com resource_metadata');
const RU = 'https://claude.ai/api/mcp/auth_callback';
const reg = await (await F('/register', { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ redirect_uris: [RU], client_name: 'Claude' }) })).json() as any;
const { createHash } = await import('node:crypto');
const ver = 'v'.repeat(50), chal = createHash('sha256').update(ver).digest('base64url');
x = await F(`/authorize?${new URLSearchParams({ response_type: 'code', client_id: reg.client_id, redirect_uri: RU, code_challenge: chal, code_challenge_method: 'S256', state: 'st1' })}`);
const toMs = x.headers.get('location') ?? '';
ok(x.status === 302 && toMs.includes('/oauth2/v2.0/authorize') && toMs.includes('Calendars.ReadWrite'), 'http: authorize leva ao login Microsoft com permissão de calendário', toMs);
x = await F(toMs); x = await F(x.headers.get('location')!);
const back = new URL(x.headers.get('location') ?? 'http://x');
ok(back.origin + back.pathname === RU && back.searchParams.get('state') === 'st1', 'http: volta ao claude.ai com code + state', back.toString());
const tokf = (b: Record<string, string>) => F('/token', { method: 'POST', headers: { 'content-type': 'application/x-www-form-urlencoded' }, body: new URLSearchParams(b) });
const tr = await (await tokf({ grant_type: 'authorization_code', code: back.searchParams.get('code')!, redirect_uri: RU, client_id: reg.client_id, code_verifier: ver })).json() as any;
ok(!!tr.access_token, 'http: token emitido');
const mcp = async (body: any) => (await F('/mcp', { method: 'POST', headers: { authorization: `Bearer ${tr.access_token}`, 'content-type': 'application/json', accept: 'application/json, text/event-stream' }, body: JSON.stringify(body) })).json() as any;
ok((await mcp({ jsonrpc: '2.0', id: 1, method: 'initialize', params: { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 'claude-ai', version: '0' } } })).result.serverInfo.name === 'outlook-calendar-lite-mcp', 'http: initialize');
const lsr = await mcp({ jsonrpc: '2.0', id: 3, method: 'tools/call', params: { name: 'ls', arguments: { periodo: 'amanha' } } });
ok(lsr.result.content[0].text.includes('#agenda'), 'http: ls com a conta conectada', JSON.stringify(lsr));
const crr = await mcp({ jsonrpc: '2.0', id: 4, method: 'tools/call', params: { name: 'cr', arguments: { titulo: 'x', inicio: `${dia(5)}T10:00`, participantes: ['ana@acme.com'] } } });
ok(crr.result.content[0].text.includes('CONFIRMAÇÃO'), 'http: convite pede confirmação');
hs.kill();

console.log(fails ? `\n${fails} falha(s)` : '\ntudo ok');
process.exit(fails ? 1 : 0);
