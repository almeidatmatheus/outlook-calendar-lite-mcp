import type { Tool } from '../mcp.ts';
import { g } from '../g.ts';
import { res } from '../ids.ts';
import { ms, PH, ev, hora, rcp, att, str, ALVO, obj, LISTA_EMAILS } from '../lib.ts';
import { datas, campos, DISP } from './ev.ts';
import { conflitos } from './cr.ts';

const MUDA_HORA = ['inicio', 'fim', 'duracao', 'dia_inteiro'];

async function plano(a: any) {
  const id0 = res(a.alvo), e = await ev(id0);
  if (!e.isOrganizer) throw new Error('só o organizador altera o evento; para responder ao convite use rsp');
  const id = a.serie && e.seriesMasterId ? e.seriesMasterId : id0;
  const { tz, tzG } = await ms();
  const p: Record<string, unknown> = campos(a);
  let d: ReturnType<typeof datas> | null = null;
  if (MUDA_HORA.some(k => a[k] !== undefined)) {
    d = datas(a, tzG, { ini: String(e.start.dateTime).slice(0, 16), fim: String(e.end.dateTime).slice(0, 16), diaTodo: !!e.isAllDay });
    Object.assign(p, { start: d.start, end: d.end, isAllDay: d.isAllDay });
  }
  const sai = new Set(rcp(a.rm_participantes).map(x => x.toLowerCase()));
  const novos = rcp(a.add_participantes);
  const ps = [...(e.attendees ?? []).filter((x: any) => !sai.has(String(x.emailAddress?.address).toLowerCase())), ...att(novos.filter(x => !(e.attendees ?? []).some((y: any) => String(y.emailAddress?.address).toLowerCase() === x.toLowerCase())))];
  if (sai.size || novos.length) p.attendees = ps;
  if (!Object.keys(p).length) throw new Error('nada para alterar');
  return { id, e, p, d, ps, tz, novos, sai };
}

const t: Tool = {
  name: 'ed',
  description: 'Altera um evento organizado pelo usuário: horário, título, local, descrição, participantes, online, lembrete. Se houver participantes, eles são avisados e a alteração pede confirmação.',
  inputSchema: obj({
    alvo: ALVO, serie: { type: 'boolean', default: false, description: 'alterar a série inteira (evento recorrente)' },
    titulo: str('novo título'), inicio: str('novo início AAAA-MM-DDTHH:mm (mantém a duração se não houver fim)'), fim: str('novo fim'),
    duracao: { type: 'integer', minimum: 5, maximum: 1440, description: 'nova duração em minutos' }, dia_inteiro: { type: 'boolean' },
    local: str('novo local'), descricao: str('nova descrição'), online: { type: 'boolean' },
    lembrete: { type: 'integer', minimum: -1, maximum: 40320, description: 'minutos antes; -1 = sem lembrete' },
    disp: str('como o horário aparece', { enum: DISP }), privado: { type: 'boolean' },
    add_participantes: { ...LISTA_EMAILS, description: 'convidar' }, rm_participantes: { ...LISTA_EMAILS, description: 'retirar' },
  }, ['alvo']),
  annotations: { destructiveHint: false, openWorldHint: true },
  async cf(a) {
    const { e, p, d, ps, tz, id, novos, sai } = await plano(a);
    if (!ps.length) return null; // evento só do usuário: ninguém é avisado
    const L = [`ALTERAR "${e.subject}" (${hora(e)})${a.serie && e.seriesMasterId ? ' — SÉRIE INTEIRA' : ''}`];
    if (p.subject) L.push(`título: ${p.subject}`);
    if (d) L.push(`novo horário: ${hora(d)}`);
    if (p.location) L.push(`local: ${a.local}`);
    if (novos.length) L.push(`convidar: ${novos.join(', ')}`);
    if (sai.size) L.push(`retirar: ${[...sai].join(', ')}`);
    L.push(`${ps.length} participante(s) serão avisados`);
    return L.join('\n') + (d ? await conflitos(d.ini, d.fim, tz, id) : '');
  },
  async run(a) {
    const { id, p, d, tz } = await plano(a);
    const r = await g('PATCH', `/me/events/${id}`, p, await PH());
    return `alterado|${hora(r)}${d ? await conflitos(d.ini, d.fim, tz, id) : ''}`;
  },
};

export default t;
