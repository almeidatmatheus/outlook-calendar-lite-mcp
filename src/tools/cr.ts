import type { Tool } from '../mcp.ts';
import { g } from '../g.ts';

import { cell } from '../fmt.ts';
import { ms, PH, cal, view, hora, REC, recTxt, rec, str, obj, LISTA_EMAILS } from '../lib.ts';
import { datas, campos, pessoas, recorre, DISP } from './ev.ts';

// eventos que já ocupam o horário (aviso, não bloqueia)
export async function conflitos(ini: string, fim: string, tz: string, ignorar = '') {
  const { es } = await view('/me/calendar', ini, fim, tz, 20);
  const c = es.filter(e => e.id !== ignorar && e.seriesMasterId !== ignorar && !e.isCancelled && e.showAs !== 'free');
  return c.length ? `\nconflito: ${c.map(e => `${e.id} ${cell(e.subject, 40)} (${hora(e)})`).join('; ')}` : '';
}

const t: Tool = {
  name: 'cr',
  description: 'Cria um evento. Sem participantes, vai só para a agenda do usuário. Com participantes, envia convite e pede confirmação. Avisa se o horário conflita com outro evento.',
  inputSchema: obj({
    titulo: str('título'),
    inicio: str('AAAA-MM-DDTHH:mm no fuso do usuário; só AAAA-MM-DD = dia inteiro'),
    fim: str('AAAA-MM-DDTHH:mm (ou use duracao)'),
    duracao: { type: 'integer', minimum: 5, maximum: 1440, default: 30, description: 'minutos, se não houver fim' },
    dia_inteiro: { type: 'boolean', description: 'evento de dia inteiro' },
    local: str('local'), descricao: str('descrição'),
    participantes: { ...LISTA_EMAILS, description: 'emails dos convidados obrigatórios' },
    opcionais: { ...LISTA_EMAILS, description: 'emails dos convidados opcionais' },
    online: { type: 'boolean', description: 'criar reunião online (Teams, em conta corporativa)' },
    lembrete: { type: 'integer', minimum: -1, maximum: 40320, description: 'minutos antes; -1 = sem lembrete' },
    disp: str('como o horário aparece para os outros', { enum: DISP }),
    privado: { type: 'boolean', description: 'esconder detalhes de quem vê sua agenda' },
    rec: str('repetição', { enum: REC }), rec_ate: str('repetir até AAAA-MM-DD'), rec_vezes: { type: 'integer', minimum: 1, maximum: 999, description: 'número de repetições' },
    calendario: str('nome do calendário (vazio = principal)'),
  }, ['titulo', 'inicio']),
  annotations: { destructiveHint: false, openWorldHint: true },
  async cf(a) {
    const ps = pessoas(a);
    if (!ps.length) return null; // só na agenda do usuário: ninguém é avisado
    const { tz, tzG } = await ms(), d = datas(a, tzG);
    return `CRIAR "${a.titulo}" ${hora(d)}${a.local ? ` em ${a.local}` : ''}${a.rec ? `\nrepete: ${recTxt(rec(a.rec, d.ini, a.rec_ate, a.rec_vezes))}` : ''}\nconvite para: ${ps.map(p => p.emailAddress.address + (p.type === 'optional' ? ' (opcional)' : '')).join(', ')}${await conflitos(d.ini, d.fim, tz)}`;
  },
  async run(a) {
    const { tz, tzG } = await ms(), d = datas(a, tzG);
    const r = await g('POST', `${await cal(a.calendario)}/events`, { ...campos(a), subject: a.titulo, start: d.start, end: d.end, isAllDay: d.isAllDay, attendees: pessoas(a), ...recorre(a, d.ini) }, await PH());
    return `criado:${r.id}|${hora(r)}${r.onlineMeeting?.joinUrl ? `\nonline:${r.onlineMeeting.joinUrl}` : ''}${await conflitos(d.ini, d.fim, tz, r.id)}`;
  },
};

export default t;
