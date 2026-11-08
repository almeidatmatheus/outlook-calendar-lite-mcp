import type { Tool } from '../mcp.ts';
import { put, res, st } from '../ids.ts';
import { who, clean, cap } from '../fmt.ts';
import { ev, hora, loc, resp, disp, recTxt, str, ALVO, obj, RO } from '../lib.ts';
import { g } from '../g.ts';

const R: Record<string, string> = { accepted: 'ok', tentativelyAccepted: 'tlv', declined: 'nao', notResponded: '?', none: '?', organizer: 'org' };

const t: Tool = {
  name: 'rd',
  description: 'Lê um evento: horário, local, organizador, participantes com a resposta de cada um, link da reunião online, recorrência e descrição.',
  inputSchema: obj({ alvo: ALVO }, ['alvo']),
  annotations: RO,
  async run(a) {
    const id = res(a.alvo);
    const e = await ev(id, ',body,onlineMeeting,recurrence,categories,sensitivity,isReminderOn,reminderMinutesBeforeStart');
    st().cur = id;
    let r = e.recurrence;
    if (!r && e.seriesMasterId) r = (await g('GET', `/me/events/${e.seriesMasterId}?$select=recurrence`).catch(() => null))?.recurrence;
    const ps = (e.attendees ?? []).map((x: any) => `${who(x)} ${R[x.status?.response] ?? '?'}${x.type === 'optional' ? ' (opcional)' : ''}`);
    return cap([
      `id:${put(id)}`, `titulo:${e.isCancelled ? '[cancelado] ' : ''}${e.subject ?? ''}`, `quando:${hora(e)}`,
      loc(e) ? `local:${loc(e)}` : '', `org:${e.isOrganizer ? 'eu' : who(e.organizer)}`, `eu:${resp(e)}|disp:${disp(e)}${e.sensitivity === 'private' ? '|privado' : ''}${e.isReminderOn ? `|lembrete:${e.reminderMinutesBeforeStart}min` : ''}`,
      ps.length ? `participantes:${ps.join('; ')}` : '',
      e.onlineMeeting?.joinUrl ? `online:${e.onlineMeeting.joinUrl}` : e.isOnlineMeeting ? 'online:1' : '',
      r ? `rec:${recTxt(r)}` : '', e.categories?.length ? `cat:${e.categories.join(',')}` : '',
      '---', clean(e.body?.content ?? '', 3000) || '(sem descrição)',
    ].filter(Boolean).join('\n'));
  },
};

export default t;
