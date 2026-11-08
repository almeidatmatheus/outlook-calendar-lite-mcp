import type { Tool } from '../mcp.ts';
import { ms, agora, dias, quando, view, table, norm, str, obj, RO } from '../lib.ts';

const t: Tool = {
  name: 'find',
  description: 'Busca eventos por texto no título, local, organizador ou participantes. Padrão: dos últimos 30 aos próximos 90 dias.',
  inputSchema: obj({
    termo: str('o que buscar (nome, email, palavra do título)'),
    de: str('data inicial AAAA-MM-DD'), ate: str('data final AAAA-MM-DD, inclusive'),
    qtd: { type: 'integer', minimum: 1, maximum: 50, default: 15 },
  }, ['termo']),
  annotations: RO,
  async run(a) {
    const { tz } = await ms(), hoje = agora(tz).slice(0, 10);
    const de = a.de ? quando(a.de, 'de').l.slice(0, 10) : dias(hoje, -30), ate = a.ate ? quando(a.ate, 'ate').l.slice(0, 10) : dias(hoje, 90);
    const { es } = await view('/me/calendar', de + 'T00:00', dias(ate, 1) + 'T00:00', tz, 500, 'id,subject,start,end,isAllDay,location,organizer,responseStatus,isOnlineMeeting,type,showAs,isCancelled,seriesMasterId,isOrganizer,attendees');
    const k = norm(a.termo);
    const hit = es.filter(e => norm([e.subject, e.location?.displayName, e.organizer?.emailAddress?.name, e.organizer?.emailAddress?.address, ...(e.attendees ?? []).flatMap((x: any) => [x.emailAddress?.name, x.emailAddress?.address])].filter(Boolean).join(' ')).includes(k));
    return hit.length ? table(hit.slice(0, a.qtd ?? 15)) : '0 resultados';
  },
};

export default t;
