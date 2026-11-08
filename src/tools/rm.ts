import type { Tool } from '../mcp.ts';
import { g } from '../g.ts';
import { res } from '../ids.ts';
import { who } from '../fmt.ts';
import { ev, hora, str, ALVO, obj } from '../lib.ts';

async function alvo(a: any) {
  const id0 = res(a.alvo), e = await ev(id0);
  return { e, id: a.serie && e.seriesMasterId ? e.seriesMasterId : id0, cancela: !!e.isOrganizer && (e.attendees ?? []).length > 0 };
}

const t: Tool = {
  name: 'rm',
  description: 'Cancela ou exclui um evento. Se o usuário organiza e há participantes, eles recebem o cancelamento (com mensagem opcional). Pede confirmação.',
  inputSchema: obj({ alvo: ALVO, serie: { type: 'boolean', default: false, description: 'a série inteira (evento recorrente)' }, mensagem: str('mensagem do cancelamento para os participantes') }, ['alvo']),
  annotations: { destructiveHint: true, openWorldHint: true },
  async cf(a) {
    const { e, cancela } = await alvo(a);
    const s = a.serie && e.seriesMasterId ? ' — SÉRIE INTEIRA' : '';
    if (cancela) return `CANCELAR "${e.subject}" (${hora(e)})${s}\n${e.attendees.length} participante(s) recebem o cancelamento${a.mensagem ? `\n---\n${a.mensagem}` : ''}`;
    return `EXCLUIR da sua agenda "${e.subject}" (${hora(e)})${s}${e.isOrganizer ? '' : `\no organizador (${who(e.organizer)}) não é avisado; para recusar o convite use rsp`}`;
  },
  async run(a) {
    const { id, cancela } = await alvo(a);
    if (cancela) await g('POST', `/me/events/${id}/cancel`, { Comment: a.mensagem ?? '' });
    else await g('DELETE', `/me/events/${id}`);
    return cancela ? 'cancelado' : 'excluido';
  },
};

export default t;
