import type { Tool } from '../mcp.ts';
import { g } from '../g.ts';
import { res } from '../ids.ts';
import { who } from '../fmt.ts';
import { ev, hora, str, ALVO, obj } from '../lib.ts';

const ACAO: Record<string, string> = { aceitar: 'accept', recusar: 'decline', talvez: 'tentativelyAccept' };

async function alvo(a: any) {
  const id0 = res(a.alvo), e = await ev(id0);
  if (e.isOrganizer) throw new Error('você organiza este evento; não há convite para responder');
  return { e, id: a.serie && e.seriesMasterId ? e.seriesMasterId : id0 };
}

const t: Tool = {
  name: 'rsp',
  description: 'Responde a um convite: aceitar, recusar ou talvez, com comentário opcional. Avisar o organizador pede confirmação.',
  inputSchema: obj({
    alvo: ALVO, resposta: str('resposta', { enum: Object.keys(ACAO) }), comentario: str('mensagem para o organizador'),
    avisar: { type: 'boolean', default: true, description: 'enviar a resposta ao organizador' },
    serie: { type: 'boolean', default: false, description: 'responder para a série inteira' },
  }, ['alvo', 'resposta']),
  annotations: { destructiveHint: false, openWorldHint: true },
  async cf(a) {
    if (a.avisar === false) return null; // só muda a sua agenda
    const { e } = await alvo(a);
    return `${a.resposta.toUpperCase()} "${e.subject}" (${hora(e)})${a.serie && e.seriesMasterId ? ' — SÉRIE INTEIRA' : ''}\nresposta vai para: ${who(e.organizer)}${a.comentario ? `\n---\n${a.comentario}` : ''}`;
  },
  async run(a) {
    const { id } = await alvo(a);
    await g('POST', `/me/events/${id}/${ACAO[a.resposta]}`, { comment: a.comentario ?? '', sendResponse: a.avisar !== false });
    return 'respondido';
  },
};

export default t;
