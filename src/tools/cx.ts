import type { Tool } from '../mcp.ts';
import { g } from '../g.ts';
import { cx } from '../ctx.ts';
import { cell, tsv } from '../fmt.ts';
import { me, ms, agora, dia, str, obj } from '../lib.ts';

const t: Tool = {
  name: 'cx',
  description: 'Conta conectada: email, fuso horário, horário de trabalho e calendários. acao trocar desconecta e pede login de outra conta.',
  inputSchema: obj({ acao: str('info ou trocar', { enum: ['info', 'trocar'], default: 'info' }) }),
  annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
  async run(a) {
    if (a?.acao === 'trocar') {
      const C = cx();
      if (!C.reset) return 'neste modo, troque a conta desconectando e reconectando o conector no app';
      C.c = {}; return C.reset();
    }
    const [m, w, c] = await Promise.all([me(), ms(), g('GET', '/me/calendars?$select=name,canEdit,isDefaultCalendar,owner&$top=100')]);
    return `conta:${m.mail}\nfuso:${w.tz}|agora:${dia(agora(w.tz))} ${agora(w.tz).slice(11)}\ntrabalho:${w.dias.map(x => x.slice(0, 3)).join(',')} ${w.ini}-${w.fim}\n`
      + tsv(['calendario', 'dono', 'edita', 'padrao'], (c?.value ?? []).map((x: any) => [cell(x.name, 40), cell(x.owner?.address ?? '', 40), x.canEdit ? 1 : 0, x.isDefaultCalendar ? 1 : 0]));
  },
};

export default t;
