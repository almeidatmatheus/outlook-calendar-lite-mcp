import type { Tool } from '../mcp.ts';
import { PER, janela, cal, view, table, agora, dia, dias, str, obj, RO } from '../lib.ts';

const t: Tool = {
  name: 'ls',
  description: 'Agenda: lista os eventos de um período, em ordem de horário, no fuso do usuário. Use para "o que tenho hoje", "minha semana", "reuniões de amanhã".',
  inputSchema: obj({
    periodo: str('janela de tempo (semana = de hoje até domingo)', { enum: PER, default: 'semana' }),
    de: str('data inicial AAAA-MM-DD (no lugar de periodo)'),
    ate: str('data final AAAA-MM-DD, inclusive'),
    calendario: str('nome do calendário (vazio = principal)'),
    qtd: { type: 'integer', minimum: 1, maximum: 100, default: 50, description: 'máx de eventos' },
  }),
  annotations: RO,
  async run(a) {
    const [de, ate, tz] = await janela(a);
    const { es, mais } = await view(await cal(a.calendario), de, ate, tz, a.qtd ?? 50);
    const cab = `#agenda ${dia(de)} → ${dia(dias(ate, -1))} (${tz}) · agora: ${dia(agora(tz))} ${agora(tz).slice(11)}${mais ? ' · +mais (use qtd ou de/ate)' : ''}`;
    return es.length ? `${cab}\n${table(es)}` : `${cab}\n0 eventos`;
  },
};

export default t;
