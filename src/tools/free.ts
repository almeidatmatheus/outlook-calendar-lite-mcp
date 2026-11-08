import type { Tool } from '../mcp.ts';
import { g } from '../g.ts';
import { tsv } from '../fmt.ts';
import { ms, me, PH, janela, view, rcp, agora, dias, dow, dia, min, soma, DOW_EN, str, obj, RO, LISTA_EMAILS } from '../lib.ts';

type I = [number, number];
const ms0 = (l: string) => Date.parse(l.slice(0, 16) + 'Z');
const hm = (t: number) => new Date(t).toISOString().slice(11, 16);

const t: Tool = {
  name: 'free',
  description: 'Horários livres para uma reunião: cruza a sua agenda (e a dos participantes, em conta corporativa) dentro do horário de trabalho.',
  inputSchema: obj({
    duracao: { type: 'integer', minimum: 5, maximum: 600, default: 30, description: 'minutos' },
    participantes: { ...LISTA_EMAILS, description: 'emails de quem precisa estar livre (além de você)' },
    periodo: str('janela de tempo', { enum: ['hoje', 'amanha', 'semana', 'proxima_semana', '7d'], default: '7d' }),
    de: str('data inicial AAAA-MM-DD'), ate: str('data final AAAA-MM-DD, inclusive'),
    das: str('início do dia HH:mm (padrão: horário de trabalho)'), ate_as: str('fim do dia HH:mm (padrão: horário de trabalho)'),
    todos_dias: { type: 'boolean', default: false, description: 'incluir dias fora do horário de trabalho (fim de semana)' },
    qtd: { type: 'integer', minimum: 1, maximum: 30, default: 10 },
  }),
  annotations: RO,
  async run(a) {
    const [de, ate, tz] = await janela(a, '7d'), W = await ms(), dur = a.duracao ?? 30, outros = rcp(a.participantes);
    const ini = a.das ?? W.ini, fim = a.ate_as ?? W.fim;
    if (!/^\d{2}:\d{2}$/.test(ini) || !/^\d{2}:\d{2}$/.test(fim)) throw new Error('das/ate_as no formato HH:mm');
    const busy: I[] = [], nota: string[] = [];
    const eu = (await me()).mail;
    try {
      const r = await g('POST', '/me/calendar/getSchedule', { schedules: [eu, ...outros], startTime: { dateTime: de + ':00', timeZone: W.tzG }, endTime: { dateTime: ate + ':00', timeZone: W.tzG }, availabilityViewInterval: 15 }, await PH());
      for (const s of r?.value ?? []) {
        if (s.error) { nota.push(`sem acesso à agenda de ${s.scheduleId}`); continue; }
        for (const x of s.scheduleItems ?? []) if (x.status !== 'free') busy.push([ms0(x.start.dateTime), ms0(x.end.dateTime)]);
      }
    } catch {
      if (outros.length) nota.push('a agenda dos participantes só pode ser consultada em conta corporativa; considerei só a sua');
      const { es } = await view('/me/calendar', de, ate, tz, 500);
      for (const e of es) if (e.showAs !== 'free' && !e.isCancelled) busy.push([ms0(e.start.dateTime), ms0(e.end.dateTime)]);
    }
    busy.sort((x, y) => x[0] - y[0]);
    const now = ms0(agora(tz)), out: unknown[][] = [];
    for (let d = de.slice(0, 10); d < ate.slice(0, 10) && out.length < (a.qtd ?? 10); d = dias(d, 1)) {
      if (!a.todos_dias && !W.dias.includes(DOW_EN[dow(d)])) continue;
      let c = Math.max(ms0(`${d}T${ini}`), Math.ceil(now / 9e5) * 9e5);
      const e = ms0(`${d}T${fim}`);
      for (const [x, y] of busy) {
        if (y <= c || x >= e) continue;
        if (x - c >= dur * 6e4) out.push([dia(d), hm(c), hm(x), (x - c) / 6e4]);
        c = Math.max(c, y);
      }
      if (e - c >= dur * 6e4) out.push([dia(d), hm(c), hm(e), (e - c) / 6e4]);
    }
    const cab = `#free ${dur}min · ${[eu, ...outros].join(', ')} · ${ini}-${fim} (${tz})${nota.length ? `\naviso: ${nota.join('; ')}` : ''}`;
    return out.length ? `${cab}\n${tsv(['dia', 'de', 'ate', 'min'], out.slice(0, a.qtd ?? 10))}` : `${cab}\nnenhum horário livre no período`;
  },
};

export default t;
