// montagem do corpo do evento (criar e alterar compartilham as regras de data)
import { quando, dias, soma, min, rec, att, rcp } from '../lib.ts';

export function datas(a: { inicio?: string; fim?: string; duracao?: number; dia_inteiro?: boolean }, tz: string /* fuso como o Graph aceita */, cur?: { ini: string; fim: string; diaTodo: boolean }) {
  const i = a.inicio ? quando(a.inicio) : cur ? { l: cur.ini, diaTodo: cur.diaTodo } : null;
  if (!i) throw new Error('inicio obrigatório');
  const diaTodo = a.dia_inteiro ?? (a.inicio ? i.diaTodo : cur!.diaTodo);
  let ini: string, fim: string;
  if (diaTodo) {
    ini = i.l.slice(0, 10) + 'T00:00';
    const f = a.fim ? quando(a.fim, 'fim').l.slice(0, 10) : null;
    fim = f ? dias(f, 1) + 'T00:00' : cur && !a.inicio ? cur.fim : dias(ini, cur?.diaTodo ? Math.max(1, min(cur.ini, cur.fim) / 1440) : 1) + 'T00:00';
  } else {
    ini = i.l;
    fim = a.fim ? quando(a.fim, 'fim').l : soma(ini, a.duracao ?? (cur && !cur.diaTodo ? min(cur.ini, cur.fim) : 30));
  }
  if (fim <= ini) throw new Error('fim precisa ser depois do inicio');
  return { start: { dateTime: ini + ':00', timeZone: tz }, end: { dateTime: fim + ':00', timeZone: tz }, isAllDay: diaTodo, ini, fim };
}

const SHOW: Record<string, string> = { livre: 'free', provisorio: 'tentative', ocupado: 'busy', fora: 'oof', remoto: 'workingElsewhere' };
export const DISP = Object.keys(SHOW);

export function campos(a: any) {
  const o: Record<string, unknown> = {};
  if (a.titulo !== undefined) o.subject = a.titulo;
  if (a.local !== undefined) o.location = { displayName: a.local };
  if (a.descricao !== undefined) o.body = { contentType: 'Text', content: a.descricao };
  if (a.online !== undefined) o.isOnlineMeeting = !!a.online;
  if (a.lembrete !== undefined) Object.assign(o, a.lembrete < 0 ? { isReminderOn: false } : { isReminderOn: true, reminderMinutesBeforeStart: a.lembrete });
  if (a.disp !== undefined) o.showAs = SHOW[a.disp];
  if (a.privado !== undefined) o.sensitivity = a.privado ? 'private' : 'normal';
  return o;
}

export const pessoas = (a: any) => [...att(rcp(a.participantes)), ...att(rcp(a.opcionais), 'optional')];
export const recorre = (a: any, ini: string) => (a.rec ? { recurrence: rec(a.rec, ini, a.rec_ate, a.rec_vezes) } : {});
