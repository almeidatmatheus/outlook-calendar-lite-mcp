// ids curtos estáveis por usuário: '1','2',... <-> id Graph
import { cx } from './ctx.ts';

export const st = () => cx().ids;

export function put(gid: string): string {
  const I = st();
  let s = I.g2s.get(gid);
  if (!s) { s = String(++I.n); I.s2g.set(s, gid); I.g2s.set(gid, s); }
  return s;
}

// mensagem movida ganha id novo no Graph: mantém o id curto
export function remap(old: string, neu: string) {
  const I = st(), s = I.g2s.get(old);
  if (!s || !neu) return;
  I.g2s.delete(old); I.g2s.set(neu, s); I.s2g.set(s, neu);
  if (I.cur === old) I.cur = neu;
  I.last = I.last.map(x => (x === old ? neu : x));
}

export function res(a: string): string {
  const I = st(), k = String(a ?? '').trim().toLowerCase();
  if (k === 'atual') { if (!I.cur) throw new Error('nenhum email atual; use rd ou ls antes'); return I.cur; }
  if (k === 'ultimo') { if (!I.last[0]) throw new Error('nenhuma listagem; use ls antes'); return I.last[0]; }
  if (/^\d+$/.test(k)) { const g = I.s2g.get(k); if (!g) throw new Error(`id ${k} desconhecido`); return g; }
  if (k.length > 20) return String(a).trim();
  throw new Error(`alvo inválido: ${a}`);
}

export function resN(a: unknown): string[] {
  if (a === 'lista') { const l = st().last; if (!l.length) throw new Error('nenhuma listagem; use ls antes'); return [...l]; }
  const v = Array.isArray(a) ? a : [a];
  if (!v.length) throw new Error('alvos vazio');
  return v.map(x => res(String(x)));
}
