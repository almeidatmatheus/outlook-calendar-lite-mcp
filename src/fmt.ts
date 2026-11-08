export const cell = (s: unknown, n = 60) => {
  const v = String(s ?? '').replace(/[​-‏﻿]/g, '').replace(/[|\r\n\t]+/g, ' ').replace(/\s+/g, ' ').trim();
  return v.length > n ? v.slice(0, n - 1) + '…' : v;
};

export const tsv = (h: string[], rows: unknown[][]) => [h.join('|'), ...rows.map(r => r.map(x => (x === undefined || x === null || x === '' ? '-' : x)).join('|'))].join('\n');

const Y = new Date().getFullYear();
export function dt(iso?: string) {
  if (!iso) return '-';
  const d = new Date(iso), p = (x: number) => String(x).padStart(2, '0');
  const s = `${p(d.getMonth() + 1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}`;
  return d.getFullYear() === Y ? s : `${d.getFullYear()}-${s}`;
}

export const addr = (r: any) => r?.emailAddress?.address ?? '';
export const nm = (r: any) => r?.emailAddress?.name || addr(r).split('@')[0] || '';
export const dom = (r: any) => addr(r).split('@')[1]?.toLowerCase() ?? '';
export const who = (r: any) => { const a = addr(r), n = r?.emailAddress?.name; return n && n !== a ? `${n} <${a}>` : a; };

const QUOTE = /^(-{2,}\s*(original message|mensagem original)\s*-{2,}|_{10,}|(de|from):\s.+|(em|on)\s.{5,120}(escreveu|wrote):?)$/i;
const SIG = /^(--|—|-- )\s*$/;
const DISC = /(confidencial|confidential|este e-?mail|this e-?mail|esta mensagem|this message).{0,200}(destinat|intended|recipient|autorizad|authori[sz]ed)/is;

// corpo limpo: sem citações, assinatura, aviso legal, links longos
export function clean(s: string, max = 2000) {
  let t = String(s ?? '').replace(/\r/g, '').replace(/[​-‏﻿ ]/g, ' ');
  const L = t.split('\n'), o: string[] = [];
  for (const l of L) {
    const x = l.trim();
    if (QUOTE.test(x) && o.join('').length > 0) break;
    if (SIG.test(x)) break;
    if (x.startsWith('>')) continue;
    o.push(l.trimEnd());
  }
  t = o.join('\n');
  const ps = t.split(/\n{2,}/);
  while (ps.length > 1 && DISC.test(ps[ps.length - 1])) ps.pop();
  t = ps.join('\n\n')
    .replace(/<?https?:\/\/([^\s/?#>]+)[/?#][^\s>]{30,}>?/g, '<$1>')
    .replace(/[ \t]+/g, ' ')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
  return t.length > max ? t.slice(0, max) + `… [+${t.length - max}]` : t;
}

export const cap = (s: string, n = 30000) => (s.length > n ? s.slice(0, n) + `\n… [cortado ${s.length - n}]` : s);
