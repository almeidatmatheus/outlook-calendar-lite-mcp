import type { Tool } from '../mcp.ts';
import { g } from '../g.ts';
import { cal, str, obj } from '../lib.ts';

// URL e segredo só do ambiente: o modelo não escolhe para onde as notificações vão
const env = () => ({ url: process.env.OUTLOOK_MCP_NOTIFY_URL ?? '', seg: process.env.OUTLOOK_MCP_NOTIFY_SECRET ?? '' });
const iso = (t: number) => new Date(t).toISOString().replace(/\.\d{3}Z$/, 'Z');
// o Graph pode devolver o resource com outra forma (barra inicial, users/{id}, maiúsculas)
const rk = (s: string) => String(s ?? '').replace(/^\//, '').toLowerCase().replace(/^users\/[^/]+\//, 'me/');

async function calendario(nome?: string) {
  const c = await g('GET', `${await cal(nome)}?$select=id,name`);
  return { nome: c.name as string, resource: `me/calendars/${c.id}/events` };
}

// assinatura deste app (mesma URL) para este calendário
async function achar(url: string, resource: string) {
  let u: string | undefined = '/subscriptions';
  while (u) {
    const r = await g('GET', u);
    const s = (r?.value ?? []).find((x: any) => x.notificationUrl === url && rk(x.resource) === rk(resource));
    if (s) return s;
    u = r?.['@odata.nextLink'];
  }
  return null;
}

const t: Tool = {
  name: 'ntf',
  description: 'Notificações de mudança do calendário para o app configurado (Microsoft Graph): info mostra a assinatura, ligar cria ou renova, desligar remove. A URL que recebe vem só da configuração do servidor.',
  inputSchema: obj({
    acao: str('info, ligar ou desligar', { enum: ['info', 'ligar', 'desligar'], default: 'info' }),
    calendario: str('nome do calendário (vazio = principal)'),
    dias: { type: 'integer', minimum: 1, maximum: 3, default: 2, description: 'validade da assinatura, em dias' },
  }),
  annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: true, openWorldHint: true },
  async run(a) {
    const { url, seg } = env();
    if (!url) return 'notificações não configuradas (defina OUTLOOK_MCP_NOTIFY_URL e OUTLOOK_MCP_NOTIFY_SECRET)';
    if (!seg) return 'notificações não configuradas (defina OUTLOOK_MCP_NOTIFY_SECRET)';
    if (seg.length > 128) throw new Error('OUTLOOK_MCP_NOTIFY_SECRET passa de 128 caracteres');
    if (!/^https:\/\//i.test(url)) throw new Error('OUTLOOK_MCP_NOTIFY_URL precisa ser HTTPS');
    const dias = a.dias ?? 2;
    if (dias < 1 || dias > 3) throw new Error('dias: use de 1 a 3'); // o validador do MCP não confere mínimo/máximo
    try {
      const c = await calendario(a.calendario), s = await achar(url, c.resource);
      if (a.acao === 'desligar') {
        if (s) await g('DELETE', `/subscriptions/${s.id}`).catch(e => { if (!/^graph 404/.test(e.message)) throw e; });
        return 'desligado';
      }
      if (a.acao !== 'ligar') return s ? `assinatura:${s.id}|expira:${iso(Date.parse(s.expirationDateTime))}|calendario:${c.nome}` : 'sem assinatura';
      const exp = iso(Date.now() + dias * 864e5);
      let r = s ? await g('PATCH', `/subscriptions/${s.id}`, { expirationDateTime: exp }).catch(e => { if (/^graph 404/.test(e.message)) return null; throw e; }) : null;
      if (!r) {
        r = await g('POST', '/subscriptions', { changeType: 'updated,deleted', resource: c.resource, notificationUrl: url, lifecycleNotificationUrl: url, clientState: seg, expirationDateTime: exp })
          // só a falha da chamada de validação à URL; outros 400 (ex.: expiração inválida) seguem o erro do g()
          .catch(e => { throw /^graph 400/.test(e.message) && /validation request|notification endpoint/i.test(e.message) ? new Error('a URL de notificação não respondeu à validação do Microsoft Graph em 10 s (o app está no ar e devolve o validationToken?)') : e; });
      }
      return `ligado:${r.id}|expira:${iso(Date.parse(r.expirationDateTime ?? exp))}`;
    } catch (e: any) {
      throw new Error(String(e?.message ?? e).split(seg).join('***')); // o segredo nunca aparece na saída
    }
  },
};

export default t;
