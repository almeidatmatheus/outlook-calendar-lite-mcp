# outlook-calendar-lite-mcp

Servidor [MCP](https://modelcontextprotocol.io) leve para o Calendário do Microsoft Outlook (Microsoft Graph).
Funciona com conta corporativa (Microsoft 365) e pessoal (Outlook.com, Hotmail).

- **Zero dependências.** Node 22.18+ roda o TypeScript direto: sem `npm install`, sem build.
- **Roda com você.** Na sua máquina, com o seu app do Azure. Nenhum servidor de terceiros vê sua agenda.
- **Econômico em tokens.** Saída tabular.
- **IDs permanentes.** Cada evento vem com o id real do Microsoft Graph, pedido como *id imutável*, então pode ser guardado num banco e usado depois.
- **Sempre no seu fuso.** Os horários são lidos e gravados no fuso da sua caixa; toda listagem mostra a data e a hora de agora, para a IA não chutar o que é "amanhã".
- **Seguro por padrão.** Tudo que avisa outras pessoas (convites, alterações, cancelamentos, respostas) mostra um preview e espera a aprovação do usuário. Evento só na sua agenda não pergunta.
- **Login ao conectar.** Na primeira conexão, o navegador abre na escolha de conta Microsoft.

> English: [README.md](README.md).

## Pré-requisitos

- **Node.js 22.18 ou mais novo** (`node --version`). Nada mais para instalar: sem `npm install`, sem build.
- **git**, para clonar o repositório (ou baixe o ZIP pelo GitHub).
- **Uma conta Microsoft** com calendário: corporativa (Microsoft 365) ou pessoal (Outlook.com, Hotmail).
- **Permissão para registrar um app no Microsoft Entra** (gratuito):
  - conta corporativa: liberado por padrão na maioria das empresas; se estiver bloqueado, peça ao TI para registrar ou aprovar o app;
  - conta pessoal: entre uma vez em <https://portal.azure.com> com uma conta gratuita do Azure, que cria o seu diretório.
- **Um cliente MCP**: Claude Desktop, Claude Code, Cursor, VS Code ou qualquer cliente com suporte a servidores stdio.

## Começo rápido

### 1. Baixe o código

```bash
git clone https://github.com/almeidatmatheus/outlook-calendar-lite-mcp.git
cd outlook-calendar-lite-mcp
node --version   # precisa ser 22.18 ou mais novo
```

### 2. Registre o seu app no Microsoft Entra (uma vez, ~5 min)

> Já usa o [outlook-lite-mcp](https://github.com/almeidatmatheus/outlook-lite-mcp)? Reaproveite o mesmo app: faça só o passo 4 (acrescentar `Calendars.ReadWrite`) e use o mesmo client ID.

1. Em <https://entra.microsoft.com> → **Applications** → **App registrations** → **New registration**.
   - **Name:** `outlook-calendar-lite-mcp`
   - **Supported account types:**
     - pessoal + qualquer empresa: **Accounts in any organizational directory and personal Microsoft accounts**
     - só a sua empresa: **Accounts in this organizational directory only**
   - **Register**.
2. Em **Overview**, copie o **Application (client) ID** (e o **Directory (tenant) ID**, se escolheu "só a sua empresa").
3. **Authentication** → **Add a platform** → **Mobile and desktop applications** → redirect URI personalizado:
   ```
   http://localhost/cb
   ```
   Ainda em **Authentication**, ative **Allow public client flows** → **Save**. Não precisa de segredo para uso local.
4. **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions**:
   `Calendars.ReadWrite`, `MailboxSettings.Read`, `User.Read`, `offline_access`. Nunca *Application permissions*.
5. Conta corporativa: se a empresa exigir, um admin clica em **Grant admin consent for <empresa>**.

### 3. Adicione ao seu cliente MCP

**Claude Desktop** (`claude_desktop_config.json`, veja [examples/](examples/claude_desktop_config.json)):

```json
{
  "mcpServers": {
    "calendar": {
      "command": "node",
      "args": ["/caminho/absoluto/outlook-calendar-lite-mcp/src/main.ts"],
      "env": { "OUTLOOK_MCP_CLIENT_ID": "<seu client id>" }
    }
  }
}
```

**Claude Code:**

```bash
claude mcp add calendar -e OUTLOOK_MCP_CLIENT_ID=<seu client id> -- node /caminho/absoluto/outlook-calendar-lite-mcp/src/main.ts
```

**Outros clientes MCP** (Cursor, VS Code, modelos locais…): comando `node /caminho/absoluto/outlook-calendar-lite-mcp/src/main.ts`, transporte stdio, mesma variável.

Reinicie o cliente. O navegador abre na escolha de conta Microsoft; entre e pronto.
O token fica em `~/.outlook-calendar-lite-mcp/default.json` (permissão 600) e renova sozinho.
Para usar outra conta, peça "trocar a conta" (ferramenta `cx`, acao `trocar`).

No macOS, deixe a pasta fora de `Downloads`/`Mesa`/`Documentos` (ex.: `~/outlook-calendar-lite-mcp`), senão o cliente pode não ter permissão de leitura.

### Experimente

> "o que tenho hoje?" · "acha uma hora livre com a ana@acme.com esta semana" · "marca e convida ela, com Teams" · "muda o dentista para as 11h" · "aceita a reunião da ACME"

## Ferramentas

| tool | faz |
|---|---|
| `ls` | agenda de um período (hoje, amanhã, esta semana, próxima semana, 7/30 dias ou datas) |
| `rd` | lê um evento: horário, local, organizador, participantes com a resposta de cada um, link do Teams, recorrência, descrição |
| `find` | busca eventos por título, local, organizador ou participante |
| `free` | horários livres para uma reunião, cruzando a sua agenda (e a dos participantes, em conta corporativa) dentro do horário de trabalho |
| `cr` | cria evento: dia inteiro, recorrente, online (Teams), lembrete, participantes. **Pede confirmação quando convida pessoas**; avisa conflitos |
| `ed` | altera evento que você organiza (horário, título, local, participantes…). **Pede confirmação quando os participantes são avisados** |
| `rm` | cancela (participantes recebem o cancelamento, com mensagem opcional) ou tira da sua agenda. **Pede confirmação** |
| `rsp` | responde a um convite: aceitar, recusar ou talvez, com comentário. **Pede confirmação quando o organizador é avisado** |
| `cx` | conta, fuso, horário de trabalho e calendários; `trocar` conta |

IDs: a coluna `id` de `ls`/`find` (e o id devolvido por `cr`) é o id permanente do evento no Graph (id imutável). Passe esse id em `alvo` nas outras ferramentas, agora ou numa sessão futura.
Eventos recorrentes: `ed`, `rm` e `rsp` agem numa ocorrência; passe `serie: true` para a série inteira.
Datas no formato `AAAA-MM-DDTHH:mm` no seu fuso; só `AAAA-MM-DD` = dia inteiro.

**Confirmação.** Se o cliente suporta *elicitation*, o MCP pergunta direto ao usuário.
Se não, a 1ª chamada devolve o preview e um código `cf`; só a chamada idêntica com esse `cf` executa.
O cliente deve mostrar o preview e só mandar o `cf` depois da aprovação explícita do usuário.

**Contas pessoais.** Ver a disponibilidade de outras pessoas (`free` com participantes) e reuniões no Teams são recursos do Microsoft 365; no Outlook.com, o `free` usa só a sua agenda e avisa isso.

## Variáveis

| var | padrão |
|---|---|
| `OUTLOOK_MCP_CLIENT_ID` | obrigatória |
| `OUTLOOK_MCP_TENANT` | `common` (use o id do tenant para travar em uma empresa) |
| `OUTLOOK_MCP_PROFILE` | `default` (um perfil por conta) |
| `OUTLOOK_MCP_LOGIN` | navegador; `device` = código de dispositivo (SSH, sem navegador) |
| `OUTLOOK_MCP_TZ` | o fuso da sua caixa; defina um nome IANA (ex.: `America/Sao_Paulo`) para forçar as contas de data |
| `OUTLOOK_MCP_HOME` | `~/.outlook-calendar-lite-mcp` |

Terminal: `node src/main.ts login | whoami | logout`.
Atrás de proxy corporativo: Node 22.21+ com `NODE_USE_ENV_PROXY=1` e `HTTPS_PROXY`.

## Modo remoto (opcional, hospedado por você)

Para usar no claude.ai como conector personalizado, suba a sua própria instância HTTPS:

```bash
export OUTLOOK_MCP_CLIENT_ID=<client id>
export OUTLOOK_MCP_URL=https://seu-dominio.exemplo        # URL pública HTTPS
export OUTLOOK_MCP_ALLOW=voce@empresa.com,empresa.com     # quem pode conectar (emails ou domínios)
node src/main.ts http                                     # escuta na PORT (8787)
```

No Entra, adicione a plataforma **Web** com o redirect `https://seu-dominio.exemplo/ms/cb` (e um segredo em
`OUTLOOK_MCP_CLIENT_SECRET`, se o Entra pedir). No claude.ai, adicione um conector personalizado com `https://seu-dominio.exemplo/mcp`.
Para testar rápido: `cloudflared tunnel --url http://localhost:8787` gera uma URL HTTPS temporária.

**Sempre defina `OUTLOOK_MCP_ALLOW`.** Sem ele, qualquer pessoa com a URL pode conectar a própria conta ao seu servidor.

## Segurança

- Só permissões delegadas: o servidor faz o que o usuário logado pode fazer, nada além.
- Os tokens ficam na sua máquina (`~/.outlook-calendar-lite-mcp`, permissão 600). No modo remoto, em `http-store.json` no servidor.
- **Convites não são confiáveis.** A descrição de um evento pode ter texto tentando instruir a IA.
  A confirmação obrigatória antes de qualquer coisa que avise outras pessoas é a principal defesa; não a remova.
- Para relatar uma vulnerabilidade, veja [SECURITY.md](SECURITY.md).

## Estrutura do projeto

```
src/
  main.ts     entrada: stdio (padrão) | http | login | logout | whoami
  mcp.ts      protocolo MCP: JSON-RPC, validação de schema, confirmação (elicitation ou código cf)
  http.ts     modo remoto: Streamable HTTP + servidor OAuth que repassa o login à Microsoft
  auth.ts     login Microsoft: navegador (PKCE), device code, refresh, token em arquivo
  ctx.ts      estado por usuário (token, caches)
  g.ts        cliente Graph mínimo: get/post/patch/delete, retry em 429/401, ids imutáveis
  ids.ts      valida os ids de evento recebidos pelas ferramentas
  fmt.ts      saída tabular, limpeza de texto
  lib.ts      fuso e contas de data, períodos, tabela de eventos, recorrência, helpers
  tools/      uma ferramenta por arquivo + index.ts (ev.ts: campos do evento usados por criar e alterar)
test/
  mock.ts     Microsoft simulada (login + calendário do Graph)
  run.ts      clientes MCP reais contra o mock, local e remoto
```

## Desenvolvimento

```bash
node test/run.ts   # 66 checagens, sem rede nem conta Microsoft
npx tsc -p .       # checagem de tipos (opcional)
```

## Licença

[MIT](LICENSE)
