import type { Tool } from '../mcp.ts';
import ls from './ls.ts';
import rd from './rd.ts';
import find from './find.ts';
import free from './free.ts';
import cr from './cr.ts';
import ed from './ed.ts';
import rm from './rm.ts';
import rsp from './rsp.ts';
import cx from './cx.ts';
import ntf from './ntf.ts';

// ordem = ordem em tools/list
export const tools: Tool[] = [ls, rd, find, free, cr, ed, rm, rsp, cx, ntf];

export const INSTR = "Calendário Microsoft (Outlook). Horários sempre no fuso do usuário (ls e cx mostram o fuso e a data/hora de agora); escreva datas como AAAA-MM-DDTHH:mm. Saída tabular separada por '|'. Cada evento tem um id permanente (coluna id de ls/find): passe esse id para as outras ferramentas; ele pode ser guardado para usar depois. Antes de marcar com outras pessoas, use free. cr/ed com participantes, rm e rsp enviam aviso a outras pessoas e pedem confirmação do usuário. ntf liga as notificações de mudança do calendário para o app configurado no servidor.";
