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

// ordem = ordem em tools/list
export const tools: Tool[] = [ls, rd, find, free, cr, ed, rm, rsp, cx];

export const INSTR = "Calendário Microsoft (Outlook). Horários sempre no fuso do usuário (ls e cx mostram o fuso e a data/hora de agora); escreva datas como AAAA-MM-DDTHH:mm. Saída tabular separada por '|'. Eventos têm id curto ('3'); 'atual' = último evento lido ou criado; 'ultimo' = 1º da última listagem. Antes de marcar com outras pessoas, use free. cr/ed com participantes, rm e rsp enviam aviso a outras pessoas e pedem confirmação do usuário.";
