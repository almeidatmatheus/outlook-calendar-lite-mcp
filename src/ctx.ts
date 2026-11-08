// estado por usuário/conexão: token, ids curtos, caches. stdio usa o padrão; http um por usuário
import { AsyncLocalStorage } from 'node:async_hooks';

export type Ids = { s2g: Map<string, string>; g2s: Map<string, string>; n: number; cur: string; last: string[] };
export type Ctx = {
  tok: (force?: boolean) => Promise<string>;
  ids: Ids;
  c: Record<string, any>;
  caps: any;
  ask?: (method: string, params: unknown) => Promise<any>;
  reset?: () => Promise<string>;
};

export const newIds = (): Ids => ({ s2g: new Map(), g2s: new Map(), n: 0, cur: '', last: [] });
export const als = new AsyncLocalStorage<Ctx>();
let def: Ctx | null = null;
export const setDef = (c: Ctx) => { def = c; };
export const cx = (): Ctx => { const c = als.getStore() ?? def; if (!c) throw new Error('sem contexto'); return c; };
