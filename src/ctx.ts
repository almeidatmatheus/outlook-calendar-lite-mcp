// estado por usuário/conexão: token, caches. stdio usa o padrão; http um por usuário
import { AsyncLocalStorage } from 'node:async_hooks';

export type Ctx = {
  tok: (force?: boolean) => Promise<string>;
  c: Record<string, any>;
  caps: any;
  ask?: (method: string, params: unknown) => Promise<any>;
  reset?: () => Promise<string>;
};

export const als = new AsyncLocalStorage<Ctx>();
let def: Ctx | null = null;
export const setDef = (c: Ctx) => { def = c; };
export const cx = (): Ctx => { const c = als.getStore() ?? def; if (!c) throw new Error('sem contexto'); return c; };
