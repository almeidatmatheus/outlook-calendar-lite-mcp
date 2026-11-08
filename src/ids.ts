// ids: o MCP usa sempre o id real do provedor, que é permanente e pode ser salvo e reutilizado
export function res(a: unknown): string {
  const v = String(a ?? '').trim();
  if (!v) throw new Error('informe o id do evento (coluna id da listagem)');
  return v;
}

export function resN(a: unknown): string[] {
  const v = (Array.isArray(a) ? a : [a]).map(x => String(x ?? '').trim()).filter(Boolean);
  if (!v.length) throw new Error('informe ao menos um id');
  return v;
}
