#!/usr/bin/env node
import { stdio, type Info } from './mcp.ts';
import { tools, INSTR } from './tools/index.ts';
import { login, logout, tok, relogin, hasLocal, startLogin, CID } from './auth.ts';
import { setDef, type Ctx } from './ctx.ts';
import { g } from './g.ts';

const E = process.env;
const info: Info = { name: 'outlook-calendar-lite-mcp', version: '1.0.0', instructions: INSTR };
const c = process.argv[2];
const local: Ctx = { tok, c: {}, caps: {}, reset: relogin };
setDef(local);

if (c === 'login') await login();
else if (c === 'logout') { logout(); process.stderr.write('ok\n'); }
else if (c === 'whoami') { const r = await g('GET', '/me?$select=displayName,mail,userPrincipalName'); console.log(`${r.displayName} <${r.mail || r.userPrincipalName}>`); }
else if (c === 'http') (await import('./http.ts')).serveHttp(tools, info);
// stdio: ao conectar sem conta, já abre o navegador para escolher a conta
else stdio(tools, info, local, { onInit: () => { if (CID && !hasLocal()) startLogin().catch(e => process.stderr.write(`${e.message}\n`)); } });
