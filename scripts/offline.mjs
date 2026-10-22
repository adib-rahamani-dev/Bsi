// Produce a self-contained, encrypted local report. No plaintext source data is embedded.
import fs from 'node:fs/promises';
import {randomBytes,pbkdf2Sync,createCipheriv} from 'node:crypto';
const password=await new Promise(resolve=>{let data='';process.stdin.setEncoding('utf8');process.stdin.on('data',c=>data+=c);process.stdin.on('end',()=>resolve(data.trim()));});
if(!password)throw Error('Pass a password through stdin.');
const source=await fs.readFile('private/workbook.json'),salt=randomBytes(16),iv=randomBytes(12);
const key=pbkdf2Sync(password,salt,600000,32,'sha256'),cipher=createCipheriv('aes-256-gcm',key,iv);
const encrypted=Buffer.concat([cipher.update(source),cipher.final(),cipher.getAuthTag()]);
const payload={salt:salt.toString('base64'),iv:iv.toString('base64'),data:encrypted.toString('base64')};
let html=await fs.readFile('public/index.html','utf8');
const css=(await fs.readFile('public/style.css','utf8'))+'\n'+(await fs.readFile('public/refinement.css','utf8')),excel=await fs.readFile('public/vendor/exceljs.min.js','utf8');
const engine=(await fs.readFile('public/analyzer.js','utf8')).replace(/^export /gm,'');
const report=(await fs.readFile('public/report.js','utf8')).replace(/^export /gm,'');
const app=(await fs.readFile('public/app.js','utf8')).replace(/^import .*?;\r?\n/gm,'');
html=html.replace('<link rel="stylesheet" href="/style.css">',()=>`<style>${css}</style>`).replace('<link rel="stylesheet" href="/refinement.css">','').replace('<script defer src="/vendor/exceljs.min.js"></script>','').replace('<script type="module" src="/app.js"></script>','');
const safeScript=s=>s.replace(/<\/script/gi,'<\\/script');
html=html.replace('</body>',()=>`<script>window.BSI_OFFLINE=${JSON.stringify(payload)};</script><script>${safeScript(excel)}</script><script type="module">${safeScript(engine+'\n'+report+'\n'+app)}</script></body>`);
await fs.mkdir('outputs',{recursive:true});await fs.writeFile('outputs/BSI-report.html',html);
console.log('Created outputs/BSI-report.html (encrypted local report, excluded from Git).');
