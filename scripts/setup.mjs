import { randomBytes,scryptSync } from 'node:crypto';
import { writeFile } from 'node:fs/promises';
import readline from 'node:readline/promises';
const rl=readline.createInterface({input:process.stdin,output:process.stdout});
const password=await rl.question('Local BSI password: ');rl.close();
if(!password)throw Error('Password required');
const salt=randomBytes(16).toString('hex');
await writeFile('.env',`BSI_PASSWORD_HASH=${salt}:${scryptSync(password,salt,64).toString('hex')}\nBSI_SESSION_SECRET=${randomBytes(48).toString('hex')}\n`,{flag:'wx'});
console.log('Created local .env (excluded from Git). Set the same two variables in Vercel.');
