import http from 'node:http';
import { readFile } from 'node:fs/promises';
import { createHmac, timingSafeEqual, scryptSync, randomBytes } from 'node:crypto';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root=path.dirname(fileURLToPath(import.meta.url));
const attempts=new Map();
const cookieName='bsi_session';
const secret=process.env.BSI_SESSION_SECRET;
const passwordHash=process.env.BSI_PASSWORD_HASH;
const ttl=60*60*4;
const sign=s=>createHmac('sha256',secret).update(s).digest('base64url');
function authenticated(req){
 if(!secret||!passwordHash)return false;
 const value=(req.headers.cookie??'').split(';').map(s=>s.trim()).find(s=>s.startsWith(cookieName+'='))?.slice(cookieName.length+1);
 if(!value)return false;
 const [expires,nonce,signature]=value.split('.');
 if(!expires||!nonce||!signature||!/^\d+$/.test(expires)||Number(expires)<Date.now())return false;
 const expected=sign(`${expires}.${nonce}`);
 return signature.length===expected.length&&timingSafeEqual(Buffer.from(signature),Buffer.from(expected));
}
function send(res,status,body,type='application/json; charset=utf-8'){
 res.writeHead(status,{'Content-Type':type,'Cache-Control':'no-store','X-Content-Type-Options':'nosniff','Referrer-Policy':'no-referrer','X-Frame-Options':'DENY','Content-Security-Policy':"default-src 'self'; script-src 'self'; style-src 'self' 'unsafe-inline'; img-src 'self' data:; connect-src 'self'; font-src 'self'; frame-ancestors 'none'; form-action 'self'; base-uri 'none'"});res.end(typeof body==='string'||Buffer.isBuffer(body)?body:JSON.stringify(body));
}
export default async function handler(req,res){
 const url=new URL(req.url,'http://localhost');
 const secure=process.env.VERCEL||req.headers['x-forwarded-proto']==='https';
 const cookie=v=>`${cookieName}=${v}; HttpOnly; SameSite=Strict; Path=/; ${secure?'Secure; ':''}`;
 const sameOrigin=()=>!req.headers.origin||req.headers.origin===`${secure?'https':'http'}://${req.headers.host}`;
 try{
  if(url.pathname==='/api/login'&&req.method==='POST'){
   if(!secret||!passwordHash)return send(res,503,{error:'رمز سرور تنظیم نشده است. راهنمای README را ببینید.'});
   if(!sameOrigin())return send(res,403,{error:'درخواست نامعتبر است.'});
   // Instance-local limiter complements expensive scrypt. Production needs a shared limiter for multiple instances.
   const ip=req.socket?.remoteAddress??'serverless';
   if(attempts.size>10000)for(const [k,v]of attempts)if(v.until<Date.now())attempts.delete(k);
   let a=attempts.get(ip);if(!a||a.until<Date.now()){a={count:0,until:Date.now()+900000};attempts.set(ip,a);}
   if(a.count>=8)return send(res,429,{error:'تعداد تلاش زیاد است. ۱۵ دقیقه دیگر امتحان کنید.'});
   let body='';for await(const chunk of req){body+=chunk;if(body.length>1024)return send(res,413,{error:'درخواست بزرگ است.'});}
   let input;try{input=JSON.parse(body);}catch{return send(res,400,{error:'درخواست نامعتبر است.'});}
   a.count++;
   const [salt,hash]=passwordHash.split(':');
   const actual=scryptSync(String(input.password??''),salt,64),expected=Buffer.from(hash??'','hex');
   if(expected.length!==64||!timingSafeEqual(actual,expected))return send(res,401,{error:'رمز ورود نادرست است.'});
   attempts.delete(ip);
   const data=`${Date.now()+ttl*1000}.${randomBytes(16).toString('hex')}`;
   res.setHeader('Set-Cookie',cookie(`${data}.${sign(data)}`)+`Max-Age=${ttl}`);
   return send(res,200,{ok:true});
  }
  if(url.pathname==='/api/logout'&&req.method==='POST'){
   if(!sameOrigin())return send(res,403,{error:'درخواست نامعتبر است.'});
   res.setHeader('Set-Cookie',cookie('')+'Max-Age=0');return send(res,200,{ok:true});
  }
  if(url.pathname==='/api/session')return send(res,authenticated(req)?200:401,{authenticated:authenticated(req)});
  if(url.pathname==='/api/sample'){
   if(!authenticated(req))return send(res,401,{error:'ابتدا وارد شوید.'});
   try{return send(res,200,await readFile(path.join(root,'private/workbook.json')));}catch{return send(res,404,{error:'فایل نمونه روی این سرور نگهداری نمی‌شود. فایل اکسل را از دستگاه انتخاب کنید.'});}
  }
  if(!['GET','HEAD'].includes(req.method))return send(res,405,{error:'روش نامعتبر'});
  const files={'/':'index.html','/style.css':'style.css','/refinement.css':'refinement.css','/app.js':'app.js','/analyzer.js':'analyzer.js','/report.js':'report.js','/vendor/exceljs.min.js':'vendor/exceljs.min.js'};
  const file=files[url.pathname];if(!file)return send(res,404,{error:'یافت نشد'});
  // Application code is public; bank records are never embedded in public assets.
  const types={'.html':'text/html; charset=utf-8','.css':'text/css; charset=utf-8','.js':'text/javascript; charset=utf-8'};
  return send(res,200,await readFile(path.join(root,'public',file)),types[path.extname(file)]);
 }catch(e){console.error('BSI request failed:',e.code??e.name);return send(res,500,{error:'خطای پردازش درخواست'});}
}
if(process.argv[1]&&path.resolve(process.argv[1])===fileURLToPath(import.meta.url))http.createServer(handler).listen(Number(process.env.PORT)||3000,'127.0.0.1',()=>console.log('BSI: http://localhost:3000'));
