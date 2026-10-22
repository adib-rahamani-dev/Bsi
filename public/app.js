import {analyze,RULES,normalize} from './analyzer.js';
const $=s=>document.querySelector(s),$$=s=>[...document.querySelectorAll(s)];
const fa=n=>String(n??'').replace(/\d/g,c=>'۰۱۲۳۴۵۶۷۸۹'[c]);
const num=v=>{if(v===null||v===undefined)return 'ناموجود';try{return BigInt(v).toLocaleString('fa-IR');}catch{return String(v);}};
const escape=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
let book,report,page=0,recordPage=0,current='overview',offline=Boolean(window.BSI_OFFLINE);
const statusNames={high:'اولویت بالا',medium:'نیازمند بررسی',info:'پیگیری مطالبات'};
function notice(message){$('#notice').hidden=!message;$('#notice').textContent=message;}
async function jsonRequest(path,options){const res=await fetch(path,options);let data;try{data=await res.json();}catch{throw Error('پاسخ سرور قابل خواندن نیست.');}if(!res.ok)throw Error(data.error??'نشست منقضی شده؛ دوباره وارد شوید.');return data;}
function showLogin(){ $('#login').hidden=false;$('#workspace').hidden=true;$('#password').value='';book=undefined;report=undefined;$('#findings-list').replaceChildren();$('#records-body').replaceChildren();$('#detail').close();$('#detail-content').replaceChildren();}
function enter(){$('#login').hidden=true;$('#workspace').hidden=false;$('#password').value='';$('#empty').hidden=Boolean(report);}
async function decrypt(password){const payload=window.BSI_OFFLINE;const decode=s=>Uint8Array.from(atob(s),c=>c.charCodeAt(0));const material=await crypto.subtle.importKey('raw',new TextEncoder().encode(password),'PBKDF2',false,['deriveKey']);const key=await crypto.subtle.deriveKey({name:'PBKDF2',salt:decode(payload.salt),iterations:600000,hash:'SHA-256'},material,{name:'AES-GCM',length:256},false,['decrypt']);return JSON.parse(new TextDecoder().decode(await crypto.subtle.decrypt({name:'AES-GCM',iv:decode(payload.iv)},key,decode(payload.data))));}
$('#login-form').addEventListener('submit',async e=>{e.preventDefault();$('#login-button').disabled=true;$('#login-error').textContent='';try{if(offline){book=await decrypt($('#password').value);enter();run();}else{await jsonRequest('/api/login',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({password:$('#password').value})});enter();}}catch(err){$('#login-error').textContent=offline?'رمز نادرست است یا مرورگر امکان رمزگشایی این فایل را ندارد.':err.message;}finally{$('#login-button').disabled=false;}});
$('#logout').onclick=async()=>{try{if(!offline)await jsonRequest('/api/logout',{method:'POST'});showLogin();}catch(err){notice(err.message);}};
function view(name){current=name;const title={overview:'نمای کلی بررسی',findings:'یافته‌ها و هشدارها',records:'داده‌های سند',method:'قواعد و روش بررسی'}[name];$('#page-title').textContent=title;$('#breadcrumb').textContent=title;$$('.view').forEach(e=>e.hidden=e.id!==name||(!report&&e.id!=='method'));$('#empty').hidden=Boolean(report)||name==='method';$$('nav button').forEach(b=>b.classList.toggle('active',b.dataset.view===name));}
$$('[data-view]').forEach(b=>b.onclick=()=>view(b.dataset.view));
$('#sample').onclick=async()=>{try{notice('در حال خواندن نمونه محلی…');if(!offline)book=await jsonRequest('/api/sample');if(!book)throw Error('نمونه محلی در دسترس نیست. فایل اکسل را انتخاب کنید.');run();}catch(err){notice(err.message);}};
$('#file').onchange=async e=>{
 const file=e.target.files[0];if(!file)return;
 try{
  if(!/\.xlsx$/i.test(file.name))throw Error('فقط فایل XLSX پشتیبانی می‌شود.');
  if(file.size>10*1024*1024)throw Error('حداکثر حجم مجاز ۱۰ مگابایت است.');
  notice('در حال خواندن و بررسی فایل…');
  if(!window.ExcelJS)throw Error('کتابخانه اکسل بارگذاری نشده است. npm run build را اجرا کنید.');
  const bytes=await file.arrayBuffer(),wb=new ExcelJS.Workbook();await wb.xlsx.load(bytes);
  const next={name:file.name,sha256:[...new Uint8Array(await crypto.subtle.digest('SHA-256',bytes))].map(b=>b.toString(16).padStart(2,'0')).join(''),sheets:[]};
  let count=0;
  for(const sheet of wb.worksheets){
   if(sheet.rowCount>100000||sheet.columnCount>200)throw Error('ابعاد شیت بیش از حد مجاز است (۱۰۰ هزار ردیف و ۲۰۰ ستون).');
   const rows=[];
   for(let i=1;i<=sheet.rowCount;i++){
    count++;if(count>100000)throw Error('حداکثر مجموع ردیف‌ها ۱۰۰ هزار است.');
    const row=[];for(let j=1;j<=sheet.columnCount;j++){
     const cell=sheet.getCell(i,j),raw=cell.value;let v=raw,formula=null,error=false,missingCache=false;
     if(raw&&typeof raw==='object'){
      if(raw.formula||raw.sharedFormula){formula=raw.formula??raw.sharedFormula;v=raw.result;missingCache=v===undefined||v===null;}
      else if(raw.richText)v=raw.richText.map(t=>t.text).join('');else if(raw.text!==undefined)v=raw.text;else if(raw.error){error=true;v=raw.error;}else if(raw instanceof Date)v=raw.toISOString();
     }
     if(v&&typeof v==='object'&&v.error){error=true;v=v.error;}
     if(typeof v==='number'&&!Number.isSafeInteger(v))v=`عدد با دقت نامطمئن: ${v}`;
     row.push({v:v===undefined||v===null?'':String(v),formula,error,missingCache});
    }rows.push(row);
   }next.sheets.push({name:sheet.name,rows});
  }
  // Replace the current report only after successful parsing and analysis.
  const result=analyze(next,{threshold:$('#threshold').value});book=next;report=result;render();notice('بررسی کامل شد. فایل در حافظه همین صفحه است و به سرور ارسال نشد.');view('overview');
 }catch(err){notice(`فایل جدید بررسی نشد: ${err.message} ${report?'گزارش قبلی همچنان نمایش داده می‌شود.':''}`);}finally{e.target.value='';}
};
function run(){if(!book)return;try{report=analyze(book,{threshold:$('#threshold').value});render();notice('');view('overview');}catch(err){notice(err.message);}}
function render(){
 const s=report.summary;$('#empty').hidden=true;$('#print').disabled=false;
 $('#subtitle').textContent=report.name;$('#document-name').textContent=report.name;$('#metadata').textContent=report.sheets.filter(s=>s.supported).map(s=>s.metadata).join(' | ');
 $('#records-count').textContent=num(s.records);$('#subtotal-count').textContent=`در ${num(s.subtotals)} سرفصل`;$('#high-count').textContent=num(s.high);$('#medium-count').textContent=num(s.medium);$('#watch-count').textContent=num(s.watch);$('#affected-count').textContent=`${num(s.affected)} ردیف دارای هشدار یا بررسی`;$('#nav-count').textContent=num(report.findings.length);
 const counts=RULES.map(([id,name])=>({id,name,count:report.findings.filter(f=>f.rule===id).length})).filter(x=>x.count);
 $('#bars').innerHTML=counts.length?counts.map(c=>`<div class="bar-row"><span>${escape(c.name)}</span><div class="track"><i style="width:${Math.max(3,c.count/Math.max(...counts.map(x=>x.count))*100)}%"></i></div><b>${num(c.count)}</b></div>`).join(''):'<p class="muted">در کنترل‌های اجراشده یافته‌ای ثبت نشد.</p>';
 const fail=report.checks.filter(c=>c.status==='fail').length,unavailable=report.checks.filter(c=>c.status==='unavailable').length,pass=report.checks.filter(c=>c.status==='pass').length;
 $('#check-badge').textContent=fail?'دارای مغایرت':unavailable?'بخشی بررسی نشد':report.checks.length?'تطبیق کنترل‌های اجراشده':'بررسی نشده';
 $('#check-summary').innerHTML=`<div class="check-row"><span>تطبیق جمع‌ها (کنترل ستونی)</span><b>${num(pass)} منطبق / ${num(report.checks.length)}</b></div><div class="check-row"><span>مغایرت جمع‌ها</span><b>${num(fail)}</b></div><div class="check-row"><span>کنترل عددی ناممکن</span><b>${num(unavailable)}</b></div><div class="check-row"><span>اختلاف مانده با گردش‌ها</span><b>${num(report.findings.filter(f=>f.rule==='balance').length)}</b></div>`;
 $('#top-findings').innerHTML=report.findings.slice(0,4).map(f=>findingHTML(f)).join('')||'<p class="muted">یافته‌ای در کنترل‌های اجراشده ثبت نشد.</p>';
 $('#coverage').innerHTML=report.sheets.map(s=>`<p>${escape(s.name)}: ${s.supported?`${num(s.records)} معین و ${num(s.subtotals)} جمع سرفصل بررسی شد.`:`بررسی نشد — ${escape(s.reason)}`}</p>`).join('');
 $('#fingerprint').textContent=`SHA-256: ${report.sha256||'unavailable'}`;
 $('#generated').textContent=`زمان بررسی: ${new Date(report.generatedAt).toLocaleString('fa-IR')}`;
 page=0;recordPage=0;renderFindings();renderRecords();view(current);
 if(report.sheets.some(s=>!s.supported))notice('بعضی شیت‌ها با قالب پشتیبانی‌شده مطابقت ندارند و بررسی مالی نشده‌اند. جزئیات در روش بررسی آمده است.');
}
function findingHTML(f,print=false){return `<div class="finding"><div><div class="finding-title">${escape(f.title)} <span class="badge ${f.severity}">${statusNames[f.severity]}</span></div><p>${escape(f.account)}</p><div class="finding-meta"><span>${escape(f.sheet)} · ردیف ${fa(f.row)}</span><span>کد ${escape(f.code||'—')}</span>${f.amount!==null?`<span>مقدار مرتبط: ${escape(num(f.amount))}</span>`:''}</div><p class="print-details">${escape(f.detail)}</p>${print?`<p>${escape(f.detail)}</p>`:''}</div><button class="secondary" data-finding="${f.id}">مشاهده مستندات ↗</button></div>`;}
function filtered(){const q=normalize($('#search').value);return report.findings.filter(f=>(!$('#severity').value||f.severity===$('#severity').value)&&(!$('#rule').value||f.rule===$('#rule').value)&&(!q||normalize(`${f.title} ${f.account} ${f.code} ${f.detail} ${f.row} ${f.sheet}`).includes(q)));}
function renderFindings(){if(!report)return;const list=filtered(),pages=Math.max(1,Math.ceil(list.length/20));page=Math.min(page,pages-1);$('#findings-list').innerHTML=list.slice(page*20,(page+1)*20).map(f=>findingHTML(f)).join('')||'<p class="muted">یافته‌ای با این فیلتر وجود ندارد.</p>';$('#filter-count').textContent=`${num(list.length)} یافته از ${num(report.findings.length)} مورد؛ یک حساب ممکن است چند یافته داشته باشد.`;$('#page-number').textContent=`صفحه ${num(page+1)} از ${num(pages)}`;$('#prev').disabled=page===0;$('#next').disabled=page>=pages-1;}
['search','severity','rule'].forEach(id=>$('#'+id).addEventListener(id==='search'?'input':'change',()=>{page=0;renderFindings();}));
$('#prev').onclick=()=>{page--;renderFindings();};$('#next').onclick=()=>{page++;renderFindings();};
function renderRecords(){if(!report)return;const q=normalize($('#record-search').value),list=report.rows.filter(r=>!q||normalize(`${r.title} ${r.code} ${r.row} ${r.group}`).includes(q)),pages=Math.max(1,Math.ceil(list.length/30));recordPage=Math.min(recordPage,pages-1);$('#records-body').innerHTML=list.slice(recordPage*30,(recordPage+1)*30).map(r=>`<tr><td>${escape(r.sheet)} / ${fa(r.row)}</td><td>${escape(r.title)}</td><td>${escape(r.code)}</td><td>${{detail:'معین',subtotal:'جمع سرفصل',grand:'جمع کل'}[r.kind]}</td><td>${escape(r.nature)}</td>${[2,3,6].map(i=>`<td class="numeric">${escape(num(r.values[i]))}</td>`).join('')}</tr>`).join('');$('#record-page').textContent=`${num(list.length)} ردیف · صفحه ${num(recordPage+1)} از ${num(pages)}`;$('#record-prev').disabled=recordPage===0;$('#record-next').disabled=recordPage>=pages-1;}
$('#record-search').oninput=()=>{recordPage=0;renderRecords();};$('#record-prev').onclick=()=>{recordPage--;renderRecords();};$('#record-next').onclick=()=>{recordPage++;renderRecords();};
document.addEventListener('click',e=>{const button=e.target.closest('[data-finding]');if(!button||!report)return;const f=report.findings.find(f=>f.id===button.dataset.finding),row=report.rows.find(r=>r.row===f.row&&r.sheet===f.sheet);$('#detail-content').innerHTML=`<h2>${escape(f.title)}</h2><p><b>${escape(f.account)}</b></p><p>${escape(f.sheet)}، ردیف ${fa(f.row)}، کد ${escape(f.code||'—')}</p><p>${escape(f.detail)}</p><h3>مقادیر ردیف منبع</h3><pre>${escape(row?['ردیف','سرفصل','معین','ماهیت','عنوان حساب','کد سپهری','تعداد بدهکار','تعداد بستانکار','بدهکار روز','بستانکار روز','گردش بدهکار','گردش بستانکار','مانده'].map((label,i)=>`${label}: ${row.source[i]}`).join('\n'):'ردیف ساختاری گزارش')}</pre><p>اقدام پیشنهادی: دریافت ریز سند و توضیح مسئول حساب، سپس تطبیق با ضوابط داخلی بانک.</p>`;$('#detail').showModal();});
$('#close-detail').onclick=()=>$('#detail').close();
$('#settings').onsubmit=e=>{e.preventDefault();if(book)run();else notice('ابتدا یک فایل انتخاب کنید.');};
$('#rules-list').innerHTML=RULES.map(([id,name,detail])=>`<div class="rule-item"><h3>${escape(name)}</h3><p>${escape(detail)}</p></div>`).join('');
$('#rule').innerHTML+=[...RULES].map(([id,name])=>`<option value="${id}">${escape(name)}</option>`).join('');
$('#export').onclick=()=>{const list=filtered(),columns=['شناسه','اولویت','قاعده','شیت','ردیف اکسل','کد سپهری','عنوان حساب','یافته','مقدار مرتبط در واحد فایل','شرح'];const safe=s=>{s=String(s??'');if(/^[=+\-@\t\r\n]/.test(s))s="'"+s;return '"'+s.replace(/"/g,'""')+'"';};const csv='\uFEFF'+[columns,...list.map(f=>[f.id,statusNames[f.severity],f.rule,f.sheet,f.row,f.code,f.account,f.title,f.amount,f.detail])].map(row=>row.map(safe).join(',')).join('\r\n');const url=URL.createObjectURL(new Blob([csv],{type:'text/csv;charset=utf-8'})),a=document.createElement('a');a.href=url;a.download='BSI-findings.csv';a.click();setTimeout(()=>URL.revokeObjectURL(url),2000);};
$('#print').onclick=()=>{if(!report)return;const original=$('#findings-list').innerHTML;$('#findings-list').innerHTML=report.findings.map(f=>findingHTML(f)).join('');window.print();$('#findings-list').innerHTML=original;};
if(!offline){jsonRequest('/api/session').then(enter).catch(()=>{});}else{$('#sample').textContent='نمایش دوباره گزارش نمونه';$('.privacy').textContent='نسخه HTML محلی؛ داده نمونه با رمز شما رمزگذاری شده است. برای میزبانی اینترنتی از نسخه سرور استفاده کنید.';}
