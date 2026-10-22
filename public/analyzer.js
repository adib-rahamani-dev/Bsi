export const RULES = [
 ['balance','مغایرت مانده','مانده باید برابر گردش بستانکار منهای گردش بدهکار باشد.'],
 ['subtotal','مغایرت جمع سرفصل','هفت ستون عددی جمع سرفصل با ردیف‌های معین همان سرفصل مقایسه می‌شوند.'],
 ['grand','مغایرت جمع کل','جمع کل با معین‌ها، بدون شمارش دوباره جمع سرفصل‌ها، تطبیق می‌یابد.'],
 ['nature','مانده خلاف ماهیت','مانده مثبت برای ماهیت بدهکار و مانده منفی برای ماهیت بستانکار، نیازمند بررسی است. ماهیت دوگانه مستثنا است.'],
 ['count','تعداد و مبلغ ناسازگار','مبلغ روز غیرصفر با تعداد صفر یا تعداد مثبت با مبلغ صفر علامت‌گذاری می‌شود؛ ممکن است ناشی از تعریف گزارش باشد.'],
 ['turnover','کنترل گردش','مبلغ یا تعداد منفی و مبلغ روز بزرگ‌تر از گردش تجمعی علامت‌گذاری می‌شود.'],
 ['duplicate','شناسه تکراری','تکرار کد سپهری یا ترکیب سرفصل و معین در یک شیت کنترل می‌شود.'],
 ['quality','کیفیت داده','داده عددی نامعتبر، شناسه خالی، خطای اکسل و فرمول فاقد مقدار ذخیره‌شده گزارش می‌شود.'],
 ['watch','مطالبات نیازمند پیگیری','مانده غیرصفر حساب‌های سررسیدگذشته، معوق و مشکوک‌الوصول بر مبنای عنوان حساب یا سرفصل مادر فهرست می‌شود؛ این مورد مغایرت حسابداری نیست.'],
 ['suspense','حساب‌های باز نیازمند پیگیری','مانده حساب‌های مغایرت و بدهکاران یا بستانکاران موقت، در حد بررسی انتخاب‌شده یا بیشتر، علامت‌گذاری می‌شود؛ قدمت مانده مشخص نیست.'],
 ['cross','تطبیق حساب‌های مقابل','حساب‌های دقیقاً هم‌نام «مغایرت سیستم متمرکز» با ماهیت مقابل، به‌عنوان پیشنهاد تطبیق کنار هم قرار می‌گیرند. عدم تهاتر به‌تنهایی مغایرت قطعی نیست.'],
 ['commitment','تطبیق تعهدات مقابل','جفت سرفصل‌های ۴۰۲/۴۲۲، ۴۱۸/۴۳۸، ۴۱۹/۴۳۹ و ۴۵۱/۴۷۶ با عنوان هم‌موضوع و ماهیت مقابل تطبیق می‌شوند؛ این نگاشت مخصوص این قالب است.']
];
export const normalize = v => String(v ?? '').replace(/[يى]/g,'ی').replace(/ك/g,'ک').replace(/[ۀةھ]/g,'ه').replace(/[\u200c\u200f\u200e]/g,' ').replace(/\s+/g,' ').trim();
export function integer(v) {
 let s=String(v??'').trim().replace(/[۰-۹]/g,c=>String('۰۱۲۳۴۵۶۷۸۹'.indexOf(c))).replace(/[٠-٩]/g,c=>String('٠١٢٣٤٥٦٧٨٩'.indexOf(c))).replace(/[٬,\s]/g,'').replace(/−/g,'-');
 if(/^\(\d+\)$/.test(s)) s='-'+s.slice(1,-1);
 if(!/^-?\d+(?:\.0+)?$/.test(s)) return null;
 return BigInt(s.split('.')[0]);
}
const labels=['تعداد بدهکار','تعداد بستانکار','بدهکار روز','بستانکار روز','گردش بدهکار','گردش بستانکار','مانده'];
const abs=n=>n<0n?-n:n;
export function analyze(book, options={}) {
 const threshold=integer(options.threshold??'1000000000');
 if(threshold===null||threshold<0n) throw Error('حد بررسی باید عدد صحیح نامنفی باشد.');
 const findings=[], rows=[], sheets=[], checks=[], crossChecks=[];
 const push=(rule,r,title,detail,severity='medium',amount=null,cells=[])=>findings.push({id:`F${findings.length+1}`,rule,sheet:r.sheet,row:r.row,code:r.code,title,account:r.title,detail,severity,amount:amount===null?null:amount.toString(),cells});
 for(const sheet of book.sheets){
  const raw=sheet.rows;
  const header=raw.findIndex(r=>r.some(c=>normalize(c.v)==='سرفصل')&&r.some(c=>normalize(c.v)==='مانده'));
  if(header<0){sheets.push({name:sheet.name,supported:false,reason:'قالب شناخته نشد؛ این شیت بررسی مالی نشده است.'});continue;}
  const names=raw[header].map(c=>normalize(c.v));
  const aliases=[['ردیف'],['سرفصل'],['معین'],['م'],['عنوان حساب'],['کد سپهری'],['تعداد بدهکار در یکروز','تعداد بدهکار در یک روز'],['تعداد بستانکار در یک روز'],['بدهکار روز جاری'],['بستانکار روز جاری'],['گردش بدهکار تا پایان روز'],['گردش بستانکار تا پایان روز'],['مانده']];
  const indexes=aliases.map(a=>names.findIndex(n=>a.includes(n)));
  if(indexes.some(i=>i<0)){sheets.push({name:sheet.name,supported:false,reason:'ستون‌های الزامی کامل نیستند؛ این شیت بررسی مالی نشده است.'});continue;}
  const local=[];
  for(let i=header+1;i<raw.length;i++){
   if(raw[i].every(c=>String(c.v??'').trim()===''&&!c.formula))continue;
   const c=indexes.map(j=>raw[i][j]??{v:''});
   const vals=c.map(x=>normalize(x.v));
   const grand=/^جمع کل\s*:?$/.test(vals[4]);
   const r={sheet:sheet.name,row:i+1,serial:vals[0],group:vals[1],sub:vals[2],nature:vals[3],title:vals[4],code:vals[5],kind:grand?'grand':integer(vals[2])===0n?'subtotal':'detail',numbers:vals.slice(6).map(integer),values:vals.slice(6),source:vals,columns:indexes.map(column)};
   local.push(r); rows.push({...r,numbers:undefined});
   c.forEach((x,j)=>{if(x.error||x.missingCache)push('quality',r,x.error?'خطای ذخیره‌شده اکسل':'فرمول بدون مقدار ذخیره‌شده',`ستون ${names[indexes[j]]}؛ فرمول در مرورگر محاسبه مجدد نمی‌شود.`,'high',null,[`${column(indexes[j])}${i+1}`]);});
   r.numbers.forEach((n,j)=>{if(n===null)push('quality',r,'داده عددی خالی یا نامعتبر',`${labels[j]}: «${vals[j+6]}». کنترل‌های وابسته انجام نشده‌اند.`,'high',null,[`${column(indexes[j+6])}${i+1}`]);});
   if(r.numbers[4]!==null&&r.numbers[5]!==null&&r.numbers[6]!==null){
    const expected=r.numbers[5]-r.numbers[4],delta=r.numbers[6]-expected;
    if(delta!==0n)push('balance',r,'مانده با گردش‌ها تطبیق ندارد',`مانده مورد انتظار ${expected}، مانده ثبت‌شده ${r.numbers[6]}، اختلاف ${delta}.`,'high',delta);
   }
   if(r.kind!=='detail')continue;
   if(!r.code||!r.group||!r.sub||!r.title)push('quality',r,'اطلاعات شناسایی ناقص','کد سپهری، سرفصل، معین و عنوان حساب باید موجود باشند.','high');
   if(!['بد','بس','بدوبس','بد و بس'].includes(r.nature))push('quality',r,'ماهیت ناشناخته',`ماهیت ثبت‌شده «${r.nature}» است؛ کنترل ماهیت انجام نشد.`);
   const n=r.numbers;
   if(n[6]!==null&&((r.nature==='بد'&&n[6]>0n)||(r.nature==='بس'&&n[6]<0n)))push('nature',r,'مانده خلاف ماهیت حساب',`ماهیت ${r.nature==='بد'?'بدهکار':'بستانکار'} و مانده ${n[6]}. قرارداد علامت فایل: بستانکار منهای بدهکار.`,abs(n[6])>=threshold?'high':'medium',n[6]);
   for(let side=0;side<2;side++){
    if(n[side]!==null&&n[side+2]!==null&&((n[side]===0n&&n[side+2]!==0n)||(n[side]>0n&&n[side+2]===0n)))push('count',r,'تفاوت تعداد و مبلغ روز',`${labels[side]}: ${n[side]}؛ ${labels[side+2]}: ${n[side+2]}. تعریف شمارنده گزارش باید تأیید شود.`,'medium',n[side+2],[`${r.columns[side+6]}${r.row}`,`${r.columns[side+8]}${r.row}`]);
    if(n[side+2]!==null&&n[side+4]!==null&&n[side+2]>n[side+4])push('turnover',r,'گردش روز بیشتر از گردش تجمعی',`${labels[side+2]}: ${n[side+2]}؛ ${labels[side+4]}: ${n[side+4]}. فرض کنترل: گردش روز در گردش تجمعی منظور شده است.`,'high',n[side+2]-n[side+4],[`${r.columns[side+8]}${r.row}`,`${r.columns[side+10]}${r.row}`]);
   }
   n.slice(0,6).forEach((v,j)=>{if(v!==null&&v<0n)push('turnover',r,'تعداد یا گردش منفی',`${labels[j]} برابر ${v} است؛ احتمال ثبت اصلاحی باید بررسی شود.`,'medium',v,[`${r.columns[j+6]}${r.row}`]);});
   if(n[6]!==null&&n[6]!==0n&&abs(n[6])>=threshold&&/مغایرت|بدهکاران موقت|بستانکاران موقت/.test(r.title))push('suspense',r,'مانده حساب باز در محدوده بررسی',`مانده ${n[6]}؛ حد بررسی ${threshold}. این حد، انتخاب تحلیلی است و الزام قانونی نیست. قدمت مانده و ریز اقلام باید دریافت شود.`,'medium',n[6]);
  }
  const details=local.filter(r=>r.kind==='detail'), totals=local.filter(r=>r.kind==='subtotal');
  for(const r of details){const parent=totals.find(t=>t.group===r.group);if(r.numbers[6]!==null&&r.numbers[6]!==0n&&/سررسید\s*گذشته|معوق|مشکوک\s*الوصول/.test(`${r.title} ${parent?.title??''}`))push('watch',r,'مانده مطالبات غیرجاری',`سرفصل مادر: ${parent?.title??'ناموجود'}. مانده ${r.numbers[6]}؛ برای تعیین وضعیت وصول به ریز تسهیلات، سررسید و وثایق نیاز است.`,'info',r.numbers[6]);}
  for(const key of ['code','pair']){
   const seen=new Map();
   for(const r of details){const k=key==='code'?r.code:`${r.group}/${r.sub}`;if(!k||k==='/')continue;if(seen.has(k)){push('duplicate',r,key==='code'?'کد سپهری تکراری':'ترکیب سرفصل و معین تکراری',`شناسه ${k} در ردیف ${seen.get(k)} نیز آمده است.`,'high',null,(key==='code'?[5]:[1,2]).map(i=>`${r.columns[i]}${r.row}`));findings.at(-1).relatedRows=[seen.get(k),r.row];}else seen.set(k,r.row);}
  }
  const reconcile=(r,parts,rule)=>{
   for(let j=0;j<7;j++){
    if(r.numbers[j]===null||parts.some(p=>p.numbers[j]===null)){checks.push({sheet:r.sheet,row:r.row,column:labels[j],status:'unavailable'});continue;}
    const sum=parts.reduce((a,p)=>a+p.numbers[j],0n),delta=r.numbers[j]-sum;
    checks.push({sheet:r.sheet,row:r.row,column:labels[j],status:delta===0n?'pass':'fail',expected:sum.toString(),actual:r.numbers[j].toString()});
    if(delta!==0n)push(rule,r,rule==='grand'?'جمع کل با معین‌ها تطبیق ندارد':'جمع سرفصل با معین‌ها تطبیق ندارد',`${labels[j]}: جمع معین‌ها ${sum}، مقدار ثبت‌شده ${r.numbers[j]}، اختلاف ${delta}. ردیف‌های مبنا: ${parts.map(p=>p.row).join('، ')}.`,'high',delta,[`${r.columns[j+6]}${r.row}`]);
   }
  };
  const groups=new Map();for(const r of details){if(!groups.has(r.group))groups.set(r.group,[]);groups.get(r.group).push(r);}
  for(const [group,parts] of groups){const matching=totals.filter(t=>t.group===group);if(matching.length!==1)push('quality',parts[0],'تعداد جمع سرفصل نامعتبر',`سرفصل ${group} دارای ${matching.length} ردیف جمع است؛ انتظار یک ردیف.`,'high');else reconcile(matching[0],parts,'subtotal');}
  for(const t of totals)if(!groups.has(t.group))push('quality',t,'جمع سرفصل بدون معین','ردیف معین متناظر برای این سرفصل موجود نیست.','high');
  const grands=local.filter(r=>r.kind==='grand');if(grands.length!==1)push('quality',{sheet:sheet.name,row:header+1,title:'ساختار گزارش',code:''},'جمع کل ناموجود یا تکراری',`${grands.length} ردیف جمع کل پیدا شد؛ انتظار یک ردیف.`,'high');else reconcile(grands[0],details,'grand');
  const system=details.filter(r=>r.title==='مغایرت سیستم متمرکز');
  if(system.length===2&&system.some(r=>r.nature==='بد')&&system.some(r=>r.nature==='بس')&&system.every(r=>r.numbers[6]!==null)){
   const net=system[0].numbers[6]+system[1].numbers[6];
   crossChecks.push({sheet:sheet.name,title:'حساب‌های مغایرت سیستم متمرکز',rows:system.map(r=>r.row),net:net.toString(),status:net===0n?'pass':'review',basis:'پیشنهاد تطبیق؛ یکسان بودن دوره، واحد و مبنای ثبت باید تأیید شود.'});
   if(net!==0n){push('cross',system[0],'مانده خالص حساب‌های مغایرت سیستم متمرکز',`ردیف‌های ${system.map(r=>r.row).join(' و ')} با مانده‌های ${system.map(r=>r.numbers[6]).join(' و ')}؛ خالص ${net}. این تطبیق، پیشنهادی است؛ هم‌نام بودن، الزام به تهاتر را اثبات نمی‌کند.`,'medium',net);findings.at(-1).relatedRows=system.map(r=>r.row);}
  }
  for(const [a,b,subject]of [['402','422','ضمانت'],['418','438','کارت'],['419','439','اسنادی'],['451','476','اسنادی']]){
   const aa=totals.filter(r=>r.group===a),bb=totals.filter(r=>r.group===b);
   if(!aa.length&&!bb.length)continue;
   const valid=aa.length===1&&bb.length===1&&aa[0].title.includes(subject)&&bb[0].title.includes(subject)&&aa[0].nature==='بد'&&bb[0].nature==='بس'&&aa[0].numbers[6]!==null&&bb[0].numbers[6]!==null;
   if(!valid){crossChecks.push({sheet:sheet.name,title:`تعهدات ${a}/${b}`,rows:[...aa,...bb].map(r=>r.row),status:'unavailable',basis:'ساختار، عنوان، ماهیت یا عدد مورد انتظار موجود نیست؛ تطبیق انجام نشد.'});continue;}
   const net=aa[0].numbers[6]+bb[0].numbers[6];crossChecks.push({sheet:sheet.name,title:`تعهدات ${a}/${b}`,rows:[aa[0].row,bb[0].row],net:net.toString(),status:net===0n?'pass':'review',basis:'نگاشت اختصاصی قالب؛ تطبیق مانده سرفصل‌های مقابل.'});
   if(net!==0n){push('commitment',aa[0],'تفاوت مانده تعهدات مقابل',`سرفصل ${a} در ردیف ${aa[0].row} و سرفصل ${b} در ردیف ${bb[0].row}؛ خالص ${net}. نگاشت سرفصل باید با راهنمای بانک تأیید شود.`,'medium',net);findings.at(-1).relatedRows=[aa[0].row,bb[0].row];}
  }
  sheets.push({name:sheet.name,supported:true,metadata:raw.slice(0,header).flat().map(c=>c.v).filter(Boolean).join(' | '),records:details.length,subtotals:totals.length,grand:grands.map(g=>({row:g.row,values:g.values})),groups:groups.size});
 }
 findings.sort((a,b)=>({high:0,medium:1,info:2}[a.severity]-{high:0,medium:1,info:2}[b.severity])||a.row-b.row);
 const rowMap=new Map(rows.map(r=>[`${r.sheet}/${r.row}`,r]));
 for(const f of findings){f.classification=['balance','subtotal','grand'].includes(f.rule)?'مغایرت محاسباتی':f.rule==='quality'||f.rule==='duplicate'?'کنترل کیفیت داده':f.rule==='watch'?'پیگیری مطالبات':'نیازمند تأیید بازرس';if(!f.cells.length){const r=rowMap.get(`${f.sheet}/${f.row}`);const cols={balance:[10,11,12],nature:[3,12],count:[6,7,8,9],turnover:[6,7,8,9,10,11],suspense:[12],watch:[12],cross:[12],commitment:[12],subtotal:[6,7,8,9,10,11,12],grand:[6,7,8,9,10,11,12]}[f.rule]??[1,2,5];f.cells=r?cols.map(i=>`${r.columns[i]}${r.row}`):[];}}
 return {name:book.name,sha256:book.sha256??'',generatedAt:new Date().toISOString(),version:'2.0',threshold:threshold.toString(),sheets,rows,findings,checks,crossChecks,summary:{records:rows.filter(r=>r.kind==='detail').length,subtotals:rows.filter(r=>r.kind==='subtotal').length,high:findings.filter(f=>f.severity==='high').length,medium:findings.filter(f=>f.severity==='medium').length,watch:findings.filter(f=>f.severity==='info').length,confirmed:findings.filter(f=>['balance','subtotal','grand'].includes(f.rule)).length,affected:new Set(findings.filter(f=>f.rule!=='watch').flatMap(f=>(f.relatedRows??[f.row]).map(row=>`${f.sheet}/${row}`))).size}};
}
function column(i){let s='';for(i++;i;i=Math.floor((i-1)/26))s=String.fromCharCode(65+(i-1)%26)+s;return s;}
